import os
import re
import time
import pickle
import unicodedata
import warnings
import requests
import pandas as pd
from rapidfuzz import fuzz, process as rfprocess
from dotenv import load_dotenv

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
load_dotenv()

EXCEL_FILE = "Bilan_Carbone_V9.01.xlsx"
CACHE_FILE  = "cache_fe.pkl"
API_KEY     = os.getenv("ADEME_API_KEY")

FE_SHEETS = [
    "FE Energie", "FE Autres émissions directes", "FE Intrants",
    "FE Déchets", "FE Déplacements", "FE Fret", "FE Immobilisations",
]

# Priorité des types de poste pour sélectionner la bonne ligne d'un FE
TYPE_PRIORITY = {
    "combustion": 10,
    "combustion à la centrale": 10,
    "usage": 9,
    "fin de vie": 8,
    "fabrication": 5,
    "amont": 3,
    "pertes": 2,
}


def normalize_text(text: str) -> str:
    if pd.isna(text):
        return ""
    text = str(text).lower().strip()
    text = unicodedata.normalize("NFD", text)
    return "".join(c for c in text if unicodedata.category(c) != "Mn")


def normalize_unite(unite: str) -> str:
    """Normalise une unité ADEME ou une unité de quantité.

    Point crucial : pour un facteur du type "kgCO2e/kg", on compare le
    dénominateur (kg), pas le préfixe "kgCO2e". Cela évite de considérer
    "kgCO2e/keuro" comme compatible avec des kg.
    """
    if unite is None:
        return ""

    u = normalize_text(unite)
    u = (u.replace("₂", "2")
           .replace("²", "2")
           .replace("³", "3")
           .replace("co₂", "co2"))
    u = re.sub(r"\s+", "", u)
    u = u.replace("kilowatt-heure", "kilowattheure")
    u = u.replace("megawatt-heure", "megawattheure")
    u = u.replace("tonnes-km", "tonneskm").replace("tonne-km", "tonnekm")

    # Si c'est une unité de facteur d'émission, on ne garde que le dénominateur.
    # Ex : kgCO2e/kWh -> kWh ; kgCO2e/keuro -> keuro ; kgCO2e/t.km -> t.km
    if "/" in u:
        u = u.split("/")[-1]

    u = u.replace("litre(s)", "litres")
    u = u.replace("m^3", "m3")
    return u


# Unités compatibles par catégorie de grandeur. Comparaison volontairement
# stricte : pas de test du type "kg" contenu dans "kgCO2e/keuro".
UNITE_COMPAT = {
    "litre":      {"l", "litre", "litres"},
    "kwh":        {"kwh", "kilowattheure", "kilowattheures", "kwhpci"},
    "mwh":        {"mwh", "megawattheure", "megawattheures"},
    "m3":         {"m3"},
    "kg":         {"kg", "kilogramme", "kilogrammes"},
    "tonne":      {"t", "tonne", "tonnes"},
    "km":         {"km", "kilometre", "kilometres", "vehicule.km", "vehiculekm", "v.km"},
    "tkm":        {"tkm", "t.km", "tonne.km", "tonnes.km", "tonnekm", "tonneskm"},
    "passagerkm": {"passager.km", "passagerkm", "km.passager", "km.passagers", "p.km", "pkm"},
    "tep":        {"tep", "teppci"},
}

FE_UNITES_INTERDITES = {
    "keuro", "keur", "euro", "euros", "eur", "€", "k€", "m€",
    "personne", "personnes", "salarie", "salaries", "etp",
    "m2", "m2.an", "m²", "m².an", "%", "pourcent",
    "100km", "kgh2", "kgh2/100km", "l100km", "l/100km",
}


def categorie_unite(unite: str) -> str | None:
    """Retourne la catégorie d'une unité ou None si elle n'est pas exploitable."""
    u = normalize_unite(unite)
    if not u:
        return None

    if u in FE_UNITES_INTERDITES:
        return None

    # Cas spécifiques avant les unités courtes pour éviter les collisions.
    if u in UNITE_COMPAT["tkm"]:
        return "tkm"
    if u in UNITE_COMPAT["passagerkm"]:
        return "passagerkm"

    for cat in ("litre", "kwh", "mwh", "m3", "tonne", "kg", "km", "tep"):
        if u in UNITE_COMPAT[cat]:
            return cat

    return None


def unite_fe_interdite(unite_fe: str) -> bool:
    """True si le FE est monétaire, surfacique, ratio, ou non applicable."""
    u = normalize_unite(unite_fe)
    return u in FE_UNITES_INTERDITES


def unites_compatibles(unite_fe: str, unite_quantite: str) -> bool:
    """Vérifie si un FE peut s'appliquer à une quantité.

    Compatible ne veut pas dire « identique » : kg/t, kWh/MWh et L/m3 sont
    convertis ensuite dans calcul_data.py. En revanche, t.km n'est jamais
    compatible avec km sans tonnage documenté.
    """
    if unite_fe_interdite(unite_fe):
        return False

    cat_fe = categorie_unite(unite_fe)
    cat_q  = categorie_unite(unite_quantite)

    if not cat_fe or not cat_q:
        return False

    if cat_fe == cat_q:
        return True

    if {cat_fe, cat_q} <= {"kwh", "mwh"}:
        return True
    if {cat_fe, cat_q} <= {"kg", "tonne"}:
        return True
    if {cat_fe, cat_q} <= {"litre", "m3"}:
        return True

    # Pour les déplacements de personnes, un km extrait est souvent un passager.km.
    # On accepte donc passager.km <-> km, mais jamais t.km <-> km.
    if cat_fe == "passagerkm" and cat_q == "km":
        return True

    return False


CORRESPONDANCES: dict[str, str] = {
    # Carburants (SCOPE 1)
    "diesel":                 "Gazole routier (B10) - contient un % de biocarburant",
    "gazole":                 "Gazole routier (B10) - contient un % de biocarburant",
    "gasoil":                 "Gazole routier (B10) - contient un % de biocarburant",
    "gazole routier":         "Gazole routier (B10) - contient un % de biocarburant",
    "gazole routier (b10)":   "Gazole routier (B10) - contient un % de biocarburant",
    "b10":                    "Gazole routier (B10) - contient un % de biocarburant",
    "carburant":              "Gazole routier (B10) - contient un % de biocarburant",
    "carburants":             "Gazole routier (B10) - contient un % de biocarburant",
    "gnr":                    "Gazole non routier",
    "gazole non routier":     "Gazole non routier",
    "gasoil non routier":     "Gazole non routier",
    "carburant engins":       "Gazole non routier",
    "essence":                "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "sp95":                   "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "sp98":                   "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "supercarburant":         "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "gaz naturel":            "Gaz naturel",
    "gaz":                    "Gaz naturel",
    "methane":                "Gaz naturel",
    "gaz de ville":           "Gaz naturel",
    "butane":                 "Butane (inclus maritime)",
    "propane":                "Propane (inclus maritime)",
    "gpl":                    "GPL pour véhicule routier",
    "gaz propane":            "Propane (inclus maritime)",
    "gaz butane":             "Butane (inclus maritime)",
    "bouteille gaz":          "Butane (inclus maritime)",
    "carburation":            "Butane (inclus maritime)",
    "antargaz":               "Butane (inclus maritime)",
    "primagaz":               "Propane (inclus maritime)",
    "fioul":                  "Fioul domestique",
    "fioul domestique":       "Fioul domestique",
    "fuel":                   "Fioul domestique",
    "kerosene":               "Carbureacteur (large coupe (jet B))",
    "kerosene jet":           "Kérosène (jet A1 ou A)",
    "jet a1":                 "Kérosène (jet A1 ou A)",

    # Électricité (SCOPE 2)
    "electricite":            "2023 - mix moyen",
    "electricité":            "2023 - mix moyen",
    "elec":                   "2023 - mix moyen",
    "mix electricite":        "2023 - mix moyen",
    "mix moyen":              "2023 - mix moyen",
    "2023 - mix moyen":       "2023 - mix moyen",
    "consommation edf":       "2023 - mix moyen",
    "compteur":               "2023 - mix moyen",
    "energie":                "2023 - mix moyen",
    "chauffage urbain":       "Chaleur et froid - réseau de chaleur/froid - France",
    "reseau chaleur":         "Chaleur et froid - réseau de chaleur/froid - France",

    # Déplacements (SCOPE 3)
    "voiture - motorisation moyenne - 2018": "Voiture - motorisation moyenne - 2018",
    "voiture":                "Voiture - motorisation moyenne - 2018",
    "voiture particuliere":   "Voiture - motorisation moyenne - 2018",
    "deplacement voiture":    "Voiture - motorisation moyenne - 2018",
    "indemnites km":          "Voiture - motorisation moyenne - 2018",
    "indemnites kilometriques": "Voiture - motorisation moyenne - 2018",
    "km domicile travail":    "Voiture - motorisation moyenne - 2018",
    "voiture diesel":         "Voiture - motorisation gazole - 2018",
    "voiture essence":        "Voiture - motorisation essence - 2018",
    "voiture electrique":     "Voiture - motorisation électrique - 2018",
    "train":                  "Train de voyageurs grandes lignes - 2018 - France",
    "tgv":                    "TGV - 2018 - France",
    "intercites":             "Intercités - 2018 - France",
    "rer":                    "RER et transilien - 2018 - Ile de France",
    "metro":                  "Métro - 2018 - Ile de France",
    "tramway":                "Tramway - 2018 - Ile de France",
    "bus":                    "Autobus - 2018 - France",
    "autocar":                "Autocar - 2018 - France",
    "avion court":            "Avion passagers, court courrier (<1000 km), avec trainées",
    "avion long":             "Avion passagers, long courrier (>3500 km), avec trainées",
    "avion":                  "Avion passagers, court courrier (<1000 km), avec trainées",
    "trajet(s) en train":     "Train de voyageurs grandes lignes - 2018 - France",
    "trajet(s) en avion":     "Avion passagers, court courrier (<1000 km), avec trainées",

    # Fret (SCOPE 3)
    "transport":              "transport",
    "fret":                   "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",
    "fret routier":           "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",
    "camion":                 "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",
    "poids lourd":            "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",
    "vehicule utilitaire":    "Véhicule utilitaire léger, diesel routier",
    "vul":                    "Véhicule utilitaire léger, diesel routier",
    "fret ferroviaire":       "Wagon - réseau ferré - France - 2018",
    "fret maritime":          "Navire porte-conteneurs, 10 000 - 50 000 tpl, moyenne mondiale",

    # Déchets (SCOPE 3)
    "dechets":                "dechets",
    "dis":                    "DIS (Déchets Industriels Spéciaux) - Stockage - Impacts",
    "dechets industriels":    "DIS (Déchets Industriels Spéciaux) - Stockage - Impacts",
    "ordures menageres":      "Déchets putrescibles - Stockage - Impacts",
    "papier":                 "Papier graphique - fin de vie - Incinération - Impacts",
    "papier carton":          "Papier & Carton mélangé - fin de vie - Incinération - Impacts",
    "papier / carton":        "Papier & Carton mélangé - fin de vie - Incinération - Impacts",
    "papier a4":              "Papier graphique - fin de vie - Incinération - Impacts",
    "papier a3":              "Papier graphique - fin de vie - Incinération - Impacts",
    "traceur":                "Papier graphique - fin de vie - Incinération - Impacts",
    "carton":                 "Carton - fin de vie - Incinération - Impacts",
    "plastique":              "Plastique - fin de vie - Incinération - Impacts",
    "verre":                  "Verre - fin de vie - Incinération - Impacts",
    "metal":                  "Métaux - Fin de vie hors recyclage - Impacts",
    "aluminium":              "Aluminium - Fin de vie hors recyclage - Impacts",
    "acier":                  "Acier - Incinération - Impacts",
    "deee":                   "DEEE moyen (par défaut) - Fin de vie moyenne filière - Impacts",
    "d3e":                    "DEEE moyen (par défaut) - Fin de vie moyenne filière - Impacts",
    "cartouches":             "Cartouche jet d'encre noir et blanc re-conditionnée",
    "capsules":               "Déchets putrescibles - Stockage - Impacts",
    "dechets verts":          "Déchets de cuisine et déchets verts - Compostage - Impacts",
    "gravats":                "Déchets inertes en mélange (Gravats) - Fin de vie hors recyclage - Impacts",

    # Intrants (SCOPE 3)
    "eau":                    "Eau de réseau - hors infrastructure",
    "eau potable":            "Eau de réseau - hors infrastructure",
    "eau de reseau":          "Eau de réseau - hors infrastructure",
    "eau de réseau - hors infrastructure": "Eau de réseau - hors infrastructure",
    "fournitures":            "Fournitures de bureau",
    "fournitures bureau":     "Fournitures de bureau",
    "ciment":                 "Ciment",
    "grave non traitee":      "Grave non traitée",
    "grave non traitée":      "Grave non traitée",
    "grave recyclee":         "Grave recyclée",
    "grave recyclée":         "Grave recyclée",
    "granulats / sable / terre": "Granulats",
    "granulat":               "Granulats",
    "granulats":              "Granulats",
    "sable":                  "Sable",
    "terre":                  "Terre",
    "paves / bordures":       "Pavés béton",
    "pavés / bordures":       "Pavés béton",

    # Gaz réfrigérants
    "r410a":                  "R410a",
    "r32":                    "R32 (HFC-32)",
    "refrigerant":            "R410a",
    "gaz refrigerant":        "R410a",
}



def _detect_header(df: pd.DataFrame) -> int | None:
    for i, row in df.iterrows():
        vals = {str(v).strip().lower() for v in row}
        if "nom" in vals and "unité" in vals:
            return i
    return None


def _load_sheet(sheet_name: str) -> list[dict]:
    """Charge une feuille FE et retourne la liste de toutes les entrées."""
    df_raw = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)
    h = _detect_header(df_raw)
    if h is None:
        raise ValueError(f"Header non trouvé dans {sheet_name}")

    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=h)
    df.columns = [str(c).strip().lower() for c in df.columns]

    col_nom   = next((c for c in ("nom", "nomenclature") if c in df.columns), None)
    col_val   = next((c for c in df.columns if "total non" in c.lower()), None)
    col_unite = next((c for c in df.columns if c == "unité"), None)
    col_type  = next((c for c in df.columns if c == "type"), None)
    col_incert = next((c for c in df.columns if "incertitude" in c.lower()), None)
    col_src   = next((c for c in df.columns if c == "source"), None)

    if not col_nom or not col_val or not col_unite:
        raise ValueError(f"Colonnes manquantes dans {sheet_name}: "
                         f"nom={col_nom} val={col_val} unité={col_unite}")

    entries = []
    for _, row in df.iterrows():
        nom = str(row[col_nom]).strip()
        if not nom or nom.lower() in ("nan", "nom", "nomenclature"):
            continue
        try:
            val = float(row[col_val])
            if pd.isna(val):
                continue
        except (ValueError, TypeError):
            continue

        unite  = str(row[col_unite]).strip() if pd.notna(row.get(col_unite, "")) else ""
        type_  = str(row.get(col_type, "")).strip().lower() if col_type else ""
        incert = 0.0
        if col_incert:
            try:
                incert = float(row[col_incert])
            except Exception:
                pass

        entries.append({
            "nom":         nom,
            "nom_norm":    normalize_text(nom),
            "valeur":      val,
            "unite":       unite,
            "unite_norm":  normalize_text(unite),
            "type":        type_,
            "priorite":    TYPE_PRIORITY.get(type_, 1),
            "incertitude": incert,
            "source":      str(row.get(col_src, "Base Carbone")).strip(),
            "sheet":       sheet_name,
        })

    return entries


def load_all_fe() -> list[dict]:
    """Charge toutes les feuilles FE et retourne une liste plate."""
    all_entries: list[dict] = []
    if not os.path.exists(EXCEL_FILE):
        print(f"[WARN] {EXCEL_FILE} introuvable — FE locaux non disponibles")
        return all_entries

    for sheet in FE_SHEETS:
        try:
            entries = _load_sheet(sheet)
            all_entries.extend(entries)
            print(f"  [{sheet}] {len(entries)} lignes chargées")
        except Exception as e:
            print(f"  [{sheet}] Erreur : {e}")

    return all_entries



def load_cache() -> dict:
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "rb") as f:
                return pickle.load(f)
        except Exception:
            return {}
    return {}


def save_cache(cache: dict) -> None:
    try:
        with open(CACHE_FILE, "wb") as f:
            pickle.dump(cache, f)
    except Exception as e:
        print(f"[WARN] Cache non sauvegardé : {e}")



def _score_entry(entry: dict, unite_q: str, query_norm: str) -> float:
    """Score composite : unité prioritaire, puis similarité, puis type de poste."""
    compat = 1.0 if unites_compatibles(entry["unite"], unite_q) else 0.0
    sim = fuzz.token_set_ratio(query_norm, entry["nom_norm"]) / 100.0
    prio = entry["priorite"] / 10.0

    # L'unité pèse très lourd : mieux vaut ne rien calculer qu'appliquer un FE faux.
    return compat * 55 + sim * 35 + prio * 10


def local_search(
    designation: str,
    unite: str | None = None,
    threshold_sim: float = 60.0,
) -> dict | None:
    """Cherche le meilleur FE local pour une désignation et une unité données."""
    if not ALL_FE:
        return None

    query_norm = normalize_text(designation)
    unite_q    = unite or ""

    noms_norm = [e["nom_norm"] for e in ALL_FE]
    matches = rfprocess.extract(
        query_norm,
        noms_norm,
        scorer=fuzz.token_set_ratio,
        limit=30,
        score_cutoff=threshold_sim,
    )
    if not matches:
        return None

    indices = {m[2] for m in matches}
    candidats = [ALL_FE[i] for i in indices]

    # Rejet des FE monétaires/surfaciques/ratios.
    candidats = [e for e in candidats if not unite_fe_interdite(e["unite"])]

    # Si une unité de quantité est fournie, on impose la compatibilité.
    # On ne garde plus de candidat incompatible « faute de mieux ».
    if unite_q:
        candidats = [e for e in candidats if unites_compatibles(e["unite"], unite_q)]

    if not candidats:
        return None

    best = max(candidats, key=lambda e: _score_entry(e, unite_q, query_norm))

    return {
        "designation":  best["nom"],
        "valeur":       best["valeur"],
        "unite":        best["unite"],
        "incertitude":  best["incertitude"],
        "source":       best["source"],
        "sheet":        best["sheet"],
        "type":         best["type"],
        "score":        fuzz.token_set_ratio(query_norm, best["nom_norm"]),
    }


def api_search(designation: str, unite: str | None = None) -> dict | None:
    cache_key = f"api::{designation}::{unite}"
    if cache_key in FE_CACHE:
        print(f"[CACHE] {designation}")
        return FE_CACHE[cache_key]

    print(f"[API] Recherche : {designation}")
    url = "https://data.ademe.fr/data-fair/api/v1/datasets/base-carboner/lines"
    try:
        time.sleep(1)
        resp = requests.get(
            url,
            headers={"X-API-Key": API_KEY} if API_KEY else {},
            params={"q": designation, "size": 10},
            timeout=10,
        )
        resp.raise_for_status()
        results = resp.json().get("results", [])
    except Exception as e:
        print(f"  [API ERREUR] {e}")
        return None

    if not results:
        return None

    query_norm = normalize_text(designation)
    best = None
    best_score = -1

    for item in results:
        nom      = str(item.get("Nom_base_français", "") or "")
        unite_fe = str(item.get("Unité_français", "") or "")
        valeur   = item.get("Total_poste_non_décomposé")
        if valeur is None:
            continue
        if unite_fe_interdite(unite_fe):
            continue
        if unite and not unites_compatibles(unite_fe, unite):
            continue

        score = fuzz.token_set_ratio(query_norm, normalize_text(nom))
        if score > best_score:
            best_score = score
            try:
                valeur_float = float(valeur)
            except (TypeError, ValueError):
                continue
            best = {
                "designation": nom or designation,
                "valeur":      valeur_float,
                "unite":       unite_fe,
                "incertitude": float(item.get("Incertitude", 0) or 0),
                "source":      "API ADEME",
                "score":       score,
            }

    if best is None or best_score < 45:
        return None

    FE_CACHE[cache_key] = best
    save_cache(FE_CACHE)
    return best


def _cible_correspondance(designation: str, unite: str | None = None) -> str:
    """Mapping prudent : ne pas transformer les libellés génériques en camions/DIS."""
    key = normalize_text(designation)
    ucat = categorie_unite(unite or "")

    # Ces mots sont trop génériques : ils ont causé les faux « articulé 40T ».
    if key in {"transport", "deplacement", "deplacements", "trajet", "trajets", "distance"} and ucat == "km":
        return designation

    # Un total déchets mélangé ne doit pas devenir automatiquement DIS.
    if key in {"dechet", "dechets", "total dechets", "total dechet"}:
        return designation

    return CORRESPONDANCES.get(key, designation)


def get_facteurs(designation: str, unite: str | None = None) -> dict | None:
    """Retourne le meilleur facteur d'émission compatible avec l'unité fournie."""
    cible = _cible_correspondance(designation, unite)

    result = local_search(cible, unite)
    if result and result["score"] >= 65:
        return result

    if cible != designation:
        result2 = local_search(designation, unite)
        if result2 and result2["score"] >= 65:
            return result2

    # Fallback API, mais toujours avec compatibilité stricte d'unité.
    api_result = api_search(cible, unite)
    if api_result:
        return api_result

    # Dernier recours local compatible si présent : utile pour les libellés proches.
    return result


print("Chargement des facteurs d'émission...")
ALL_FE:   list[dict] = load_all_fe() if os.path.exists(EXCEL_FILE) else []
FE_CACHE: dict       = load_cache()
print(f"  {len(ALL_FE)} lignes FE chargées | {len(FE_CACHE)} entrées en cache")