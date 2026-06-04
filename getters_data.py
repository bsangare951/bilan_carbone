import pandas as pd
import re
import os
import time
import pickle
import unicodedata
import warnings
import requests
from dotenv import load_dotenv
from rapidfuzz import fuzz, process

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
load_dotenv()

API_KEY   = os.getenv("ADEME_API_KEY")
EXCEL_FILE = "Bilan_Carbone_V9.01.xlsx"
CACHE_FILE = "cache_fe.pkl"

# ---------------------------------------------------------------------------
# RÉFÉRENTIELS
# ---------------------------------------------------------------------------

referentiel: dict[str, str] = {
    # Litres
    "litres": "L", "litre": "L", "l": "L", "litre(s)": "L", "litres ": "L",
    # Kilowattheures
    "kwh": "kWh", "kilowattheure": "kWh", "kilowattheures": "kWh",
    "kwh/h": "kWh", "kw/h": "kWh",
    # Mégawattheures
    "mwh": "MWh", "megawattheure": "MWh", "megawattheures": "MWh",
    # Mètres cubes
    "m3": "m³", "m³": "m³", "metre cube": "m³", "metres cubes": "m³", "m^3": "m³",
    # Kilogrammes
    "kilogramme": "kg", "kilogrammes": "kg", "kilo": "kg", "kilos": "kg", "kg ": "kg",
    # Tonnes
    "tonne": "t", "tonnes": "t", "tonne(s)": "t", "t ": "t",
    # Kilomètres
    "kilometre": "km", "kilometres": "km", "kms": "km", "km ": "km",
    # Tonnes.km
    "t.km": "t.km", "tonnes.km": "t.km", "tkm": "t.km", "t.km ": "t.km",
    # Passagers.km
    "km.passagers": "km.passager", "passager.km": "km.passager", "pass.km": "km.passager",
    # CO2e
    "kgco2e": "kgCO₂e", "kg co2e": "kgCO₂e", "kg co2 e": "kgCO₂e", "kgco2": "kgCO₂e",
    "gco2e": "gCO₂e", "g co2e": "gCO₂e",
}

CORRESPONDANCES: dict[str, str] = {
    # ── SCOPE 1 — Combustibles ────────────────────────────────────────────
    # Noms vérifiés contre FE Energie du fichier Bilan_Carbone_V9.01.xlsx
    "gaz":                    "Gaz naturel",
    "gaz naturel":            "Gaz naturel",
    "gaz natural":            "Gaz naturel",
    "methane":                "Gaz naturel",
    "méthane":                "Gaz naturel",
    "gaz de ville":           "Gaz naturel",
    "gaz type h":             "Gaz naturel, Type H, sauf nord",
    "gaz type b":             "Gaz naturel, Type B, nord",
    "butane":                 "Butane (inclus maritime)",
    "propane":                "Propane (inclus maritime)",
    "gpl":                    "Butane (inclus maritime)",
    "gaz propane":            "Propane (inclus maritime)",
    "gaz butane":             "Butane (inclus maritime)",
    "bouteille gaz":          "Butane (inclus maritime)",
    "fioul":                  "Fioul domestique",
    "fioul domestique":       "Fioul domestique",
    "fod":                    "Fioul domestique",
    "fuel":                   "Fioul domestique",
    "fuel oil":               "Fioul domestique",
    "fioul lourd":            "Fioul Lourd (commercial)",
    "chv":                    "Combustible haute viscosité (CHV)",
    "diesel":                 "Gazole routier (B10)",
    "gazole":                 "Gazole routier (B10)",
    "gazole routier":         "Gazole routier (B10)",
    "gasoil":                 "Gazole routier (B10)",
    "go":                     "Gazole routier (B10)",
    "gnr":                    "Gazole non routier",
    "gazole non routier":     "Gazole non routier",
    "gasoil non routier":     "Gazole non routier",
    "gnr engins":             "Gazole non routier",
    "carburant engins":       "Gazole non routier",
    "carburant camions":      "Gazole routier (B10)",
    "carburant poids lourds": "Gazole routier (B10)",
    "gazole poids lourds":    "Gazole routier (B10)",
    "diesel camions":         "Gazole routier (B10)",
    "b10":                    "Gazole routier (B10)",
    "b30":                    "Gazole (B30) - contient un % de biocarburant",   # nom exact FE
    "b100":                   "Gazole routier (B100)",
    "essence":                "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "sp95":                   "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "sp98":                   "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "e10":                    "Essence (E10) - contient un % de biocarburant",  # nom exact FE
    "e85":                    "Essence (E85) - contient un % de biocarburant",  # nom exact FE
    "supercarburant":         "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "sans plomb":             "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "carburant":              "Gazole routier (B10)",
    "kerosene":               "Carbureacteur (large coupe (jet B))",
    "kérosène":               "Carbureacteur (large coupe (jet B))",
    "jet a1":                 "Carbureacteur (large coupe (jet B))",
    "jet b":                  "Carbureacteur (large coupe (jet B))",
    "carburant aviation":     "Carbureacteur (large coupe (jet B))",
    "avgas":                  "Essence aviation (AvGas)",
    "charbon":                "Charbon sous-bitumineux (PCS entre 17 435 et 23 865 kJ par kg)",  # nom exact FE
    "coke":                   "Coke de pétrole",
    "biomasse":               "Biomasse",
    "bois":                   "Bois",
    "granules bois":          "Granulés - blancs français issus de connexe de scierie",  # nom exact FE
    "pellets":                "Granulés - blancs français issus de connexe de scierie",  # nom exact FE

    # ── SCOPE 2 — Électricité / Chaleur ──────────────────────────────────
    # "2023 - mix moyen" est le nom exact dans FE Energie pour l'électricité France 2023
    "electricite":            "2023 - mix moyen",
    "électricité":            "2023 - mix moyen",
    "elec":                   "2023 - mix moyen",
    "courant":                "2023 - mix moyen",
    "suivi consomation edf":  "2023 - mix moyen",
    "consommation edf":       "2023 - mix moyen",
    "compteur":               "2023 - mix moyen",
    "relevé compteur":        "2023 - mix moyen",
    "total kw/h":             "2023 - mix moyen",
    "mix electricite":        "2023 - mix moyen",
    "mix electricité":        "2023 - mix moyen",
    # Chaleur urbaine : les noms FE sont par réseau (département + nom réseau).
    # On garde la valeur générique qui sera affinée par fuzzy match en local_search.
    "chauffage urbain":       "Chaleur - réseau de chaleur urbain, France 2023",
    "reseau chaleur":         "Chaleur - réseau de chaleur urbain, France 2023",
    "vapeur":                 "Chaleur - réseau de chaleur urbain, France 2023",
    # Climatisation : nom exact FE
    "climatisation":          "2023 - usage (moyenne) : Climatisation",
    "refroidissement":        "2023 - usage (moyenne) : Climatisation",

    # ── SCOPE 3 — Déchets ────────────────────────────────────────────────
    # Noms vérifiés contre FE Déchets
    "dechets ménagers et assimilés": "DIS (Déchets Industriels Spéciaux) - Stockage - Impacts",
    "dechets menagers":              "DIS (Déchets Industriels Spéciaux) - Stockage - Impacts",
    "déchets":                       "DIS (Déchets Industriels Spéciaux) - Stockage - Impacts",
    "dechets":                       "DIS (Déchets Industriels Spéciaux) - Stockage - Impacts",
    "ordures menageres":             "DIS (Déchets Industriels Spéciaux) - Stockage - Impacts",
    "om":                            "Déchets putrescibles - Stockage - Impacts",   # nom exact FE
    "papier":                        "Papier",
    "papier carton":                 "Papier & Carton",                             # nom exact FE
    "carton":                        "Carton",
    "plastique":                     "Plastique",
    "verre":                         "Verre",
    "metal":                         "Métaux - Fin de vie hors recyclage - Impacts", # nom exact FE
    "métaux":                        "Métaux - Fin de vie hors recyclage - Impacts",
    "aluminium":                     "Aluminium",
    "canettes":                      "Aluminium",
    "acier":                         "Acier - Incinération - Impacts",              # nom exact FE
    "ferraille":                     "Acier - Incinération - Impacts",
    "d3e":                           "DEEE moyen (par défaut) - Fin de vie moyenne filière - Impacts",  # nom exact FE
    "deee":                          "DEEE moyen (par défaut) - Fin de vie moyenne filière - Impacts",
    "dechets electroniques":         "DEEE moyen (par défaut) - Fin de vie moyenne filière - Impacts",
    "cartouches":                    "Cartouche jet d'encre noir et blanc re-conditionnée",  # meilleure correspondance FE
    "capsules":                      "Déchets putrescibles - Stockage - Impacts",
    "dechets verts":                 "Déchets de cuisine et déchets verts - Compostage domestique en bac - Impacts",  # nom exact FE
    "biomasse dechets":              "Déchets de cuisine et déchets verts - Compostage domestique en bac - Impacts",
    "gravats":                       "Déchets inertes en mélange (Gravats) - Fin de vie hors recyclage - Impacts",  # nom exact FE
    "btp":                           "Déchets bâtiment",                            # nom exact FE
    # Matériaux BTP collectés pour clients (NON des émissions propres)
    "terres inertes":                "_MATERIAUX_CLIENT_",
    "sable":                         "_MATERIAUX_CLIENT_",
    "sable 0/4":                     "_MATERIAUX_CLIENT_",
    "calcaire":                      "_MATERIAUX_CLIENT_",
    "beton non valorisable":         "_MATERIAUX_CLIENT_",
    "blocs beton":                   "_MATERIAUX_CLIENT_",
    "souches":                       "_MATERIAUX_CLIENT_",
    "deblais":                       "_MATERIAUX_CLIENT_",
    "remblais":                      "_MATERIAUX_CLIENT_",
    "enrobe":                        "_MATERIAUX_CLIENT_",
    "granulat":                      "_MATERIAUX_CLIENT_",

    # ── SCOPE 3 — Fret / Transport ───────────────────────────────────────
    # Noms vérifiés contre FE Fret
    "transport":          "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",  # nom exact FE (% manquait)
    "fret":               "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",
    "fret routier":       "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",
    "camion":             "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",
    "poids lourd":        "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",
    "pl":                 "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",
    "vehicule utilitaire": "Véhicule utilitaire léger",
    "vul":                "Véhicule utilitaire léger",
    "fret ferroviaire":   "Ferroviaire",                                            # nom exact FE
    "train fret":         "Train de marchandises",                                  # nom exact FE (manquait le s)
    "fret maritime":      "Train de marchandises",                                  # fallback : pas de FE maritime dédié
    "bateau":             "Train de marchandises",
    "fret aerien":        "Avion passagers, court courrier, avec trainées",         # meilleure approximation FE
    "avion fret":         "Avion passagers, court courrier, avec trainées",
    "messagerie":         "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",
    "livraison":          "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",
    "logistique":         "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",
    "transporteur":       "Articulé, 34 à 40 T, diesel routier, 7% de biodiesel",

    # ── SCOPE 3 — Déplacements ───────────────────────────────────────────
    # Noms vérifiés contre FE Déplacements
    "voiture particuliere":     "Voiture - motorisation moyenne - 2018",
    "voiture":                  "Voiture - motorisation moyenne - 2018",
    "indemnites kilometriques": "Voiture - motorisation moyenne - 2018",
    "indemnites km":            "Voiture - motorisation moyenne - 2018",
    "deplacement voiture":      "Voiture - motorisation moyenne - 2018",
    "vp":                       "Voiture - motorisation moyenne - 2018",            # hydrogène retiré : trop spécifique
    "voiture diesel":           "Voiture - motorisation gazole - 2018",             # nom exact FE
    "voiture essence":          "Voiture essence - mixte - 2018",                   # nom exact FE
    "voiture electrique":       "Voiture particulière - cœur de gamme - véhicule compact - électrique",  # nom exact FE
    "velo":                     "Vélo à assistance électrique",                     # meilleure correspondance FE (vélo simple absent)
    "trottinette":              "Trottinette électrique",
    "bus":                      "Autobus GNV",                                      # nom exact FE (Bus générique absent)
    "car":                      "Autocar gazole",                                   # nom exact FE
    "autocar":                  "Autocar gazole",
    "metro":                    "Métro - 2018 - Ile de France",                     # nom exact FE
    "rer":                      "RER et transilien - 2018 - Ile de France",         # nom exact FE
    "tramway":                  "Tramway - 2018 - Ile de France",                   # nom exact FE
    "train":                    "Train de voyageurs",                               # nom exact FE
    "tgv":                      "TGV - 2018",                                       # nom exact FE
    "intercites":               "Intercités - 2018",                                # nom exact FE
    "avion":                    "Avion passagers, court courrier, avec trainées",   # nom exact FE
    "avion court":              "Avion passagers, court courrier, avec trainées",
    "avion long":               "Avion passagers, long courrier, avec trainées",    # nom exact FE
    "domicile travail":         "Voiture - motorisation moyenne - 2018",
    "deplacement domicile":     "Voiture - motorisation moyenne - 2018",
    "trajet domicile":          "Voiture - motorisation moyenne - 2018",
    "mission":                  "Voiture - motorisation moyenne - 2018",
    "deplacement professionnel": "Voiture - motorisation moyenne - 2018",
    # ── Sorties typiques du modèle Groq llama-3.1-8b-instant ─────────────
    # Le modèle 8b retourne souvent des désignations génériques ou avec
    # parenthèses — on les capture ici pour éviter le fuzzy match hasardeux.
    "trajet(s) en avion":       "Avion passagers, court courrier, avec trainées",
    "trajet(s) en train":       "Train de voyageurs",
    "trajet en avion":          "Avion passagers, court courrier, avec trainées",
    "trajet en train":          "Train de voyageurs",
    "deplacements en avion":    "Avion passagers, court courrier, avec trainées",
    "deplacements en train":    "Train de voyageurs",
    "deplacements en voiture":  "Voiture - motorisation moyenne - 2018",
    "km parcourus":             "Voiture - motorisation moyenne - 2018",
    "kilometrage":              "Voiture - motorisation moyenne - 2018",

    # ── SCOPE 3 — Immobilisations ────────────────────────────────────────
    # Noms vérifiés contre FE Immobilisations
    "immobilisation":       "Bâtiments de bureaux",
    "batiment":             "Bâtiments de bureaux",                                 # nom exact FE
    "bureau":               "Bâtiments de bureaux",
    "ordinateur":           "Ordinateur portable",
    "pc portable":          "Ordinateur portable",
    "pc fixe":              "Ordinateur fixe - bureautique",                        # nom exact FE
    "ecran":                "Ecran 21,5 pouces",                                    # nom exact FE
    "imprimante":           "Imprimante jet d'encre",                               # nom exact FE
    "serveur":              "Serveurs informatiques",                               # nom exact FE
    "photocopieur":         "Photocopieurs",                                        # nom exact FE
    "copies couleur":       "Papier",
    "copies nb":            "Papier",
    "consommation copies":  "Papier",
    "impression":           "Papier",
    "papier bureau":        "Papier",
    "telephone":            "Smartphone 5 pouces",                                  # nom exact FE
    "smartphone":           "Smartphone 5 pouces",
    "tablette":             "Tablette classique - 9 à 11 pouces",                   # nom exact FE
    "vehicule societe":     "Véhicule",                                             # nom exact FE (générique)
    "machine":              "Machines",                                             # nom exact FE
    "outillage":            "Machines",

    # ── SCOPE 3 — Intrants ───────────────────────────────────────────────
    # Noms vérifiés contre FE Intrants
    "achat":              "Achats de biens",
    "matieres premieres":  "Matières premières",
    "fournitures":         "Fournitures de bureau",
    "fournitures bureau":  "Fournitures de bureau",
    "eau":                 "Eau de réseau - hors infrastructure",                   # nom exact FE
    "consommation eau":    "Eau de réseau - hors infrastructure",
    "eau potable":         "Eau de réseau - hors infrastructure",
    # Sorties Groq 8b pour relevés compteurs et achats divers
    "energie":             "2023 - mix moyen",                                      # relevés compteurs sans libellé précis
    "electricite batiment": "2023 - mix moyen",
    "logiciels":           "Achats de biens",                                       # grand-livre comptable -> intrant générique
    "materiel informatique": "Ordinateur portable",
    "mobilier":            "Machines",

    # ── SCOPE 3 — Gaz réfrigérants ──────────────────────────────────────
    # Noms vérifiés contre FE Autres émissions directes
    "refrigerant":    "R410a",                                                      # fallback générique -> R410a (le plus courant)
    "gaz refrigerant": "R410a",
    "r410a":           "R410a",                                                     # nom exact FE (casse différente de l'ancien)
    "r32":             "R32 (HFC-32)",                                              # nom exact FE
    "hfc":             "HFC-41",                                                    # nom exact FE (HFC générique -> HFC-41)
    "fuite":           "R410a",
    "climatiseur":     "R410a",
}

CONVERSIONS: dict[tuple[str, str], float] = {
    ("t",    "kg"):  1_000,
    ("kg",   "t"):   0.001,
    ("MWh",  "kWh"): 1_000,
    ("kWh",  "MWh"): 0.001,
    ("m³",   "L"):   1_000,
    ("L",    "m³"):  0.001,
    ("GJ PCI",         "L"):          23.88,
    ("kgCO2e/GJ PCI",  "kgCO2e/L"):   3.16,
}

SHEETS = ["FE Energie", "FE Fret", "FE Déchets", "FE Déplacements", "FE Intrants"]

# Pré-compile le pattern "année" une seule fois
_YEAR_RE = re.compile(r"\d{4} - ")

# ---------------------------------------------------------------------------
# UTILITAIRES
# ---------------------------------------------------------------------------

def normalize_text(text) -> str:
    if pd.isna(text):
        return ""
    text = str(text).lower().strip()
    text = unicodedata.normalize("NFD", text)
    return "".join(c for c in text if unicodedata.category(c) != "Mn")


def normalize_designation(designation: str, unite: str | None = None) -> str:
    key = designation.lower().strip()
    if key in CORRESPONDANCES:
        return CORRESPONDANCES[key]

    unite_norm = normalize_text(unite or "")

    if "diesel" in key or "gazole" in key:
        return "Gazole routier" if unite_norm in {"l", "litre", "litres"} else (
            "Ensemble articulé" if unite_norm == "km" else designation
        )

    if "articul" in key and unite_norm == "km":
        return "Ensemble articulé"

    if "diesel routier" in key or "gazole routier" in key:
        return "Gazole routier"

    return designation


def convert_unit(value: float, from_unit: str, to_unit: str) -> float:
    from_u = referentiel.get(from_unit.lower().strip(), from_unit.lower().strip())
    to_u   = referentiel.get(to_unit.lower().strip(),   to_unit.lower().strip())

    if from_u == to_u:
        return value

    key = (from_u, to_u)
    if key in CONVERSIONS:
        return value * CONVERSIONS[key]

    raise ValueError(f"Conversion de {from_unit} vers {to_unit} non définie.")

# ---------------------------------------------------------------------------
# CHARGEMENT DES FEUILLES EXCEL
# ---------------------------------------------------------------------------

def detect_header_row(df: pd.DataFrame) -> int | None:
    for i, row in df.iterrows():
        values = {str(v).strip().lower() for v in row}
        if "nom" in values and "unité" in values:
            return i
    return None


def load_fe_sheet(sheet_name: str) -> pd.DataFrame:
    print(f"Chargement de la feuille : {sheet_name}")
    df_raw = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)

    header_row = detect_header_row(df_raw)
    if header_row is None:
        raise ValueError(f"En-tête introuvable dans {sheet_name}")

    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=header_row)
    df.columns = [str(c).strip().lower() for c in df.columns]
    df = df.loc[:, ~df.columns.str.contains("^unnamed", case=False, na=False)]
    df = df.dropna(how="all").reset_index(drop=True)

    col = next((c for c in ("nom", "nomenclature") if c in df.columns), None)
    if col is None:
        raise ValueError(f"Colonne 'nom' ou 'nomenclature' introuvable dans {sheet_name}")

    df["search"] = df[col].apply(normalize_text)
    return df


def load_all_sheets(sheets_list: list[str]) -> dict[str, pd.DataFrame]:
    result = {}
    for sheet in sheets_list:
        try:
            result[sheet] = load_fe_sheet(sheet)
            print(f"  Feuille {sheet} chargée.")
        except Exception as e:
            print(f"  Erreur chargement {sheet} : {e}")
    return result

# ---------------------------------------------------------------------------
# CACHE
# ---------------------------------------------------------------------------

def load_cache() -> dict:
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "rb") as f:
            return pickle.load(f)
    return {}


def save_cache(cache: dict) -> None:
    with open(CACHE_FILE, "wb") as f:
        pickle.dump(cache, f)

# ---------------------------------------------------------------------------
# COMPATIBILITÉ D'UNITÉS
# ---------------------------------------------------------------------------

def compatibilite_unite(unite_fe: str, unite_q: str) -> bool:
    fe = normalize_text(unite_fe)
    q  = normalize_text(unite_q)
    rules = {
        ("/m3", "/m³"):        {"m3", "m³", "l", "litre", "litres"},
        ("/l",):               {"l", "litre", "litres", "m3", "m³"},
        ("/kg",):              {"kg", "t", "tonne", "tonnes"},
        ("/t", "/tonne", "/tonnes"): {"kg", "t", "tonne", "tonnes"},
        ("/mwh", "/kwh"):      {"kwh", "mwh"},
        ("/t.km", "/tonne.km"): {"t.km", "tonne.km"},
        ("/km",):              {"km", "km.passager", "km.passagers", "passager.km"},
    }
    for suffixes, valid_units in rules.items():
        if any(fe.endswith(s) for s in suffixes):
            return q in valid_units
    return True

# ---------------------------------------------------------------------------
# RECHERCHE LOCALE
# ---------------------------------------------------------------------------

def _score_compatibilite(item: tuple, unite: str | None) -> tuple:
    year, row, score = item
    fe_unite = str(row.get("unité", "")).lower()
    q_unite  = (unite or "").lower()

    compatible = 0
    pairs = [
        (q_unite == "l",    "litre" in fe_unite or fe_unite.endswith("/l")),
        (q_unite == "kg",   fe_unite.endswith("/kg")),
        (q_unite == "t",    fe_unite.endswith("/t") or "/tonne" in fe_unite),
        (q_unite == "kwh",  "kwh" in fe_unite and "gj" not in fe_unite),
        (q_unite in {"km", "km.passager", "km.passagers"}, fe_unite.endswith("/km") and not fe_unite.endswith(("t.km", "tonne.km"))),
        (q_unite == "m3",   "/m3" in fe_unite or "/m³" in fe_unite),
        (q_unite == "m³",   "/m3" in fe_unite or "/m³" in fe_unite),
    ]
    if any(a and b for a, b in pairs):
        compatible = 1_000
    # Pénaliser fortement les FE en t.km quand on cherche du /km
    if q_unite in {"km", "km.passager", "km.passagers"} and (
        "t.km" in fe_unite or "tonne.km" in fe_unite
    ):
        compatible = -500

    return (compatible, year, score)


def local_search(designation: str, unite: str | None = None, threshold: int = 80) -> dict | None:
    designation_norm = normalize_text(designation)

    for sheet_name, df in loaded_sheets.items():
        if df is None:
            continue

        matches = process.extract(
            designation_norm,
            df["search"],
            scorer=fuzz.token_set_ratio,
            limit=10,
        )
        if not matches:
            continue

        valid: list[tuple] = []
        for match in matches:
            _, score, idx = match if isinstance(match[0], str) else (None, match[1], match[2])
            if int(score) < threshold:
                continue
            row = df.iloc[idx]
            if pd.isna(row.get("total non décomposé")) or pd.isna(row.get("unité")):
                continue
            year_m = re.search(r"\d{4}", str(row["nom"]))
            valid.append((int(year_m.group()) if year_m else 0, row, int(score)))

        if not valid:
            continue

        valid.sort(key=lambda item: _score_compatibilite(item, unite), reverse=True)
        _, best_row, best_score = valid[0]

        valeur    = best_row["total non décomposé"]
        unite_row = best_row["unité"]

        # Correction papier/carton (t -> kg)
        nom_norm = normalize_text(best_row["nom"])
        if ("papier" in nom_norm or "carton" in nom_norm):
            ul = str(unite_row).lower()
            if ("/t" in ul or "tonne" in ul) and valeur > 100:
                valeur    = valeur / 1_000
                unite_row = "kgCO₂e/kg"
                print(f"    [FIX] Papier/Carton converti : {valeur} {unite_row}")

        if unite and pd.notna(unite_row):
            try:
                valeur    = convert_unit(valeur, unite_row, unite)
                unite_row = unite
            except ValueError:
                pass

        return {
            "designation":  best_row["nom"],
            "valeur":       valeur,
            "unite":        unite_row,
            "source":       best_row.get("source", ""),
            "localisation": best_row.get("localisation", ""),
            "incertitude":  best_row.get("incertitude", ""),
            "sheet":        sheet_name,
            "score":        best_score,
        }

    return None

# ---------------------------------------------------------------------------
# RECHERCHE API ADEME
# ---------------------------------------------------------------------------

def api_search(designation: str, unite: str | None = None) -> dict | None:
    key = f"{designation}_{unite}"
    if key in FE_CACHE:
        print(f"[CACHE] {designation}")
        return FE_CACHE[key]

    print(f"[API] Recherche : {designation}")
    url     = "https://data.ademe.fr/data-fair/api/v1/datasets/base-carboner/lines"
    headers = {"X-API-Key": API_KEY}
    params  = {"q": designation, "size": 8}

    try:
        time.sleep(1)
        response = requests.get(url, headers=headers, params=params, timeout=10)
        response.raise_for_status()
        results = response.json().get("results", [])
    except requests.exceptions.RequestException as e:
        print(f"Erreur API : {e}")
        return None

    if not results:
        return None

    designation_norm = normalize_text(designation)
    best: dict | None = None
    best_score = -1

    for item in results:
        nom      = str(item.get("Nom_base_français", "") or "")
        unite_fe = str(item.get("Unité_français", "") or "")
        valeur   = item.get("Total_poste_non_décomposé")
        score    = fuzz.token_set_ratio(designation_norm, normalize_text(nom))

        if unite and not compatibilite_unite(unite_fe, unite):
            score -= 40
        if "kgco2e" in normalize_text(unite_fe) and "kgco2e" not in normalize_text(unite or ""):
            score -= 5

        if score > best_score:
            best_score = score
            best = {
                "designation": nom or designation,
                "valeur":      valeur,
                "unite":       unite_fe,
                "incertitude": item.get("Incertitude", 0),
                "source":      "API ADEME",
                "score":       score,
            }

    if best is None or best_score < 45:
        return None

    FE_CACHE[key] = best
    save_cache(FE_CACHE)
    return best

# ---------------------------------------------------------------------------
# POINT D'ENTRÉE PRINCIPAL
# ---------------------------------------------------------------------------

def get_facteurs(designation: str, unite: str | None = None) -> dict | None:
    norm = normalize_designation(designation, unite)

    # 1. Recherche locale sur la désignation normalisée
    result = local_search(norm, unite)
    if result:
        return result

    # 2. Recherche locale sur la désignation originale (si différente)
    if norm != designation:
        result = local_search(designation, unite)
        if result:
            return result

    # 3. Recherche API
    result = api_search(norm, unite)
    if result:
        return result

    # 4. Fallback km/t.km : essayer en forçant l'unité '/km'
    unite_norm = normalize_text(unite or "")
    if unite_norm in {"km", "t.km", "tonne.km"}:
        result = local_search(norm, "km") or api_search(norm, "km")
        if result:
            return result

    return api_search(designation, unite)


# ---------------------------------------------------------------------------
# INITIALISATION AU CHARGEMENT DU MODULE
# ---------------------------------------------------------------------------

loaded_sheets: dict = load_all_sheets(SHEETS) if os.path.exists(EXCEL_FILE) else {}
FE_CACHE: dict      = load_cache()

if __name__ == "__main__":
    pass