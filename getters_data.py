import pandas as pd
import time
import os
import requests
import pickle
import unicodedata
import warnings
from dotenv import load_dotenv
from rapidfuzz import fuzz, process
import re


warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
load_dotenv()

API_KEY = os.getenv("ADEME_API_KEY")
EXCEL_FILE = "Bilan_Carbone_V9.01.xlsx"
CACHE_FILE = "cache_fe.pkl"

# Dictionnaire de normalisation des unités
referentiel = {
    # Litres
    "litres": "L", "litre": "L", "l": "L", "litre(s)": "L", "litres ": "L",
    # Kilowattheures
    "kwh": "kWh", "kilowattheure": "kWh", "kilowattheures": "kWh", "kwh/h": "kWh", "Kw/h": "kWh",
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

# Dictionnaire de correspondances
CORRESPONDANCES = {
    # SCOPE 1 — Énergie / Combustibles
    "gaz": "Gaz naturel",
    "gaz naturel": "Gaz naturel",
    "gaz natural": "Gaz naturel",
    "methane": "Gaz naturel",
    "méthane": "Gaz naturel",
    "gaz de ville": "Gaz naturel",
    "gaz type h": "Gaz naturel, Type H, sauf nord",
    "gaz type b": "Gaz naturel, Type B, nord",
    "butane": "Butane (inclus maritime)",
    "propane": "Propane (inclus maritime)",
    "gpl": "Butane (inclus maritime)",
    "gaz propane": "Propane (inclus maritime)",
    "gaz butane": "Butane (inclus maritime)",
    "bouteille gaz": "Butane (inclus maritime)",
    "fioul": "Fioul domestique",
    "fioul domestique": "Fioul domestique",
    "fod": "Fioul domestique",
    "fuel": "Fioul domestique",
    "fuel oil": "Fioul domestique",
    "fioul lourd": "Fioul Lourd (commercial)",
    "chv": "Combustible haute viscosité (CHV)",
    "diesel": "Gazole routier (B10)",
    "gazole": "Gazole routier (B10)",
    "gazole routier": "Gazole routier (B10)",
    "gasoil": "Gazole routier (B10)",
    "go": "Gazole routier (B10)",
    "gnr": "Gazole non routier",
    "gazole non routier": "Gazole non routier",
    "b10": "Gazole routier (B10)",
    "b30": "Gazole (B30)",
    "b100": "Gazole routier (B100)",
    "essence": "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "sp95": "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "sp98": "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "e10": "Essence (E10)",
    "e85": "Essence (E85)",
    "supercarburant": "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "sans plomb": "Essence (Supercarburant sans plomb (95, 95-E10, 98))",
    "carburant": "Gazole routier (B10)",
    "kerosene": "Carbureacteur (large coupe (jet B))",
    "kérosène": "Carbureacteur (large coupe (jet B))",
    "jet a1": "Carbureacteur (large coupe (jet B))",
    "jet b": "Carbureacteur (large coupe (jet B))",
    "carburant aviation": "Carbureacteur (large coupe (jet B))",
    "avgas": "Essence aviation (AvGas)",
    "charbon": "Charbon",
    "coke": "Coke de pétrole",
    "biomasse": "Biomasse",
    "bois": "Bois",
    "granules bois": "Granulés de bois (pellets)",
    "pellets": "Granulés de bois (pellets)",

    # SCOPE 2 — Électricité / Chaleur
    "electricite": "mix electricité",
    "électricité": "mix electricité",
    "elec": "mix electricité",
    "courant": "mix electricité",
    "suivi consomation edf": "mix electricité",
    "consommation edf": "mix electricité",
    "compteur": "mix electricité",
    "relevé compteur": "mix electricité",
    "total kw/h": "mix electricité",
    "chauffage urbain": "Chaleur - réseau de chaleur urbain, France 2023",
    "reseau chaleur": "Chaleur - réseau de chaleur urbain, France 2023",
    "vapeur": "Chaleur - réseau de chaleur urbain, France 2023",
    "climatisation": "Climatisation 2023",  
    "refroidissement": "Climatisation 2023",


    # SCOPE 3 — Déchets
    "dechets": "Déchets ménagers et assimilés",
    "déchets": "Déchets ménagers et assimilés",
    "ordures menageres": "Déchets ménagers et assimilés",
    "om": "Déchets ménagers et assimilés",
    "papier": "Papier",
    "papier carton": "Papier carton",
    "papier / carton": "Papier carton",
    "carton": "Carton",
    "plastique": "Plastique",
    "verre": "Verre",
    "metal": "Métaux",
    "métaux": "Métaux",
    "aluminium": "Aluminium",
    "canettes": "Aluminium",
    "acier": "Acier",
    "ferraille": "Acier",
    "d3e": "Déchets d'équipements électriques et électroniques (DEEE)",
    "deee": "Déchets d'équipements électriques et électroniques (DEEE)",
    "dechets electroniques": "Déchets d'équipements électriques et électroniques (DEEE)",
    "cartouches": "Cartouches d'impression",
    "capsules": "Déchets ménagers et assimilés",
    "dechets verts": "Déchets verts",
    "biomasse dechets": "Déchets verts",
    "gravats": "Déchets inertes",
    "btp": "Déchets du BTP",

    # SCOPE 3 — Fret / Transport
    "transport": "Articulé, 34 à 40 T, diesel routier, 7 de biodiesel",
    "fret": "Transport routier de marchandises",
    "fret routier": "Articulé, 34 à 40 T, diesel routier, 7 de biodiesel",
    "camion": "Articulé, 34 à 40 T, diesel routier, 7 de biodiesel",
    "poids lourd": "Articulé, 34 à 40 T, diesel routier, 7 de biodiesel",
    "pl": "Articulé, 34 à 40 T, diesel routier, 7 de biodiesel",
    "vehicule utilitaire": "Véhicule utilitaire léger",
    "vul": "Véhicule utilitaire léger",
    "fret ferroviaire": "Transport ferroviaire de marchandises",
    "train fret": "Train de marchandise",
    "fret maritime": "Transport maritime de marchandises",
    "bateau": "Transport maritime de marchandises",
    "fret aerien": "Transport aérien de marchandises",
    "avion fret": "Transport aérien de marchandises",
    "messagerie": "Transport routier de marchandises",
    "livraison": "Transport routier de marchandises",
    "logistique": "Transport routier de marchandises",
    "transporteur": "Transport routier de marchandises",

    # SCOPE 3 — Déplacements
    "deplacement": "Voiture particulière",
    "déplacement": "Voiture particulière",
    "voiture": "Voiture particulière",
    "vehicule": "Voiture particulière",
    "vp": "Voiture particulière",
    "voiture diesel": "Voiture particulière, diesel",
    "voiture essence": "Voiture particulière, essence",
    "voiture electrique": "Voiture particulière, électrique",
    "velo": "Vélo",
    "trottinette": "Trottinette électrique",
    "bus": "Bus",
    "car": "Autocar",
    "autocar": "Autocar",
    "metro": "Métro",
    "rer": "RER",
    "tramway": "Tramway",
    "train": "Train grande ligne",
    "tgv": "TGV",
    "intercites": "Intercités",
    "avion": "Avion court courrier",
    "avion court": "Avion court courrier",
    "avion long": "Avion long courrier",
    "domicile travail": "Voiture particulière",
    "deplacement domicile": "Voiture particulière",
    "trajet domicile": "Voiture particulière",
    "mission": "Voiture particulière",
    "deplacement professionnel": "Voiture particulière",

    # SCOPE 3 — Immobilisations
    "immobilisation": "Immobilisation",
    "batiment": "Bâtiment",
    "bureau": "Bâtiment tertiaire",
    "ordinateur": "Ordinateur portable",
    "pc portable": "Ordinateur portable",
    "pc fixe": "Ordinateur fixe",
    "ecran": "Écran",
    "imprimante": "Imprimante",
    "serveur": "Serveur informatique",
    "photocopieur": "Photocopieur",
    "copies couleur": "Papier bureau",
    "copies nb": "Papier bureau",
    "consommation copies": "Papier bureau",
    "impression": "Papier bureau",
    "papier bureau": "Papier bureau",
    "telephone": "Téléphone portable",
    "smartphone": "Téléphone portable",
    "tablette": "Tablette",
    "vehicule societe": "Véhicule léger",
    "machine": "Machines et équipements",
    "outillage": "Machines et équipements",

    # SCOPE 3 — Intrants
    "achat": "Achats de biens",
    "matieres premieres": "Matières premières",
    "fournitures": "Fournitures de bureau",
    "fournitures bureau": "Fournitures de bureau",
    "eau": "Eau potable",
    "consommation eau": "Eau potable",
    "eau potable": "Eau potable",

    # SCOPE 3 — Autres émissions directes
    "refrigerant": "Gaz réfrigérant",
    "gaz refrigerant": "Gaz réfrigérant",
    "r410a": "R410A",
    "r32": "R32",
    "hfc": "HFC",
    "fuite": "Gaz réfrigérant",
    "climatiseur": "Gaz réfrigérant",
}

# Dictionnaire de conversions manuelles
CONVERSIONS = {
    ("t", "kg"): 1000,
    ("kg", "t"): 0.001,
    ("MWh", "kWh"): 1000,
    ("kWh", "MWh"): 0.001,
    ("m³", "L"): 1000,
    ("L", "m³"): 0.001,
    ("GJ PCI", "L"): 23.88,  
    ("kgCO2e/GJ PCI", "kgCO2e/L"): 3.16
}

def normalize_text(text):
    if pd.isna(text):
        return ""
    text = str(text).lower().strip()
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return text

def normalize_designation(designation):
    """Normalise une désignation en utilisant CORRESPONDANCES."""
    designation_lower = designation.lower()
    return CORRESPONDANCES.get(designation_lower, designation)

def detect_header_row(df):
    for i, row in df.iterrows():
        values = [str(v).strip().lower() for v in row.tolist()]
        if ("nom" in values) and ("unité" in values):
            return i
    return None

def load_fe_sheet(sheet_name):
    print(f"Chargement de la feuille : {sheet_name}")
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)

    header_row = detect_header_row(df)
    if header_row is None:
        raise ValueError(f"En-tête introuvable dans {sheet_name}")

    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=header_row)

    df.columns = [str(c).strip().lower() for c in df.columns]
    df = df.loc[:, ~df.columns.str.contains("^unnamed", case=False, na=False)]
    df = df.dropna(how="all")
    df = df.reset_index(drop=True)

    if "nom" in df.columns:
        df["search"] = df["nom"].apply(normalize_text)
    elif "nomenclature" in df.columns:
        df["search"] = df["nomenclature"].apply(normalize_text)
    else:
        raise ValueError(f"Colonne 'nom' ou 'nomenclature' introuvable dans {sheet_name}")

    return df

def load_all_sheets(sheets_list):
    loaded_sheets = {}
    for sheet in sheets_list:
        try:
            loaded_sheets[sheet] = load_fe_sheet(sheet)
            print(f"Feuille {sheet} chargée avec succès.")
        except Exception as e:
            print(f"Erreur lors du chargement de {sheet} : {e}")
    return loaded_sheets

def convert_unit(value, from_unit, to_unit):
    from_unit = referentiel.get(from_unit.lower().strip(), from_unit.lower().strip())
    to_unit = referentiel.get(to_unit.lower().strip(), to_unit.lower().strip())

    if from_unit == to_unit:
        return value

    key = (from_unit, to_unit)
    if key in CONVERSIONS:
        return value * CONVERSIONS[key]

    raise ValueError(f"Conversion de {from_unit} vers {to_unit} non définie.")

def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "rb") as f:
            return pickle.load(f)
    return {}

def save_cache(cache):
    with open(CACHE_FILE, "wb") as f:
        pickle.dump(cache, f)

def local_search(designation, unite=None, threshold=80):
    designation_normalized = normalize_text(designation)
    search_terms = [designation_normalized]

    # Si la désignation contient une année, essayer aussi sans l'année
    year_pattern = re.compile(r"\d{4} - ")
    if year_pattern.search(designation):
        base_designation = year_pattern.sub("", designation)
        search_terms.append(normalize_text(base_designation))

    for sheet_name, df in loaded_sheets.items():
        if df is None:
            continue

        matches = process.extract(
            designation_normalized,
            df["search"],
            scorer=fuzz.token_set_ratio,
            limit=10
        )

        if matches:
            valid_results = []
            for match in matches:
                if isinstance(match[0], str):
                    _, score, idx = match
                else:
                    score, idx, _ = match

                if int(score) >= threshold:
                    row = df.iloc[idx]
                    if pd.isna(row["total non décomposé"]) or pd.isna(row["unité"]):
                        continue

                    # Extraire l'année si présente
                    nom = row["nom"]
                    year_match = re.search(r"\d{4}", nom)
                    year = int(year_match.group()) if year_match else 0
                    valid_results.append((year, row, int(score)))

            if valid_results:
                # Trier par année décroissante (priorité aux années récentes)
                valid_results.sort(key=lambda x: x[0], reverse=True)
                best_year, best_row, best_score = valid_results[0]

                valeur = best_row["total non décomposé"]
                unite_row = best_row["unité"]

                if unite and pd.notna(unite_row):
                    try:
                        valeur = convert_unit(valeur, unite_row, unite)
                        unite_row = unite
                    except ValueError:
                        pass

                return {
                    "designation": best_row["nom"],
                    "valeur": valeur,
                    "unite": unite_row,
                    "source": best_row.get("source", ""),
                    "localisation": best_row.get("localisation", ""),
                    "incertitude": best_row.get("incertitude", ""),
                    "sheet": sheet_name,
                    "score": best_score
                }

    return None
def api_search(designation, unite=None):
    key = f"{designation}_{unite}"
    if key in FE_CACHE:
        print(f"[CACHE] {designation}")
        return FE_CACHE[key]

    print(f"[API] Recherche : {designation}")
    url = "https://data.ademe.fr/data-fair/api/v1/datasets/base-carboner/lines"
    headers = {"X-API-Key": API_KEY}
    params = {"q": designation, "size": 5}

    try:
        time.sleep(1)
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        data = response.json()
        results = data.get("results", [])

        if not results:
            return None

        premier = results[0]
        resultat = {
            "designation": premier.get("Nom_base_français", designation),
            "valeur": premier.get("Total_poste_non_décomposé"),
            "unite": premier.get("Unité_français"),
            "incertitude": premier.get("Incertitude", 0),
            "source": "API ADEME",
        }

        FE_CACHE[key] = resultat
        save_cache(FE_CACHE)
        return resultat

    except requests.exceptions.RequestException as e:
        print(f"Erreur API : {e}")
        return None

def get_facteurs(designation, unite=None):
    normalized_designation = normalize_designation(designation)

    # 1. Recherche locale avec priorité aux années récentes
    local_result = local_search(normalized_designation, unite)
    if local_result:
        return local_result

    # 2. Recherche locale avec la désignation originale
    if normalized_designation != designation:
        local_result = local_search(designation, unite)
        if local_result:
            return local_result

    # 3. Recherche API
    api_result = api_search(normalized_designation, unite)
    if api_result:
        return api_result

    return api_search(designation, unite)

SHEETS = ["FE Energie", "FE Fret", "FE Déchets", "FE Déplacements", "FE Intrants"]
loaded_sheets = load_all_sheets(SHEETS)
FE_CACHE = load_cache()


print(get_facteurs("livraison"))  