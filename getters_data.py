import pandas as pd
import numpy as np
import os
import requests
import pickle
import unicodedata
import warnings
from dotenv import load_dotenv
from rapidfuzz import fuzz, process

# --- Configuration initiale ---
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
load_dotenv()

API_KEY = os.getenv("ADEME_API_KEY")
EXCEL_FILE = "Bilan_Carbone_V9.01.xlsx"
CACHE_FILE = "cache_fe.pkl"  # Cache en pickle

# Dictionnaire de normalisation des unités
referentiel = {
    # Litres
    "litres": "L", "litre": "L", "l": "L", "litre(s)": "L", "litres ": "L",
    # Kilowattheures
    "kwh": "kWh", "kilowattheure": "kWh", "kilowattheures": "kWh", "kwh/h": "kWh",
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
    # CO2e (ajout pour les unités de bilan carbone)
    "kgco2e": "kgCO₂e", "kg co2e": "kgCO₂e", "kg co2 e": "kgCO₂e", "kgco2": "kgCO₂",
    "gco2e": "gCO₂e", "g co2e": "gCO₂e",
}

SYNONYMS = {
    "électricité": [
        "électricité",
        "electricite",
        "électricité (mix fr)",
        "électricité (mix france)",
        "mix électrique",
        "mix électrique français",
        "électricité basse tension",
        "électricité haute tension",
    ],
    "essence": ["essence", "carburant essence", "sp95", "sp98", "e10"],
    "gaz naturel": ["gaz naturel", "gaz", "méthane"],
}

# Dictionnaire de conversions manuelles
CONVERSIONS = {
    ("t", "kg"): 1000,
    ("kg", "t"): 0.001,
    ("MWh", "kWh"): 1000,
    ("kWh", "MWh"): 0.001,
    ("m³", "L"): 1000,
    ("L", "m³"): 0.001,
}


def normalize_text(text):
    if pd.isna(text):
        return ""
    text = str(text).lower().strip()
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return text

def detect_header_row(df):
    for i, row in df.iterrows():
        values = [str(v).strip().lower() for v in row.tolist()]
        if ("nom" in values) and ("unité" in values):
            return i
    return None

def load_fe_sheet(sheet_name):
    """Charge une feuille Excel et la nettoie."""
    print(f"Chargement de la feuille : {sheet_name}")
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)

    # Détecter l'en-tête
    header_row = detect_header_row(df)
    if header_row is None:
        raise ValueError(f"En-tête introuvable dans {sheet_name}")

    # Recharger avec l'en-tête détecté
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=header_row)

    # Nettoyage
    df.columns = [str(c).strip().lower() for c in df.columns]
    df = df.loc[:, ~df.columns.str.contains("^unnamed", case=False, na=False)]
    df = df.dropna(how="all")
    df = df.reset_index(drop=True)

    # Ajouter une colonne de recherche normalisée
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
    """Charge le cache depuis un fichier."""
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "rb") as f:
            return pickle.load(f)
    return {}

def save_cache(cache):
    """Sauvegarde le cache dans un fichier."""
    with open(CACHE_FILE, "wb") as f:
        pickle.dump(cache, f)

def local_search(designation, unite=None, threshold=80):
    designation_normalized = normalize_text(designation)
    search_terms = [designation_normalized]
    if designation.lower() in SYNONYMS:
        search_terms.extend(normalize_text(syn) for syn in SYNONYMS[designation.lower()])

    for sheet_name, df in loaded_sheets.items():
        if df is None:
            continue

        matches = process.extract(
            designation_normalized,
            df["search"],
            scorer=fuzz.token_set_ratio,
            limit=1
        )

        if matches:
            first_element = matches[0][0]

            if isinstance(first_element, str):
                _, score, idx = matches[0]
            else:
                score, idx, _ = matches[0]

            score = int(score)

            if score >= threshold:
                row = df.iloc[idx]

                if pd.isna(row["total non décomposé"]) or pd.isna(row["unité"]):
                    continue

                valeur = row["total non décomposé"]
                unite_row = row["unité"]

                if unite and pd.notna(unite_row):
                    try:
                        valeur = convert_unit(valeur, unite_row, unite)
                        unite_row = unite
                    except ValueError:
                        pass

                result = {
                    "designation": row["nom"],
                    "valeur": valeur,
                    "unite": unite_row,
                    "source": row.get("source", ""),
                    "localisation": row.get("localisation", ""),
                    "sheet": sheet_name,
                    "score": score,
                }
                return result

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
    """Récupère les facteurs d'émission (local ou API)."""
    local_result = local_search(designation, unite)
    if local_result:
        return local_result

    return api_search(designation, unite)

SHEETS = ["FE Energie", "FE Fret", "FE Déchets", "FE Déplacements", "FE Intrants"]
loaded_sheets = load_all_sheets(SHEETS)
FE_CACHE = load_cache()