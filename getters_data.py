import math
import pandas as pd
import numpy as np
import time
import os
import requests
from dotenv import load_dotenv
import json
import warnings

warnings.filterwarnings(
    "ignore",
    category=UserWarning,
    module="openpyxl"
)

load_dotenv()

API_KEY = os.getenv("ADEME_API_KEY")

url = "https://data.ademe.fr/data-fair/api/v1/datasets/base-carboner/lines"
headers = {"X-API-Key": API_KEY}

EXCEL_FILE = "Bilan_Carbone_V9.01.xlsx"



referentiel = {
    # Litres
    "litres": "L",
    "litre": "L",
    "l": "L",

    # Kilowattheures
    "kwh": "kWh",
    "kilowattheure": "kWh",
    "kilowattheures": "kWh",

    # Mégawattheures
    "mwh": "MWh",
    "megawattheure": "MWh",
    "megawattheures": "MWh",

    # Mètres cubes
    "m3": "m³",
    "m³": "m³",
    "metre cube": "m³",
    "metres cubes": "m³",

    # Kilogrammes
    "kilogramme": "kg",
    "kilogrammes": "kg",
    "kilo": "kg",
    "kilos": "kg",

    # Tonnes
    "tonne": "t",
    "tonnes": "t",
    "t": "t",

    # Kilomètres
    "kilometre": "km",
    "kilometres": "km",
    "kms": "km",

    # Tonnes.kilomètres
    "t.km": "t.km",
    "tonnes.km": "t.km",
    "tkm": "t.km",

    # Kilomètres passagers
    "km.passagers": "km.passager",
    "passager.km": "km.passager",
    "pass.km": "km.passager",
}


# =========================================================
# CONVERSIONS
# =========================================================

CONVERSIONS = {
    ("t", "kg"): 1000,
    ("kg", "t"): 0.001,
    ("MWh", "kWh"): 1000,
    ("kWh", "MWh"): 0.001,
    ("m³", "L"): 1000,
    ("L", "m³"): 0.001,
}



def detect_header_row(raw_df):
    """
    Détecte automatiquement la vraie ligne de header.
    """

    for i, row in raw_df.iterrows():

        values = [str(v).strip() for v in row.tolist()]

        # Cas ADEME classique
        if "NOM" in values and "Unité" in values:
            return i

        # fallback
        if "Nom" in values and "Unité" in values:
            return i

    return None



def load_fe_sheet(sheet_name):
    """
    Charge automatiquement une feuille FE
    avec détection du vrai header.
    """

    print(f"\nChargement : {sheet_name}")

    raw = pd.read_excel(
        EXCEL_FILE,
        sheet_name=sheet_name,
        header=None
    )

    header_row = detect_header_row(raw)

    if header_row is None:
        raise ValueError(f"Impossible de détecter le header dans {sheet_name}")

    df = pd.read_excel(
        EXCEL_FILE,
        sheet_name=sheet_name,
        header=header_row
    )

    # Nettoyage colonnes
    df.columns = [str(c).strip() for c in df.columns]

    # Suppression colonnes unnamed
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", na=False)]

    # Suppression lignes totalement vides
    df = df.dropna(how="all")

    # Reset index
    df = df.reset_index(drop=True)

    return df



SHEETS = [
    "FE Energie",
    "FE Fret",
    "FE Déchets",
    "FE Déplacements",
    "FE Intrants",
]

loaded_sheets = {}

for sheet in SHEETS:

    try:
        loaded_sheets[sheet] = load_fe_sheet(sheet)

    except Exception as e:
        print(f"Erreur {sheet} : {e}")


# Accès rapide

df_energie = loaded_sheets.get("FE Energie")
df_fret = loaded_sheets.get("FE Fret")
df_dechets = loaded_sheets.get("FE Déchets")
df_deplacements = loaded_sheets.get("FE Déplacements")
df_intrants = loaded_sheets.get("FE Intrants")



def conversion(value, from_unit, to_unit):

    from_unit = referentiel.get(
        from_unit.lower().strip(),
        from_unit.lower().strip()
    )

    to_unit = referentiel.get(
        to_unit.lower().strip(),
        to_unit.lower().strip()
    )

    if from_unit == to_unit:
        return value

    key = (from_unit, to_unit)

    if key in CONVERSIONS:
        return value * CONVERSIONS[key]

    raise ValueError(
        f"Conversion de {from_unit} vers {to_unit} non définie"
    )




def local_search(designation, unite=None):

    designation = designation.lower().strip()

    for sheet_name, df in loaded_sheets.items():

        if df is None:
            continue

        for _, row in df.iterrows():

            nom = str(row.get("NOM", "")).strip()

            if not nom:
                continue

            if designation in nom.lower():

                unite_row = row.get("Unité", None)
                valeur_row = row.get("Total non décomposé", None)

                result = {
                    "designation": nom,
                    "unite": unite_row,
                    "valeur": valeur_row,
                    "source": row.get("Source", None),
                    "localisation": row.get("Localisation", None),
                    "sheet": sheet_name,
                }

                return result

    return None


CACHE_FILE = "cache_fe.json"


# Chargement cache
if os.path.exists(CACHE_FILE):

    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        FE_CACHE = json.load(f)

else:
    FE_CACHE = {}



def save_cache():

    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(FE_CACHE, f, ensure_ascii=False, indent=2)




def get_facteurs(designation, unite=None):

    key = f"{designation}_{unite}"

    if key in FE_CACHE:
        print(f"[CACHE] {designation}")
        return FE_CACHE[key]

    print(f"[API] Recherche : {designation}")

    time.sleep(1)

    params = {
        "q": designation,
        "size": 5
    }

    response = requests.get(
        url,
        headers=headers,
        params=params
    )

    if response.status_code != 200:
        print("Erreur API :", response.text)
        return None

    data = response.json()

    results = data.get("results", [])

    if not results:
        return None

    premier = results[0]

    resultat = {
        "designation": premier.get("Nom_base_français"),
        "valeur": premier.get("Total_poste_non_décomposé"),
        "unite": premier.get("Unité_français"),
        "incertitude": premier.get("Incertitude", 0),
        "source": "API ADEME"
    }

    FE_CACHE[key] = resultat
    save_cache()

    return resultat

print(local_search("gaz naturel"))