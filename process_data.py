import pandas as pd
import numpy as np
import os
import load_data as ld
import openpyxl

# Chargement de tous les fichiers
datas, errors = ld.charger_tout_le_dossier(os.path.dirname(__file__))

# Fichiers à exclure
EXCLUDED_FILES = {"support_bc.xlsx", "export_bc.xlsx"}

def clean_excel(datas):
    cleaned_datas = {}
    output_dir = "cleaned_files"
    for filename, content in datas.items():
        # Ignorer les fichiers déjà traités
        if filename.startswith("cleaned_"):
            continue

        if content["type"] != "excel":
            cleaned_datas[filename] = content
            continue
        
        if filename in EXCLUDED_FILES:
            cleaned_datas[filename] = content
            print(f"[INFO] Fichier '{filename}' exclu")
            continue

        cleaned_sheets = {}
        for sheet_name, df in content["sheets"].items():
            try:
                # 1. Copie brute, aucune modification de structure
                df_clean = pd.DataFrame(df).copy()

                # 2. Suppression uniquement des lignes totalement vides
                df_clean = df_clean.dropna(how='all', axis=0)
                print(f"[INFO] '{filename}' - '{sheet_name}': {len(df) - len(df_clean)} lignes vides supprimées.")

            except Exception as e:
                print(f"[ERREUR] {filename} : {e}")

test = clean_excel(datas)