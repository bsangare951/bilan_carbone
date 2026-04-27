import pandas as pd 
import numpy as np 
from google import genai 
from dotenv import load_dotenv 
import os 
import time 
import load_data as ld 

# Chargement de tous les fichiers
datas, errors = ld.charger_tout_le_dossier(os.path.dirname(__file__))

def clean_excel_secure(datas):
    cleaned_datas = {}
    
    for filename, content in datas.items():
        if content["type"] != "excel":
            cleaned_datas[filename] = content
            continue
            
        cleaned_sheets = {}
        for sheet_name, df in content["sheets"].items():
            df_clean = df.copy()
            
            # 1. Suppression des lignes et colonnes entièrement vides
            df_clean = df_clean.dropna(how='all').dropna(axis=1, how='all')
            
            # 2. Nettoyage des noms de colonnes (Correction de l'erreur)
            # On force la conversion en str pour éviter l'erreur sur les entiers
            df_clean.columns = (df_clean.columns
                               .astype(str)
                               .str.strip()
                               .str.lower()
                               .str.replace(r'[^\w]+', '_', regex=True)
                               .str.strip('_'))
            
            # 3. Trim des espaces dans les cellules texte (Correction du warning)
            # Ajout de 'string' dans include et cast explicite en str
            for col in df_clean.select_dtypes(include=['object', 'string']).columns:
                df_clean[col] = df_clean[col].astype(str).str.strip()
            
            # 4. Conversion sécurisée des dates
            for col in df_clean.columns:
                if 'date' in col:
                    # 'dayfirst=True' aide souvent à résoudre le UserWarning sur les formats
                    converted = pd.to_datetime(df_clean[col], errors='coerce', dayfirst=True)
                    if converted.isnull().sum() <= (len(df_clean) * 0.5):
                        df_clean[col] = converted
            
            
            cleaned_sheets[sheet_name] = df_clean
            print(f"[INFO] Nettoyage fini pour {filename} ({sheet_name})")
            
        cleaned_datas[filename] = {**content, "sheets": cleaned_sheets}
        
    return cleaned_datas

test = clean_excel_secure(datas)