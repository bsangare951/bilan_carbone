import pandas as pd 
import os 
import pymupdf 
import re 
from pathlib import Path 
import glob 
import warnings 

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def log(msg):
    print(f"[INFO] {msg}")

def charger_tout_le_dossier(dir_path="."):
    data = {}
    errors = []
    log("Scan du dossier...")

    # 1. EXCEL (MODIFIÉ POUR PLUS DE ROBUSTESSE)
    excel_paths = glob.glob(os.path.join(dir_path, "**", "*.xlsx"), recursive=True) + \
                  glob.glob(os.path.join(dir_path, "**", "*.xls"), recursive=True) + \
                  glob.glob(os.path.join(dir_path, "**", "*.xlsm"), recursive=True)
    
    for path in excel_paths:
        name = Path(path).name
        # Ignorer les fichiers temporaires Excel
        if name.startswith("~$"): continue
        
        try:
            # Ouverture propre du fichier Excel
            xls = pd.ExcelFile(path, engine="openpyxl")
            processed_sheets = {}
            
            for sheet_name in xls.sheet_names:
                # Lecture de la feuille
                df = pd.read_excel(xls, sheet_name=sheet_name)
                
                # Vérification de sécurité : on ignore les feuilles vides
                if df.empty or df.dropna(how='all').empty:
                    continue
                
                processed_sheets[sheet_name] = df
            
            # Enregistrement des résultats
            if processed_sheets:
                data[name] = {
                    "type": "excel", 
                    "sheets": processed_sheets, 
                    "nb_sheets": len(processed_sheets)
                }
                log(f"Excel '{name}' chargé : {len(processed_sheets)} feuilles détectées.")
            
        except Exception as e:
            errors.append((name, str(e)))
            log(f"Erreur lors du chargement de '{name}': {e}")

    # 2. PDF
    pdf_paths = glob.glob(os.path.join(dir_path, "**", "*.pdf"), recursive=True) # chargement de tous les fichiers pdf du dossier courant et sous-dossiers
    for path in pdf_paths: # boucle de chaque fichier pdf 
        name = Path(path).name # report du nom par le nom du fichier pdf
        try:
            doc = pymupdf.open(path)
            full_text = ""
            for page in doc:
                full_text += page.get_text("text")
            doc.close()
            # gestion des tabulations existantes dans un fichier pdf
            tables_regex = re.findall(r'(\d{2}/\d{2})\s+(\d{5,6})\s+([\d,]+\.?\d*)', full_text)
            data[name] = {"type": "pdf", "text": full_text, "tables_regex": tables_regex}
            log(f"PDF '{name}' chargé")
        except Exception as e:
            errors.append((name, str(e)))
            log(f"Erreur PDF '{name}'")
            
    # 3 . CSV
    csv_paths = glob.glob(os.path.join(dir_path, "**", "*.csv"), recursive=True) # chargement de tous les fichiers csv du dossier courant
    for path in csv_paths: # boucle de chaque fichier csv
        name = Path(path).name # report du nom par le nom du fichier csv
        try:
            df = pd.read_csv(path) # lecture du fichier csv
            data[name] = {"type": "csv", "dataframe": df} # report du nom de chaque feuille d'un fichier excel
            log(f"CSV '{name}' chargé")
        except Exception as e:
            errors.append((name, str(e)))
            log(f"Erreur CSV '{name}'")
    

    # 4. TXT
    text_paths = glob.glob(os.path.join(dir_path, "**", "*.txt"), recursive=True) # chargement de tous les fichiers txt du dossier courant et sous-dossiers
    for path in text_paths: # boucle de chaque fichier txt
        name = Path(path).name # report du nom par le nom du fichier txt
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f: # lecture du fichier txt, en ignorant les erreurs d'encodage
                content = f.read()
            data[name] = {"type": "text", "content": content} # report du nom de chaque feuille d'un fichier excel
            log(f"TXT '{name}' chargé") 
        except Exception as e:
            errors.append((name, str(e)))
            log(f"Erreur TXT '{name}'")

    log(f"{len(data)} fichiers chargés")
    return data, errors

if __name__ == "__main__": # point d'entrée du programme
    print("Chargement des fichiers ...")
    data, errors = charger_tout_le_dossier(".") # chargement de tous les fichiers du dossier courant
    print("\n Fichiers :", list(data.keys()))
    if errors:
        print("\ Erreurs :", [e[0] for e in errors])

