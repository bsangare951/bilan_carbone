import pandas as pd 
import os 
import pymupdf 
import re 
from pathlib import Path 
import glob 
import warnings
from PIL import Image
from docx import Document

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

def log(msg):
    print(f"[INFO] {msg}")

def charger_tout_le_dossier(dir_path="."):
    data = {}
    errors = []
    log("Scan du dossier...")

    # 1. EXCEL 
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
    pdf_paths = glob.glob(os.path.join(dir_path, "**", "*.pdf"), recursive=True)
    for path in pdf_paths:
        name = Path(path).name
        try:
            doc = pymupdf.open(path)
            full_text = ""
            for page in doc:
                full_text += page.get_text("text")
            doc.close()
            # Gestion des tabulations existantes dans un fichier pdf
            tables_regex = re.findall(r'(\d{2}/\d{2})\s+(\d{5,6})\s+([\d,]+\.?\d*)', full_text)
            data[name] = {"type": "pdf", "text": full_text, "tables_regex": tables_regex}
            log(f"PDF '{name}' chargé")
        except Exception as e:
            errors.append((name, str(e)))
            log(f"Erreur PDF '{name}'")

    # 3. CSV
    csv_paths = glob.glob(os.path.join(dir_path, "**", "*.csv"), recursive=True)
    for path in csv_paths:
        name = Path(path).name
        try:
            df = pd.read_csv(path)
            data[name] = {"type": "csv", "dataframe": df}
            log(f"CSV '{name}' chargé")
        except Exception as e:
            errors.append((name, str(e)))
            log(f"Erreur CSV '{name}'")

    # 4. TXT
    text_paths = glob.glob(os.path.join(dir_path, "**", "*.txt"), recursive=True)
    for path in text_paths:
        name = Path(path).name
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
            data[name] = {"type": "txt", "text": content} 
            log(f"TXT '{name}' chargé")
        except Exception as e:
            errors.append((name, str(e)))
            log(f"Erreur TXT '{name}'")

    # 5. DOCX
    docx_paths = glob.glob(os.path.join(dir_path, "**", "*.docx"), recursive=True)
    for path in docx_paths:
        name = Path(path).name
        try:
            doc = Document(path)
            data[name] = {"type": "docx", "document": doc}
            log(f"DOCX '{name}' chargé")
        except Exception as e:
            errors.append((name, str(e)))
            log(f"Erreur DOCX '{name}' : {e}")

    # 6. JPEG / JPG
    jpg_paths = glob.glob(os.path.join(dir_path, "**", "*.jpg"), recursive=True) + \
                glob.glob(os.path.join(dir_path, "**", "*.jpeg"), recursive=True)
    for path in jpg_paths:
        name = Path(path).name
        try:
            img = Image.open(path)
            data[name] = {"type": "jpg", "image_object": img}
            log(f"JPEG '{name}' chargé")
        except Exception as e:
            errors.append((name, str(e)))
            log(f"Erreur JPEG '{name}' : {e}")

    log(f"{len(data)} fichiers chargés")

    return data, errors


if __name__ == "__main__":
    print("Chargement des fichiers ...")
    data, errors = charger_tout_le_dossier(".")
    print("\nFichiers :", list(data.keys()))
    if errors:
        print("\nErreurs :", [e[0] for e in errors])