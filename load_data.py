import pandas as pd # ouverture des fichiers Excel/tableur
import os # gestion des chemins de fichiers
import pymupdf # lecture de fichiers PDF
import re # gestion des expressions régulières pour extraire les données tabulaires des PDF
from pathlib import Path # gestion des chemins de fichiers de manière plus portable
import glob # pour trouver tous les fichiers d'un type donné dans un dossier
import warnings # pour gérer les avertissements liés à l'ouverture de fichiers



warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl') # Ignore les avertissements spécifiques d'openpyxl(les avertissements ne bloquent pas l'éxécution du programme)

def log(msg): # Fonction à utiliser pour le suivi des algorithmes
    print(f"[INFO] {msg}")

def charger_tout_le_dossier(dir_path="."): # Fonction permettant de charger tous les fichiers du dossier courant
    data = {} # initialisation d'un dictionnaire vide de fichier
    errors = [] # Initialisation d'une liste de fichier erroné
    log("Scan du dossier...")

    # 1. EXCEL
    excel_paths = glob.glob(os.path.join(dir_path, "*.xlsx")) + glob.glob(os.path.join(dir_path, "*.xls")) # sélection des fichiers excel par formats
    for path in excel_paths: # boucle pour chaque fichier excel
        name = Path(path).name # report du nom apr le nom du fichier
        try:
            sheets = pd.read_excel(path, sheet_name=None, engine="openpyxl") # lecture des feuilles excel
            data[name] = {"type": "excel", "sheets": sheets, "nb_sheets": len(sheets)} # report de nom de chaque feuille d'un fichier excel
            log(f"Excel '{name}' chargé ({len(sheets)} feuilles)")
        except Exception as e:
            errors.append((name, str(e))) # ajout si erreur, à la liste des fichiers erronés
            log(f"Erreur Excel '{name}'")

    # 2. PDF
    pdf_paths = glob.glob(os.path.join(dir_path, "*.pdf")) # chargement de tous les fichiers pdf du dossier courant
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
    csv_paths = glob.glob(os.path.join(dir_path, "*.csv")) # chargement de tous les fichiers csv du dossier courant
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
    text_paths = glob.glob(os.path.join(dir_path, "*.txt"))
    for path in text_paths:
        name = Path(path).name
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
            data[name] = {"type": "text", "content": content}
            log(f"TXT '{name}' chargé")
        except Exception as e:
            errors.append((name, str(e)))
            log(f"Erreur TXT '{name}'")

    log(f"{len(data)} fichiers chargés")
    return data, errors

if __name__ == "__main__":
    print("Chargement des fichiers ...")
    data, errors = charger_tout_le_dossier(".")
    print("\n Fichiers :", list(data.keys()))
    if errors:
        print("\ Erreurs :", [e[0] for e in errors])

