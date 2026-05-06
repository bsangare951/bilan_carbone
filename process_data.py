import pandas as pd
import pytesseract
import os
import load_data as ld
import docx
import re
import csv
import shutil

# Chargement de tous les fichiers
datas, errors = ld.charger_tout_le_dossier(os.path.dirname(__file__) or ".")

# Fichiers à exclure
EXCLUDED_FILES = {"support_bc.xlsx", "export_bc.xlsx"}

def clean_excel(datas):
    cleaned_datas = {}
    output_dir = "cleaned_files"

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"[INFO] Dossier '{output_dir}' créé.")

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

                # 3. Nettoyage des cellules
                for col in df_clean.columns:
                    if df_clean[col].dtype == "object":
                        df_clean[col] = (
                            df_clean[col]
                            .astype(str)
                            .str.replace("\n", " ", regex=False)
                            .str.replace("\r", " ", regex=False)
                            .str.replace(";", ",", regex=False)
                        )

                # 4. Sauvegarde
                clean_filename = f"cleaned_{filename.replace('.xlsx', '').replace('.xls', '')}_{sheet_name}.csv"
                save_path = os.path.join(output_dir, clean_filename)

                # Export CSV
                df_clean.to_csv(
                    save_path,
                    index=False,
                    sep=';',
                    encoding='utf-8-sig',
                    quoting=csv.QUOTE_ALL
                )

                cleaned_sheets[sheet_name] = df_clean
                print(f"[OK] Sauvegardé sans perte : {save_path}")

            except Exception as e:
                print(f"[ERREUR] {filename} : {e}")
            cleaned_datas[filename] = {
                "type": "excel",
                "sheets": cleaned_sheets,
                "nb_sheets": len(cleaned_sheets)
            }

    return cleaned_datas


def clean_PDF(datas):
    output_dir = "cleaned_files"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    cleaned_datas = {}

    for filename, content in datas.items():
        if content.get("type") != "pdf":
            cleaned_datas[filename] = content
            continue
        if filename.startswith("cleaned_"):
            cleaned_datas[filename] = content
            continue

        try:
            # Récupération des données
            text = content.get("text", "")
            tables_regex = content.get("tables_regex", [])

            # Nettoyage du texte
            text = re.sub(r'\n+', '\n', text)
            text = re.sub(r' +', ' ', text)
            text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', ' ', text) 
            cleaned_content = text.strip()

            # Sauvegarde en TXT
            txt_path = os.path.join(output_dir, f"cleaned_{filename.replace('.pdf', '.txt')}")
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(cleaned_content)

            if tables_regex:
                csv_path = os.path.join(output_dir, f"cleaned_{filename.replace('.pdf', '.csv')}")
                with open(csv_path, "w", encoding="utf-8", newline='') as f:
                    writer = csv.writer(f, delimiter=';')
                    writer.writerow(['Date', 'ID', 'Montant'])
                    writer.writerows(tables_regex)
                print(f"[OK] Données tabulaires extraites en CSV pour '{filename}'.")

            cleaned_datas[filename] = {
                "type": "pdf",
                "content": cleaned_content
            }
            print(f"[OK] PDF '{filename}' nettoyé.")

        except Exception as e:
            print(f"[ERREUR] {filename} : {e}")

    return cleaned_datas


def clean_CSV(datas):
    output_dir = "cleaned_files"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    cleaned_datas = {}

    for filename, content in datas.items():
        if content.get("type") != "csv":
            cleaned_datas[filename] = content
            continue
        if filename.startswith("cleaned_"):
            cleaned_datas[filename] = content
            continue

        try:
            df = content.get("dataframe")
            if df is None:
                print(f"[ERREUR] {filename} : DataFrame manquant.")
                continue

            # Nettoyage des données
            for col in df.columns:
                if df[col].dtype == "object":
                    df[col] = (
                        df[col]
                        .astype(str)
                        .str.replace("\n", " ", regex=False)
                        .str.replace("\r", " ", regex=False)
                        .str.replace(";", ",", regex=False)
                    )

            # Sauvegarde en CSV
            clean_filename = f"cleaned_{filename}"
            save_path = os.path.join(output_dir, clean_filename)
            df.to_csv(
                save_path,
                index=False,
                sep=';',
                encoding='utf-8-sig',
                quoting=csv.QUOTE_ALL
            )
            print(f"[OK] CSV '{filename}' nettoyé et sauvegardé : {save_path}")

            cleaned_datas[filename] = {
                "type": "csv",
                "dataframe": df
            }

        except Exception as e:
            print(f"[ERREUR] {filename} : {e}")

    return cleaned_datas


def clean_DOCX(datas):
    output_dir = "cleaned_files"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for filename, content in datas.items():
        f_type = str(content.get("type", "")).lower()

        if f_type == "docx":
            source_path = filename
            new_path = os.path.join(output_dir, f"cleaned_{filename}")

            try:
                shutil.copy2(source_path, new_path)
                print(f"[OK] {filename} copié vers {output_dir}")
            except FileNotFoundError:
                print(f"[ERREUR] Impossible de trouver le fichier '{filename}'.")
            except Exception as e:
                print(f"[ERREUR] {filename} : {e}")

    return datas


def clean_JPEG(datas):
    output_dir = "cleaned_files"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for filename, content in datas.items():
        f_type = str(content.get("type", "")).lower()

        if f_type in ["jpg", "jpeg"]:
            source_path = content.get("image_object").filename if content.get("image_object") else None
            if not source_path:
                print(f"[ERREUR] {filename} : Chemin de l'image introuvable.")
                continue
            new_path = os.path.join(output_dir, f"cleaned_{filename}")

            try:
                shutil.copy2(source_path, new_path)
                print(f"[OK] {filename} copié vers {output_dir}")
            except FileNotFoundError:
                print(f"[ERREUR] Le fichier {filename} est introuvable à la racine.")
            except Exception as e:
                print(f"[ERREUR] {filename} : {e}")

    return datas


test = clean_excel(datas)
test2 = clean_PDF(test)
test3 = clean_CSV(test2)
test4 = clean_DOCX(test3)
test5 = clean_JPEG(test4)
