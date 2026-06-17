import pandas as pd
import os
import re
from pathlib import Path
import glob
import warnings
from PIL import Image
from docx import Document

try:
    import pytesseract
except Exception:
    pytesseract = None

try:
    import pymupdf
except Exception:
    pymupdf = None

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

GENERATED_DIRS = {
    "cleaned_files",
    "SCOPE_1",
    "SCOPE_2",
    "SCOPE_3",
    "USELESS",
    "__pycache__",
}


def log(msg):
    print(f"[INFO] {msg}")


def _is_generated_or_temp(path: Path, root: Path) -> bool:
    try:
        rel_parts = path.relative_to(root).parts
    except Exception:
        rel_parts = path.parts

    if any(part in GENERATED_DIRS for part in rel_parts):
        return True

    name = path.name
    return (
        name.startswith("~$")
        or name.startswith(".")
        or name.lower().endswith((".tmp", ".bak"))
    )


def _rel_name(path: Path, root: Path) -> str:
    return path.relative_to(root).as_posix()


def _read_csv_safely(path: Path) -> pd.DataFrame:
    last_error = None
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            return pd.read_csv(path, sep=None, engine="python", encoding=enc)
        except Exception as e:
            last_error = e

    # Dernier essai, fréquent dans les exports français
    for enc in ("utf-8-sig", "latin-1", "cp1252"):
        try:
            return pd.read_csv(path, sep=";", encoding=enc)
        except Exception as e:
            last_error = e

    raise last_error


def charger_tout_le_dossier(dir_path="."):
    """
    Charge récursivement un dossier source en conservant les sous-dossiers.

    Correction importante :
    - Avant, les fichiers étaient stockés avec Path(path).name uniquement.
      Deux fichiers portant le même nom dans deux sous-dossiers pouvaient donc
      s'écraser dans le dictionnaire.
    - Maintenant, la clé est le chemin relatif :
        ClientA/facture.csv
        ClientB/facture.csv
      donc aucun fichier n'est perdu.
    """
    root = Path(dir_path).resolve()
    data = {}
    errors = []

    log(f"Scan récursif du dossier : {root}")

    all_files = sorted(
        p for p in root.rglob("*")
        if p.is_file() and not _is_generated_or_temp(p, root)
    )

    for path in all_files:
        suffix = path.suffix.lower()
        rel = _rel_name(path, root)
        name = path.name

        try:
            if suffix in {".xlsx", ".xls", ".xlsm"}:
                xls = pd.ExcelFile(path)
                processed_sheets = {}

                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    if df.empty or df.dropna(how="all").empty:
                        continue
                    processed_sheets[sheet_name] = df

                if processed_sheets:
                    data[rel] = {
                        "type": "excel",
                        "sheets": processed_sheets,
                        "nb_sheets": len(processed_sheets),
                        "original_path": str(path),
                        "relative_path": rel,
                        "name": name,
                    }
                    log(f"Excel '{rel}' chargé : {len(processed_sheets)} feuille(s).")

            elif suffix == ".pdf":
                if pymupdf is None:
                    raise RuntimeError("pymupdf n'est pas disponible")

                doc = pymupdf.open(path)
                full_text = ""
                for page in doc:
                    full_text += page.get_text("text")

                # Fallback OCR optionnel : utile pour les PDF scannés.
                if len(full_text.strip()) < 50 and pytesseract is not None:
                    try:
                        ocr_parts = []
                        for page in doc:
                            pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
                            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                            ocr_parts.append(pytesseract.image_to_string(img, lang="fra+eng"))
                        full_text = "\n".join(ocr_parts)
                        if full_text.strip():
                            log(f"PDF '{rel}' lu par OCR")
                    except Exception as ocr_e:
                        log(f"OCR impossible pour '{rel}' : {ocr_e}")

                doc.close()

                tables_regex = re.findall(
                    r"(\d{2}/\d{2})\s+(\d{5,6})\s+([\d,]+\.?\d*)",
                    full_text
                )

                data[rel] = {
                    "type": "pdf",
                    "text": full_text,
                    "tables_regex": tables_regex,
                    "original_path": str(path),
                    "relative_path": rel,
                    "name": name,
                }
                log(f"PDF '{rel}' chargé")

            elif suffix == ".csv":
                df = _read_csv_safely(path)
                data[rel] = {
                    "type": "csv",
                    "dataframe": df,
                    "original_path": str(path),
                    "relative_path": rel,
                    "name": name,
                }
                log(f"CSV '{rel}' chargé")

            elif suffix == ".txt":
                content = ""
                for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
                    try:
                        content = path.read_text(encoding=enc, errors="ignore")
                        break
                    except Exception:
                        continue

                data[rel] = {
                    "type": "txt",
                    "text": content,
                    "original_path": str(path),
                    "relative_path": rel,
                    "name": name,
                }
                log(f"TXT '{rel}' chargé")

            elif suffix == ".docx":
                doc = Document(path)
                data[rel] = {
                    "type": "docx",
                    "document": doc,
                    "original_path": str(path),
                    "relative_path": rel,
                    "name": name,
                }
                log(f"DOCX '{rel}' chargé")

            elif suffix in {".jpg", ".jpeg", ".png"}:
                img = Image.open(path)
                data[rel] = {
                    "type": "image",
                    "image_object": img,
                    "original_path": str(path),
                    "relative_path": rel,
                    "name": name,
                }
                log(f"Image '{rel}' chargée")

        except Exception as e:
            errors.append((rel, str(e)))
            log(f"Erreur lors du chargement de '{rel}' : {e}")

    log(f"{len(data)} fichier(s) chargé(s)")
    return data, errors


if __name__ == "__main__":
    print("Chargement des fichiers ...")
    data, errors = charger_tout_le_dossier(".")
    print("\nFichiers :", list(data.keys()))
    if errors:
        print("\nErreurs :", [e[0] for e in errors])
