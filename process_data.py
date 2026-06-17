import pandas as pd
import os
import re
import csv
import shutil
from pathlib import Path

import load_data as ld


EXCLUDED_FILES = {
    "support_bc.xlsx",
    "export_bc.xlsx",
    "Bilan_Carbone_V9.01.xlsx",
}

OUTPUT_DIR = "cleaned_files"


def _safe_name(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r'[<>:"/\\|?*]+', "_", name)
    name = re.sub(r"\s+", " ", name)
    return name[:120] if len(name) > 120 else name


def _rel_parent(filename: str) -> Path:
    p = Path(filename)
    if len(p.parts) <= 1:
        return Path("")
    return Path(*p.parts[:-1])


def _original_path(content: dict, filename: str) -> str:
    return content.get("original_path") or filename


def _out_path(filename: str, cleaned_basename: str) -> Path:
    """
    Conserve les sous-dossiers dans cleaned_files.

    Exemple :
    source : ClientA/Factures/facture.pdf
    sortie : cleaned_files/ClientA/Factures/cleaned_facture.txt
    """
    out = Path(OUTPUT_DIR) / _rel_parent(filename) / cleaned_basename
    out.parent.mkdir(parents=True, exist_ok=True)
    return out


def _clean_cell_text(df: pd.DataFrame) -> pd.DataFrame:
    df_clean = pd.DataFrame(df).copy()
    df_clean = df_clean.dropna(how="all", axis=0)

    for col in df_clean.columns:
        if df_clean[col].dtype == "object":
            df_clean[col] = (
                df_clean[col]
                .astype(str)
                .str.replace("\n", " ", regex=False)
                .str.replace("\r", " ", regex=False)
                .str.replace(";", ",", regex=False)
            )
    return df_clean


def _write_csv(df: pd.DataFrame, path: Path):
    df.to_csv(
        path,
        index=False,
        sep=";",
        encoding="utf-8-sig",
        quoting=csv.QUOTE_ALL,
    )


def charger_et_nettoyer(dossier_source):
    print(f"[INFO] Chargement depuis : {dossier_source}")

    # Nettoyage de l'ancien dossier cleaned_files pour éviter les vieux fichiers fantômes.
    out = Path(dossier_source) / OUTPUT_DIR
    if out.exists():
        shutil.rmtree(out)
        print(f"[INFO] Ancien dossier '{OUTPUT_DIR}' supprimé.")
    out.mkdir(parents=True, exist_ok=True)

    old_cwd = os.getcwd()
    os.chdir(dossier_source)
    try:
        datas, errors = ld.charger_tout_le_dossier(dossier_source)

        cleaned = clean_excel(datas)
        cleaned = clean_PDF(cleaned)
        cleaned = clean_TXT(cleaned)
        cleaned = clean_CSV(cleaned)
        cleaned = clean_DOCX(cleaned)
        cleaned = clean_IMAGES(cleaned)

        return cleaned, errors
    finally:
        os.chdir(old_cwd)


def clean_excel(datas):
    cleaned_datas = {}

    for filename, content in datas.items():
        basename = Path(filename).name

        if basename.startswith("cleaned_"):
            continue

        if content.get("type") != "excel":
            cleaned_datas[filename] = content
            continue

        if basename in EXCLUDED_FILES:
            cleaned_datas[filename] = content
            print(f"[INFO] Fichier '{filename}' exclu")
            continue

        cleaned_sheets = {}
        stem = Path(basename).stem

        for sheet_name, df in content.get("sheets", {}).items():
            try:
                df_clean = _clean_cell_text(df)
                safe_sheet = _safe_name(sheet_name)
                clean_basename = f"cleaned_{stem}_{safe_sheet}.csv"
                save_path = _out_path(filename, clean_basename)

                _write_csv(df_clean, save_path)

                rel_out = save_path.as_posix()
                cleaned_sheets[sheet_name] = df_clean
                print(f"[OK] Excel '{filename}' - '{sheet_name}' sauvegardé : {save_path}")

            except Exception as e:
                print(f"[ERREUR] Excel {filename} / {sheet_name} : {e}")

        cleaned_datas[filename] = {
            "type": "excel",
            "sheets": cleaned_sheets,
            "nb_sheets": len(cleaned_sheets),
            "original_path": content.get("original_path"),
            "relative_path": content.get("relative_path", filename),
        }

    return cleaned_datas


def clean_PDF(datas):
    cleaned_datas = {}

    for filename, content in datas.items():
        if content.get("type") != "pdf":
            cleaned_datas[filename] = content
            continue

        basename = Path(filename).name
        if basename.startswith("cleaned_"):
            continue

        try:
            text = content.get("text", "")
            tables_regex = content.get("tables_regex", [])

            text = re.sub(r"\n+", "\n", text)
            text = re.sub(r" +", " ", text)
            text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", " ", text)
            cleaned_content = text.strip()

            txt_name = f"cleaned_{Path(basename).stem}.txt"
            txt_path = _out_path(filename, txt_name)
            txt_path.write_text(cleaned_content, encoding="utf-8")

            if tables_regex:
                csv_name = f"cleaned_{Path(basename).stem}.csv"
                csv_path = _out_path(filename, csv_name)
                with csv_path.open("w", encoding="utf-8", newline="") as f:
                    writer = csv.writer(f, delimiter=";")
                    writer.writerow(["Date", "ID", "Montant"])
                    writer.writerows(tables_regex)
                print(f"[OK] Données tabulaires extraites en CSV pour '{filename}'.")

            cleaned_datas[filename] = {
                "type": "pdf",
                "content": cleaned_content,
                "original_path": content.get("original_path"),
                "relative_path": content.get("relative_path", filename),
            }
            print(f"[OK] PDF '{filename}' nettoyé : {txt_path}")

        except Exception as e:
            print(f"[ERREUR] PDF {filename} : {e}")
            cleaned_datas[filename] = content

    return cleaned_datas


def clean_TXT(datas):
    cleaned_datas = {}

    for filename, content in datas.items():
        f_type = str(content.get("type", "")).lower()
        if f_type != "txt":
            cleaned_datas[filename] = content
            continue

        basename = Path(filename).name
        if basename.startswith("cleaned_"):
            continue

        try:
            raw_text = content.get("text", "") or content.get("content", "")

            if not raw_text:
                src = _original_path(content, filename)
                for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
                    try:
                        raw_text = Path(src).read_text(encoding=enc, errors="ignore")
                        break
                    except Exception:
                        continue

            if not raw_text:
                print(f"[WARN] TXT vide : {filename}")
                cleaned_datas[filename] = content
                continue

            lines = raw_text.splitlines()
            non_empty = [l for l in lines if l.strip()]
            short_lines = [l for l in non_empty if len(l.strip()) <= 15]
            is_fragmented = len(non_empty) > 10 and len(short_lines) / max(len(non_empty), 1) > 0.5

            if is_fragmented:
                rebuilt = []
                buffer = []
                for line in lines:
                    stripped = line.strip()
                    if not stripped:
                        if buffer:
                            rebuilt.append(" ".join(buffer))
                            buffer = []
                        continue

                    has_number = bool(re.search(r"\d{2,}", stripped))
                    is_long = len(stripped) > 20

                    if has_number or is_long:
                        if buffer:
                            rebuilt.append(" | ".join(buffer + [stripped]))
                            buffer = []
                        else:
                            rebuilt.append(stripped)
                    else:
                        buffer.append(stripped)

                if buffer:
                    rebuilt.append(" ".join(buffer))

                cleaned_text = "\n".join(rebuilt)
                print(f"  [TXT] Fichier fragmenté reconstruit : {filename} ({len(lines)} -> {len(rebuilt)} lignes)")
            else:
                cleaned_lines = []
                for line in lines:
                    line = line.replace("\r", "").strip()
                    if re.match(r"^[-=_*]{3,}$", line):
                        continue
                    cleaned_lines.append(line)
                cleaned_text = re.sub(r"\n{3,}", "\n\n", "\n".join(cleaned_lines))

            cleaned_text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", " ", cleaned_text)
            cleaned_text = re.sub(r" {3,}", "  ", cleaned_text).strip()

            clean_basename = f"cleaned_{basename}"
            save_path = _out_path(filename, clean_basename)
            save_path.write_text(cleaned_text, encoding="utf-8")

            cleaned_datas[filename] = {
                "type": "txt",
                "text": cleaned_text,
                "original_path": content.get("original_path"),
                "relative_path": content.get("relative_path", filename),
            }
            print(f"[OK] TXT '{filename}' nettoyé : {save_path}")

        except Exception as e:
            print(f"[ERREUR] TXT {filename} : {e}")
            cleaned_datas[filename] = content

    return cleaned_datas


def clean_CSV(datas):
    cleaned_datas = {}

    for filename, content in datas.items():
        if content.get("type") != "csv":
            cleaned_datas[filename] = content
            continue

        basename = Path(filename).name
        if basename.startswith("cleaned_"):
            continue

        try:
            df = content.get("dataframe")
            if df is None:
                print(f"[ERREUR] {filename} : DataFrame manquant.")
                continue

            df_clean = _clean_cell_text(df)
            clean_basename = f"cleaned_{basename}"
            save_path = _out_path(filename, clean_basename)
            _write_csv(df_clean, save_path)

            cleaned_datas[filename] = {
                "type": "csv",
                "dataframe": df_clean,
                "original_path": content.get("original_path"),
                "relative_path": content.get("relative_path", filename),
            }
            print(f"[OK] CSV '{filename}' nettoyé : {save_path}")

        except Exception as e:
            print(f"[ERREUR] CSV {filename} : {e}")
            cleaned_datas[filename] = content

    return cleaned_datas


def clean_DOCX(datas):
    for filename, content in datas.items():
        if str(content.get("type", "")).lower() != "docx":
            continue

        basename = Path(filename).name
        clean_basename = f"cleaned_{basename}"
        save_path = _out_path(filename, clean_basename)

        try:
            source_path = _original_path(content, filename)
            shutil.copy2(source_path, save_path)
            print(f"[OK] DOCX '{filename}' copié : {save_path}")
        except Exception as e:
            print(f"[ERREUR] DOCX {filename} : {e}")

    return datas


def clean_IMAGES(datas):
    for filename, content in datas.items():
        if str(content.get("type", "")).lower() not in {"image", "jpg", "jpeg", "png"}:
            continue

        basename = Path(filename).name
        clean_basename = f"cleaned_{basename}"
        save_path = _out_path(filename, clean_basename)

        try:
            source_path = _original_path(content, filename)
            shutil.copy2(source_path, save_path)
            print(f"[OK] Image '{filename}' copiée : {save_path}")
        except Exception as e:
            print(f"[ERREUR] Image {filename} : {e}")

    return datas
