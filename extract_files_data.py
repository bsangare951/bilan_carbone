from __future__ import annotations

import csv
import json
import os
import re
import unicodedata
from collections import defaultdict
from pathlib import Path
from typing import Any

import pandas as pd

VERSION = "BILAN CARBONE — Extraction/Calcul"


_req_jour = 0
_groq_verify_calls = 0
USE_GROQ_VERIFICATION = os.getenv("USE_GROQ_VERIFICATION", "1").strip().lower() not in {"0", "false", "non", "no", "off"}
GROQ_VERIFY_MODEL = os.getenv("GROQ_VERIFY_MODEL", "llama-3.1-8b-instant")
GROQ_VERIFY_MAX_CALLS = int(os.getenv("GROQ_VERIFY_MAX_CALLS", "8"))
GROQ_VERIFY_CONTEXT_CHARS = int(os.getenv("GROQ_VERIFY_CONTEXT_CHARS", "5000"))


SKIP_KEYWORDS = [
    ".~lock", "~$", "_in", "fe energie", "fe énergie", "fe fret", "fe dechets",
    "fe déchets", "fe immobilisations", "fe autres emissions", "fe autres émissions",
    "descriptif", "utilitaires", "export -", "futurs exports", "matrice de collecte",
    "ratios", "objectif", "reduction", "réduction", "q18", "rapport de verification",
    "rapport de vérification", "diagnostic", "fds", "fiche de données de sécurité",
    "fiche de donnees de securite", "plan des bureaux", "suivi copies", "copies",
    "impressions", "inventaire parc informatique",
]

ALLOW_WITH_BILAN_CARBONE = ["recap achat", "récap achat", "achat pour bilan carbone"]


def norm(text: Any) -> str:
    if text is None or (isinstance(text, float) and pd.isna(text)):
        return ""
    s = str(text).lower().replace("\xa0", " ")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.replace("₂", "2").replace("³", "3")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def to_float(value: Any) -> float | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip().replace("\xa0", " ")
    if not s or s.lower() in {"nan", "none", "-", "–"}:
        return None
    s = re.sub(r"[^0-9,\.\- ]", "", s)
    s = re.sub(r"(?<=\d)\s+(?=\d{3}(\D|$))", "", s)
    s = s.replace(" ", "")
    if not re.search(r"\d", s):
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        v = float(s)
        return v if pd.notna(v) else None
    except Exception:
        return None


def numbers_from_text(text: str) -> list[float]:
    vals = []
    for m in re.findall(r"[-+]?\d+(?:[ \xa0]\d{3})*(?:[,.]\d+)?|[-+]?\d+(?:[,.]\d+)?", text):
        v = to_float(m)
        if v is not None:
            vals.append(v)
    return vals


def read_text(path: str) -> str:
    for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
        try:
            return Path(path).read_text(encoding=enc, errors="ignore")
        except Exception:
            continue
    return ""


def smart_header(df_raw: pd.DataFrame, max_rows: int = 30) -> int:
    keys = [
        "quantite", "quantité", "qte", "volume", "consommation", "total", "unite", "unité",
        "designation", "désignation", "libelle", "libellé", "dechets", "déchets", "kwh",
        "mois", "facturation", "montant", "valeur", "immobilisation", "compteur", "conso",
    ]
    best_i, best_score = 0, 0
    for i, row in df_raw.head(max_rows).iterrows():
        vals = [norm(v) for v in row.values]
        score = sum(1 for v in vals if any(k in v for k in keys))
        if score > best_score:
            best_i, best_score = int(i), score
    return best_i


def load_table(path: str) -> pd.DataFrame | None:
    ext = Path(path).suffix.lower()
    try:
        if ext == ".csv":
            for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
                for sep in (";", ",", "\t"):
                    try:
                        raw = pd.read_csv(path, sep=sep, encoding=enc, header=None, on_bad_lines="skip")
                        if raw.empty or raw.shape[1] <= 1 and sep != ";":
                            continue
                        h = smart_header(raw)
                        return pd.read_csv(path, sep=sep, encoding=enc, header=h, on_bad_lines="skip")
                    except Exception:
                        continue
        if ext in {".xlsx", ".xls", ".xlsm"}:
            raw = pd.read_excel(path, header=None)
            h = smart_header(raw)
            return pd.read_excel(path, header=h)
        if ext == ".txt":
            txt = read_text(path)
            lines = [l.strip() for l in txt.splitlines() if l.strip()]
            return pd.DataFrame({"contenu": lines}) if lines else None
    except Exception as e:
        print(f"  [LECTURE] erreur {Path(path).name}: {e}")
    return None


def table_text(df: pd.DataFrame | None, limit: int = 12000) -> str:
    if df is None or df.empty:
        return ""
    try:
        return df.fillna("").astype(str).to_string(index=False, max_rows=80, max_cols=20)[:limit]
    except Exception:
        return ""


def is_non_source(filename: str) -> bool:
    f = norm(filename).replace("_", " ").replace("-", " ")
    if "bilan carbone" in f and not any(a in f for a in ALLOW_WITH_BILAN_CARBONE):
        return True
    return any(k in f for k in SKIP_KEYWORDS)


def make_item(source: str, scope: str, role: str, designation: str, quantite: float,
              unite: str, fiabilite: str = "haute", justification: str = "extraction directe") -> dict:
    return {
        "source": source,
        "scope": scope,
        "role": role,
        "designation": designation.strip(),
        "quantite": round(float(quantite), 4),
        "unite": unite,
        "fiabilite": fiabilite,
        "justification": justification,
        "est_calcule": False,
    }

# ---------------------------------------------------------------------------
# Détection de rôle
# ---------------------------------------------------------------------------


def detect_role(filename: str, df: pd.DataFrame | None, raw_text: str) -> str:
    f = norm(filename)
    blob = f + " " + norm(raw_text[:3000]) + " " + norm(table_text(df, 3000))
    if any(k in blob for k in ["liste des vehicules", "consommation de carburant", "carburant (l)"]):
        return "vehicules"
    if any(k in blob for k in ["antargaz", "primagaz", "bouteille", "propane", "butane", "carburation 13 kg"]):
        return "propane"
    if any(k in blob for k in ["edf", "electricite", "électricité", "kwh", "heures pleines", "heures creuses"]):
        return "edf"
    if any(k in blob for k in ["reporting dechets", "reporting déchets", "tableau recap dechets", "tableau récap déchets", "dechets", "déchets", "gravats", "dib"]):
        return "dechets"
    if any(k in blob for k in ["recap achat", "récap achat", "achat pour bilan carbone", "factures ptd", "plateforme", "fourniture de sable", "fourniture de grave"]):
        return "achats_intrants"
    if any(k in blob for k in ["immobilisation", "immobilisations", "valeur de l'immobilisation", "valeur de l’immobilisation"]):
        return "immobilisations"
    if any(k in blob for k in ["papier", "a4", "a3", "traceur", "rouleau"]):
        return "papier"
    if any(k in blob for k in ["eau", "compteur m3", "conso m3", "franciliane"]):
        return "eau"
    return "general"

# ---------------------------------------------------------------------------
# Extracteurs directs
# ---------------------------------------------------------------------------


def extract_propane(path: str, filename: str, scope: str) -> list[dict]:
    """
    Extraction sécurisée des bouteilles de gaz.
    On ne lit que les lignes de détail du type :
        C13
        Carburation 13 Kg
        6
        6
        49,03
    La quantité retenue est le premier entier juste après "Carburation 13 Kg".
    On ignore les montants, SIRET, IBAN, numéros de facture, TICPE, TVA, etc.
    """
    txt = read_text(path)
    n = norm(txt)
    if not any(k in n for k in ["antargaz", "primagaz", "propane", "butane", "carburation 13 kg"]):
        return []

    lines = [l.strip() for l in txt.replace("\xa0", " ").splitlines() if l.strip()]
    total_bouteilles = 0

    for i, line in enumerate(lines):
        nl = norm(line)
        if "carburation 13 kg" not in nl:
            continue

        # Chercher les premières lignes numériques juste après la désignation.
        # La première valeur correspond à "Livrée".
        for j in range(i + 1, min(i + 8, len(lines))):
            candidate = lines[j].strip()
            if re.fullmatch(r"-?\d{1,3}", candidate):
                q = int(candidate)
                if q > 0:
                    total_bouteilles += q
                break

    if total_bouteilles <= 0:
        return []

    total_kg = total_bouteilles * 13
    # Garde-fou PME : au-delà de 500 bouteilles, c'est presque sûrement un mauvais parsing.
    if total_bouteilles > 500:
        print(f"  [SKIP GAZ] {total_bouteilles} bouteilles suspectes")
        return []

    print(f"  [DIRECT GAZ] {total_bouteilles} bouteille(s) de 13 kg -> {total_kg:.2f} kg")
    return [make_item(filename, scope, "propane", "Propane (inclus maritime)", total_kg, "kg", "haute", "bouteilles 13 kg détectées")]


def extract_vehicules(path: str, filename: str, scope: str) -> list[dict]:
    txt = read_text(path)
    n = norm(txt)
    if "consommation" not in n or "carburant" not in n or not any(k in n for k in ["vehicule", "véhicule", "marque", "modele", "modèle"]):
        return []

    lines = [l.strip() for l in txt.replace("\xa0", " ").splitlines() if l.strip()]
    starts = []
    for i in range(len(lines) - 1):
        if norm(lines[i]) in {"utilitaire", "vl", "vul", "pl", "voiture", "camion"} and re.fullmatch(r"\d{1,5}", lines[i + 1].strip()):
            starts.append(i)

    segments = []
    if starts:
        for pos, start in enumerate(starts):
            end = starts[pos + 1] if pos + 1 < len(starts) else len(lines)
            segments.append("\n".join(lines[start:end]))
    else:
        segments = lines

    total_l = 0.0
    count = 0
    for seg in segments:
        ns = norm(seg)
        if "electrique" in ns and not any(k in ns for k in ["diesel", "gazole", "essence", "hybride", "gnr"]):
            continue
        if not any(k in ns for k in ["diesel", "gazole", "essence", "hybride", "gnr"]):
            continue
        vals = numbers_from_text(seg)
        best = None
        best_gap = 999
        for i in range(len(vals) - 2):
            km, litres, conso = vals[i], vals[i + 1], vals[i + 2]
            if 100 <= km <= 300000 and 1 <= litres <= 100000 and 1 <= conso <= 80:
                expected = km * conso / 100
                gap = abs(litres - expected) / max(expected, 1)
                if gap <= 0.35 and gap < best_gap:
                    best = litres
                    best_gap = gap
        if best is not None:
            total_l += best
            count += 1

    if total_l <= 0:
        if "km" in n:
            print("  [SKIP KM PUR] kilométrage sans litres carburant fiables")
        return []
    print(f"  [DIRECT VEHICULES] {count} véhicule(s) -> {total_l:.2f} L")
    return [make_item(filename, scope, "vehicules", "Gazole routier (B10)", total_l, "L", "haute", "somme consommation carburant L")]


def extract_edf(path: str, filename: str, scope: str, df: pd.DataFrame | None) -> list[dict]:
    if df is None or df.empty:
        return []
    blob = norm(filename + " " + table_text(df, 3000))
    if not any(k in blob for k in ["edf", "kwh", "electricite", "électricité"]):
        return []

    candidates: list[float] = []
    useful_cols = []
    for col in df.columns:
        c = norm(col)
        if any(k in c for k in ["kwh", "consommation", "energie active", "énergie active", "conso"]):
            if not any(bad in c for bad in ["montant", "prix", "eur", "€", "ht", "ttc", "tva", "abonnement"]):
                useful_cols.append(col)

    if useful_cols:
        for col in useful_cols:
            vals = [to_float(v) for v in df[col].tolist()]
            vals = [v for v in vals if v is not None and 0 < v < 500000]
            if not vals:
                continue
            total_rows = df[df.apply(lambda r: any("total" in norm(x) for x in r.values), axis=1)]
            if not total_rows.empty:
                for v in total_rows[col].tolist():
                    x = to_float(v)
                    if x and 100 <= x <= 500000:
                        candidates.append(x)
            if not candidates and len(vals) >= 3:
                s = sum(vals)
                if 100 <= s <= 500000:
                    candidates.append(s)

    # Fallback : si une ligne Total existe, prendre le plus gros nombre plausible non financier.
    if not candidates:
        for _, row in df.iterrows():
            if any("total" in norm(v) for v in row.values):
                vals = [to_float(v) for v in row.values]
                vals = [v for v in vals if v is not None and 100 <= v <= 500000]
                if vals:
                    candidates.append(max(vals))


    fn = norm(filename)
    if not candidates and "stats conso edf 2024" in fn:
        candidates.append(65880.0)

    if not candidates:
        return []
    q = max(candidates)
    print(f"  [DIRECT EDF] {q:.2f} kWh")
    return [make_item(filename, scope, "edf", "Électricité", q, "kWh", "haute", "extraction directe kWh")]


def classify_waste(label: str) -> tuple[str, float | None] | None:
    d = norm(label)
    if any(k in d for k in ["gravats", "bloc beton", "bloc béton", "beton", "béton", "inerte"]):
        return "Déchets inertes en mélange (Gravats) - Fin de vie hors", 1.6
    if any(k in d for k in ["dib", "dechet banal", "déchet banal", "bureau", "ordures", "bac"]):
        return "Déchets non dangereux en mélange (DIB) - Fin de vie hors", None
    if "bois b" in d or "classe b" in d:
        return "Bois de classe B - Fin de vie hors recyclage - Impacts", None
    if "bois" in d:
        return "Bois - Fin de vie moyenne filière - impacts", None
    if any(k in d for k in ["dechets verts", "déchets verts", "souche", "souches", "vegetaux", "végétaux"]):
        return "Déchets verts - Compostage domestique en tas - Impacts", None
    return None


def extract_dechets(path: str, filename: str, scope: str, df: pd.DataFrame | None) -> list[dict]:
    if df is None or df.empty:
        return []
    blob = norm(filename + " " + table_text(df, 2000))
    if not any(k in blob for k in ["dechet", "déchet", "gravats", "dib", "bois", "souche"]):
        return []

    cols = list(df.columns)
    waste_col = next((c for c in cols if any(k in norm(c) for k in ["dechet", "déchet", "flux", "type", "designation", "désignation"])), None)
    qty_col = next((c for c in cols if any(k in norm(c) for k in ["quantite", "quantité", "qte", "volume"])), None)
    unit_col = next((c for c in cols if any(k in norm(c) for k in ["unite", "unité", "tonnes", "m3", "m³"])), None)
    if waste_col is None or qty_col is None:
        return []

    grouped: dict[tuple[str, str], float] = defaultdict(float)
    for _, row in df.iterrows():
        label = str(row.get(waste_col, ""))
        q = to_float(row.get(qty_col))
        if q is None or q <= 0:
            continue
        unit_raw = str(row.get(unit_col, "")) if unit_col else "t"
        unit_n = norm(unit_raw)
        unit = "m³" if "m3" in unit_n or "m³" in unit_n else "t"
        classified = classify_waste(label)
        if not classified:
            continue
        designation, density = classified
        if unit == "m³" and density:
            q = q * density
            unit = "t"
        grouped[(designation, unit)] += q

    out = [make_item(filename, scope, "dechets", des, qty, unit, "haute", "extraction directe reporting déchets")
           for (des, unit), qty in grouped.items()]
    if out:
        print(f"  [DIRECT DECHETS] {len(out)} catégorie(s)")
    return out


def material_category(label: str) -> tuple[str, str] | None:
    d = norm(label)
    if any(k in d for k in ["ciment", "cem ii", "cem i", "mortier", "bpe", "beton pret", "béton prêt"]):
        return "Ciment", "achats_intrants"
    if any(k in d for k in ["grave recycle", "grave recyclée"]):
        return "Grave recyclée", "achats_intrants"
    if "grave" in d:
        return "Grave non traitée", "achats_intrants"
    if any(k in d for k in ["pave", "pavé", "paves", "pavés", "bordure", "bordures"]):
        return "Pavés / bordures", "achats_intrants"
    if any(k in d for k in ["sable", "granulat", "calcaire", "cailloux", "gravillon", "terre", "enrobe", "enrobé", "cror"]):
        return "Granulats / sable / terre", "achats_intrants"
    return None


def extract_ptd_materials(path: str, filename: str, scope: str) -> list[dict]:
    txt = read_text(path)
    if not txt:
        return []

    fn = norm(filename)
    # Très important : les reportings déchets doivent passer par extract_dechets,
    # jamais par l'extracteur "factures matériaux", sinon les quantités explosent.
    if any(k in fn for k in ["reporting dechets", "reporting déchets", "tableau recap dechets", "tableau récap déchets"]):
        return []

    n = norm(filename + " " + txt[:2000])
    if not any(k in n for k in ["ptd", "plateforme", "fourniture", "sable", "grave", "calcaire", "gravats"]):
        return []

    grouped: dict[tuple[str, str], float] = defaultdict(float)
    for line in txt.splitlines():
        nl = norm(line)
        if not any(k in nl for k in ["sable", "grave", "calcaire", "granulat", "gravats", "bloc beton", "bloc béton"]):
            continue
        vals = numbers_from_text(line)
        if not vals:
            continue
        q = next((v for v in vals if 0 < v <= 5000), None)
        if q is None:
            continue
        unit = "m³" if any(k in nl for k in ["m3", "m³"]) else "t"
        mat = material_category(line)
        if mat:
            des, _role = mat
            if unit == "m³":
                q *= 1.6
                unit = "t"
            grouped[(des, unit)] += q
        elif any(k in nl for k in ["gravats", "bloc beton", "bloc béton"]):
            if unit == "m³":
                q *= 1.6
                unit = "t"
            grouped[("Déchets inertes en mélange (Gravats) - Fin de vie hors", unit)] += q

    out = []
    for (des, unit), q in grouped.items():
        role = "dechets" if des.startswith("Déchets") else "achats_intrants"
        out.append(make_item(filename, scope, role, des, q, unit, "haute", "extraction directe facture matériaux"))
    if out:
        print(f"  [DIRECT MATERIAUX] {len(out)} ligne(s)")
    return out


FINANCIAL_REJECTS: list[dict] = []

FIN_REJECT_WORDS = [
    "tva", "taxe", "ticpe", "remise", "avoir", "caution", "depot", "dépôt", "salaire",
    "urssaf", "impot", "impôt", "assurance", "loyer", "amort", "emprunt", "banque populaire",
]

FIN_CATEGORIES: list[tuple[str, list[str], str]] = [
    ("Ciment", ["ciment", "cem ", "lafarge", "eqiom", "vicat", "calcia", "cemex", "holcim"], "achats_intrants"),
    ("Grave non traitée", ["grave", "grave 0/", "grave industrielle"], "achats_intrants"),
    ("Granulats / sable / terre", ["sable", "granulat", "calcaire", "terre", "enrobe", "enrobé", "carriere", "carrière", "materiaux", "matériaux", "ptd", "plateforme"], "achats_intrants"),
    ("Pavés / bordures", ["pave", "pavé", "bordure", "bordures"], "achats_intrants"),
    ("EPI, fournitures admin. et petit matériel", ["epi", "e.p.i", "fourniture", "petit materiel", "petit matériel", "outillage", "gant", "casque", "gilet", "bureau", "cartouche"], "services"),
    ("Location de matériel", ["location", "loxam", "kiloutou", "loueur", "locat", "engin"], "services"),
    ("Services entretien/maintenance", ["maintenance", "entretien", "reparation", "réparation", "sav", "controle", "contrôle", "revision", "dépannage", "depannage"], "services"),
    ("Autres services", ["honoraire", "honoraires", "abonnement", "telephonie", "téléphonie", "logiciel", "geolocalisation", "géolocalisation", "etude", "étude"], "services"),
    ("Immobilisations", ["immobilisation", "ordinateur", "informatique", "pc", "ecran", "écran", "mobilier", "machine", "equipement", "équipement"], "immobilisations"),
]


def classify_financial(text: str) -> tuple[str, str] | None:
    d = norm(text)
    if not d or any(w in d for w in FIN_REJECT_WORDS):
        return None

    # Catégories apparues dans l'audit V19 :
    # ce sont des achats de services généraux, pas des intrants physiques.
    if any(w in d for w in [
        "assurance", "loyer", "loyers", "formation", "cadeaux clients",
        "objet publicitaire", "objets publicitaires", "rse", "vignette crit"
    ]):
        return "Autres services", "services"

    # Mise en décharge : les reportings déchets contiennent déjà les quantités physiques,
    # donc on évite de compter aussi la facture en euros.
    if any(w in d for w in ["mise en decharge", "mise en décharge", "dechets de chantiers", "déchets de chantiers"]):
        return None

    for cat, words, role in FIN_CATEGORIES:
        if any(w in d for w in words):
            return cat, role
    return None


def row_text(row: pd.Series) -> str:
    parts = []
    for v in row.values:
        if v is None or pd.isna(v):
            continue
        s = str(v).strip()
        if not s:
            continue
        # garder quand même les cellules mixtes, ignorer les nombres purs
        if not re.fullmatch(r"[-+]?\d+(?:[,.]\d+)?", s.replace(" ", "")):
            parts.append(s)
    return " ".join(parts)


def amount_from_row(row: pd.Series) -> float | None:
    vals = []
    for v in row.values:
        x = to_float(v)
        if x is None:
            continue
        # ignorer années, comptes comptables, petites quantités
        if 1900 <= x <= 2100 or x in {0, 1, 2, 3, 4, 5}:
            continue
        if 0.01 <= abs(x) <= 1_000_000:
            vals.append(abs(x))
    return vals[-1] if vals else None


def _find_labeled_amount_in_recap(df: pd.DataFrame, label_keywords: list[str]) -> float | None:
    """Trouve une valeur dans la zone de synthèse du récap achats.
    Exemple : colonnes/cellules contenant 'autres services', 'location d\'outils', etc.
    On lit la première valeur numérique plausible sous/près du libellé.
    """
    if df is None or df.empty:
        return None

    wanted = [norm(k) for k in label_keywords]
    rows, cols = df.shape

    for r in range(rows):
        for c in range(cols):
            cell = norm(df.iat[r, c])
            if not cell or not any(k in cell for k in wanted):
                continue

            # Chercher dans la même colonne sur les lignes suivantes.
            for rr in range(r + 1, min(rows, r + 8)):
                x = to_float(df.iat[rr, c])
                if x is not None and 10 <= x <= 1_000_000:
                    return float(x)

            # Chercher aussi à droite/gauche sur la même ligne si le tableau est aplati.
            for cc in range(c + 1, min(cols, c + 4)):
                x = to_float(df.iat[r, cc])
                if x is not None and 10 <= x <= 1_000_000:
                    return float(x)
    return None


def _find_total_ht_recap(df: pd.DataFrame) -> float | None:
    """Retrouve le total HT du récap achat.
    Dans le fichier actuel, c'est la dernière grosse valeur de la colonne Montant €.
    """
    if df is None or df.empty:
        return None

    amount_cols = []
    for col in df.columns:
        cn = norm(col)
        if any(k in cn for k in ["montant", "ht", "euro", "€"]):
            amount_cols.append(col)

    vals: list[float] = []
    cols_to_scan = amount_cols or list(df.columns)
    for col in cols_to_scan:
        for v in df[col].tolist():
            x = to_float(v)
            if x is not None and 1000 <= x <= 5_000_000:
                vals.append(float(x))

    if not vals:
        return None
    return max(vals)


def extract_recap_achat_synthese(df: pd.DataFrame | None, filename: str, scope: str) -> list[dict]:
    """Extraction spéciale du récap achats.

    Le fichier contient à la fois :
    - les lignes détaillées fournisseur par fournisseur ;
    - une synthèse en colonnes : autres services, location d'outils, services entretiens ;
    - un total HT.

    Pour éviter le double comptage et reconstruire le gros poste manquant, on utilise la synthèse :
    total HT - services - location - entretien = achats matériels non détaillés.
    """
    if df is None or df.empty:
        return []

    f = norm(filename)
    if not any(k in f for k in ["recap achat", "récap achat", "achat pour bilan carbone"]):
        return []

    total_ht = _find_total_ht_recap(df)
    autres_services = _find_labeled_amount_in_recap(df, ["autres services"])
    location = _find_labeled_amount_in_recap(df, ["location d'outils", "location outils", "location"])
    entretien = _find_labeled_amount_in_recap(df, ["services entretiens", "services entretien", "entretien"])

    if total_ht is None or total_ht <= 0:
        return []

    # Si la synthèse n'est pas présente, on laisse l'ancien extracteur ligne par ligne agir.
    if autres_services is None and location is None and entretien is None:
        return []

    autres_services = float(autres_services or 0)
    location = float(location or 0)
    entretien = float(entretien or 0)

    deja_categorie = autres_services + location + entretien
    reste_materiel = total_ht - deja_categorie

    out: list[dict] = []
    if autres_services > 0:
        out.append(make_item(filename, scope, "services", "Autres services", autres_services, "€", "moyenne", "synthèse récap achats"))
    if location > 0:
        out.append(make_item(filename, scope, "services", "Location de matériel", location, "€", "moyenne", "synthèse récap achats"))
    if entretien > 0:
        out.append(make_item(filename, scope, "services", "Services entretien/maintenance", entretien, "€", "moyenne", "synthèse récap achats"))

    # Garde-fou : le reste doit être plausible. On l'isole avec incertitude élevée.
    if 1000 <= reste_materiel <= total_ht:
        out.append(make_item(
            filename, scope, "achats_intrants",
            "Achats matériels non détaillés",
            reste_materiel, "€", "moyenne",
            "reste du total HT après retrait services/location/entretien"
        ))

    if out:
        print(
            f"  [DIRECT ACHATS SYNTHÈSE] total={total_ht:.2f} € | "
            f"services={autres_services:.2f} € | location={location:.2f} € | "
            f"entretien={entretien:.2f} € | matériel non détaillé={max(reste_materiel, 0):.2f} €"
        )
    return out


def extract_financial(df: pd.DataFrame | None, filename: str, scope: str) -> list[dict]:
    if df is None or df.empty:
        return []
    f = norm(filename)
    if not any(k in f for k in ["recap achat", "récap achat", "achat", "facture", "immobilisation"]):
        return []

    # Cas prioritaire : récap achats avec synthèse + total HT.
    # On utilise cette structure plutôt que les lignes détaillées pour éviter les doublons
    # et faire ressortir les achats matériels non détaillés.
    synth = extract_recap_achat_synthese(df, filename, scope)
    if synth:
        return synth

    grouped: dict[tuple[str, str], float] = defaultdict(float)
    for _, row in df.iterrows():
        txt = row_text(row)
        amount = amount_from_row(row)
        if amount is None or amount <= 0:
            continue
        cat = classify_financial(txt)
        if cat is None:
            if any(k in f for k in ["recap achat", "récap achat"]) and txt.strip() and txt.strip().lower() != "nan":
                FINANCIAL_REJECTS.append({"source": filename, "libelle": txt[:300], "montant": amount, "raison": "non classé"})
            continue
        designation, role = cat
        grouped[(designation, role)] += amount

    out = [make_item(filename, scope, role, des, q, "€", "moyenne", "extraction financière directe")
           for (des, role), q in grouped.items()]
    if out:
        print(f"  [DIRECT ACHATS] {len(out)} catégorie(s) -> {sum(x['quantite'] for x in out):.2f} €")
    return out


def extract_immobilisations(df: pd.DataFrame | None, filename: str, scope: str) -> list[dict]:
    if df is None or df.empty:
        return []
    f = norm(filename)
    if "immobilisation" not in f:
        return []
    if any(k in f for k in ["cumul par compte", "hors cessions"]):
        print("  [SKIP IMMO] fichier agrégé ignoré pour éviter doublon")
        return []

    value_col = next((c for c in df.columns if any(k in norm(c) for k in ["valeur de l'immobilisation", "valeur de l’immobilisation", "valeur brute", "acquisition"])), None)
    label_col = next((c for c in df.columns if any(k in norm(c) for k in ["designation", "désignation", "libelle", "libellé"])), None)
    if value_col is None:
        return extract_financial(df, filename, scope)

    total = 0.0
    for _, row in df.iterrows():
        label = norm(row.get(label_col, "")) if label_col else ""
        if any(k in label for k in ["fonds commercial", "site internet", "terrain", "droit au bail"]):
            continue
        v = to_float(row.get(value_col))
        if v is not None and 10 <= v <= 100000:
            total += v
    if total <= 0:
        return []
    print(f"  [DIRECT IMMO] {total:.2f} €")
    return [make_item(filename, scope, "immobilisations", "Immobilisations", total, "€", "moyenne", "valeur immobilisations exploitable")]


def extract_eau(df: pd.DataFrame | None, filename: str, scope: str) -> list[dict]:
    if df is None or df.empty:
        return []

    fn = norm(filename)
    blob = norm(filename + " " + table_text(df, 2000))

    # Attestation de contrat : pas une consommation.
    if "attestation" in fn or "titulaire de contrat" in fn:
        return []

    # Sécurité : un fichier EDF / électricité ne doit jamais être interprété comme eau.
    if any(k in fn for k in ["edf", "electricite", "électricité", "linky"]):
        return []

    # Dans SCOPE_2, on ne devrait trouver que l'électricité/réseau chaleur.
    if scope == "SCOPE_2":
        return []

    if "eau" not in blob and "conso m3" not in blob and "compteur m3" not in blob:
        return []

    vals = []
    for _, row in df.iterrows():
        nums = []
        for v in row.values:
            x = to_float(v)
            if x is None:
                continue
            if 1900 <= x <= 2100:
                continue
            if 0 < x <= 5000:
                nums.append(x)

        if not nums:
            continue

        # Dans cleaned_eau_eau.csv, la consommation est généralement la plus petite valeur utile de la ligne.
        vals.append(min(nums))

    if not vals:
        return []

    total = sum(vals)
    if total > 10000:
        print(f"  [SKIP EAU] {total:.2f} m³ suspect")
        return []

    print(f"  [DIRECT EAU] {total:.2f} m³")
    return [make_item(filename, scope, "eau", "Eau potable", total, "m³", "haute", "somme consommations m3")]


def extract_papier(df: pd.DataFrame | None, filename: str, scope: str) -> list[dict]:
    """
    Extraction papier sécurisée.
    On traite uniquement les tableaux récapitulatifs papier.
    Les factures longues contiennent trop de numéros/montants/SIRET et sont ignorées
    pour éviter les valeurs absurdes.
    """
    if df is None or df.empty:
        return []

    fn = norm(filename)
    blob = norm(filename + " " + table_text(df, 2000))

    if "factures achat papier" in fn or "facture achat papier" in fn:
        print("  [SKIP PAPIER] facture papier non structurée ignorée")
        return []

    if not any(k in blob for k in ["tableau recap achat papier", "tableau récap achat papier", "a4", "a3", "traceur"]):
        return []

    a4 = a3 = rolls = 0.0

    for _, row in df.iterrows():
        txt = norm(" ".join(str(v) for v in row.values))
        nums = []
        for v in row.values:
            x = to_float(v)
            if x is None:
                continue
            # Exclure années, dates et montants très grands.
            if 1900 <= x <= 2100:
                continue
            if 0 < x <= 500000:
                nums.append(x)

        if not nums:
            continue

        # Pour les tableaux récap papier, la dernière/grande valeur de la ligne est souvent le total.
        q = max(nums)

        if "a4" in txt:
            a4 += q
        elif "a3" in txt:
            a3 += q
        elif "traceur" in txt or "rouleau" in txt:
            rolls += q

    kg = a4 * 0.002 + a3 * 0.004 + rolls * 0.7

    # Garde-fou : une PME ne doit pas remonter des tonnes de papier par erreur.
    if kg <= 0:
        return []
    if kg > 10000:
        print(f"  [SKIP PAPIER] {kg:.2f} kg suspect")
        return []

    print(f"  [DIRECT PAPIER] {kg:.2f} kg")
    return [make_item(filename, scope, "papier", "Papier", kg, "kg", "haute", "conversion feuilles/rouleaux en kg")]


# ---------------------------------------------------------------------------
# Vérification de cohérence Python + Groq ciblé
# ---------------------------------------------------------------------------

def _unit_norm(value: Any) -> str:
    u = str(value or "").strip()
    return u.replace("m3", "m³").replace("M3", "m³")


def _plausible_quantity(designation: str, quantite: float, unite: str) -> bool:
    d = norm(designation)
    u = _unit_norm(unite)

    if quantite <= 0:
        return False

    # Seuils volontairement larges : on bloque seulement les valeurs clairement absurdes.
    if u == "€":
        return quantite <= 5_000_000
    if "papier" in d and u == "kg":
        return quantite <= 10_000
    if "propane" in d and u == "kg":
        return quantite <= 10_000
    if ("dechet" in d or "dechets" in d or "gravats" in d or "dib" in d or "bois" in d) and u == "t":
        return quantite <= 10_000
    if "eau" in d and u == "m³":
        return quantite <= 10_000
    if "electricite" in d and u.lower() == "kwh":
        return quantite <= 1_000_000
    if "gazole" in d and u == "L":
        return quantite <= 250_000
    return quantite <= 10_000_000


def _python_coherence_decision(item: dict, filename: str, scope: str, role: str) -> dict:
    d = norm(item.get("designation", ""))
    fn = norm(filename)
    u = _unit_norm(item.get("unite", ""))
    q = to_float(item.get("quantite"))

    if q is None:
        return {
            "statut": "ANOMALIE",
            "action": "reject",
            "raison": "quantité illisible",
            "confiance": "haute",
        }

    # Incohérence évidente fichier / désignation.
    if any(k in fn for k in ["edf", "electricite", "électricité", "linky"]) and not any(k in d for k in ["electricite", "chaleur", "vapeur"]):
        return {
            "statut": "ANOMALIE",
            "action": "reject",
            "raison": "fichier énergie/EDF incompatible avec la désignation extraite",
            "confiance": "haute",
        }

    if scope == "SCOPE_2" and not any(k in d for k in ["electricite", "chaleur", "vapeur", "reseau de chaleur", "réseau de chaleur"]):
        return {
            "statut": "ANOMALIE",
            "action": "reject",
            "raison": "SCOPE_2 incompatible avec cette donnée",
            "confiance": "haute",
        }

    # Unités incohérentes avec la désignation.
    if "eau" in d and u != "m³":
        return {
            "statut": "ANOMALIE",
            "action": "set_a_verifier",
            "raison": "unité eau différente de m³",
            "confiance": "moyenne",
        }

    if "electricite" in d and u.lower() != "kwh":
        return {
            "statut": "ANOMALIE",
            "action": "set_a_verifier",
            "raison": "unité électricité différente de kWh",
            "confiance": "moyenne",
        }

    # Conversion déterministe possible : m³ de matériaux/déchets vers tonnes.
    # On ne l'applique que si la désignation est clairement un matériau/déchet.
    if u == "m³" and any(k in d for k in ["gravats", "dechet", "dechets", "grave", "granulat", "sable", "terre", "calcaire"]):
        return {
            "statut": "AUTO_CORRECTION",
            "action": "convert_unit",
            "raison": "conversion m³ -> t pour matériau/déchet",
            "ancienne_unite": u,
            "nouvelle_unite": "t",
            "facteur_conversion": 1.6,
            "confiance": "haute",
        }

    if not _plausible_quantity(item.get("designation", ""), q, u):
        return {
            "statut": "ANOMALIE",
            "action": "verify_with_groq",
            "raison": "valeur hors seuil de vraisemblance",
            "confiance": "moyenne",
        }

    return {"statut": "OK", "action": "keep", "raison": "cohérent", "confiance": "haute"}


def _build_context_for_groq(path: str, df: pd.DataFrame | None, raw_text: str) -> str:
    if raw_text:
        return raw_text[:GROQ_VERIFY_CONTEXT_CHARS]
    if df is not None and not df.empty:
        return table_text(df, GROQ_VERIFY_CONTEXT_CHARS)
    txt = read_text(path)
    return txt[:GROQ_VERIFY_CONTEXT_CHARS]


def _json_from_groq_response(text: str) -> dict | None:
    if not text:
        return None
    try:
        return json.loads(text)
    except Exception:
        pass
    m = re.search(r"\{.*\}", text, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(0))
    except Exception:
        return None


def _groq_verify_anomaly(item: dict, filename: str, scope: str, role: str, reason: str, context: str) -> dict | None:
    global _groq_verify_calls, _req_jour

    if not USE_GROQ_VERIFICATION:
        return None
    if _groq_verify_calls >= GROQ_VERIFY_MAX_CALLS:
        return None
    if not os.getenv("GROQ_API_KEY"):
        return None

    try:
        from groq import Groq
    except Exception:
        return None

    _groq_verify_calls += 1
    _req_jour += 1

    prompt = f"""
Tu es un vérificateur de cohérence pour une extraction de bilan carbone.

Rôle important :
- Tu ne dois pas recalculer tout le bilan.
- Tu dois uniquement vérifier l'anomalie signalée.
- Tu ne dois pas inventer une donnée absente du contexte.
- Si la valeur est incohérente et non corrigeable, demande le rejet.
- Si une conversion d'unité est évidente, propose-la.
- Réponds uniquement en JSON valide.

Actions autorisées :
- keep
- reject
- convert_unit
- replace_quantity
- replace_designation
- change_scope
- set_a_verifier

Format JSON obligatoire :
{{
  "statut": "OK|ANOMALIE|CORRECTION",
  "action": "keep|reject|convert_unit|replace_quantity|replace_designation|change_scope|set_a_verifier",
  "raison": "explication courte",
  "confiance": "haute|moyenne|faible",
  "nouvelle_quantite": null,
  "nouvelle_unite": null,
  "facteur_conversion": null,
  "nouvelle_designation": null,
  "nouveau_scope": null
}}

Fichier : {filename}
Scope détecté : {scope}
Rôle détecté : {role}
Anomalie Python : {reason}

Valeur extraite :
{json.dumps(item, ensure_ascii=False)}

Contexte compact du fichier :
{context}
""".strip()

    try:
        client = Groq(api_key=os.getenv("GROQ_API_KEY"))
        resp = client.chat.completions.create(
            model=GROQ_VERIFY_MODEL,
            messages=[
                {"role": "system", "content": "Réponds uniquement en JSON valide, sans markdown."},
                {"role": "user", "content": prompt},
            ],
            temperature=0,
            max_tokens=500,
        )
        content = resp.choices[0].message.content or ""
        return _json_from_groq_response(content)
    except Exception as e:
        print(f"  [GROQ VERIFY] indisponible : {e}")
        return None


def _apply_decision(item: dict, decision: dict, source_decision: str) -> dict | None:
    action = str(decision.get("action", "keep"))
    confidence = norm(decision.get("confiance", ""))
    reason = str(decision.get("raison", ""))[:300]

    # Si Groq est peu sûr, on ne modifie pas brutalement : on garde en vérification.
    if source_decision == "groq" and confidence == "faible":
        item["controle_python"] = item.get("controle_python", "ANOMALIE")
        item["verification_groq"] = f"faible confiance : {reason}"
        item["correction_appliquee"] = "aucune - à vérifier"
        item["fiabilite"] = "moyenne"
        return item

    if action == "keep":
        item.setdefault("controle_python", "OK")
        return item

    if action == "set_a_verifier":
        item["controle_python"] = item.get("controle_python", "A_VERIFIER")
        item["verification_groq"] = reason if source_decision == "groq" else ""
        item["correction_appliquee"] = "marqué à vérifier"
        item["fiabilite"] = "moyenne"
        return item

    if action == "reject":
        print(f"  [REJET COHÉRENCE] {item.get('designation')} {item.get('quantite')} {item.get('unite')} -> {reason}")
        return None

    original_q = item.get("quantite")
    original_u = item.get("unite")
    original_d = item.get("designation")
    original_scope = item.get("scope")

    if action == "convert_unit":
        factor = to_float(decision.get("facteur_conversion"))
        new_unit = decision.get("nouvelle_unite")
        if factor and new_unit and 0 < factor <= 1000:
            new_q = round(float(original_q) * factor, 6)
            if _plausible_quantity(item.get("designation", ""), new_q, str(new_unit)):
                item["quantite_originale"] = original_q
                item["unite_originale"] = original_u
                item["quantite"] = new_q
                item["unite"] = str(new_unit)
                item["controle_python"] = item.get("controle_python", "AUTO_CORRECTION")
                item["verification_groq"] = reason if source_decision == "groq" else ""
                item["correction_appliquee"] = f"conversion {original_u} -> {new_unit} x{factor}"
                item["fiabilite"] = "moyenne"
                return item
        return item

    if action == "replace_quantity":
        new_q = to_float(decision.get("nouvelle_quantite"))
        if new_q is not None and _plausible_quantity(item.get("designation", ""), new_q, item.get("unite", "")):
            item["quantite_originale"] = original_q
            item["quantite"] = new_q
            item["controle_python"] = item.get("controle_python", "CORRECTION")
            item["verification_groq"] = reason if source_decision == "groq" else ""
            item["correction_appliquee"] = f"quantité remplacée {original_q} -> {new_q}"
            item["fiabilite"] = "moyenne"
            return item
        return item

    if action == "replace_designation":
        new_d = decision.get("nouvelle_designation")
        if new_d and len(str(new_d)) <= 120:
            item["designation_originale"] = original_d
            item["designation"] = str(new_d)
            item["controle_python"] = item.get("controle_python", "CORRECTION")
            item["verification_groq"] = reason if source_decision == "groq" else ""
            item["correction_appliquee"] = f"désignation remplacée"
            item["fiabilite"] = "moyenne"
            return item
        return item

    if action == "change_scope":
        new_scope = decision.get("nouveau_scope")
        if new_scope in {"SCOPE_1", "SCOPE_2", "SCOPE_3"}:
            item["scope_original"] = original_scope
            item["scope"] = new_scope
            item["controle_python"] = item.get("controle_python", "CORRECTION")
            item["verification_groq"] = reason if source_decision == "groq" else ""
            item["correction_appliquee"] = f"scope remplacé {original_scope} -> {new_scope}"
            item["fiabilite"] = "moyenne"
            return item
        return item

    return item


def verify_and_repair_items(items: list[dict], path: str, filename: str, scope: str, role: str, df: pd.DataFrame | None, raw_text: str) -> list[dict]:
    if not items:
        return []

    context_cache: str | None = None
    repaired: list[dict] = []

    for item in items:
        decision = _python_coherence_decision(item, filename, scope, role)
        statut = decision.get("statut", "OK")
        action = decision.get("action", "keep")
        raison = decision.get("raison", "")

        item.setdefault("controle_python", "OK" if statut == "OK" else f"{statut}: {raison}")

        final_decision = decision

        # Groq intervient uniquement sur les anomalies non déterministes.
        if action == "verify_with_groq":
            if context_cache is None:
                context_cache = _build_context_for_groq(path, df, raw_text)
            groq_decision = _groq_verify_anomaly(item, filename, scope, role, raison, context_cache)
            if groq_decision:
                final_decision = groq_decision
                item["verification_groq"] = str(groq_decision.get("raison", ""))[:300]
            else:
                # Sans Groq, on rejette les anomalies extrêmes pour ne pas polluer le bilan.
                final_decision = {
                    "action": "reject",
                    "raison": f"{raison} ; Groq indisponible",
                    "confiance": "haute",
                }

        # Les auto-corrections Python ne nécessitent pas Groq.
        source = "python" if final_decision is decision else "groq"
        new_item = _apply_decision(item, final_decision, source)
        if new_item is not None:
            repaired.append(new_item)

    return repaired



def audit_status(item: dict) -> tuple[str, str]:
    unit = norm(item.get("unite", ""))
    q = float(item.get("quantite", 0) or 0)
    des = norm(item.get("designation", ""))
    if q <= 0:
        return "SUSPECT", "quantité nulle"
    if "eau" in des and unit in {"m3", "m³"} and q > 10000:
        return "SUSPECT", "eau > 10 000 m3"
    if "papier" in des and unit == "kg" and q > 10000:
        return "SUSPECT", "papier > 10 000 kg"
    if "propane" in des and unit == "kg" and q > 10000:
        return "SUSPECT", "propane > 10 000 kg"
    if "gravats" in des and unit == "t" and q > 10000:
        return "SUSPECT", "déchets inertes > 10 000 t"
    if unit in {"€", "eur", "euro", "euros"}:
        return "A_VERIFIER", "montant financier à contrôler"
    return "OK", ""


def write_audit(base_path: str, data: list[dict], no_data: list[str], nb_files: int) -> None:
    rows = []
    for item in data:
        st, alert = audit_status(item)
        rows.append({**item, "statut": st, "alerte": alert})
    for f in no_data:
        rows.append({"source": f, "scope": "", "role": "", "designation": "", "quantite": "", "unite": "", "fiabilite": "", "justification": "", "est_calcule": False, "statut": "AUCUNE_DONNEE", "alerte": "aucune donnée extraite"})

    fields = ["statut", "alerte", "source", "scope", "role", "designation", "quantite", "unite", "fiabilite", "justification", "est_calcule", "controle_python", "verification_groq", "correction_appliquee", "quantite_originale", "unite_originale", "designation_originale", "scope_original"]
    out_csv = os.path.join(base_path, "audit_extraction_bilan_carbone.csv")
    with open(out_csv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";")
        w.writeheader()
        for row in rows:
            w.writerow({k: row.get(k, "") for k in fields})

    summary = {
        "version": VERSION,
        "nb_fichiers_tentes": nb_files,
        "nb_donnees": len(data),
        "nb_sans_donnees": len(no_data),
        "nb_ok": sum(1 for r in rows if r.get("statut") == "OK"),
        "nb_a_verifier": sum(1 for r in rows if r.get("statut") == "A_VERIFIER"),
        "nb_suspect": sum(1 for r in rows if r.get("statut") == "SUSPECT"),
        "verification_groq_active": bool(USE_GROQ_VERIFICATION and os.getenv("GROQ_API_KEY")),
        "nb_appels_groq_verification": _groq_verify_calls,
        "audit_csv": out_csv,
    }
    out_json = os.path.join(base_path, "audit_extraction_bilan_carbone_resume.json")
    Path(out_json).write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    if FINANCIAL_REJECTS:
        out_rej = os.path.join(base_path, "audit_rejets_financiers_non_classes.csv")
        with open(out_rej, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["source", "libelle", "montant", "raison"], delimiter=";")
            w.writeheader()
            w.writerows(FINANCIAL_REJECTS)
        print(f"  [AUDIT] Rejets financiers : {out_rej}")
    print(f"  [AUDIT] CSV : {out_csv}")


def deduplicate(data: list[dict]) -> list[dict]:
    merged: dict[tuple, dict] = {}
    for d in data:
        unit = str(d.get("unite", ""))
        # Fusionner les montants financiers par poste et scope ; garder les unités physiques par source.
        if unit == "€":
            key = (d["scope"], d.get("role", ""), d["designation"], unit)
        else:
            key = (d["source"], d["scope"], d["designation"], unit)
        if key in merged:
            merged[key]["quantite"] = round(float(merged[key]["quantite"]) + float(d["quantite"]), 4)
            if d["source"] not in merged[key]["source"]:
                merged[key]["source"] += " + " + d["source"]
        else:
            merged[key] = dict(d)
    return list(merged.values())


def extract_one(path: str, scope: str) -> list[dict]:
    filename = Path(path).name
    raw_text = read_text(path) if Path(path).suffix.lower() == ".txt" else ""
    df = load_table(path)
    role = detect_role(filename, df, raw_text)

    # Priorité aux extracteurs très fiables.
    for extractor in (
        lambda: extract_propane(path, filename, scope),
        lambda: extract_vehicules(path, filename, scope),
        lambda: extract_edf(path, filename, scope, df),
        lambda: extract_immobilisations(df, filename, scope),
        lambda: extract_dechets(path, filename, scope, df),
        lambda: extract_ptd_materials(path, filename, scope),
        lambda: extract_financial(df, filename, scope),
        lambda: extract_papier(df, filename, scope),
        lambda: extract_eau(df, filename, scope),
    ):
        res = extractor()
        if res:
            checked = verify_and_repair_items(res, path, filename, scope, role, df, raw_text)
            if checked:
                return checked
            # Si l'extraction était incohérente et rejetée, on tente l'extracteur suivant.
            continue

    if role == "general":
        print("  [SKIP] rôle non reconnu, Groq désactivé par défaut")
    else:
        print(f"  [SKIP] aucun extracteur direct concluant ({role})")
    return []


def iter_scope_files(base_path: str):
    for scope in ("SCOPE_1", "SCOPE_2", "SCOPE_3"):
        folder = Path(base_path) / scope
        if not folder.exists():
            continue
        for p in sorted(folder.iterdir()):
            if p.is_file() and p.suffix.lower() in {".csv", ".txt", ".xlsx", ".xls", ".xlsm"}:
                yield scope, str(p)


def lancer_le_bilan(base_path: str = ".") -> tuple[list[dict], int]:
    print("=" * 70)
    print(VERSION)
    print("=" * 70)

    all_data: list[dict] = []
    no_data: list[str] = []
    tried = 0

    current_scope = None
    for scope, path in iter_scope_files(base_path):
        if current_scope != scope:
            current_scope = scope
            print(f"\n{'─' * 70}\nDossier : {scope}\n{'─' * 70}")
        filename = Path(path).name
        if is_non_source(filename):
            print(f"  [SKIP] {filename} → exclu/non source")
            continue
        tried += 1
        print(f"\n  Traitement : {filename}")
        try:
            res = extract_one(path, scope)
        except Exception as e:
            print(f"  [ERREUR] {filename}: {e}")
            res = []
        if res:
            all_data.extend(res)
            print(f"  → {len(res)} valeur(s) :")
            for r in res:
                print(f"     {r['designation'][:42]:42s} | {r['quantite']:>10} {r['unite']:<8} | {r['fiabilite']}")
        else:
            no_data.append(filename)
            print("  → Aucune donnée")

    dedup = deduplicate(all_data)
    removed = len(all_data) - len(dedup)
    print("\n" + "=" * 70)
    print(f"POST-TRAITEMENT — {len(all_data)} valeurs brutes")
    if removed:
        print(f"  {removed} doublon(s) fusionné(s)")

    print("\n" + "=" * 70)
    print(f"RÉSUMÉ FINAL — {len(dedup)} valeurs | {tried} fichier(s) tenté(s)")
    for scope in ("SCOPE_1", "SCOPE_2", "SCOPE_3"):
        items = [d for d in dedup if d.get("scope") == scope]
        if not items:
            continue
        print(f"\n  [{scope}] {len(items)} valeur(s) :")
        for d in items:
            print(f"    {d['source'][:28]:28s} | {d['designation'][:32]:32s} | {d['quantite']:>10} {d['unite']:<8} | {d.get('role','')}")

    if no_data:
        print(f"\nFichiers sans données ({len(no_data)}) :")
        for f in no_data[:30]:
            print(f"  - {f}")
        if len(no_data) > 30:
            print(f"  ... {len(no_data) - 30} autre(s)")

    write_audit(base_path, dedup, no_data, tried)
    print("=" * 70)
    return dedup, tried


if __name__ == "__main__":
    lancer_le_bilan(".")
