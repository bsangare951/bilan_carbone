import getters_data as gd
import os
import pandas as pd
import re
from datetime import datetime

# --- CONSTANTES ---
UNITES_VALIDES = {
    "L": ["l", "litres", "litre"],
    "m³": ["m3", "m³"],
    "kg": ["kg", "kilogrammes"],
    "t": ["t", "tonnes", "tonne"],
    "kWh": ["kwh", "kilowatt-heure", "kilowattheure"],
    "MWh": ["mwh", "megawatt-heure"],
    "t.km": ["t.km", "tonnes.km", "tonnes-km", "tonne.km", "tkm"],
    "km": ["km", "kilometres", "kilometre"],
    "kgCO2e": ["kgco2e", "kg co2e", "kgco2"]
}

FICHIERS_EXCLUS = ["synthese", "recap", "bilan_carbone", "ratios", "utilitaires",
                   "export", "matrice", "descriptif", "soc_", "rgpd", "liste",
                   ".~lock", "~$"]

MOTS_CLES_HEADER = [
    "quantite", "qte", "volume", "consommation", "total", "valeur",
    "designation", "libelle", "nom", "article", "nature", "description",
    "unite", "unites", "papier", "d3e", "cartouche", "capsule", "canette",
    "plastique", "exercice", "annee", "mois", "distance", "tonnes", "kwh",
    "energie", "periodicite", "type", "flux"
]


# --- UTILITAIRES ---

def get_annee_cible():
    annee = datetime.now().year
    mois = datetime.now().month
    if mois >= 10:
        return f"{annee % 100:02d}-{(annee + 1) % 100:02d}"
    return f"{(annee - 1) % 100:02d}-{annee % 100:02d}"

ANNEE_CIBLE = get_annee_cible()


def extract_numeric(value):
    if pd.isna(value):
        return None
    try:
        s = str(value).replace("\xa0", "").replace(" ", "")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        val = float(s)
        return val if not pd.isna(val) else None
    except (ValueError, TypeError):
        return None


def detecter_unite(text):
    if not text:
        return None
    text_norm = gd.normalize_text(str(text))
    for unite, mots in UNITES_VALIDES.items():
        for mot in mots:
            if mot == text_norm or f"({mot})" in text_norm or text_norm.endswith(f" {mot}"):
                return unite
    return None


def smart_header(df_raw):
    best_idx, best_score = 0, 0
    for i, row in df_raw.iterrows():
        if i > 10:
            break
        vals = [gd.normalize_text(str(v)) for v in row.values]
        score = sum(1 for v in vals if any(k in v for k in MOTS_CLES_HEADER))
        score += sum(1 for v in vals if re.search(r'\d{2}-\d{2}', v))
        if score > best_score:
            best_score, best_idx = score, i
    return best_idx


def load_file(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == ".csv":
            df_raw = pd.read_csv(filepath, sep=";", encoding="utf-8-sig",
                                 on_bad_lines='skip', header=None)
            h = smart_header(df_raw)
            return pd.read_csv(filepath, sep=";", encoding="utf-8-sig",
                               on_bad_lines='skip', header=h)
        elif ext in [".xlsx", ".xls"]:
            df_raw = pd.read_excel(filepath, header=None)
            h = smart_header(df_raw)
            return pd.read_excel(filepath, header=h)
    except Exception as e:
        print(f"  [ERREUR LECTURE] {e}")
    return None


# --- DÉTECTION DU TYPE ---

def detecter_type(filename, df):
    cols = [gd.normalize_text(str(c)) for c in df.columns]
    cols_str = " ".join(cols)

    # ⚠️ DECHETS en premier — "Exercice 24-25" ne doit pas être pris pour une année temporelle
    mots_dechets = ["papier", "d3e", "cartouche", "capsule", "canette", "plastique"]
    if sum(1 for m in mots_dechets if m in cols_str) >= 2:
        return "dechets_pivote"

    # Fichier avec kgCO2e déjà calculé
    if any("kgco2e" in c and "fe" not in c for c in cols):
        return "already_calculated"

    # Suivi temporel (années en colonnes)
    if any(re.search(r'\d{2}-\d{2}', c) for c in cols):
        return "suivi_temporel"

    # Facture standard / outillage
    if any(any(m in gd.normalize_text(c) for m in
               ["quantite", "qte", "volume", "energie consommee"]) for c in df.columns):
        return "facture_standard"

    # Suivi temporel avec années dans les lignes
    for _, row in df.head(5).iterrows():
        if any(re.search(r'\d{2}-\d{2}', str(v)) for v in row.values):
            return "suivi_temporel_lignes"

    return "generique"


# --- EXTRACTEURS ---

def extraire_already_calculated(df, filename, scope):
    col_co2 = next((c for c in df.columns
                    if "kgco2e" in gd.normalize_text(c)
                    and "fe" not in gd.normalize_text(c)), None)
    if not col_co2:
        return []
    total = pd.to_numeric(df[col_co2], errors='coerce').sum()
    if total <= 0:
        return []
    return [{
        "source": filename, "scope": scope,
        "designation": "Transport routier de marchandises",
        "quantite": round(total, 4), "unite": "kgCO2e",
        "est_calcule": True, "confiance": "haute"
    }]


def get_colonne_recente(df, cols_annees):
    for col in reversed(cols_annees):
        vals = df[col].apply(extract_numeric).dropna()
        if not vals.empty and vals.max() > 0:
            return col
    return cols_annees[-1]


def extraire_suivi_temporel(df, filename, scope):
    cols_annees = [c for c in df.columns if re.search(r'\d{2}-\d{2}', str(c))]
    col_cible = get_colonne_recente(df, cols_annees)
    vals = df[col_cible].apply(extract_numeric).dropna()
    total = vals.max() if not vals.empty else 0
    if total <= 0:
        return []
    unite = detecter_unite(filename) or "kWh"
    designation = next(
        (v for k, v in gd.CORRESPONDANCES.items()
         if k in gd.normalize_text(filename)), "mix electricité")
    return [{
        "source": filename, "scope": scope, "designation": designation,
        "quantite": round(total, 2), "unite": unite,
        "est_calcule": False, "confiance": "moyenne"
    }]


def extraire_suivi_temporel_lignes(df, filename, scope):
    header_idx = None
    for i, row in df.iterrows():
        if any(re.search(r'\d{2}-\d{2}', str(v)) for v in row.values):
            header_idx = i
            break
    if header_idx is None:
        return []
    df_new = df.iloc[header_idx + 1:].copy()
    df_new.columns = [str(h) for h in df.iloc[header_idx].tolist()]
    df_new = df_new.reset_index(drop=True)
    cols_annees = [c for c in df_new.columns if re.search(r'\d{2}-\d{2}', str(c))]
    if not cols_annees:
        return []
    col_cible = get_colonne_recente(df_new, cols_annees)
    vals = df_new[col_cible].apply(extract_numeric).dropna()
    total = vals.max() if not vals.empty else 0
    if total <= 0:
        return []
    unite = detecter_unite(filename) or "kWh"
    designation = next(
        (v for k, v in gd.CORRESPONDANCES.items()
         if k in gd.normalize_text(filename)), "mix electricité")
    return [{
        "source": filename, "scope": scope, "designation": designation,
        "quantite": round(total, 2), "unite": unite,
        "est_calcule": False, "confiance": "moyenne"
    }]


def extraire_dechets_pivote(df, filename, scope):
    mots_dechets = ["papier", "d3e", "cartouche", "capsule", "canette",
                    "plastique", "verre", "metal", "aluminium", "carton"]
    cols_cat = [c for c in df.columns
                if any(m in gd.normalize_text(c) for m in mots_dechets)]
    results = []
    for col in cols_cat:
        vals = df[col].apply(extract_numeric).dropna()
        total = vals[vals > 0].sum()
        if total > 0:
            designation = gd.CORRESPONDANCES.get(gd.normalize_text(col), col)
            results.append({
                "source": filename, "scope": scope, "designation": designation,
                "quantite": round(total, 2), "unite": "kg",
                "est_calcule": False, "confiance": "moyenne"
            })
    return results


def extraire_facture_standard(df, filename, scope):
    col_desig = next((c for c in df.columns if any(
        m in gd.normalize_text(c) for m in
        ["designation", "libelle", "nom", "article", "nature", "description"])), None)
    col_qte = next((c for c in df.columns if any(
        m in gd.normalize_text(c) for m in
        ["quantite", "qte", "volume", "total", "energie consommee"])), None)
    col_unite = next((c for c in df.columns if any(
        m in gd.normalize_text(c) for m in
        ["unite", "unites", "u", "energie consommee"])), None)
    if not col_qte:
        return []
    results = []
    for _, row in df.iterrows():
        qte = extract_numeric(row[col_qte])
        if not qte or qte <= 0:
            continue
        desig_brute = str(row[col_desig]).strip() if col_desig else filename
        designation = gd.CORRESPONDANCES.get(gd.normalize_text(desig_brute), desig_brute)
        unite_brute = str(row[col_unite]).strip() if col_unite else ""
        unite = detecter_unite(unite_brute) or detecter_unite(col_qte) or "kg"
        results.append({
            "source": filename, "scope": scope, "designation": designation,
            "quantite": round(qte, 4), "unite": unite,
            "est_calcule": False,
            "confiance": "haute" if col_desig and col_unite else "moyenne"
        })
    return results


def extraire_generique(df, filename, scope):
    cols_num = df.select_dtypes(include=['number']).columns.tolist()
    col_desig = df.columns[0]
    results = []
    for col in cols_num:
        unite = detecter_unite(col)
        if not unite:
            continue
        for _, row in df.iterrows():
            qte = extract_numeric(row[col])
            if not qte or qte <= 0:
                continue
            desig = str(row[col_desig]) if not pd.isna(row[col_desig]) else col
            designation = gd.CORRESPONDANCES.get(gd.normalize_text(desig), desig)
            results.append({
                "source": filename, "scope": scope, "designation": designation,
                "quantite": round(qte, 4), "unite": unite,
                "est_calcule": False, "confiance": "faible"
            })
    return results


# --- ORCHESTRATEUR ---

def extract_from_file(filepath, scope):
    filename = os.path.basename(filepath)
    df = load_file(filepath)
    if df is None or df.empty:
        return []

    type_fichier = detecter_type(filename, df)
    print(f"  → Type : {type_fichier}")

    extracteurs = {
        "already_calculated":    extraire_already_calculated,
        "suivi_temporel":        extraire_suivi_temporel,
        "suivi_temporel_lignes": extraire_suivi_temporel_lignes,
        "dechets_pivote":        extraire_dechets_pivote,
        "facture_standard":      extraire_facture_standard,
        "generique":             extraire_generique,
    }
    return extracteurs.get(type_fichier, extraire_generique)(df, filename, scope)


def lancer_le_bilan(base_path="."):
    final_data, non_traites = [], []
    print(f"Exercice cible : {ANNEE_CIBLE}")
    print("=" * 60)

    for root, dirs, files in os.walk(base_path):
        folder = os.path.basename(root).upper()
        if "SCOPE" not in folder:
            continue
        scope = folder
        for f in sorted(files):
            # Ignorer fichiers temporaires et exclus
            if any(exclu in f.lower() for exclu in FICHIERS_EXCLUS):
                continue
            if f.startswith("~$") or f.startswith(".~"):
                continue
            if f.lower().endswith((".py", ".pkl", ".docx", ".txt", ".jpg")):
                continue

            filepath = os.path.join(root, f)
            print(f"\n[{scope}] {f}")
            results = extract_from_file(filepath, scope)
            if results:
                print(f"  → {len(results)} valeur(s) extraite(s)")
                final_data.extend(results)
            else:
                print(f"  → Aucune valeur extraite")
                non_traites.append(f)

    # Déduplication par données (pas par nom de fichier)
    seen = set()
    final_data_dedup = []
    for d in final_data:
        key = (d["designation"], d["quantite"], d["unite"])
        if key not in seen:
            seen.add(key)
            final_data_dedup.append(d)

    print("\n" + "=" * 60)
    print(f"RÉSULTAT : {len(final_data_dedup)} valeurs extraites")
    if non_traites:
        print(f"FICHIERS NON TRAITÉS ({len(non_traites)}) :")
        for f in non_traites:
            print(f"  - {f}")
    print("=" * 60)
    return final_data_dedup


if __name__ == "__main__":
    data = lancer_le_bilan()
    print(f"\n{'SOURCE':<35} | {'DÉSIGNATION':<30} | {'QUANTITÉ':>12} | UNITÉ")
    print("-" * 90)
    for d in data:
        print(f"{d['source'][:35]:<35} | {d['designation'][:30]:<30} | "
              f"{d['quantite']:>12.2f} | {d['unite']}")