import os
from pathlib import Path
import re
import shutil
import unicodedata
import numpy as np
import pandas as pd
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

MODEL_NAME = "paraphrase-multilingual-MiniLM-L12-v2"
model = SentenceTransformer(MODEL_NAME)

SCOPES_THEMES = {
    "SCOPE_1": [
        "Liste des véhicules de société avec consommation carburant annuelle en litres gazole diesel GNR",
        "Consommation carburant flotte véhicules utilitaires camions engins litres",
        "Facture carburant gazole diesel fioul véhicules et engins entreprise",
        "Facture de gaz naturel chaudière chauffage bâtiment consommation kWh PCI m3",
        "Bouteille gaz propane butane Antargaz Primagaz usage professionnel cuisson chauffage",
        "Remplissage cuve fioul domestique litres chauffage",
        "Relevé consommation GNR gazole non routier engins de chantier travaux",
        "Fuite fluides frigorigènes climatisation R410a gaz réfrigérant",
    ],
    "SCOPE_2": [
        "Facture EDF électricité consommation kWh puissance souscrite index heures pleines heures creuses",
        "Statistiques consommation électrique annuelle bâtiment tertiaire kWh",
        "Relevé compteur électrique Linky consommation mensuelle kWh",
        "Tableau de bord consommation énergétique électricité bâtiment exercice annuel",
        "Consommation réseau de chaleur urbain vapeur MWh abonnement",
    ],
    "SCOPE_3": [
        "Kilométrage annuel conducteurs véhicules relevé odomètre tableau récapitulatif",
        "Déplacements professionnels billets train SNCF TGV avion notes de frais",
        "Indemnités kilométriques déplacements domicile travail salariés",
        "Tableau suivi déchets entreprise papier carton plastique métal D3E DEEE tri sélectif",
        "Bordereau suivi déchets dangereux DIS produits chimiques entreprise bureau",
        "Factures collecte déchets entreprise papier carton DEEE recyclage",
        "Achat fournitures de bureau papier cartouches encre matériel informatique",
        "Facture eau potable consommation m3 litres abonnement réseau",
        "Récapitulatif achats biens et services fournitures année",
        "Liste immobilisations état amortissement matériel équipement véhicules",
        "Bordereau de suivi déchets BSD collecte matériaux chantier gravats terres client BTP",
        "Tableau récapitulatif collectes déchets inertes béton sable calcaire client BTP",
        "Fiche de données de sécurité FDS matériaux chantier travaux publics client",
    ],
}

USELESS_KEYWORDS = [
    "export - beges", "export - iso", "export - ghg", "export - occf",
    "export boites", "export sous-postes", "export donnees utilisees",
    "export données utilisées", "futurs exports", "matrice de collecte",
    "ratios", "objectif", "reduction", "réduction",
    "q18", "rapport de verification", "rapport de vérification",
    "diagnostic", "diagnostics", "fds", "fiche de donnees de securite",
    "fiche de données de sécurité", "plan des bureaux",
    "rgpd", "violation", "soc_abs", "soc_pdf",
]

def normalize_text(text: str) -> str:
    if text is None:
        return ""
    text = unicodedata.normalize("NFD", str(text).lower())
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def classify_by_rules(filename: str, text: str) -> str | None:
    fn = normalize_text(filename)
    tx = normalize_text(text[:2500])
    blob = f"{fn} {tx}"

    # Fichiers générés par l'outil : jamais source d'activité.
    if any(k in fn for k in ["audit_extraction", "audit_rejets", "resume", "résumé"]):
        return "USELESS"

    # Certains noms de fichiers sont suffisants même si le texte est vide
    # (PDF scanné, OCR absent, extraction PyMuPDF vide).
    if any(k in fn for k in ["plan d evacuation", "plan d'evacuation", "evacuation", "évacuation"]):
        return "USELESS"

    if any(k in fn for k in ["fournisseur", "fournisseurs", "achats fournisseurs", "bilan fournisseurs", "compte resultat", "compte de resultat", "grand livre", "balance generale", "journal achats"]):
        return "SCOPE_3"

    # Règle généralisante : les sous-dossiers de factures ne doivent pas tomber en USELESS
    # uniquement parce que l'OCR est imparfait ou que le score embedding est faible.
    if (
        "/factures" in fn
        or "\\factures" in fn
        or "factures " in fn
        or re.search(r"(^|/)cleaned_f[ _-]", fn) is not None
    ):
        if any(k in blob for k in ["edf", "electricite", "électricité", "linky", "kwh"]):
            return "SCOPE_2"
        if any(k in blob for k in ["gazole", "diesel", "gnr", "fioul", "carburant", "propane", "butane", "gaz naturel"]):
            return "SCOPE_1"
        return "SCOPE_3"

    if "facture" in fn and any(k in fn for k in ["climatisation", "clim", "pompe a chaleur", "pompe à chaleur"]):
        return "SCOPE_3"
    if any(k in fn for k in ["tableau conso edf", "conso edf", "electricite", "électricité"]):
        return "SCOPE_2"
    if any(k in fn for k in ["tableau conso eau", "conso eau"]):
        return "SCOPE_3"

    if any(normalize_text(k) in blob for k in USELESS_KEYWORDS):
        return "USELESS"

    # Scope 2 d'abord : électricité achetée
    if any(k in blob for k in [
        "edf", "electricite", "linky", "consommation facturee",
        "kwh", "heures pleines", "heures creuses"
    ]):
        if any(k in blob for k in ["rapport", "verification", "q18", "installation electrique"]):
            return "USELESS"
        return "SCOPE_2"

    # Scope 1 : carburant / gaz / combustion directe
    if any(k in blob for k in [
        "consommation de carburant", "gazole", "diesel", "gnr", "fioul",
        "carburant (l)", "litres carburant", "liste des vehicules",
        "bouteille de gaz", "antargaz", "primagaz", "propane", "butane", "gaz naturel"
    ]):
        return "SCOPE_1"

    # Scope 3 : déchets, achats, papier, eau, déplacements
    if any(k in blob for k in [
        "dechet", "dechets", "gravats", "dib", "bois b", "bois a",
        "papier", "a4", "a3", "traceur", "rouleau",
        "eau", "facture eau", "consommation eau",
        "km annuel", "kilometrage", "domicile travail",
        "achat", "achats", "fournisseur", "fournisseurs", "facture fournisseur",
        "fournitures", "materiel informatique", "immobilisation",
        "climatisation", "pompe a chaleur", "pompe à chaleur", "maintenance clim",
        "mondocunique", "bureau", "bureaux", "atelier", "depot", "dépôt", "chantier", "chantiers"
    ]):
        return "SCOPE_3"

    return None

ref_vectors: list[np.ndarray] = []
ref_scopes: list[str] = []
for scope, phrases in SCOPES_THEMES.items():
    for phrase in phrases:
        ref_vectors.append(model.encode(normalize_text(phrase)))
        ref_scopes.append(scope)
REF_MATRIX = np.array(ref_vectors)
REF_SCOPES = np.array(ref_scopes)

def extract_text(path: str, max_chars: int = 1500) -> str:
    ext = path.lower().rsplit(".", 1)[-1]
    try:
        if ext == "txt":
            for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
                try:
                    with open(path, encoding=enc) as f:
                        return f.read(max_chars)
                except UnicodeDecodeError:
                    continue
        elif ext == "csv":
            df = pd.read_csv(path, sep=None, engine="python", encoding="latin1", on_bad_lines="skip", nrows=50)
            tokens = df.fillna("").astype(str).values.flatten()
            return " ".join(t for t in tokens if t.strip())[:max_chars]
        elif ext in ("xls", "xlsx", "xlsm"):
            df = pd.read_excel(path, nrows=50)
            tokens = df.fillna("").astype(str).values.flatten()
            return " ".join(t for t in tokens if t.strip())[:max_chars]
        elif ext == "docx":
            from docx import Document
            doc = Document(path)
            return "\n".join(p.text for p in doc.paragraphs)[:max_chars]
        elif ext == "pdf":
            import pymupdf
            doc = pymupdf.open(path)
            return "".join(page.get_text("text") for page in doc)[:max_chars]
    except Exception as e:
        print(f"  [ERREUR EXTRACT] {path}: {e}")
    return ""

def run_test(folder: str = "cleaned_files", threshold: float = 0.40) -> dict[str, str]:
    results: dict[str, str] = {}
    if not os.path.exists(folder):
        print(f"[ERREUR] Dossier introuvable : {folder}")
        return results

    base_dir = Path(folder)
    files = sorted(
        str(p.relative_to(base_dir)).replace("\\", "/")
        for p in base_dir.rglob("*")
        if p.is_file()
    )
    if not files:
        return results

    print(f"\nClassification de {len(files)} fichier(s) (seuil={threshold})...\n")
    texts: list[tuple[str, str]] = []

    for fname in files:
        text = extract_text(os.path.join(folder, fname))
        rule_scope = classify_by_rules(fname, text)
        if rule_scope:
            results[fname] = rule_scope
            icon = "✓" if rule_scope != "USELESS" else "✗"
            print(f"  {icon} {fname:<55} règle métier → {rule_scope}")
            continue
        if text and len(text.strip()) >= 20:
            texts.append((fname, text))
        else:
            results[fname] = "USELESS"
            print(f"  ✗ {fname:<55} vide/illisible → USELESS")

    if texts:
        norms = [normalize_text(t) for _, t in texts]
        vecs = model.encode(norms, batch_size=32, convert_to_numpy=True)
        sims = cosine_similarity(vecs, REF_MATRIX)
        for i, (fname, _) in enumerate(texts):
            best_idx = int(np.argmax(sims[i]))
            best_score = float(sims[i][best_idx])
            best_scope = REF_SCOPES[best_idx]
            if best_score < threshold:
                best_scope = "USELESS"
            icon = "✓" if best_scope != "USELESS" else "✗"
            print(f"  {icon} {fname:<55} score={best_score:.3f} → {best_scope}")
            results[fname] = best_scope

    print()
    for scope in ("SCOPE_1", "SCOPE_2", "SCOPE_3", "USELESS"):
        n = sum(1 for v in results.values() if v == scope)
        if n:
            print(f"  {scope:<10} : {n} fichier(s)")
    return results

def clear_scope_dirs(dest_base: str = ".") -> None:
    for scope in ("SCOPE_1", "SCOPE_2", "SCOPE_3"):
        dest_dir = os.path.join(dest_base, scope)
        os.makedirs(dest_dir, exist_ok=True)
        for name in os.listdir(dest_dir):
            path = os.path.join(dest_dir, name)
            try:
                if os.path.isfile(path) or os.path.islink(path):
                    os.remove(path)
                elif os.path.isdir(path):
                    shutil.rmtree(path)
            except Exception as e:
                print(f"  [WARN CLEAN] Impossible de supprimer {path}: {e}")

def _safe_scope_filename(relative_name: str) -> str:
    """
    Aplatit un chemin relatif pour éviter de copier un dossier entier dans SCOPE_X.

    Exemple :
        Factures EPI 2023-2024/cleaned_F CMO 202301062.txt
    devient :
        Factures EPI 2023-2024__cleaned_F CMO 202301062.txt

    Ça garde l'information du sous-dossier sans recréer le sous-dossier.
    """
    p = Path(relative_name)
    if len(p.parts) <= 1:
        return p.name

    parent = "__".join(p.parts[:-1])
    candidate = f"{parent}__{p.name}"
    candidate = re.sub(r'[<>:"/\\|?*]+', "_", candidate)
    candidate = re.sub(r"\s+", " ", candidate).strip()
    return candidate


def _unique_destination(dest_dir: str, filename: str) -> str:
    """
    Évite d'écraser un fichier si deux sources donnent le même nom final.
    """
    path = Path(dest_dir) / filename
    if not path.exists():
        return str(path)

    stem = path.stem
    suffix = path.suffix
    i = 2
    while True:
        candidate = Path(dest_dir) / f"{stem}__{i}{suffix}"
        if not candidate.exists():
            return str(candidate)
        i += 1


def export_per_scope(
    results: dict[str, str],
    source_folder: str = "cleaned_files",
    dest_base: str = ".",
    clean_before_export: bool = True,
    preserve_subfolders: bool = False,
) -> None:
    """
    Exporte les fichiers classés vers SCOPE_1 / SCOPE_2 / SCOPE_3.

    Par défaut, les fichiers sont aplatis :
        cleaned_files/Factures/fichier.txt
        -> SCOPE_3/Factures__fichier.txt

    Pourquoi ?
    - éviter que le mapping donne l'impression de classer un dossier complet ;
    - avoir directement tous les fichiers visibles dans le scope ;
    - éviter les problèmes dans les traitements qui ne parcourent pas récursivement.

    Si besoin, on peut conserver l'ancienne structure :
        export_per_scope(..., preserve_subfolders=True)
    """
    if clean_before_export:
        print("\nNettoyage des anciens dossiers SCOPE...")
        clear_scope_dirs(dest_base)

    for fname, scope in results.items():
        if scope not in ("SCOPE_1", "SCOPE_2", "SCOPE_3"):
            continue

        dest_dir = os.path.join(dest_base, scope)
        os.makedirs(dest_dir, exist_ok=True)

        src = os.path.join(source_folder, fname)

        if preserve_subfolders:
            dst = os.path.join(dest_dir, fname)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
        else:
            flat_name = _safe_scope_filename(fname)
            dst = _unique_destination(dest_dir, flat_name)

        try:
            shutil.copy2(src, dst)
            print(f"  Copié : {fname} → {scope}/{Path(dst).name}")
        except Exception as e:
            print(f"  [ERREUR COPIE] {fname}: {e}")

if __name__ == "__main__":
    results = run_test(threshold=0.40)
    export_per_scope(results, clean_before_export=True)
    print("\nRÉSULTATS FINAUX")
    for scope in ("SCOPE_1", "SCOPE_2", "SCOPE_3"):
        fichiers = [k for k, v in results.items() if v == scope]
        if fichiers:
            print(f"\n  {scope} ({len(fichiers)}) :")
            for f in fichiers:
                print(f"    • {f}")
