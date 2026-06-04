import os
import unicodedata
import numpy as np
import pandas as pd
import pymupdf
from PIL import Image
from docx import Document
import shutil
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity as sklearn_cosine_similarity

model = SentenceTransformer("paraphrase-multilingual-MiniLM-L12-v2")

SCOPES_THEMES = {
    "SCOPE_1": [
        "Facture de gaz naturel",
        "Bouteille de gaz propane butane pour usage domestique ou professionnel",
        "Achat de carburant gazole fioul diesel pour flotte de véhicules de l'entreprise",
        "Consommation essence sans plomb pour véhicules de service",
        "Recharge de gaz propane butane en citerne ou bouteille",
        "Remplissage cuve fioul domestique chauffage",
        "Relevé de consommation fioul lourd ou GNR engins de chantier",
    ],
    "SCOPE_2": [
        "Facture d'électricité EDF, abonnement et consommation en kWh",
        "Relevé de compteur électrique Linky",
        "Consommation électrique des bâtiments, électricité réseau",
        "Consommation réseau de chaleur urbain ou réseau de froid"
    ],
    "SCOPE_3": [
        "Transport de marchandises, logistique, fret routier maritime aérien",
        "Facture de livraison prestataire externe camion poids lourd",
        "Gestion des déchets, ordures ménagères, DEEE, recyclage papier carton",
        "Déplacements professionnels, billets de train SNCF TGV, avion",
        "Notes de frais, indemnités kilométriques IK, déplacements collaborateurs",
        "Trajets domicile travail des salariés, abonnement transports en commun",
        "Achat de fournitures de bureau, papier, cartouches d'encre, matériel informatique",
        "Facture d'eau potable, consommation en m3 ou litres",
        "Achat de matières premières, intrants, composants industriels",
        "Prestations de services, honoraires, frais bancaires, assurances"
        "Déchets , papier, carton, matières, DEEE, recyclage, béton, gravats, ordures ménagères",
    ]
}

# 2. PRÉCALCUL DES VECTEURS INDIVIDUELS


def normalize_text(text):
    if text is None:
        return ""
    text = unicodedata.normalize("NFD", str(text).lower())
    return "".join(c for c in text if unicodedata.category(c) != "Mn")


reference_vectors = []
reference_scopes = []
for scope, phrases in SCOPES_THEMES.items():
    for phrase in phrases:
        reference_vectors.append(model.encode(normalize_text(phrase)))
        reference_scopes.append(scope)

# Convertir en matrice numpy pour opérations vectorisées
reference_vectors = np.array(reference_vectors)
reference_scopes = np.array(reference_scopes)

def extract_text(path): 
    ext = path.lower().split(".")[-1]
    try:
        if ext == "txt":
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read(300)

        elif ext == "csv":
            df = pd.read_csv(path, sep=None, engine="python", encoding="latin1", on_bad_lines='skip', nrows=50)
            content = df.fillna('').astype(str).values.flatten()
            return " ".join([item for item in content if item.strip()])

        elif ext in ["xls", "xlsx", "xlsm"]:
            df = pd.read_excel(path, nrows=50)
            content = df.fillna('').astype(str).values.flatten()
            return " ".join([item for item in content if item.strip()])

        elif ext == "docx":
            doc = Document(path)
            return "\n".join([p.text for p in doc.paragraphs])

        elif ext == "pdf":
            doc = pymupdf.open(path)
            text = ""
            for page in doc:
                text += page.get_text("text")
            return text
        else:
            return ""
    except Exception as e:
        print("Erreur extraction:", path, e)
        return ""

def run_test(folder="cleaned_files", threshold=0.30): 
    results = {}
    rejected_files = []  # Tracker les fichiers rejetés
    
    if not os.path.exists(folder):
        print(f"Dossier {folder} introuvable.")
        return results

    # 1. Extraction + normalisation de tous les textes
    files_to_classify = []
    for file in sorted(os.listdir(folder)):  # tri pour reproductibilité
        path = os.path.join(folder, file)
        if not os.path.isfile(path):
            continue

        text = extract_text(path)
        if text and len(text.strip()) >= 20:
            files_to_classify.append((file, text))

    if not files_to_classify:
        return results

    print(f"\n Classification {len(files_to_classify)} fichier(s)...")
    print(f"   Seuil de confiance : {threshold}\n")
    
    # 2. Batch encoding : traiter tous les textes en une fois
    texts_normalized = [normalize_text(text) for _, text in files_to_classify]
    vectors = model.encode(texts_normalized, batch_size=32, convert_to_numpy=True)

    # 3. Classification vectorisée
    similarities = sklearn_cosine_similarity(vectors, reference_vectors)
    
    for i, (file, _) in enumerate(files_to_classify):
        # Meilleur score et scope pour ce fichier
        best_idx = np.argmax(similarities[i])
        best_score = similarities[i][best_idx]
        best_scope = reference_scopes[best_idx]
        
        if best_score < threshold:
            best_scope = "USELESS"
            rejected_files.append((file, best_score, reference_scopes[best_idx]))
        
        status = "" if best_score >= threshold else ""
        print(f"  {status} {file:<50} score={best_score:.3f} -> {best_scope}")
        results[file] = best_scope

    # Afficher un résumé des fichiers rejetés
    if rejected_files:
        print(f"\n  {len(rejected_files)} fichier(s) rejeté(s) (score < {threshold}) :")
        for file, score, scope in sorted(rejected_files, key=lambda x: x[1], reverse=True):
            print(f"     • {file:<50} score={score:.3f} (proche de {scope})")

    return results

def export_per_scope(results, source_folder="cleaned_files", dest_base="."):
    for file_name, scope in results.items():
        if scope in ["SCOPE_1", "SCOPE_2", "SCOPE_3"]:
            scope_dir = os.path.join(dest_base, scope)  
            if not os.path.exists(scope_dir):
                os.makedirs(scope_dir)
            src_path = os.path.join(source_folder, file_name)
            dst_path = os.path.join(scope_dir, file_name)
            try:
                shutil.copy(src_path, dst_path)
                print(f"Copié : {file_name} -> {scope_dir}/")
            except Exception as e:
                print(f"Erreur lors de la copie de {file_name}: {e}")

if __name__ == "__main__":
    results = run_test(threshold=0.30)
    export_per_scope(results)
    print("\ RÉSULTATS FINAUX")
    for k, v in results.items():
        print(k, ":", v)