import os
import numpy as np
from sentence_transformers import SentenceTransformer
import pandas as pd
import pymupdf
from PIL import Image
from docx import Document
import load_data
import process_data

model = SentenceTransformer("paraphrase-multilingual-MiniLM-L12-v2") # modèle léger et rapide pour les tests, à remplacer par un modèle plus puissant si besoin

# Les textes de référence pour chaque scope
SCOPES = {
    "SCOPE_1": (
        "Facture de carburant diesel essence pour véhicule de société. "
        "Consommation de gaz naturel fioul propane butane charbon biomasse. "
        "Émissions directes combustion station service plein litres. "
        "Flotte véhicules entreprise carburant fossile."
    ),
    "SCOPE_2": (
        "Facture d'électricité EDF fournisseur d'énergie consommation kWh MWh. "
        "Abonnement électrique compteur relevé de compteur chauffage urbain. "
        "Énergie achetée refroidissement vapeur consommation électrique bâtiment."
    ),
    "SCOPE_3": (
        "Transport de marchandises logistique sous-traitant prestataire livraison. "
        "Gestion des déchets collecte tri recyclage valorisation tonnes kilogrammes. "
        "Immobilisations achat matières premières fournisseur bilan carbone indirect. "
        "Déplacements domicile travail déchets suivi tonnes transporteurs sites."
    ),
}

def cosine_similarity(a, b): # calcul de la similarité cosinus entre deux vecteurs
    return np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b))

def chunk_text(text, size=150): # découpage du texte en morceaux de 150 mots pour une meilleure classification
    words = text.split()
    return [" ".join(words[i:i+size]) for i in range(0, len(words), size)]

# Précalcul des vecteurs de référence
scope_vectors = {scope: model.encode(text) for scope, text in SCOPES.items()}

def classify_scope(text): # classification du texte dans un scope en fonction de la similarité avec les textes de référence
    if not text or len(text.strip()) < 30:  
        return "USELESS"

    chunks = chunk_text(text)
    best_scope = "USELESS"
    best_score = -1

    for chunk in chunks: # calcul du vecteur du chunk et comparaison avec les vecteurs de référence pour trouver le scope le plus similaire
        vec = model.encode(chunk)
        for scope, scope_vec in scope_vectors.items():
            score = cosine_similarity(vec, scope_vec)
            if score > best_score:
                best_score = score
                best_scope = scope

    print(f" score={best_score:.3f} → {best_scope}")

    if best_score < 0.20: 
        return "USELESS"

    return best_scope


def extract_text(path): # extraction du texte d'un fichier en fonction de son extension, avec gestion des erreurs et des encodages pour les fichiers texte et CSV
    ext = path.lower().split(".")[-1]

    try:
        if ext == "txt":
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()

        elif ext == "csv":
            df = pd.read_csv(
                path, sep=None, engine="python",
                encoding="latin1", on_bad_lines='skip', nrows=100
            )
            content = df.fillna('').astype(str).values.flatten()
            return " ".join([item for item in content if item.strip() not in ('', 'nan')])

        elif ext in ["xls", "xlsx", "xlsm"]:
            df = pd.read_excel(path)
            content = df.fillna('').astype(str).values.flatten()
            return " ".join([item for item in content if item.strip() not in ('', 'nan')])

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


def run_test(folder="cleaned_files"): # test de classification sur les fichiers nettoyés, avec extraction du texte et classification dans les scopes
    results = {}

    for file in os.listdir(folder): # parcours des fichiers du dossier, extraction du texte et classification
        path = os.path.join(folder, file)
        if not os.path.isfile(path):
            continue

        print("\n---", file)
        text = extract_text(path)
        scope = classify_scope(text)
        print("→", scope)
        results[file] = scope

    return results

import shutil # Pour copier ou déplacer les fichiers

def export_per_scope(results, source_folder="cleaned_files"):
    for file_name, scope in results.items():
        # Si le fichier est utile (SCOPE_1, 2 ou 3)
        if scope in ["SCOPE_1", "SCOPE_2", "SCOPE_3"]:
            
            # Créer le dossier du scope s'il n'existe pas
            if not os.path.exists(scope):
                os.makedirs(scope)
            
            # chemins source et destination
            src_path = os.path.join(source_folder, file_name)
            dst_path = os.path.join(scope, file_name)
            
            # Copier le fichier dans le bon dossier
            try:
                shutil.copy(src_path, dst_path)
                print(f"Copié : {file_name} -> {scope}/")
            except Exception as e:
                print(f"Erreur lors de la copie de {file_name}: {e}")



if __name__ == "__main__": # exécution du test de classification et affichage des résultats
    results = run_test()
    sorted = export_per_scope(results)
    print("\nRÉSULTATS")
    for k, v in results.items(): # affichage du nom du fichier et du scope classifié pour chaque fichier testé
        print(k, ":", v)
