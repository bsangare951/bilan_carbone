import os
from dotenv import load_dotenv
import load_data as ld
import process_data as pt
import numpy as np
from sentence_transformers import SentenceTransformer
import pandas as pd
import pytesseract
import pymupdf
from PIL import Image
from docx import Document
 
 
def get_units_referential():
    return {
        "ENERGIE": ["kWh", "MWh", "m3", "L"],
        "TRANSPORT": ["km", "L", "passager.km", "t.km"],
        "DECHETS": ["kg", "t"],
        "IMMOBILISATION": ["m2", "unite", "kg"],
        "AUTRES": ["€", "k€"]
    }
 
def get_vehicle_attributes():
    return ["type", "marque", "modèle", "année", "carburant", "puissance_fiscale"]
 
def get_carbon_schema():
    return {
        "Nom-Fichier": None,
        "Date": None,
        "ID_Facture": None,
        "Designation": None,
        "Prix": 0.0,
        "Quantite": 0.0,
        "Unite": None,
        "Details_Vehicule": {
            "Marque_Modele": None,
            "Carburant": None,
            "Type": None
        },
        "Emissions_tCO2e": 0.0
    }
 
def get_scope():
    return {
        "SCOPE_1": {
            "nom": "émissions directes",
            "mot-clés": ["gaz naturel", "fioul", "diesel", "essence", "propane", "butane", "charbon", "biomasse"],
            "description": "Émissions directes provenant de sources possédées ou contrôlées par l'entreprise, telles que la combustion de carburant dans les véhicules de l'entreprise ou les émissions fugitives."
        },
        "SCOPE_2": {
            "nom": "émissions indirectes liées à l'énergie",
            "mot-clés": ["électricité", "chauffage", "refroidissement", "vapeur"],
            "description": "Émissions indirectes provenant de la consommation d'électricité, de chauffage ou de refroidissement achetés par l'entreprise."
        },
        "SCOPE_3": {
            "nom": "autres émissions indirectes",
            "mot-clés": ["transport de marchandises", "déchets", "immobilisations", "autres"],
            "description": "Toutes les autres émissions indirectes qui se produisent dans la chaîne de valeur de l'entreprise, y compris les émissions liées au transport de marchandises, à la gestion des déchets, à l'utilisation des produits vendus, etc."
        }
    }
 
model = SentenceTransformer("paraphrase-multilingual-MiniLM-L12-v2")
 
# Exemples de référence enrichis pour chaque scope
SCOPE_EXAMPLES = {
    "SCOPE_1": (
        "diesel essence carburant gaz naturel véhicule combustion fioul propane butane "
        "charbon biomasse émissions directes plein carburant station service litres"
    ),
    "SCOPE_2": (
        "électricité kWh MWh chauffage énergie achetée consommation électrique facture EDF "
        "fournisseur énergie refroidissement vapeur compteur abonnement"
    ),
    "SCOPE_3": (
        "transport marchandises déchets fournisseur logistique sous-traitant livraison "
        "immobilisations achat matières premières prestataire bilan carbone indirect"
    ),
}
 
def cosine_similarity(a, b):
    return np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b))

def chunk_text(text, size=150):
    words = text.split()
    return [" ".join(words[i:i+size]) for i in range(0, len(words), size)]
 
 
# Calcul des vecteurs d'embedding pour les mots-clés de chaque scope
scope_vectors = {}
for scope, text in SCOPE_EXAMPLES.items():
    scope_vectors[scope] = model.encode(text)
 
 
def classify_scope(text):
    if not text or len(text.strip()) < 100:
        return "UNKNOWN"
 
    chunks = chunk_text(text)
 
    best_scope = "UNKNOWN"
    best_score = -1
 
    for chunk in chunks:
        vec = model.encode(chunk)
 
        for scope, scope_vec in scope_vectors.items():
            score = cosine_similarity(vec, scope_vec)
 
            if score > best_score:
                best_score = score
                best_scope = scope
 
    print(f"  [DEBUG] score={best_score:.3f} → {best_scope}")
    if best_score < 0.10:
        return "Fichier non traitable"
 
    return best_scope
 
 
def extract_text(path):
    ext = path.lower().split(".")[-1]
 
    try:
        # TXT
        if ext == "txt":
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                return f.read()
 
        # CSV
        elif ext == "csv":
            df = pd.read_csv(path, sep=None, engine="python", encoding="latin1", on_bad_lines='skip')
            content = df.astype(str).fillna('').values.flatten()
            return " ".join([str(item) for item in content if str(item).lower() != 'nan'])
 
        # EXCEL
        elif ext in ["xls", "xlsx", "xlsm"]:
            df = pd.read_excel(path)
            content = df.astype(str).fillna('').values.flatten()
            return " ".join([str(item) for item in content if str(item).lower() != 'nan'])
 
        # DOCX
        elif ext == "docx":
            doc = Document(path)
            return "\n".join([p.text for p in doc.paragraphs])
 
        # PDF (PyMuPDF)
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
 
 
def run_test(folder="cleaned_files"):
    results = {}
 
    for file in os.listdir(folder):
        path = os.path.join(folder, file)
 
        if not os.path.isfile(path):
            continue
 
        print("\n---", file)
 
        text = extract_text(path)
        scope = classify_scope(text)
 
        print("→", scope)
 
        results[file] = scope
 
    return results
 
 
if __name__ == "__main__":
    results = run_test()
 
    print("\nRÉSULTATS")
    for k, v in results.items():
        print(k, ":", v)