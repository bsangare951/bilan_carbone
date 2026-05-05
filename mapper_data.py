import os
from dotenv import load_dotenv
import load_data as ld
import process_data as pt
import google.genai as genai
import time

load_dotenv()
api_key = os.getenv("GEMINI_API_KEY")
ai = genai.Client(api_key=api_key)



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
        return{
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

