import load_data as ld
import process_data as pt

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
        "Date": None,
        "ID_Facture": None,
        "Designation": None, 
        "Prix": 0.0,
        "Quantite": 0.0,
        "Unite": None,
        "Details_Vehicule": {
            "Marque_Modele": None,
            "Carburant": None,
            "Type": None # VP ou VUL
        },
        "Emissions_tCO2e": 0.0
    }