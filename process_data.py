import load_data as ld # importation du module de chargement des données
import pandas as pd # pour la manipulation des données tabulaires
import numpy as np # pour la manipulation des données numériques
from google import genai # pour l'utilisation de l'API de Google Gemini (nouveau package)
from dotenv import load_dotenv # pour charger les variables d'environnement à partir d'un fichier .env
import os # pour accéder aux variables d'environnement
import time # pour gérer les délais entre les requêtes à l'API

load_dotenv() # Charger les variables d'environnement à partir du fichier .env (assurez-vous que le fichier .env est dans le même répertoire que ce script et contient la clé GEMINI_API_KEY)
api = genai.Client(api_key=os.getenv("GEMINI_API_KEY")) 

