import getters_data as gd
import re
import pandas as pd
import os

def extract_from_table(filename, content, scope):
    resultats = []

    # Unifier CSV et Excel
    if content.get("type") == "csv":
        sheets = {"sheet1": content.get("dataframe")}
    else:
        sheets = content.get("sheets", {})

    for sheet_name, df in sheets.items():
        if df is None or df.empty:
            continue

        # Normaliser les noms de colonnes
        df.columns = [gd.normalize_text(col) for col in df.columns]

        # Chercher les colonnes clés
        col_designation = next((c for c in df.columns if c in ["designation", "libelle", "nature", "produit", "description"]), None)
        col_quantite    = next((c for c in df.columns if c in ["quantite", "qte", "volume", "consommation", "total"]), None)
        col_unite       = next((c for c in df.columns if c in ["unite", "unites", "u"]), None)

        if col_quantite is None:
            continue  # pas de colonne quantité → feuille inutile

        for _, row in df.iterrows():
            try:
                # 1. Récupérer la quantité
                quantite = float(str(row[col_quantite]).replace(",", ".").strip())
                if quantite <= 0:
                    continue

                # 2. Récupérer la désignation
                designation_brute = str(row[col_designation]).strip() if col_designation else sheet_name
                designation_norm  = gd.normalize_text(designation_brute)
                designation       = gd.CORRESPONDANCES.get(designation_norm, designation_brute)

                # 3. Récupérer l'unité
                unite = str(row[col_unite]).strip() if col_unite else ""
                unite = gd.referentiel.get(unite.lower().strip(), unite)

                resultats.append({
                    "fichier":     filename,
                    "scope":       scope,
                    "designation": designation,
                    "quantite":    quantite,
                    "unite":       unite,
                    "sheet":       sheet_name,
                    "confiance":   "haute" if col_designation else "moyenne"
                })

            except (ValueError, TypeError):
                continue

    return resultats