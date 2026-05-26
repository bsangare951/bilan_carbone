import getters_data as gd
import extract_files_data as ev
import os
from collections import defaultdict
import pickle

# Unités qui signifient que l'émission est déjà calculée
UNITES_DEJA_CALCULEES = ["kgco2e", "tco2e", "kgco2", "co2e"]

# Mots clés pour ignorer les lignes total (double comptage)
MOTS_TOTAL_DESIG = ["total", "sous-total", "cumul", "somme"]

# Facteur de conversion
KG_TO_T = 0.001


def convertir_facteur_si_necessaire(valeur_fe, unite_fe, unite_quantite):

    unite_fe_norm = unite_fe.lower().replace(" ", "")
    unite_q_norm  = unite_quantite.lower().strip()

    # FE en /tonne mais quantité en kg
    if any(u in unite_fe_norm for u in ["/t", "/tonne", "/tonnes"]) \
       and unite_q_norm == "kg":
        return valeur_fe / 1000, f"{unite_fe} (converti /kg)"

    # FE en /kg mais quantité en tonne
    if "/kg" in unite_fe_norm and unite_q_norm in ["t", "tonne", "tonnes"]:
        return valeur_fe * 1000, f"{unite_fe} (converti /t)"

    # FE en /MWh mais quantité en kWh
    if "/mwh" in unite_fe_norm and unite_q_norm == "kwh":
        return valeur_fe / 1000, f"{unite_fe} (converti /kWh)"

    # FE en /kWh mais quantité en MWh
    if "/kwh" in unite_fe_norm and unite_q_norm == "mwh":
        return valeur_fe * 1000, f"{unite_fe} (converti /MWh)"

    # FE en /km mais quantité en t.km → incompatible, on signale
    if "/km" in unite_fe_norm and unite_q_norm == "t.km":
        return None, "incompatible"

    return valeur_fe, unite_fe


def calculer_emissions(data_extraite):
    resultats    = []
    non_calcules = []

    for entry in data_extraite:
        designation = entry.get("designation", "")
        quantite    = entry.get("quantite", 0)
        unite       = entry.get("unite", "").strip()
        scope       = entry.get("scope", "")
        source      = entry.get("source", "")
        fiabilite   = entry.get("fiabilite", "faible")

        desig_norm = gd.normalize_text(designation)
        if any(m in desig_norm for m in MOTS_TOTAL_DESIG):
            print(f"  [SKIP] Total ignore : {designation}")
            continue

        if unite.lower().replace(" ", "") in UNITES_DEJA_CALCULEES:
            emissions_kg = quantite
            resultats.append({
                "source":           source,
                "scope":            scope,
                "designation":      designation,
                "quantite":         quantite,
                "unite":            unite,
                "facteur_emission": 1.0,
                "unite_facteur":    unite,
                "emissions_kgCO2e": round(emissions_kg, 4),
                "emissions_tCO2e":  round(emissions_kg * KG_TO_T, 6),
                "incertitude":      entry.get("incertitude", 0),
                "fiabilite":        fiabilite,
                "source_facteur":   "deja calcule"
            })
            print(f"  [OK] {designation[:40]} -> {emissions_kg:.2f} kgCO2e (deja calcule)")
            continue

        print(f"  [FE] Recherche facteur pour '{designation}' ({unite})...")
        facteur = gd.get_facteurs(designation, unite)

        if facteur is None or facteur.get("valeur") is None:
            print(f"  [ERREUR] Facteur introuvable : {designation} ({unite})")
            non_calcules.append({
                "source":      source,
                "scope":       scope,
                "designation": designation,
                "quantite":    quantite,
                "unite":       unite,
                "raison":      "Facteur d emission introuvable"
            })
            continue

        valeur_fe   = facteur.get("valeur", 0)
        unite_fe    = facteur.get("unite", "")
        incertitude = facteur.get("incertitude", 0)
        source_fe   = facteur.get("source", "Base Carbone ADEME")

        if ("diesel" in designation.lower() or "gazole" in designation.lower()) and unite == "L":
            if valeur_fe < 1:
                print(f"    [CONV DIESEL] Facteur erroné {valeur_fe} {unite_fe} -> remplacement par 2.41 kgCO2e/L")
                valeur_fe = 2.41
                unite_fe = "kgCO2e/L"

        if ("papier" in designation.lower() or "carton" in designation.lower()) and unite == "kg":
            if valeur_fe > 100:
                print(f"    [CONV] Papier/Carton : {valeur_fe} {unite_fe} -> /1000")
                valeur_fe = valeur_fe / 1000
                unite_fe = "kgCO2e/kg"

        if valeur_fe is None or valeur_fe <= 0:
            print(f"  [ERREUR] Facteur invalide ({valeur_fe}) pour : {designation}")
            non_calcules.append({
                "source":      source,
                "scope":       scope,
                "designation": designation,
                "quantite":    quantite,
                "unite":       unite,
                "raison":      f"Facteur negatif ou nul : {valeur_fe}"
            })
            continue

        valeur_fe_conv, unite_fe_conv = convertir_facteur_si_necessaire(valeur_fe, unite_fe, unite)

        if valeur_fe_conv is None:
            print(f"  [ERREUR] Unite incompatible : {unite_fe} vs {unite} pour {designation}")
            non_calcules.append({
                "source":      source,
                "scope":       scope,
                "designation": designation,
                "quantite":    quantite,
                "unite":       unite,
                "raison":      f"Unite incompatible : FE={unite_fe}, quantite={unite}"
            })
            continue

        emissions_kg = quantite * valeur_fe_conv
        print(f"  [OK] {designation[:35]} -> {quantite} {unite} x "
              f"{valeur_fe_conv:.4f} = {emissions_kg:.2f} kgCO2e")

        resultats.append({
            "source":           source,
            "scope":            scope,
            "designation":      designation,
            "quantite":         quantite,
            "unite":            unite,
            "facteur_emission": valeur_fe_conv,
            "unite_facteur":    unite_fe_conv,
            "emissions_kgCO2e": round(emissions_kg, 4),
            "emissions_tCO2e":  round(emissions_kg * KG_TO_T, 6),
            "incertitude":      incertitude,
            "fiabilite":        fiabilite,
            "source_facteur":   source_fe
        })

    return resultats, non_calcules

def agreger_par_scope(resultats):
    agregat = defaultdict(lambda: {
        "emissions_kgCO2e": 0.0,
        "emissions_tCO2e":  0.0,
        "nb_sources":       0,
        "incertitudes":     [],
        "details":          []
    })

    for r in resultats:
        scope = r["scope"]
        agregat[scope]["emissions_kgCO2e"] += r["emissions_kgCO2e"]
        agregat[scope]["emissions_tCO2e"]  += r["emissions_tCO2e"]
        agregat[scope]["nb_sources"]       += 1
        if r.get("incertitude"):
            try:
                agregat[scope]["incertitudes"].append(float(r["incertitude"]))
            except (ValueError, TypeError):
                pass
        agregat[scope]["details"].append(r)

    bilan    = {}
    total_kg = 0.0
    for scope, data in sorted(agregat.items()):
        incert_moy = (sum(data["incertitudes"]) / len(data["incertitudes"])
                      if data["incertitudes"] else 0)
        bilan[scope] = {
            "emissions_kgCO2e": round(data["emissions_kgCO2e"], 4),
            "emissions_tCO2e":  round(data["emissions_tCO2e"], 6),
            "nb_sources":       data["nb_sources"],
            "incertitude_moy":  round(incert_moy, 2),
            "details":          data["details"]
        }
        total_kg += data["emissions_kgCO2e"]

    bilan["TOTAL"] = {
        "emissions_kgCO2e": round(total_kg, 4),
        "emissions_tCO2e":  round(total_kg * KG_TO_T, 6),
    }
    return bilan


def calculer_incertitude_bilan(bilan, data_extraite, fichiers_attendus=None):
    toutes_incertitudes = []
    for scope, data in bilan.items():
        if scope == "TOTAL":
            continue
        toutes_incertitudes.extend(
            [d.get("incertitude", 0) for d in data.get("details", [])
             if d.get("incertitude")]
        )

    incert_facteurs = (sum(toutes_incertitudes) / len(toutes_incertitudes)
                       if toutes_incertitudes else 0)

    if fichiers_attendus:
        nb_presents     = len(set(d["source"] for d in data_extraite))
        taux_completude = min(nb_presents / fichiers_attendus, 1.0)
        incert_fichiers = (1 - taux_completude) * 100
    else:
        incert_fichiers = 0

    incertitude_globale = round((incert_facteurs + incert_fichiers) / 2, 2)

    return {
        "incertitude_facteurs_%": round(incert_facteurs, 2),
        "incertitude_fichiers_%": round(incert_fichiers, 2),
        "incertitude_globale_%":  incertitude_globale,
        "niveau_confiance":       "haute"   if incertitude_globale < 20
                                  else "moyenne" if incertitude_globale < 40
                                  else "faible"
    }


def afficher_bilan(bilan, incertitude):
    """Affiche le bilan formaté dans le terminal."""
    print("\n" + "═" * 65)
    print("  BILAN CARBONE — RÉSULTATS")
    print("═" * 65)

    for scope, data in bilan.items():
        if scope == "TOTAL":
            continue
        print(f"\n  {scope}")
        print(f"  {'─' * 40}")
        for d in data.get("details", []):
            print(f"    {d['designation'][:35]:<35} "
                  f"{d['emissions_tCO2e']:>10.4f} tCO2e")
        print(f"  {'SOUS-TOTAL':>40} : "
              f"{data['emissions_tCO2e']:>10.4f} tCO2e")

    print("\n" + "═" * 65)
    print(f"  TOTAL GÉNÉRAL : {bilan['TOTAL']['emissions_tCO2e']:.4f} tCO2e")
    print(f"                  {bilan['TOTAL']['emissions_kgCO2e']:.2f} kgCO2e")
    print("═" * 65)
    print(f"\n  INCERTITUDE")
    print(f"  Facteurs d'émission : ±{incertitude['incertitude_facteurs_%']}%")
    print(f"  Fichiers manquants  : ±{incertitude['incertitude_fichiers_%']}%")
    print(f"  Globale             : ±{incertitude['incertitude_globale_%']}% "
          f"[{incertitude['niveau_confiance']}]")
    print("═" * 65)


if __name__ == "__main__":
    # Supprimer le cache pour forcer un recalcul des facteurs
    if os.path.exists("cache_fe.pkl"):
        os.remove("cache_fe.pkl")
        print("[CACHE] cache_fe.pkl supprimé — recalcul des facteurs")

    print("Extraction des données...")
    data_extraite = ev.lancer_le_bilan()

    print(f"\n{len(data_extraite)} valeurs extraites. Calcul des émissions...")
    print("─" * 65)

    resultats, non_calcules = calculer_emissions(data_extraite)

    print(f"\n{len(resultats)} émissions calculées")
    if non_calcules:
        print(f"{len(non_calcules)} valeur(s) non calculée(s) :")
        for nc in non_calcules:
            print(f"  - {nc['designation']} ({nc['unite']}) : {nc['raison']}")

    bilan       = agreger_par_scope(resultats)
    incertitude = calculer_incertitude_bilan(bilan, data_extraite)

    afficher_bilan(bilan, incertitude)