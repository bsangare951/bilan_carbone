import getters_data as gd
import extract_files_data as ev
import os
from collections import defaultdict

# Unités qui signifient que l'émission est déjà calculée
UNITES_DEJA_CALCULEES = ["kgco2e", "tco2e", "kgco2", "co2e"]

# Facteurs de conversion tCO2e
KG_TO_T = 0.001


def calculer_emissions(data_extraite):
    resultats = []
    non_calcules = []

    for entry in data_extraite:
        designation = entry.get("designation", "")
        quantite    = entry.get("quantite", 0)
        unite       = entry.get("unite", "").strip()
        scope       = entry.get("scope", "")
        source      = entry.get("source", "")
        fiabilite   = entry.get("fiabilite", "faible")

        # CAS 1 : émissions déjà calculées
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
                "source_facteur":   "déjà calculé"
            })
            print(f"  [OK] {designation[:40]} → {emissions_kg:.2f} kgCO2e (déjà calculé)")
            continue

        # CAS 2
        print(f"  [FE] Recherche facteur pour '{designation}' ({unite})...")
        facteur = gd.get_facteurs(designation, unite)

        if facteur is None or facteur.get("valeur") is None:
            print(f"Facteur introuvable : {designation} ({unite})")
            non_calcules.append({
                "source":      source,
                "scope":       scope,
                "designation": designation,
                "quantite":    quantite,
                "unite":       unite,
                "raison":      "Facteur d'émission introuvable"
            })
            continue

        valeur_fe   = facteur.get("valeur", 0)
        unite_fe    = facteur.get("unite", "")
        incertitude = facteur.get("incertitude", 0)
        source_fe   = facteur.get("source", "Base Carbone ADEME")

        if not valeur_fe or valeur_fe <= 0:
            print(f"Facteur nul ou invalide : {designation}")
            non_calcules.append({
                "source":      source,
                "scope":       scope,
                "designation": designation,
                "quantite":    quantite,
                "unite":       unite,
                "raison":      f"Facteur nul : {valeur_fe}"
            })
            continue

        # CAS 3
        emissions_kg = quantite * valeur_fe
        print(f"  [OK] {designation[:35]} → {quantite} {unite} × {valeur_fe} = {emissions_kg:.2f} kgCO2e")

        resultats.append({
            "source":           source,
            "scope":            scope,
            "designation":      designation,
            "quantite":         quantite,
            "unite":            unite,
            "facteur_emission": valeur_fe,
            "unite_facteur":    unite_fe,
            "emissions_kgCO2e": round(emissions_kg, 4),
            "emissions_tCO2e":  round(emissions_kg * KG_TO_T, 6),
            "incertitude":      incertitude,
            "fiabilite":        fiabilite,
            "source_facteur":   source_fe
        })

    return resultats, non_calcules


def agreger_par_scope(resultats):
    """
    Agrège les émissions par scope et calcule :
    - Total kgCO2e par scope
    - Total tCO2e par scope
    - Incertitude moyenne pondérée
    - Total général
    """
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

    # Calcul incertitude moyenne par scope
    bilan = {}
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
    """
    Calcule le taux d'incertitude global du bilan :
    - Incertitude des facteurs d'émission (ADEME)
    - Incertitude liée aux fichiers manquants
    """
    # Incertitude facteurs
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

    # Incertitude fichiers manquants
    if fichiers_attendus:
        nb_presents  = len(set(d["source"] for d in data_extraite))
        taux_completude = min(nb_presents / fichiers_attendus, 1.0)
        incert_fichiers = (1 - taux_completude) * 100
    else:
        incert_fichiers = 0

    incertitude_globale = round((incert_facteurs + incert_fichiers) / 2, 2)

    return {
        "incertitude_facteurs_%":  round(incert_facteurs, 2),
        "incertitude_fichiers_%":  round(incert_fichiers, 2),
        "incertitude_globale_%":   incertitude_globale,
        "niveau_confiance":        "haute" if incertitude_globale < 20
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
        print(f"  {'SOUS-TOTAL':>40} : {data['emissions_tCO2e']:>10.4f} tCO2e")

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

    bilan      = agreger_par_scope(resultats)
    incertitude = calculer_incertitude_bilan(bilan, data_extraite)

    afficher_bilan(bilan, incertitude)