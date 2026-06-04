import getters_data as gd
import extract_files_data as ev
import os
from collections import defaultdict

# Unités qui signifient que l'émission est déjà calculée
UNITES_DEJA_CALCULEES = frozenset(["kgco2e", "tco2e", "kgco2", "co2e"])

# Mots-clés pour ignorer les lignes total (double comptage)
MOTS_TOTAL_DESIG = frozenset(["total", "sous-total", "cumul", "somme"])

# Mots-clés indiquant un déplacement de personne ou kilométrage véhicule
# → le FE en t.km ne doit JAMAIS être appliqué sur ces désignations
MOTS_DEPLACEMENT = frozenset([
    "voiture", "vp", "moto", "velo", "bus", "car", "autocar",
    "metro", "rer", "tramway", "train", "tgv", "intercit", "avion",
    "domicile", "mission", "professionnel", "personnel",
    "vehicule utilitaire", "vul", "utilitaire leger", "motorisation",
    "deplacement", "km annuel", "kilometrage",
])

# Mots-clés indiquant du fret avec tonnage (seul cas où t.km -> km est licite)
MOTS_FRET = frozenset([
    "fret", "marchandise", "livraison", "logistique", "camion",
    "poids lourd", "articule", "transporteur", "benne", "porteur",
    "semi", "tracteur", "gnr", "gazole non routier",
])

KG_TO_T = 0.001

# ---------------------------------------------------------------------------
# CONVERSION DE FACTEUR D'ÉMISSION
# ---------------------------------------------------------------------------

def convertir_facteur_si_necessaire(valeur_fe: float, unite_fe: str, unite_quantite: str):
    """Convertit le FE si les unités sont compatibles mais d'échelle différente.
    Retourne (valeur_convertie, unite_convertie) ou (None, 'incompatible')."""
    fen  = unite_fe.lower().replace(" ", "")
    uqn  = unite_quantite.lower().strip()

    # Vérification de compatibilité
    if "/m3" in fen or "/m³" in fen:
        if uqn not in {"m3", "m³", "l", "litre", "litres"}:
            return None, "incompatible"
    elif "/l" in fen:
        if uqn not in {"l", "litre", "litres", "m3", "m³"}:
            return None, "incompatible"
    elif "/kg" in fen:
        if uqn not in {"kg", "t", "tonne", "tonnes"}:
            return None, "incompatible"
    elif any(u in fen for u in ("/t", "/tonne", "/tonnes")):
        if uqn not in {"kg", "t", "tonne", "tonnes"}:
            return None, "incompatible"
    elif "/mwh" in fen or "/kwh" in fen:
        if uqn not in {"kwh", "mwh"}:
            return None, "incompatible"
    elif "/km" in fen:
        if uqn not in {"km", "t.km", "tonne.km", "tonnes.km",
                       "km.passager", "km.passagers", "passager.km"}:
            return None, "incompatible"

    # Conversions d'échelle
    if any(u in fen for u in ("/t", "/tonne", "/tonnes")) and uqn == "kg":
        return valeur_fe / 1_000, f"{unite_fe} (converti /kg)"
    if "/kg" in fen and uqn in {"t", "tonne", "tonnes"}:
        return valeur_fe * 1_000, f"{unite_fe} (converti /t)"
    if "/mwh" in fen and uqn == "kwh":
        return valeur_fe / 1_000, f"{unite_fe} (converti /kWh)"
    if "/kwh" in fen and uqn == "mwh":
        return valeur_fe * 1_000, f"{unite_fe} (converti /MWh)"
    if ("/m3" in fen or "/m³" in fen) and uqn in {"l", "litre", "litres"}:
        return valeur_fe / 1_000, f"{unite_fe} (converti /L)"
    if "/l" in fen and uqn in {"m3", "m³"}:
        return valeur_fe * 1_000, f"{unite_fe} (converti /m³)"

    return valeur_fe, unite_fe



def _estimate_tonnes(designation: str, data_extraite: list) -> float | None:
    """Tente d'inférer la charge transportée en tonnes depuis les données extraites.
    Utilisé UNIQUEMENT pour convertir un FE en kgCO2e/t.km vers kgCO2e/km
    quand la désignation est bien du fret (pas un déplacement)."""
    d = designation.lower()

    # 1. Chercher un poids/tonnage documenté dans les extractions
    for e in data_extraite:
        try:
            ue   = str(e.get("unite", "")).lower()
            name = str(e.get("designation", "")).lower()
            if ue in {"kg", "t", "tonne", "tonnes"} and any(
                k in name for k in ("poids", "tonnage", "charge", "capacit", "payload")
            ):
                q = float(e.get("quantite", 0) or 0)
                if q > 0:
                    return q / 1_000 if ue == "kg" else q
        except Exception:
            continue

    # 2. Valeurs empiriques par type de véhicule de fret
    if "articul" in d or "poids lourd" in d:
        return 20.0
    if "camion" in d or "fret" in d:
        return 12.0
    if "vehicule utilitaire" in d or "vul" in d or "utilitaire" in d:
        return 1.5

    return None



def calculer_emissions(data_extraite: list) -> tuple[list, list]:
    resultats:    list[dict] = []
    non_calcules: list[dict] = []

    for entry in data_extraite:
        designation = entry.get("designation", "")
        quantite    = entry.get("quantite", 0)
        unite       = entry.get("unite", "").strip()
        scope       = entry.get("scope", "")
        source      = entry.get("source", "")
        fiabilite   = entry.get("fiabilite", "faible")

        # Filtres financiers 
        mots_fin = {"€", "eur", "euro", "euros", "montant", "cout", "coût",
                    "prix", "tarif", "frais", "abonnement", "tva", "ht", "ttc"}
        desig_norm = gd.normalize_text(designation)
        unite_norm = gd.normalize_text(unite)
        # Note : on ne filtre QUE sur les symboles monétaires dans l'unité,
        # pas sur les mots de la désignation pour éviter les faux positifs
        if any(m in desig_norm for m in mots_fin) or any(m in unite_norm for m in {"€", "eur", "euro", "euros"}):
            print(f"  [SKIP] Donnée financière ignorée : {designation} ({unite})")
            continue

        #  Filtres unités non exploitables
        unites_invalides = {
            "unite-feuille", "unites-feuille", "feuille", "feuilles",
            "rouleau", "rouleaux", "copie", "copies", "page", "pages",
            "impression", "impressions", "m².an", "m2.an",
            "kgéq", "kgeq", "/m²", "/m2", "%", "ratio", "taux",
        }
        if any(t in unite_norm for t in unites_invalides):
            print(f"  [SKIP] Unité non exploitable : {designation} ({unite})")
            continue

        # Filtres de vraisemblance
        q = float(quantite)
        if "kwh" in unite_norm and not (1 <= q <= 5_000_000):
            print(f"  [SKIP] Quantité non réaliste : {designation} ({quantite} {unite})")
            continue
        # Litres : élargi pour camions lourds (100 000+ L/an pour une flotte)
        if any(t in unite_norm for t in ("m³", "m3", "l", "litre", "litres")) and not (1 <= q <= 1_000_000):
            print(f"  [SKIP] Quantité non réaliste : {designation} ({quantite} {unite})")
            continue
        if any(t in unite_norm for t in ("kg", "t", "km", "t.km")) and not (1 <= q <= 50_000_000):
            print(f"  [SKIP] Quantité non réaliste : {designation} ({quantite} {unite})")
            continue

        # ── Ignorer les lignes de total ───────────────────────────────────
        if any(m in desig_norm for m in MOTS_TOTAL_DESIG):
            print(f"  [SKIP] Total ignoré : {designation}")
            continue
        # ── Émissions déjà calculées ──────────────────────────────────────
        if unite.lower().replace(" ", "") in UNITES_DEJA_CALCULEES:
            emissions_kg = q
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
                "source_facteur":   "deja calcule",
            })
            print(f"  [OK] {designation[:40]} -> {emissions_kg:.2f} kgCO2e (déjà calculé)")
            continue

        # ── Recherche du facteur d'émission ──────────────────────────────
        print(f"  [FE] Recherche facteur pour '{designation}' ({unite})...")
        facteur = gd.get_facteurs(designation, unite)

        if facteur is None or facteur.get("valeur") is None:
            print(f"  [ERREUR] Facteur introuvable : {designation} ({unite})")
            non_calcules.append({
                "source": source, "scope": scope,
                "designation": designation, "quantite": quantite, "unite": unite,
                "raison": "Facteur d'émission introuvable",
            })
            continue

        valeur_fe   = facteur.get("valeur", 0)
        unite_fe    = facteur.get("unite", "")
        incertitude = facteur.get("incertitude", 0)
        source_fe   = facteur.get("source", "Base Carbone ADEME")

        if ("t.km" in unite_fe.lower() or "tonne.km" in unite_fe.lower()) \
                and unite.lower().strip() == "km":

            desig_lower      = designation.lower()
            est_deplacement  = any(m in desig_lower for m in MOTS_DEPLACEMENT)
            est_fret_documente = (
                not est_deplacement
                and any(m in desig_lower for m in MOTS_FRET)
            )

            if est_fret_documente:
                # Fret avec tonnage estimable → conversion licite
                tonnes_est = _estimate_tonnes(designation, data_extraite)
                if tonnes_est:
                    print(f"    [ESTIM] Conversion FE {unite_fe} -> kgCO2e/km "
                          f"(charge estimée : {tonnes_est} t)")
                    valeur_fe = valeur_fe * tonnes_est
                    unite_fe  = "kgCO2e/km"
                else:
                    print(f"    [INFO] Charge non estimable pour {designation} — abandon")
                    non_calcules.append({
                        "source": source, "scope": scope,
                        "designation": designation, "quantite": quantite, "unite": unite,
                        "raison": f"FE en {unite_fe} : charge transportée non documentée",
                    })
                    continue
            else:
                # Déplacement ou km véhicule : chercher un FE en kgCO2e/km
                print(f"    [INFO] FE en {unite_fe} non applicable à des km de "
                      f"déplacement — recherche FE /km...")
                facteur_km = gd.local_search(designation, "km") or gd.api_search(designation, "km")
                if facteur_km and facteur_km.get("valeur"):
                    valeur_fe = facteur_km["valeur"]
                    unite_fe  = facteur_km["unite"]
                    print(f"    [OK] FE alternatif trouvé : {valeur_fe:.4f} {unite_fe}")
                else:
                    print(f"    [SKIP] Aucun FE /km disponible pour {designation} — ignoré")
                    non_calcules.append({
                        "source": source, "scope": scope,
                        "designation": designation, "quantite": quantite, "unite": unite,
                        "raison": f"FE uniquement en {unite_fe}, incompatible avec km de déplacement",
                    })
                    continue

        if unite.upper() in ("M3", "M³") and any(
            u in unite_fe.lower() for u in ("/t", "/tonne", "/tonnes")
        ):
            desig_low = designation.lower()
            densites = {
                "gravat": 1.8, "beton": 2.4, "terre": 1.6, "sable": 1.7,
                "calcaire": 2.7, "enrobe": 2.3, "pierre": 2.5,
            }
            densite = next(
                (d for k, d in densites.items() if k in desig_low), None
            )
            if densite:
                quantite = quantite * densite
                q = float(quantite)
                unite = "t"
                print(f"    [CONV M3->T] {designation} : ×{densite} t/m³ -> {quantite:.2f} t")

        if ("diesel" in designation.lower() or "gazole" in designation.lower()) \
                and unite == "L" and valeur_fe < 1:
            print(f"    [CONV DIESEL] FE erroné {valeur_fe} {unite_fe} "
                  f"-> remplacement par 2.41 kgCO2e/L")
            valeur_fe = 2.41
            unite_fe  = "kgCO2e/L"

        if ("papier" in designation.lower() or "carton" in designation.lower()) \
                and unite == "kg" and valeur_fe > 100:
            print(f"    [CONV] Papier/Carton : {valeur_fe} {unite_fe} -> /1000")
            valeur_fe = valeur_fe / 1_000
            unite_fe  = "kgCO2e/kg"

        if valeur_fe is None or valeur_fe <= 0:
            print(f"  [ERREUR] Facteur invalide ({valeur_fe}) pour : {designation}")
            non_calcules.append({
                "source": source, "scope": scope,
                "designation": designation, "quantite": quantite, "unite": unite,
                "raison": f"Facteur négatif ou nul : {valeur_fe}",
            })
            continue

        # ── Conversion d'unité finale ─────────────────────────────────────
        valeur_fe_conv, unite_fe_conv = convertir_facteur_si_necessaire(
            valeur_fe, unite_fe, unite
        )

        if valeur_fe_conv is None:
            print(f"  [ERREUR] Unité incompatible : {unite_fe} vs {unite} pour {designation}")
            non_calcules.append({
                "source": source, "scope": scope,
                "designation": designation, "quantite": quantite, "unite": unite,
                "raison": f"Unité incompatible : FE={unite_fe}, quantité={unite}",
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
            "source_facteur":   source_fe,
        })

    return resultats, non_calcules



def agreger_par_scope(resultats: list[dict]) -> dict:
    agregat = defaultdict(lambda: {
        "emissions_kgCO2e": 0.0,
        "emissions_tCO2e":  0.0,
        "nb_sources":       0,
        "incertitudes":     [],
        "details":          [],
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

    bilan:    dict = {}
    total_kg: float = 0.0

    for scope, data in sorted(agregat.items()):
        incert_moy = (
            sum(data["incertitudes"]) / len(data["incertitudes"])
            if data["incertitudes"] else 0
        )
        bilan[scope] = {
            "emissions_kgCO2e": round(data["emissions_kgCO2e"], 4),
            "emissions_tCO2e":  round(data["emissions_tCO2e"], 6),
            "nb_sources":       data["nb_sources"],
            "incertitude_moy":  round(incert_moy, 2),
            "details":          data["details"],
        }
        total_kg += data["emissions_kgCO2e"]

    bilan["TOTAL"] = {
        "emissions_kgCO2e": round(total_kg, 4),
        "emissions_tCO2e":  round(total_kg * KG_TO_T, 6),
    }
    return bilan




def calculer_incertitude_bilan(
    bilan: dict,
    data_extraite: list,
    fichiers_attendus: int | None = None,
) -> dict:
    toutes_incertitudes = [
        d.get("incertitude", 0)
        for scope, data in bilan.items()
        if scope != "TOTAL"
        for d in data.get("details", [])
        if d.get("incertitude")
    ]
    incert_facteurs = (
        sum(toutes_incertitudes) / len(toutes_incertitudes)
        if toutes_incertitudes else 0
    )

    if fichiers_attendus:
        nb_presents     = len({d["source"] for d in data_extraite})
        taux_completude = min(nb_presents / fichiers_attendus, 1.0)
        incert_fichiers = (1 - taux_completude) * 100
    else:
        incert_fichiers = 0

    incertitude_globale = round((incert_facteurs + incert_fichiers) / 2, 2)

    return {
        "incertitude_facteurs_%": round(incert_facteurs, 2),
        "incertitude_fichiers_%": round(incert_fichiers, 2),
        "incertitude_globale_%":  incertitude_globale,
        "niveau_confiance": (
            "haute"   if incertitude_globale < 20 else
            "moyenne" if incertitude_globale < 40 else
            "faible"
        ),
    }



def afficher_bilan(bilan: dict, incertitude: dict) -> None:
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
    print(f"  Incertitude Globale : ±{incertitude['incertitude_globale_%']}% "
          f"[{incertitude['niveau_confiance']}]")
    print("═" * 65)



if __name__ == "__main__":
    if os.path.exists("cache_fe.pkl"):
        os.remove("cache_fe.pkl")
        print("[CACHE] cache_fe.pkl supprimé — recalcul des facteurs")

    print("Extraction des données...")
    data_extraite, nb_fichiers_tentes = ev.lancer_le_bilan()

    print(f"\n{len(data_extraite)} valeurs extraites. Calcul des émissions...")
    print("─" * 65)

    resultats, non_calcules = calculer_emissions(data_extraite)

    print(f"\n{len(resultats)} émissions calculées")
    if non_calcules:
        print(f"{len(non_calcules)} valeur(s) non calculée(s) :")
        for nc in non_calcules:
            print(f"  - {nc['designation']} ({nc['unite']}) : {nc['raison']}")

    bilan       = agreger_par_scope(resultats)
    incertitude = calculer_incertitude_bilan(bilan, data_extraite,
                                             fichiers_attendus=nb_fichiers_tentes)

    afficher_bilan(bilan, incertitude)