from __future__ import annotations

import os
from collections import defaultdict
from typing import Any

import extract_files_data as ev
import getters_data as gd

KG_TO_T = 0.001

UNITES_CALCULEES_KG = frozenset(["kgco2e", "kgco2", "co2e"])
UNITES_CALCULEES_T = frozenset(["tco2e", "tco2"])

MOTS_FINANCIERS = frozenset([
    "€", "eur", "euro", "euros", "montant", "cout", "coût", "prix", "tarif",
    "frais", "abonnement", "tva", "ht", "ttc", "amort", "caution", "solde",
])

UNITES_NON_PHYSIQUES = frozenset([
    "kw", "w", "mw", "l/100", "l/100km", "kwh/m2", "kwh/m²", "%", "ratio", "taux",
    "copie", "copies", "page", "pages", "impression", "impressions", "nombre", "nb",
    "feuille", "feuilles", "unite-feuille", "unites-feuille", "rouleau", "rouleaux",
])

MATERIAUX_CLIENTS = frozenset([
    "terres inertes", "sable", "calcaire", "beton non valorisable", "béton non valorisable",
    "blocs beton", "blocs béton", "souches", "deblais", "déblais", "remblais",
    "enrobe", "enrobé", "granulat", "materiaux client", "matériaux client",
])

MOTS_FRET_LOURD = frozenset([
    "articule", "articulé", "poids lourd", "semi", "camion", "benne", "porteur",
    "tracteur routier", "34 a 40", "34 à 40",
])

MOTS_CONTEXTE_FRET = frozenset([
    "fret", "marchandise", "marchandises", "poids lourd", "camion", "benne", "semi",
    "porteur", "transport routier", "logistique",
])

MOTS_CARBURANT = frozenset(["gazole", "diesel", "gasoil", "gnr", "essence", "carburant", "fioul"])
MOTS_VEHICULE = frozenset(["vehicule", "véhicule", "voiture", "bus", "km", "deplacement", "déplacement", "domicile", "mission"])

# Fallbacks utilisés seulement si la recherche FE échoue ou retourne une unité incompatible.
# Ils couvrent les cas simples et très fréquents afin d'éviter "Carburants (L) : FE introuvable".
FACTEURS_MANUELS = [
    {
        "mots": ("gazole", "diesel", "gasoil", "carburant"),
        "unites": ("l", "litre", "litres"),
        "designation": "Gazole routier (fallback)",
        "valeur": 2.4108,
        "unite": "kgCO2e/L",
        "source": "fallback contrôlé",
        "incertitude": 20,
    },
    {
        "mots": ("gnr", "gazole non routier"),
        "unites": ("l", "litre", "litres"),
        "designation": "Gazole non routier (fallback)",
        "valeur": 2.70,
        "unite": "kgCO2e/L",
        "source": "fallback contrôlé",
        "incertitude": 25,
    },
    {
        "mots": ("essence", "sp95", "sp98", "e10"),
        "unites": ("l", "litre", "litres"),
        "designation": "Essence (fallback)",
        "valeur": 2.28,
        "unite": "kgCO2e/L",
        "source": "fallback contrôlé",
        "incertitude": 25,
    },
    {
        "mots": ("electric", "électric", "electricite", "électricité", "energie", "énergie", "edf"),
        "unites": ("kwh",),
        "designation": "Mix électricité France (fallback)",
        "valeur": 0.0609,
        "unite": "kgCO2e/kWh",
        "source": "fallback contrôlé",
        "incertitude": 10,
    },
]


# Facteurs monétaires contrôlés (kgCO2e / euro dépensé).
# Ils sont utilisés uniquement pour les postes Scope 3 catégorisés par extract_files_data.py.
FACTEURS_MONETAIRES_DIRECTS = {
    "ciment": (0.85, "kgCO2e/€", "facteur monétaire contrôlé - ciment"),
    "grave non traitee": (0.08, "kgCO2e/€", "facteur monétaire contrôlé - grave"),
    "calcaire": (0.06, "kgCO2e/€", "facteur monétaire contrôlé - calcaire"),
    "sable": (0.06, "kgCO2e/€", "facteur monétaire contrôlé - sable"),
    "granulat": (0.06, "kgCO2e/€", "facteur monétaire contrôlé - granulats"),
    "materiaux": (0.06, "kgCO2e/€", "facteur monétaire contrôlé - matériaux"),
    "matériaux": (0.06, "kgCO2e/€", "facteur monétaire contrôlé - matériaux"),
    "grave non traitée": (0.08, "kgCO2e/€", "facteur monétaire contrôlé - grave"),
    "grave recyclee": (0.05, "kgCO2e/€", "facteur monétaire contrôlé - grave recyclée"),
    "grave recyclée": (0.05, "kgCO2e/€", "facteur monétaire contrôlé - grave recyclée"),
    "granulats / sable / terre": (0.06, "kgCO2e/€", "facteur monétaire contrôlé - granulats/sable/terre"),
    "paves / bordures": (0.10, "kgCO2e/€", "facteur monétaire contrôlé - pavés/bordures"),
    "pavés / bordures": (0.10, "kgCO2e/€", "facteur monétaire contrôlé - pavés/bordures"),
    "epi, fournitures admin. et petit materiel": (0.367, "kgCO2e/€", "facteur monétaire contrôlé - fournitures"),
    "epi, fournitures admin. et petit matériel": (0.367, "kgCO2e/€", "facteur monétaire contrôlé - fournitures"),
    "location de materiel": (0.196, "kgCO2e/€", "facteur monétaire contrôlé - services"),
    "location de matériel": (0.196, "kgCO2e/€", "facteur monétaire contrôlé - services"),
    "services entretien/maintenance": (0.196, "kgCO2e/€", "facteur monétaire contrôlé - maintenance"),
    "autres services": (0.196, "kgCO2e/€", "facteur monétaire contrôlé - services"),
    "services": (0.196, "kgCO2e/€", "facteur monétaire contrôlé - services"),
    "immobilisations": (0.273, "kgCO2e/€", "facteur monétaire contrôlé - immobilisations"),
    "achats materiels non detailles": (1.88, "kgCO2e/€", "facteur monétaire contrôlé - achats matériels non détaillés"),
    "achats matériels non détaillés": (1.88, "kgCO2e/€", "facteur monétaire contrôlé - achats matériels non détaillés"),
}

FACTEURS_MANUELS_COMPLEMENTAIRES = [
    {
        "mots": ("dechets verts", "déchets verts"),
        "unites": ("t", "tonne", "tonnes"),
        "designation": "Déchets verts - compostage (fallback)",
        "valeur": 25.0,
        "unite": "kgCO2e/t",
        "source": "fallback contrôlé",
        "incertitude": 40,
    },
    {
        "mots": ("ciment",),
        "unites": ("t", "tonne", "tonnes"),
        "designation": "Ciment (fallback)",
        "valeur": 866.0,
        "unite": "kgCO2e/t",
        "source": "fallback contrôlé",
        "incertitude": 40,
    },
    {
        "mots": ("grave", "granulat", "sable", "terre"),
        "unites": ("t", "tonne", "tonnes"),
        "designation": "Granulats (fallback)",
        "valeur": 7.0,
        "unite": "kgCO2e/t",
        "source": "fallback contrôlé",
        "incertitude": 50,
    },
]



def norm(txt: Any) -> str:
    return gd.normalize_text(str(txt or ""))


def norm_unit(unit: str) -> str:
    u = norm(unit).replace(" ", "")
    aliases = {
        "litre": "l", "litres": "l", "ltr": "l",
        "m³": "m3", "metrecube": "m3", "metrescubes": "m3",
        "tonne": "t", "tonnes": "t",
        "kilometre": "km", "kilometres": "km", "kms": "km",
        "kw/h": "kwh", "kilowattheure": "kwh", "kilowattheures": "kwh",
        "megawattheure": "mwh", "megawattheures": "mwh",
        "tkm": "t.km", "tonne.km": "t.km", "tonnes.km": "t.km", "tonnekilometre": "t.km",
        "kgco₂e": "kgco2e", "kgco2": "kgco2e", "kgco2e": "kgco2e",
        "tco₂e": "tco2e", "tco2": "tco2e", "tco2e": "tco2e",
    }
    return aliases.get(u, u)


def unit_is_financial(unite: str) -> bool:
    u = norm(unite)
    return any(m in u for m in ("€", "eur", "euro", "keuro", "k€"))


def facteur_financier_direct(designation: str, unite: str) -> tuple[float, str, str] | None:
    if not unit_is_financial(unite):
        return None
    d = norm(designation)
    for key, val in FACTEURS_MONETAIRES_DIRECTS.items():
        if key in d:
            return val
    return None


def role_achat_intrant(entry: dict) -> bool:
    role = norm(entry.get("role", ""))
    src = norm(entry.get("source", ""))
    des = norm(entry.get("designation", ""))
    return role == "achats_intrants" or "recap achat" in src or any(m in des for m in ["ciment", "grave", "granulat", "sable", "terre", "pave", "pavé", "bordure"])


def designation_financiere(designation: str, unite: str) -> bool:
    d = norm(designation)
    u = norm(unite)
    return any(m in d for m in MOTS_FINANCIERS) or unit_is_financial(u)


def unite_non_physique(unite: str) -> bool:
    """
    Vérifie les unités non exploitables sans bloquer les unités physiques.

    Ancien bug : on testait `kw` comme sous-chaîne, donc `kWh` était rejeté
    car `kw` est contenu dans `kwh`. Ici, on travaille par égalité stricte
    après normalisation.
    """
    u = norm_unit(unite)

    unites_valides = {
        "l", "m3", "kwh", "mwh", "kg", "t", "km", "t.km",
        "kgco2e", "tco2e",
    }
    if u in unites_valides:
        return False

    unites_invalides_exactes = {
        "kw", "w", "mw", "l/100", "l/100km", "kwh/m2", "kwh/m²",
        "%", "ratio", "taux", "copie", "copies", "page", "pages",
        "impression", "impressions", "nombre", "nb", "feuille", "feuilles",
        "unite-feuille", "unites-feuille", "rouleau", "rouleaux",
    }
    return u in unites_invalides_exactes


def est_materiau_client(designation: str) -> bool:
    d = norm(designation)
    return any(m in d for m in MATERIAUX_CLIENTS)


def est_fret_lourd_hallucine(designation: str, source: str) -> bool:
    d = norm(designation)
    s = norm(source)
    if not any(m in d for m in MOTS_FRET_LOURD):
        return False
    # Si le fichier lui-même ne parle pas clairement de fret lourd, on ne calcule pas.
    return not any(m in s for m in MOTS_CONTEXTE_FRET)


def kg_deja_calcule(quantite: float, unite: str) -> float | None:
    u = norm_unit(unite)
    if u in UNITES_CALCULEES_KG:
        return quantite
    if u in UNITES_CALCULEES_T:
        return quantite * 1000
    return None


def fe_unit_financial(unite_fe: str) -> bool:
    f = norm(unite_fe).replace(" ", "")
    return any(m in f for m in ("eur", "euro", "€", "keuro", "k€"))


def fe_unit_tkm(unite_fe: str) -> bool:
    f = norm(unite_fe).replace(" ", "")
    return "t.km" in f or "tonne.km" in f or "tonnes.km" in f


def fe_unit_100km_or_h2(unite_fe: str) -> bool:
    f = norm(unite_fe).replace(" ", "")
    return "100km" in f or "kgh2" in f or "h2/" in f


def unites_compatibles_local(unite_fe: str, unite_q: str) -> bool:
    """Compatibilité stricte entre unité du FE et unité de quantité.
    On refuse les FE financiers, les FE /100km et les FE t.km appliqués à km.
    """
    f = norm(unite_fe).replace(" ", "")
    q = norm_unit(unite_q)

    if fe_unit_financial(f) or fe_unit_100km_or_h2(f):
        return False

    # Ordre important : t.km avant /km.
    if fe_unit_tkm(f):
        return q == "t.km"
    if "/km" in f or "vehicule.km" in f or "véhicule.km" in f:
        return q == "km"
    if "/kwh" in f or "kwhpci" in f:
        return q in {"kwh", "mwh"}
    if "/mwh" in f:
        return q in {"kwh", "mwh"}
    if "/litre" in f or "/l" in f:
        return q in {"l", "m3"}
    if "/m3" in f or "/m³" in f:
        return q in {"l", "m3"}
    if "/kg" in f:
        return q in {"kg", "t"}
    if "/tonne" in f or f.endswith("/t") or "/tco2" in f:
        return q in {"kg", "t"}

    # Facteur sans dénominateur : uniquement accepté si quantité déjà kgCO2e/tCO2e,
    # qui est traitée ailleurs. Donc ici on refuse.
    return False


def convertir_si_necessaire(valeur_fe: float, unite_fe: str, unite_q: str) -> tuple[float | None, str]:
    f = norm(unite_fe).replace(" ", "")
    q = norm_unit(unite_q)

    if not unites_compatibles_local(unite_fe, unite_q):
        return None, "incompatible"

    # gCO2e -> kgCO2e si nécessaire
    if f.startswith("gco2") or f.startswith("gco₂"):
        valeur_fe = valeur_fe / 1000
        unite_fe = unite_fe.replace("gCO2e", "kgCO2e").replace("gCO₂e", "kgCO2e")
        f = norm(unite_fe).replace(" ", "")

    if ("/kwh" in f or "kwhpci" in f) and q == "mwh":
        return valeur_fe * 1000, f"{unite_fe} (converti /MWh)"
    if "/mwh" in f and q == "kwh":
        return valeur_fe / 1000, f"{unite_fe} (converti /kWh)"

    if ("/litre" in f or "/l" in f) and q == "m3":
        return valeur_fe * 1000, f"{unite_fe} (converti /m³)"
    if ("/m3" in f or "/m³" in f) and q == "l":
        return valeur_fe / 1000, f"{unite_fe} (converti /L)"

    if "/kg" in f and q == "t":
        return valeur_fe * 1000, f"{unite_fe} (converti /t)"
    if ("/tonne" in f or f.endswith("/t")) and q == "kg":
        return valeur_fe / 1000, f"{unite_fe} (converti /kg)"

    return valeur_fe, unite_fe


def facteur_fallback(designation: str, unite: str) -> dict | None:
    d = norm(designation)
    u = norm_unit(unite)
    for rule in FACTEURS_MANUELS + FACTEURS_MANUELS_COMPLEMENTAIRES:
        if u in rule["unites"] and any(m in d for m in rule["mots"]):
            return {
                "designation": rule["designation"],
                "valeur": rule["valeur"],
                "unite": rule["unite"],
                "source": rule["source"],
                "incertitude": rule["incertitude"],
                "score": 100,
            }
    return None


def variantes_designation(designation: str, unite: str) -> list[str]:
    d = norm(designation)
    variants = [designation]

    if norm_unit(unite) in {"l", "litre", "litres"} and any(m in d for m in ("carburant", "diesel", "gazole", "gasoil")):
        variants.extend(["Gazole routier (B10)", "Gazole routier", "Diesel", "Gazole"])
    if norm_unit(unite) == "kwh" and any(m in d for m in ("energie", "énergie", "electric", "électric", "edf", "compteur")):
        variants.extend(["mix electricité", "Electricité France"])
    if norm_unit(unite) == "km" and any(m in d for m in ("voiture", "vehicule", "véhicule", "domicile", "mission")):
        variants.extend(["Voiture - motorisation moyenne - 2018", "Voiture particulière"])

    # dédoublonnage en conservant l'ordre
    out: list[str] = []
    seen: set[str] = set()
    for v in variants:
        key = norm(v)
        if key not in seen:
            seen.add(key)
            out.append(v)
    return out


def chercher_facteur_securise(designation: str, unite: str) -> dict | None:
    fallback_prioritaire = facteur_fallback(designation, unite)
    if fallback_prioritaire and (norm_unit(unite) in {"l", "litre", "litres"} or "dechets verts" in norm(designation) or "déchets verts" in norm(designation)):
        print(f"    [FE fallback prioritaire] {fallback_prioritaire['designation']} : {fallback_prioritaire['valeur']} {fallback_prioritaire['unite']}")
        return fallback_prioritaire

    for des in variantes_designation(designation, unite):
        facteur = gd.get_facteurs(des, unite)
        if facteur and facteur.get("valeur") is not None:
            unite_fe = str(facteur.get("unite", ""))
            if unites_compatibles_local(unite_fe, unite):
                return facteur
            print(f"    [FE rejeté] {facteur.get('designation', des)} | {facteur.get('valeur')} {unite_fe} incompatible avec {unite}")

    fallback = facteur_fallback(designation, unite)
    if fallback:
        print(f"    [FE fallback] {fallback['designation']} : {fallback['valeur']} {fallback['unite']}")
        return fallback
    return None


def source_avec_litres_carburant(data_extraite: list[dict]) -> set[str]:
    sources: set[str] = set()
    for d in data_extraite:
        des = norm(d.get("designation", ""))
        u = norm_unit(str(d.get("unite", "")))
        if u == "l" and any(m in des for m in MOTS_CARBURANT):
            sources.add(str(d.get("source", "")))
    return sources


def calculer_emissions(data_extraite: list[dict]) -> tuple[list[dict], list[dict]]:
    resultats: list[dict] = []
    non_calcules: list[dict] = []
    sources_litres = source_avec_litres_carburant(data_extraite)

    for entry in data_extraite:
        designation = str(entry.get("designation", "")).strip()
        unite = str(entry.get("unite", "")).strip()
        source = str(entry.get("source", ""))
        scope = str(entry.get("scope", ""))
        fiabilite = entry.get("fiabilite", "faible")
        role = str(entry.get("role", ""))

        try:
            q = float(entry.get("quantite", 0))
        except (TypeError, ValueError):
            continue
        if q <= 0:
            continue

        desig_norm = norm(designation)
        unite_norm = norm_unit(unite)

        # Montants Scope 3 : calcul contrôlé uniquement si le poste est catégorisé.
        if unit_is_financial(unite):
            facteur_fin = facteur_financier_direct(designation, unite)
            if scope == "SCOPE_3" and facteur_fin is not None:
                valeur_fin, unite_fin, source_fin = facteur_fin
                montant_eur = q * 1000 if norm_unit(unite) in {"keuro", "k€"} else q
                em_kg = montant_eur * valeur_fin
                print(f"  [OK €] {designation[:45]} → {montant_eur:g} € × {valeur_fin:.6f} = {em_kg:.2f} kgCO2e")
                incert_fin = 80 if "achats materiels non detailles" in norm(designation) or "achats matériels non détaillés" in norm(designation) else 50
                resultats.append(_make_result(source, scope, designation, montant_eur, "€", valeur_fin, unite_fin, em_kg, fiabilite, incert_fin, source_fin))
                continue
            print(f"  [SKIP FIN] {designation} ({unite})")
            continue

        if designation_financiere(designation, unite):
            print(f"  [SKIP FIN] {designation} ({unite})")
            continue

        if unite_non_physique(unite):
            print(f"  [SKIP UNITÉ] {designation} ({unite})")
            continue

        if est_materiau_client(designation) and not role_achat_intrant(entry):
            print(f"  [SKIP CLIENT] {designation}")
            continue

        if est_fret_lourd_hallucine(designation, source):
            print(f"  [SKIP HALLU FRET] {designation} ({source})")
            non_calcules.append(_make_nc(source, scope, designation, q, unite, "Fret lourd non confirmé par le fichier source"))
            continue

        # Si un fichier contient litres carburant + km véhicules, les litres priment.
        if source in sources_litres and unite_norm == "km" and any(m in desig_norm for m in MOTS_VEHICULE):
            print(f"  [SKIP KM] {designation} : litres carburant déjà présents dans {source}")
            continue

        em_calc = kg_deja_calcule(q, unite)
        if em_calc is not None:
            resultats.append(_make_result(source, scope, designation, q, unite, 1.0, unite, em_calc, fiabilite, 0, "déjà calculé"))
            print(f"  [OK déjà] {designation[:45]} → {em_calc:.2f} kgCO2e")
            continue

        if "eau potable" in desig_norm and unite_norm == "m3" and q > 10000:
            print(f"  [SKIP EAU] {designation} : {q} m³ suspect (index/cumul probable)")
            non_calcules.append(_make_nc(source, scope, designation, q, unite, "Eau > 10 000 m³ : index/cumul probable"))
            continue

        print(f"  [FE] {designation[:55]} ({unite})...")
        facteur = chercher_facteur_securise(designation, unite)

        if facteur is None:
            print(f"  [ERREUR] FE introuvable ou incompatible : {designation} ({unite})")
            non_calcules.append(_make_nc(source, scope, designation, q, unite, "FE introuvable ou incompatible"))
            continue

        valeur_fe = facteur.get("valeur")
        unite_fe = str(facteur.get("unite", ""))
        incertitude = facteur.get("incertitude", 0)
        source_fe = facteur.get("source", "Base Carbone")
        designation_fe = facteur.get("designation", "")
        designation_fe_norm = norm(designation_fe)

        # Garde-fou famille : déchets verts ne doivent jamais matcher fioul/carburants.
        if any(m in desig_norm for m in ["dechets verts", "déchets verts"]) and any(m in designation_fe_norm for m in ["fioul", "gazole", "diesel", "carburant"]):
            fallback_dv = facteur_fallback("dechets verts", unite)
            if fallback_dv:
                facteur = fallback_dv
                valeur_fe = facteur.get("valeur")
                unite_fe = str(facteur.get("unite", ""))
                incertitude = facteur.get("incertitude", 40)
                source_fe = facteur.get("source", "fallback contrôlé")
                designation_fe = facteur.get("designation", "Déchets verts - compostage (fallback)")
                designation_fe_norm = norm(designation_fe)
                print("    [FE corrigé] Déchets verts : rejet du match fioul, fallback compostage")
            else:
                non_calcules.append(_make_nc(source, scope, designation, q, unite, "Mauvais FE déchets verts"))
                continue

        try:
            valeur_fe = float(valeur_fe)
        except (TypeError, ValueError):
            non_calcules.append(_make_nc(source, scope, designation, q, unite, f"FE non numérique : {valeur_fe}"))
            continue

        if valeur_fe <= 0:
            non_calcules.append(_make_nc(source, scope, designation, q, unite, f"FE nul ou négatif : {valeur_fe}"))
            continue

        # Garde-fou central : jamais de t.km appliqué à des km sans charge réelle.
        if fe_unit_tkm(unite_fe) and unite_norm == "km":
            print(f"  [SKIP t.km] {designation} : FE en {unite_fe}, quantité en km, charge inconnue")
            non_calcules.append(_make_nc(source, scope, designation, q, unite, "FE t.km incompatible avec km sans tonnage réel"))
            continue

        valeur_conv, unite_conv = convertir_si_necessaire(valeur_fe, unite_fe, unite)
        if valeur_conv is None:
            print(f"  [ERREUR unité] FE={unite_fe} vs quantité={unite} pour {designation}")
            non_calcules.append(_make_nc(source, scope, designation, q, unite, f"Unité incompatible : FE={unite_fe}, qté={unite}"))
            continue

        # Garde-fous de vraisemblance sur les FE unitaires.
        if unite_norm == "kwh" and valeur_conv > 5:
            non_calcules.append(_make_nc(source, scope, designation, q, unite, f"FE électricité aberrant : {valeur_conv} kgCO2e/kWh"))
            print(f"  [SKIP FE] FE électricité aberrant : {valeur_conv}")
            continue
        if unite_norm == "l" and any(m in desig_norm for m in MOTS_CARBURANT) and not (1.0 <= valeur_conv <= 5.0):
            non_calcules.append(_make_nc(source, scope, designation, q, unite, f"FE carburant aberrant : {valeur_conv} kgCO2e/L"))
            print(f"  [SKIP FE] FE carburant aberrant : {valeur_conv}")
            continue
        if unite_norm == "km" and valeur_conv > 1.0 and not any(m in desig_norm for m in MOTS_CONTEXTE_FRET):
            non_calcules.append(_make_nc(source, scope, designation, q, unite, f"FE /km trop élevé pour déplacement non-fret : {valeur_conv}"))
            print(f"  [SKIP FE] FE /km trop élevé hors fret : {valeur_conv}")
            continue

        em_kg = q * valeur_conv
        print(
            f"  [OK] {designation[:40]} → {q:g} {unite} × {valeur_conv:.6f} = {em_kg:.2f} kgCO2e"
            f" | FE: {str(designation_fe)[:45]} ({unite_conv})"
        )

        resultats.append(_make_result(source, scope, designation, q, unite, valeur_conv, unite_conv, em_kg, fiabilite, incertitude, source_fe))

    return resultats, non_calcules


def _make_result(source: str, scope: str, designation: str, quantite: float, unite: str,
                 fe: float, unite_fe: str, em_kg: float, fiabilite: str,
                 incert: Any, source_fe: str) -> dict:
    return {
        "source": source,
        "scope": scope,
        "designation": designation,
        "quantite": round(float(quantite), 4),
        "unite": unite,
        "facteur_emission": round(float(fe), 8),
        "unite_facteur": unite_fe,
        "emissions_kgCO2e": round(float(em_kg), 4),
        "emissions_tCO2e": round(float(em_kg) * KG_TO_T, 6),
        "incertitude": incert,
        "fiabilite": fiabilite,
        "source_facteur": source_fe,
    }


def _make_nc(source: str, scope: str, designation: str, quantite: float, unite: str, raison: str) -> dict:
    return {
        "source": source,
        "scope": scope,
        "designation": designation,
        "quantite": round(float(quantite), 4),
        "unite": unite,
        "raison": raison,
    }


def agreger_par_scope(resultats: list[dict]) -> dict:
    agregat = defaultdict(lambda: {
        "emissions_kgCO2e": 0.0,
        "emissions_tCO2e": 0.0,
        "nb_sources": 0,
        "incertitudes": [],
        "details": [],
    })

    for r in resultats:
        s = r.get("scope", "SCOPE_INCONNU")
        agregat[s]["emissions_kgCO2e"] += float(r["emissions_kgCO2e"])
        agregat[s]["emissions_tCO2e"] += float(r["emissions_tCO2e"])
        agregat[s]["nb_sources"] += 1
        if r.get("incertitude") not in (None, ""):
            try:
                agregat[s]["incertitudes"].append(float(r["incertitude"]))
            except Exception:
                pass
        agregat[s]["details"].append(r)

    bilan: dict = {}
    total_kg = 0.0
    for scope, data in sorted(agregat.items()):
        incert_moy = sum(data["incertitudes"]) / len(data["incertitudes"]) if data["incertitudes"] else 0.0
        bilan[scope] = {
            "emissions_kgCO2e": round(data["emissions_kgCO2e"], 4),
            "emissions_tCO2e": round(data["emissions_tCO2e"], 6),
            "nb_sources": data["nb_sources"],
            "incertitude_moy": round(incert_moy, 2),
            "details": data["details"],
        }
        total_kg += data["emissions_kgCO2e"]

    bilan["TOTAL"] = {
        "emissions_kgCO2e": round(total_kg, 4),
        "emissions_tCO2e": round(total_kg * KG_TO_T, 6),
    }
    return bilan


def calculer_incertitude_bilan(bilan: dict, data_extraite: list[dict], fichiers_attendus: int | None = None,
                               non_calcules: list[dict] | None = None) -> dict:
    incerts = []
    for scope, data in bilan.items():
        if scope == "TOTAL":
            continue
        for d in data.get("details", []):
            try:
                if d.get("incertitude") not in (None, ""):
                    incerts.append(float(d["incertitude"]))
            except Exception:
                pass

    incert_fe = sum(incerts) / len(incerts) if incerts else 0.0

    if fichiers_attendus:
        nb_sources = len({d.get("source") for d in data_extraite if d.get("source")})
        completude = min(nb_sources / fichiers_attendus, 1.0)
        incert_fichiers = (1 - completude) * 100
    else:
        incert_fichiers = 0.0

    # Ajoute une pénalité si beaucoup de données extraites n'ont pas pu être calculées.
    if data_extraite:
        taux_non_calc = len(non_calcules or []) / max(len(data_extraite), 1)
        incert_non_calc = min(taux_non_calc * 50, 30)
    else:
        incert_non_calc = 0.0

    globale = round((incert_fe + incert_fichiers + incert_non_calc) / 3, 2)
    return {
        "incertitude_facteurs_%": round(incert_fe, 2),
        "incertitude_fichiers_%": round(incert_fichiers, 2),
        "incertitude_non_calcule_%": round(incert_non_calc, 2),
        "incertitude_globale_%": globale,
        "niveau_confiance": "haute" if globale < 20 else "moyenne" if globale < 40 else "faible",
    }


def afficher_bilan(bilan: dict, incertitude: dict) -> None:
    print("\n" + "═" * 72)
    print("  BILAN CARBONE — RÉSULTATS")
    print("═" * 72)

    for scope, data in bilan.items():
        if scope == "TOTAL":
            continue
        print(f"\n  {scope}")
        print(f"  {'─' * 48}")
        for d in data.get("details", []):
            print(f"    {d['designation'][:38]:<38} {d['emissions_tCO2e']:>10.4f} tCO2e")
        print(f"  {'SOUS-TOTAL':>48} : {data['emissions_tCO2e']:>10.4f} tCO2e")

    total = bilan.get("TOTAL", {})
    print("\n" + "═" * 72)
    print(f"  TOTAL GÉNÉRAL : {total.get('emissions_tCO2e', 0):.4f} tCO2e")
    print(f"                  {total.get('emissions_kgCO2e', 0):.2f} kgCO2e")
    print("═" * 72)
    print("\n  INCERTITUDE")
    print(f"  Facteurs d'émission : ±{incertitude['incertitude_facteurs_%']}%")
    print(f"  Fichiers manquants  : ±{incertitude['incertitude_fichiers_%']}%")
    print(f"  Données non calculées : ±{incertitude['incertitude_non_calcule_%']}%")
    print(f"  Incertitude globale : ±{incertitude['incertitude_globale_%']}% [{incertitude['niveau_confiance']}]")
    print("═" * 72)


if __name__ == "__main__":
    if os.path.exists("cache_fe.pkl"):
        os.remove("cache_fe.pkl")
        print("[CACHE] Supprimé — recalcul des FE")

    print("Extraction des données...")
    sortie = ev.lancer_le_bilan()
    if isinstance(sortie, tuple):
        data_extraite, nb_fichiers = sortie
    else:
        data_extraite, nb_fichiers = sortie, None

    print(f"\n{len(data_extraite)} valeurs extraites. Calcul des émissions...")
    print("─" * 72)

    resultats, non_calcules = calculer_emissions(data_extraite)

    print(f"\n{len(resultats)} émissions calculées")
    if non_calcules:
        print(f"{len(non_calcules)} valeur(s) non calculée(s) :")
        for nc in non_calcules:
            print(f"  - {nc['designation']} ({nc['unite']}) [{nc['source']}] : {nc['raison']}")

    bilan = agreger_par_scope(resultats)
    incertitude = calculer_incertitude_bilan(bilan, data_extraite, fichiers_attendus=nb_fichiers, non_calcules=non_calcules)
    afficher_bilan(bilan, incertitude)
