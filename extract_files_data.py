import getters_data as gd
import os
import re
import json
import time
import pandas as pd
from datetime import datetime
from groq import Groq
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# CONSTANTES
# ---------------------------------------------------------------------------

UNITES_VALIDES = {
    "L":      ["l", "litres", "litre"],
    "m³":     ["m3", "m³"],
    "kg":     ["kg", "kilogrammes"],
    "t":      ["t", "tonnes", "tonne"],
    "kWh":    ["kwh", "kilowatt-heure", "kilowattheure"],
    "MWh":    ["mwh", "megawatt-heure"],
    "t.km":   ["t.km", "tonnes.km", "tonnes-km", "tonne.km", "tkm"],
    "km":     ["km", "kilometres", "kilometre"],
    "kgCO2e": ["kgco2e", "kg co2e", "kgco2"],
}

# Index inversé pré-calculé : alias normalisé -> unité canonique
_UNITE_LOOKUP: dict[str, str] = {}
for _canon, _aliases in UNITES_VALIDES.items():
    _UNITE_LOOKUP[gd.normalize_text(_canon)] = _canon
    for _a in _aliases:
        _UNITE_LOOKUP[gd.normalize_text(_a)] = _canon

COLONNES_IGNOREES = frozenset([
    "€", "eur", "euros", "m€", "k€", "etp", "%", "taux", "ratio",
    "date", "ref", "reference", "commentaire", "note", "observation",
    "fournisseur", "adresse", "siren", "siret", "prenom",
    "cout", "coût", "prix", "tarif", "montant", "tva", "ht", "ttc",
])

FICHIERS_EXCLUS = frozenset([
    "synthese", "bilan_carbone", "ratios", "utilitaires",
    "matrice", "descriptif", "soc_", "rgpd",
    "analyse env", "analyse_env", ".~lock", "~$",
    "_in",  # doublons de relevés compteurs
    # Note : "liste", "export", "recap" retirés car peuvent contenir
    # des données carbone exploitables (conso véhicules, données utilisées...)
])

MOTS_TOTAL_RE = re.compile(
    r"\b(total|sous-total|cumul|somme|total kgs|total annuel)\b"
)

FICHIERS_NON_CARBONE = frozenset(["copies", "impression", "imprimante"])

# Mots-clés indiquant un suivi kilométrique de flotte de véhicules légers
# -> la désignation doit être forcée à "Voiture - motorisation moyenne - 2018"
MOTS_FLOTTE_VEHICULES = frozenset([
    "km annuel", "km par vehicule", "kilometrage", "odometre",
    "conducteur", "suivi km", "releve km", "flotte",
])

# Mots-clés dans le contenu texte indiquant des relevés odomètre
MOTS_CONTENU_FLOTTE = frozenset([
    "conducteur", "n° int", "totaux", "remplacement",
    "total annuel", "total mensuel",
])

UNITES_INVALIDES = frozenset([
    "kw", "w", "mw", "l/100", "l/100km",
    "copies", "pages", "nombre", "nb", "impressions",
])

COLONNES_UTILES_KW = frozenset([
    "kwh", "mwh", "conso", "consommation", "energie", "énergie",
    "heures", "hp", "hc", "index", "factur", "quantite", "qte",
    "volume", "total", "designation", "libelle", "design",
    "dechet", "papier", "carton", "plastique", "metal",
    "distance", "km", "carburant", "gasoil", "essence", "contenu",
    "gnr", "diesel", "litres", "litre", "conso l", "conso litres",
    "carburant consomme", "consommation carburant",
])

DESIGNATIONS_COPIES = frozenset(["copie", "impression", "imprimante", "page"])

MOTS_CLES_HEADER = frozenset([
    "quantite", "qte", "volume", "consommation", "total", "unite",
    "designation", "libelle", "nom", "nature", "description",
    "papier", "d3e", "cartouche", "capsule", "plastique",
    "exercice", "annee", "mois", "distance", "tonnes", "kwh",
    "energie", "flux", "donnee", "carburant", "type",
    "facturation", "facturee", "facture", "acheminement",
    "fourniture", "index", "heures creuses", "heures pleines",
    "hc", "hp", "kgs", "dechets", "copies", "impression", "outillage",
])

LIMITE_RPM = 25
LIMITE_RPD = 1_000

# ---------------------------------------------------------------------------
# SYSTEM PROMPT
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """Tu es une API d'extraction de données carbone. Tu ne fournis aucune explication, aucun raisonnement, aucun commentaire. Tu retournes uniquement du JSON valide.

Ta mission : analyser des données brutes de fichiers d'entreprise et extraire uniquement les grandeurs physiques necessaires au calcul d'emissions de CO2.

---

GRANDEURS ACCEPTEES (les seules a retourner) :

- Energie       : kWh, MWh
- Carburants    : L (litres), m³
- Masses        : kg, t (tonnes)
- Transport     : km, t.km (tonnes.kilometres)
- Eau           : L, m³
- Emissions     : kgCO2e (si deja calculees)

---

REGLE 1 - PRIORITE ABSOLUE AUX TOTAUX

Si une ligne ou colonne contient "TOTAL", "Total", "total", "TOTAL 2025", "Total general",
"Total annuel", "Sous-total", "Cumul", "Somme", "Total KGS", "Total kgs" avec une valeur > 0,
alors extrais uniquement cette valeur. N'extrais pas les lignes de detail.

Quelques exemples concrets :
- 12 lignes mensuelles (179, 158, 180...) + 1 ligne "TOTAL 2025 | 2417" -> retourne "Electricite: 2417 kWh"
- 20 lignes de dechets detailles + 1 ligne "Total kgs: 47.50" -> retourne "Total dechets: 47.50 kg"
- 12 lignes de conso + 1 ligne "Cumul annuel: 15000" -> retourne "Carburant: 15000 L"
- Une colonne "Total" dans un tableau -> utilise uniquement cette colonne, ignore les colonnes mensuelles

Si aucun total n'existe, alors seulement tu additionnes toutes les valeurs de meme designation.

---

REGLE 2 - NE JAMAIS PRENDRE UNE VALEUR AU HASARD

Ne prends jamais une ligne isolee si un total existe.
Ne prends jamais un mois au hasard dans un tableau mensuel.
Si tu vois 12 mois + un total, prends le total, pas le mois de janvier ni decembre.
Si tu ne trouves pas de total, additionne tous les mois, ne prends pas un seul mois.

---

REGLE 3 - EAU ET BONBONNES : LITRES, PAS KILOGRAMMES

Pour l'eau potable, les bonbonnes d'eau, les fontaines a eau :
- Extrais le volume en litres (L) ou m³, jamais le poids en kg.
- Une bonbonne de 18 L -> c'est 18 L d'eau, pas 18 kg.
- 6 bonbonnes de 18 L -> 108 L d'eau, pas 108 kg.
- Ignore le "poids total du colis" ou "poids total" si le fichier parle d'eau ou de bonbonnes.
- Le poids d'expedition n'est pas une consommation d'eau.

Exemples :
- "BONBONNE D'EAU 18 L x6" + "poids total : 108 kg" -> extraire "Eau potable: 108 L"
- "EAU DE SOURCE 18L" -> extraire "Eau potable: 18 L"
- "poids total de : 108,00 kg" dans un fichier d'eau -> ignorer, ce n'est pas la consommation d'eau

---

REGLE 4 - DONNEES LES PLUS RECENTES

Si plusieurs exercices (ex: 23-24 et 24-25), prends uniquement les donnees de l'exercice le plus recent.

---

REGLE 5 - COHERENCE DES VALEURS

Verifie que les valeurs sont realistes pour une TPE/PME francaise :
- Electricite : entre 100 et 500 000 kWh/an
- Carburant voiture : entre 100 et 100 000 L/an
- Carburant camions/poids lourds : entre 1 000 et 500 000 L/an
- Dechets bureau : entre 10 et 10 000 kg/an
- Eau : entre 1 et 10 000 m³/an
- Transport : entre 100 et 10 000 000 t.km/an
Si une valeur depasse largement ces seuils, indique fiabilite = "faible".

---

REGLE 6 - NE PAS EXTRAIRE LES MATERIAUX COLLECTES POUR DES CLIENTS

Pour les entreprises de collecte de dechets, travaux publics, BTP, environnement :
- Les materiaux collectes (terres, gravats, beton, sable, calcaire, souches...)
  sont l'ACTIVITE du client, pas les emissions propres de l'entreprise.
- NE PAS extraire : terres inertes, gravats, blocs beton, sable, calcaire,
  souches, remblais, deblais, materiaux de chantier collectes pour des tiers.
- EXTRAIRE uniquement : carburant consomme (L), electricite (kWh),
  dechets produits par l'entreprise elle-meme, deplacements des salaries.
- Si le fichier contient un tableau recapitulatif de collectes client avec
  des materiaux en tonnes ou m3 -> retourner [].

---

REGLE 7 - CARBURANT CAMIONS LOURDS

Pour les fichiers de suivi de flotte de camions/poids lourds :
- Extraire les litres de gazole/GNR consommes (L).
- Si le fichier contient a la fois des km et des litres pour le meme vehicule,
  prendre UNIQUEMENT les litres (plus precis).
- Les litres de carburant sont dans les colonnes : conso, consommation,
  litres, carburant, gazole, GNR, diesel.

---

CE QU'IL FAUT IGNORER ABSOLUMENT :

- Montants financiers : euros, € HT, TTC, TVA, abonnement, tarif, prix, coûts
- Puissances : kW, MW (different de l'energie consommee)
- Ratios : L/100km, L/100, consommation aux 100km
- Compteurs impression : nombre de copies, pages imprimees, impressions NB, impressions couleur
- Effectifs : ETP, nombre de salaries, personnes
- Administratif : dates, references, SIREN, adresses
- Objectifs de reduction, pourcentages d'evolution
- Poids de colis/expedition quand il s'agit d'eau ou de livraison
- Doublons : si deux colonnes representent la meme energie, prends la plus complete

---

TYPES DE FICHIERS ET COMMENT LES TRAITER :

- Facture EDF / Electricite / Energie :
  Cherche une ligne "TOTAL 2025" ou "Total" avec des kWh.
  Si elle existe, prends uniquement cette ligne.
  Sinon, additionne tous les mois. Ne prends jamais un seul mois.

- Suivi dechets :
  Cherche les kg par categorie (papier, D3E, plastique...).
  Si une ligne "Total KGS" ou "Total kgs" existe, prends uniquement cette ligne.

- Transport / Deplacements :
  Cherche les km parcourus ou t.km, ignore les couts et indemnites.
  Si une ligne "Total km" existe, prends-la.

- Carburant :
  Cherche les litres consommes, ignore les euros/L et L/100km.
  Si une ligne "Total" ou "Cumul" existe, prends-la.

- Eau / Bonbonnes d'eau :
  Extrais le volume en litres (L) ou m³.
  Multiplie le nombre de bonbonnes par leur contenance.
  Ignore le poids du colis en kg.

- Suivi copies / Impressions :
  Ignore complement : un nombre de copies n'est pas une masse en kg,
  ce n'est pas une donnee carbone. Meme si la colonne s'appelle
  "Consommation totale", si le fichier parle d'impressions, retourne [].

- Fichier outillage :
  Cherche les kg ou L de matieres consommees, ignore les nombres d'unites.

- Fichier texte (TXT) :
  Cherche dans le texte toute mention de kWh, L, kg, km, t avec des valeurs numeriques.
  Priorite aux totaux si presents. Pour l'eau, convertir en litres.

---

FORMAT DE REPONSE - REGLES ABSOLUES :
- Si le fichier contient plusieurs années, prends SEULEMENT la plus récente
- NE PAS additionner des relevés historiques

- Retourne uniquement du JSON brut, rien d'autre
- Pas de texte avant le JSON
- Pas de texte apres le JSON
- Pas d'explication, pas de raisonnement, pas de markdown
- Si aucune donnee valide, retourne exactement : []
- Format : [{"designation": "...", "quantite": 0.0, "unite": "...", "fiabilite": "haute|moyenne|faible", "justification": "..."}]"""


_groq_client: Groq | None = None
_req_minute = 0
_req_jour = 0
_derniere_minute = time.time()


def get_groq_client() -> Groq:
    global _groq_client
    if _groq_client is None:
        _groq_client = Groq(api_key=os.getenv("GROQ_API_KEY"))
    return _groq_client


def appel_groq_securise(prompt: str) -> str | None:
    global _req_minute, _req_jour, _derniere_minute

    elapsed = time.time() - _derniere_minute
    if elapsed > 60:
        _req_minute = 0
        _derniere_minute = time.time()

    if _req_jour >= LIMITE_RPD:
        print(f"  [GROQ] Limite journalière atteinte ({LIMITE_RPD} requêtes)")
        return None

    if _req_minute >= LIMITE_RPM:
        attente = 60 - (time.time() - _derniere_minute)
        print(f"  [GROQ] Pause {attente:.0f}s (limite {LIMITE_RPM} req/min)")
        time.sleep(attente + 1)
        _req_minute = 0
        _derniere_minute = time.time()

    try:
        response = get_groq_client().chat.completions.create(
            model="llama-3.1-8b-instant",
            max_tokens=1_000,
            temperature=0,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user",   "content": prompt},
            ],
        )
        _req_minute += 1
        _req_jour += 1
        print(f"  [GROQ] Requête {_req_jour}/{LIMITE_RPD} ({_req_minute}/{LIMITE_RPM}/min)")
        return response.choices[0].message.content
    except Exception as e:
        print(f"  [GROQ] Erreur : {e}")
        return None



def extract_numeric(value) -> float | None:
    if pd.isna(value):
        return None
    s = str(value).strip().replace("\xa0", "").replace(" ", "")
    if not s:
        return None

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")

    try:
        val = float(s)
        return None if pd.isna(val) else val
    except (ValueError, TypeError):
        return None


def normaliser_unite(unite) -> str:
    if not unite:
        return ""
    return _UNITE_LOOKUP.get(gd.normalize_text(str(unite)), str(unite).strip())


def colonne_est_utile(col_name: str) -> bool:
    col_norm = gd.normalize_text(str(col_name))
    if any(kw in col_norm for kw in COLONNES_UTILES_KW):
        return True
    return not any(ig in col_norm for ig in COLONNES_IGNOREES)


def est_ligne_total(row) -> bool:
    return any(MOTS_TOTAL_RE.search(gd.normalize_text(str(v))) for v in row.values)


def est_fichier_copies(filename: str) -> bool:
    fn = gd.normalize_text(filename)
    return any(mot in fn for mot in FICHIERS_NON_CARBONE)


def est_fichier_flotte(filename: str, texte: str = "") -> bool:
    """Détecte si le fichier est un suivi kilométrique de flotte de véhicules légers.
    Dans ce cas la désignation doit être forcée à Voiture motorisation moyenne."""
    fn = gd.normalize_text(filename)
    if any(mot in fn for mot in MOTS_FLOTTE_VEHICULES):
        return True
    # Détection sur le contenu texte
    if texte:
        texte_norm = gd.normalize_text(texte)
        hits = sum(1 for mot in MOTS_CONTENU_FLOTTE if mot in texte_norm)
        return hits >= 2
    return False


def smart_header(df_raw: pd.DataFrame) -> int:
    best_idx, best_score = 0, 0
    for i, row in df_raw.iterrows():
        if i > 25:
            break
        vals = [gd.normalize_text(str(v)) for v in row.values]
        score = sum(1 for v in vals if any(k in v for k in MOTS_CLES_HEADER))
        score += sum(1 for v in vals if re.search(r"\d{2}-\d{2}", v))
        if score > best_score:
            best_score, best_idx = score, i

    label = f"ligne {best_idx} (score: {best_score})" if best_idx > 0 else "ligne 0 (défaut)"
    print(f"  [HEADER] En-tête détecté : {label}")
    return best_idx


def get_annee_cible() -> str:
    now = datetime.now()
    y, m = now.year, now.month
    if m >= 10:
        return f"{y % 100:02d}-{(y + 1) % 100:02d}"
    return f"{(y - 1) % 100:02d}-{y % 100:02d}"


ANNEE_CIBLE = get_annee_cible()


def load_txt_as_df(filepath: str) -> pd.DataFrame | None:
    with open(filepath, encoding="utf-8") as f:
        lignes = [l.strip() for l in f if l.strip()]
    return pd.DataFrame(lignes, columns=["contenu"]) if lignes else None


def load_file(filepath: str) -> pd.DataFrame | None:
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == ".csv":
            for enc in ("utf-8-sig", "latin-1"):
                try:
                    df_raw = pd.read_csv(filepath, sep=";", encoding=enc,
                                         on_bad_lines="skip", header=None)
                    h = smart_header(df_raw)
                    return pd.read_csv(filepath, sep=";", encoding=enc,
                                       on_bad_lines="skip", header=h)
                except UnicodeDecodeError:
                    continue

        elif ext in (".xlsx", ".xls"):
            df_raw = pd.read_excel(filepath, header=None)
            h = smart_header(df_raw)
            return pd.read_excel(filepath, header=h)

        elif ext == ".txt":
            return load_txt_as_df(filepath)

    except Exception as e:
        print(f"  [ERREUR LECTURE] {e}")

    return None


def preparer_contexte(df: pd.DataFrame, filename: str, scope: str) -> tuple[str, str]:
    # Fichier texte brut
    if list(df.columns) == ["contenu"]:
        texte = "\n".join(df["contenu"].astype(str))
        if len(texte) > 5_000:
            texte = texte[:5_000] + "\n... (tronqué)"
        print(f"  [PREP] Fichier texte : {len(texte)} caractères")
        for line in texte.split("\n")[:5]:
            print(f"    {line[:150]}")
        return "FICHIER TEXTE - Cherche les données carbone (kWh, L, kg, km, t). Priorité aux totaux.", texte

    # Tableau structuré sélection des colonnes utiles
    cols_utiles = [c for c in df.columns if colonne_est_utile(c)]
    print(f"  [PREP] Colonnes : {len(df.columns)} totales | {len(cols_utiles)} utiles")

    if not cols_utiles:
        cols_utiles = [
            c for c in df.columns
            if not any(ig in gd.normalize_text(str(c)) for ig in ("€", "euros", "tva"))
        ]
        print(f"  [PREP] Fallback : {len(cols_utiles)} colonnes")

    if not cols_utiles:
        cols_utiles = list(df.columns)
        print(f"  [PREP] Fallback ultime : toutes les colonnes ({len(cols_utiles)})")

    df_clean = df[cols_utiles].dropna(how="all")
    if df_clean.empty:
        print("  [PREP] DataFrame vide après filtrage")
        return "AUCUNE DONNEE", ""

    # Lignes de total
    lignes_total = [
        row for _, row in df_clean.iterrows()
        if est_ligne_total(row) and any(
            (v := extract_numeric(cell)) and v > 0 for cell in row.values
        )
    ]

    if lignes_total:
        df_contexte = pd.DataFrame(lignes_total)
        contexte_type = "TOTAL DETECTÉ - Retourne uniquement ces valeurs"
        print(f"  [PREP] {len(lignes_total)} total(aux) trouvé(s)")
    else:
        nb = len(df_clean)
        limite = nb if nb < 80 else 60
        df_contexte = df_clean.head(limite)
        contexte_type = f"PAS DE TOTAL ({nb} lignes) - Additionne les valeurs de même désignation"
        print(f"  [PREP] {nb} lignes -> {limite} envoyées à Groq")

    contexte_str = df_contexte.to_string(index=False)
    print(f"  [PREP] Aperçu ({len(contexte_str)} caractères) :")
    for line in contexte_str.split("\n")[:5]:
        print(f"    {line[:150]}")

    return contexte_type, contexte_str


def extraire_avec_groq(filepath: str, scope: str) -> list[dict]:
    filename = os.path.basename(filepath)

    if est_fichier_copies(filename):
        print("  [SKIP] Fichier copies/impressions -> ignoré")
        return []

    df = load_file(filepath)
    if df is None or df.empty:
        print("  [EXTRACT] Fichier vide ou illisible")
        return []

    contexte_type, contexte_str = preparer_contexte(df, filename, scope)
    if not contexte_str.strip():
        print("  [EXTRACT] Contexte vide après préparation")
        return []


    # Détection flotte véhicules légers 

    is_flotte = est_fichier_flotte(filename, contexte_str)
    if is_flotte:
        # Détecter si c'est une flotte légère ou lourde
        contexte_low = contexte_str.lower()
        is_poids_lourd = any(m in contexte_low for m in (
            "poids lourd", "pl ", "camion", "benne", "porteur", "semi",
            "articulé", "articule", "gnr", "tracteur"
        ))
        if is_poids_lourd:
            print("  [FLOTTE PL] Flotte poids lourds détectée -> extraction litres prioritaire")
            instruction_flotte = (
                "IMPORTANT : Ce fichier contient des données de flotte de camions/poids lourds. "
                "Extraire UNIQUEMENT les litres de carburant consommés (gazole, GNR, diesel). "
                "Si les km et les litres sont tous les deux présents, retourner SEULEMENT les litres. "
                "Désignation : 'Gazole routier (B10)' ou 'Gazole non routier' selon le type."
            )
        else:
            print("  [FLOTTE VL] Suivi kilométrique véhicules légers -> désignation forcée : Voiture")
            instruction_flotte = (
                "IMPORTANT : Ce fichier contient des relevés kilométriques de véhicules "
                "légers de société (voitures, utilitaires <3.5T). "
                'La désignation DOIT être "Voiture - motorisation moyenne - 2018". '
                "Ne jamais retourner Articulé, camion ou poids lourd."
            )
        contexte_type = f"{contexte_type} | {instruction_flotte}"

    prompt = (
        f"FICHIER : {filename}\n"
        f"SCOPE GHG : {scope}\n"
        f"EXERCICE CIBLE : {ANNEE_CIBLE}\n"
        f"INSTRUCTION : {contexte_type}\n\n"
        f"DONNEES BRUTES :\n{contexte_str}\n\n"
        "Applique le protocole d'extraction défini dans tes instructions système.\n"
        "Retourne les données CO2 de ce fichier au format JSON."
    )

    reponse = appel_groq_securise(prompt)
    if not reponse:
        print("  [EXTRACT] Pas de réponse de Groq")
        return []

    try:
        reponse_clean = re.sub(r"```json|```", "", reponse).strip()
        match = re.search(r"\[.*?\]", reponse_clean, re.DOTALL)
        if not match:
            print("  [EXTRACT] Réponse vide ou non JSON")
            return []
        data = json.loads(match.group(0))
    except json.JSONDecodeError as e:
        print(f"  [GROQ] JSON invalide : {e}")
        print(f"  [GROQ] Réponse brute : {reponse[:500]}")
        return []

    if not data:
        print("  [EXTRACT] JSON vide")
        return []

    agregat: dict[tuple, dict] = {}
    for item in data:
        if not isinstance(item, dict):
            continue

        qte_raw = item.get("quantite")
        try:
            qte = float(qte_raw)
        except (TypeError, ValueError):
            continue
        if qte <= 0:
            continue

        unite = str(item.get("unite", "")).strip()
        designation_brute = item.get("designation", filename)
        fiabilite = item.get("fiabilite", "faible")
        justification = item.get("justification", "")
        unite_low = unite.lower().strip()
        desig_low = gd.normalize_text(designation_brute)
        if "feuille" in unite_low or "unite-feuille" in unite_low:
            if "a3" in desig_low:
                qte = qte * 0.010  # 10g/feuille A3
            else:
                qte = qte * 0.005  # 5g/feuille A4 par défaut
            unite = "kg"
            print(f"  [CONV PAPIER] {designation_brute} : feuilles -> {qte:.2f} kg")
        elif "rouleau" in unite_low:
            qte = qte * 0.5  # ~500g/rouleau traceur
            unite = "kg"
            print(f"  [CONV PAPIER] {designation_brute} : rouleaux -> {qte:.2f} kg")
        desig_lower_check = gd.normalize_text(designation_brute)
        if unite.lower() == "kg" and any(
            mot in desig_lower_check for mot in ("eau", "water", "bonbonne", "fontaine", "reseau")
        ):
            unite = "L"
            print(f"  [FIX EAU] {designation_brute} : kg -> L (eau volumétrique)")

        if unite.lower() in UNITES_INVALIDES:
            print(f"  [FILTRE] Unité invalide : {qte} {unite} ({designation_brute})")
            continue

        design_norm = gd.normalize_text(designation_brute)
        if any(mot in design_norm for mot in DESIGNATIONS_COPIES):
            print(f"  [FILTRE] Désignation copies/impressions : {designation_brute}")
            continue

        if qte > 10_000_000:
            print(f"  [FILTRE] Valeur suspecte : {qte} {unite} ({designation_brute})")
            continue

        designation = gd.CORRESPONDANCES.get(design_norm, designation_brute)
        cle = (designation, unite.lower())

        if cle in agregat:
            agregat[cle]["quantite"] += qte
            if fiabilite == "faible":
                agregat[cle]["fiabilite"] = "faible"
        else:
            agregat[cle] = {
                "source":      filename,
                "scope":       scope,
                "designation": designation,
                "quantite":    qte,
                "unite":       unite,
                "fiabilite":   fiabilite,
                "justification": justification,
                "est_calcule": False,
            }

    results = [
        {**r, "quantite": round(r["quantite"], 4)}
        for r in agregat.values()
    ]

    unites_volume = {"l", "litre", "litres", "m3", "m³"}
    desigs_avec_litres = {
        gd.normalize_text(r["designation"])
        for r in results
        if r["unite"].lower().strip() in unites_volume
    }
    avant = len(results)
    results = [
        r for r in results
        if not (
            r["unite"].lower().strip() == "km"
            and any(
                mot in gd.normalize_text(r["designation"])
                for mot in desigs_avec_litres
                if mot  # éviter les chaînes vides
            )
        )
    ]

    MOTS_CARBURANT = {"gazole", "diesel", "essence", "carburant", "gnr", "b10", "fioul"}
    a_des_litres_carburant = any(
        any(m in gd.normalize_text(r["designation"]) for m in MOTS_CARBURANT)
        and r["unite"].lower().strip() in unites_volume
        for r in results
    )
    if a_des_litres_carburant:
        results = [
            r for r in results
            if not (
                r["unite"].lower().strip() == "km"
                and any(m in gd.normalize_text(r["designation"])
                        for m in {"articul", "camion", "poids lourd", "vehicule"})
            )
        ]

    nb_supprimes = avant - len(results)
    if nb_supprimes:
        print(f"  [DEDUP L/KM] {nb_supprimes} entrée(s) km supprimée(s) "
              f"(litres disponibles pour le même véhicule)")

    print(f"  [EXTRACT] {len(results)} valeur(s) extraite(s)")
    return results


def lancer_le_bilan(base_path: str = ".") -> list[dict]:
    final_data: list[dict] = []
    non_traites: list[str] = []
    nb_fichiers_tentes: int = 0

    print("=" * 70)
    print("BILAN CARBONE - Extraction automatique")
    print(f"Exercice cible : {ANNEE_CIBLE}")
    print(f"Limites Groq   : {LIMITE_RPM} req/min | {LIMITE_RPD} req/jour")
    print("=" * 70)

    for root, _dirs, files in os.walk(base_path):
        folder = os.path.basename(root).upper()
        if "SCOPE" not in folder:
            continue
        scope = folder

        print(f"\n{'─' * 70}")
        print(f"Dossier : {scope}")
        print(f"{'─' * 70}")

        for f in sorted(files):
            if f.startswith(("~$", ".~")):
                continue
            if any(exclu in f.lower() for exclu in FICHIERS_EXCLUS):
                print(f"  [SKIP] {f} -> Exclu")
                continue
            if f.lower().endswith((".py", ".pkl", ".jpg", ".jpeg", ".png")):
                continue

            filepath = os.path.join(root, f)
            if not os.path.isfile(filepath):
                continue

            print(f"\n  Traitement : {f}")
            nb_fichiers_tentes += 1
            results = extraire_avec_groq(filepath, scope)

            if results:
                print(f"  -> {len(results)} valeur(s) extraite(s) :")
                for r in results:
                    print(f"     {r['designation'][:40]:<40} | {r['quantite']:>10.2f} {r['unite']:<8} | {r['fiabilite']}")
                final_data.extend(results)
            else:
                print("  -> Aucune donnée extraite")
                non_traites.append(f)

    print(f"\n{'=' * 70}")
    print("POST-TRAITEMENT")
    print(f"{'=' * 70}")
    print(f"  Données brutes : {len(final_data)}")

    seen: set[tuple] = set()
    dedup_primary: list[dict] = []
    for d in final_data:
        key = (d["source"], d["designation"], round(d["quantite"], 2), d["unite"].lower())
        if key not in seen:
            seen.add(key)
            dedup_primary.append(d)
    _fiab_rank = {"haute": 2, "moyenne": 1, "faible": 0}
    seen2: dict[tuple, dict] = {}
    for d in dedup_primary:
        key2 = (d["designation"], round(d["quantite"], 2), d["unite"].lower())
        if key2 not in seen2:
            seen2[key2] = d
        else:
            # Garder la version la plus fiable
            existing_rank = _fiab_rank.get(seen2[key2].get("fiabilite", "faible"), 0)
            new_rank      = _fiab_rank.get(d.get("fiabilite", "faible"), 0)
            if new_rank > existing_rank:
                seen2[key2] = d
    dedup = list(seen2.values())

    nb_doublons = len(final_data) - len(dedup)
    if nb_doublons:
        print(f"  {nb_doublons} doublon(s) supprimé(s)")

    print(f"\n{'=' * 70}")
    print("RÉSUMÉ FINAL")
    print(f"{'=' * 70}")
    print(f"Total valeurs extraites : {len(dedup)}")
    print(f"Requêtes Groq utilisées : {_req_jour}/{LIMITE_RPD}")

    if non_traites:
        print(f"\nFichiers sans données ({len(non_traites)}) :")
        for f in non_traites:
            print(f"  - {f}")

    for scope in ("SCOPE_1", "SCOPE_2", "SCOPE_3"):
        scope_data = [d for d in dedup if d["scope"] == scope]
        if scope_data:
            print(f"\n  [{scope}] {len(scope_data)} valeur(s) :")
            for d in scope_data:
                print(
                    f"    {d['source'][:30]:<30} | {d['designation'][:30]:<30} "
                    f"| {d['quantite']:>10.2f} {d['unite']:<8} | {d['fiabilite']}"
                )

    print(f"\n{'=' * 70}")
    return dedup, nb_fichiers_tentes