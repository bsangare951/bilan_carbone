import getters_data as gd
import os
import pandas as pd
import re
import json
import time
from datetime import datetime
from groq import Groq
from dotenv import load_dotenv


load_dotenv()

# Unités physiques acceptées pour le bilan carbone
UNITES_VALIDES = {
    "L": ["l", "litres", "litre"],
    "m³": ["m3", "m³"],
    "kg": ["kg", "kilogrammes"],
    "t": ["t", "tonnes", "tonne"],
    "kWh": ["kwh", "kilowatt-heure", "kilowattheure"],
    "MWh": ["mwh", "megawatt-heure"],
    "t.km": ["t.km", "tonnes.km", "tonnes-km", "tonne.km", "tkm"],
    "km": ["km", "kilometres", "kilometre"],
    "kgCO2e": ["kgco2e", "kg co2e", "kgco2"]
}

# Colonnes à ignorer car non pertinentes pour le bilan carbone
COLONNES_IGNOREES = [
    "€", "eur", "euros", "m€", "k€", "etp", "%", "taux", "ratio",
    "date", "ref", "reference", "commentaire", "note", "observation",
    "fournisseur", "adresse", "siren", "siret", "prenom",
    "cout", "coût", "prix", "tarif", "montant", "tva", "ht", "ttc"
]

# Fichiers à exclure du traitement
FICHIERS_EXCLUS = [
    "synthese", "recap", "bilan_carbone", "ratios", "utilitaires",
    "export", "matrice", "descriptif", "soc_", "rgpd", "liste",
    "analyse env", "analyse_env", ".~lock", "~$"
]

# Mots-clés indiquant une ligne de total dans un tableau
MOTS_TOTAL = ["total", "sous-total", "cumul", "somme", "total kgs", "total annuel"]

# Mois en français pour la détection de tableaux mensuels
MOIS_FR = ["janvier", "fevrier", "mars", "avril", "mai", "juin",
           "juillet", "aout", "septembre", "octobre", "novembre", "decembre"]

# Unités explicitement invalides pour le bilan carbone
UNITES_INVALIDES = {"kw", "w", "mw", "l/100", "l/100km", "copies", "pages", "nombre", "nb", "impressions"}

# Fichiers qui ne contiennent pas de données carbone exploitables
FICHIERS_NON_CARBONE = [
    "copies", "impression", "imprimante", "suivi copies", "suivi impression"
]



_groq_client = None
_req_minute = 0
_req_jour = 0
_derniere_minute = time.time()
LIMITE_RPM = 25
LIMITE_RPD = 1000


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
- Electricite : entre 1 000 et 500 000 kWh/an
- Carburant voiture : entre 100 et 100 000 L/an
- Dechets bureau : entre 10 et 10 000 kg/an
- Eau : entre 1 et 10 000 m³/an
- Transport : entre 100 et 10 000 000 t.km/an
Si une valeur depasse largement ces seuils, indique fiabilite = "faible".

---

CE QU'IL FAUT IGNORER ABSOLUMENT :

- Montants financiers : euros, HT, TTC, TVA, abonnement, tarif, prix, cout
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

- Retourne uniquement du JSON brut, rien d'autre
- Pas de texte avant le JSON
- Pas de texte apres le JSON
- Pas d'explication, pas de raisonnement, pas de markdown
- Si aucune donnee valide, retourne exactement : []
- Format : [{"designation": "...", "quantite": 0.0, "unite": "...", "fiabilite": "haute|moyenne|faible", "justification": "..."}]"""



def get_groq_client():
    """Retourne le client Groq, en le creant si necessaire."""
    global _groq_client
    if not _groq_client:
        _groq_client = Groq(api_key=os.getenv("GROQ_API_KEY"))
    return _groq_client


def appel_groq_securise(prompt):
    """
    Appelle l'API Groq avec gestion automatique des limites.
    Gere les quotas par minute et par jour pour eviter les erreurs 429.
    """
    global _req_minute, _req_jour, _derniere_minute

    # Reinitialiser le compteur minute si necessaire
    if time.time() - _derniere_minute > 60:
        _req_minute = 0
        _derniere_minute = time.time()

    # Verifier les limites
    if _req_jour >= LIMITE_RPD:
        print(f"  [GROQ] Limite journaliere atteinte ({LIMITE_RPD} requetes)")
        return None

    if _req_minute >= LIMITE_RPM:
        attente = 60 - (time.time() - _derniere_minute)
        print(f"  [GROQ] Pause {attente:.0f}s (limite {LIMITE_RPM} req/min)")
        time.sleep(attente + 1)
        _req_minute = 0
        _derniere_minute = time.time()

    try:
        client = get_groq_client()
        response = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            max_tokens=1000,
            temperature=0,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt}
            ]
        )
        _req_minute += 1
        _req_jour += 1
        print(f"  [GROQ] Requete {_req_jour}/{LIMITE_RPD} ({_req_minute}/{LIMITE_RPM}/min)")
        return response.choices[0].message.content
    except Exception as e:
        print(f"  [GROQ] Erreur : {e}")
        return None


def extract_numeric(value):
    """Extrait une valeur numerique d'une cellule, gere les formats francais."""
    if pd.isna(value):
        return None
    try:
        s = str(value).replace("\xa0", "").replace(" ", "")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        val = float(s)
        return val if not pd.isna(val) else None
    except (ValueError, TypeError):
        return None


def colonne_est_utile(col_name):
    """
    Determine si une colonne peut contenir des donnees carbone.
    Protege les colonnes d'energie et de consommation.
    """
    col_norm = gd.normalize_text(str(col_name))
    
    # Toujours garder les colonnes liees a l'energie ou aux matieres
    if any(mot in col_norm for mot in [
        "kwh", "mwh", "conso", "consommation", "energie", "énergie",
        "heures", "hp", "hc", "index", "factur", "quantite", "qte",
        "volume", "total", "designation", "libelle", "design",
        "dechet", "papier", "carton", "plastique", "metal",
        "distance", "km", "carburant", "gasoil", "essence", "contenu"
    ]):
        return True
    
    # Sinon, filtrer les colonnes non pertinentes
    return not any(ig in col_norm for ig in COLONNES_IGNOREES)


def est_ligne_total(row):
    """Verifie si une ligne contient un mot-cle indiquant un total."""
    for val in row.values:
        if any(mot in gd.normalize_text(str(val)) for mot in MOTS_TOTAL):
            return True
    return False


def est_fichier_copies(filename):
    """Detecte si le fichier est un suivi de copies/impressions."""
    filename_norm = gd.normalize_text(filename)
    return any(mot in filename_norm for mot in FICHIERS_NON_CARBONE)


def smart_header(df_raw):
    """Detecte automatiquement la ligne d'en-tete dans un fichier brut."""
    mots_cles = [
        "quantite", "qte", "volume", "consommation", "total", "unite",
        "designation", "libelle", "nom", "nature", "description",
        "papier", "d3e", "cartouche", "capsule", "plastique",
        "exercice", "annee", "mois", "distance", "tonnes", "kwh",
        "energie", "flux", "donnee", "carburant", "type",
        "facturation", "facturee", "facture", "acheminement",
        "fourniture", "index", "heures creuses", "heures pleines",
        "hc", "hp", "kgs", "dechets", "dechets", "copies",
        "impression", "outillage"
    ]
    
    best_idx, best_score = 0, 0
    for i, row in df_raw.iterrows():
        if i > 25:
            break
        vals = [gd.normalize_text(str(v)) for v in row.values]
        score = sum(1 for v in vals if any(k in v for k in mots_cles))
        score += sum(1 for v in vals if re.search(r'\d{2}-\d{2}', v))
        if score > best_score:
            best_score, best_idx = score, i

    if best_idx > 0:
        print(f"  [HEADER] En-tete detecte : ligne {best_idx} (score: {best_score})")
    else:
        print(f"  [HEADER] Aucun en-tete detecte, utilisation de la ligne 0")
    
    return best_idx



def load_txt_as_df(filepath):
    """Charge un fichier TXT et le convertit en DataFrame."""
    with open(filepath, 'r', encoding='utf-8') as f:
        lignes = f.readlines()
    
    lignes = [l.strip() for l in lignes if l.strip()]
    
    if not lignes:
        return None
    
    return pd.DataFrame(lignes, columns=["contenu"])


def load_file(filepath):
    """Charge un fichier selon son extension (CSV, Excel, TXT)."""
    ext = os.path.splitext(filepath)[1].lower()
    
    try:
        if ext == ".csv":
            df_raw = pd.read_csv(filepath, sep=";", encoding="utf-8-sig",
                                 on_bad_lines='skip', header=None)
            h = smart_header(df_raw)
            return pd.read_csv(filepath, sep=";", encoding="utf-8-sig",
                               on_bad_lines='skip', header=h)
        
        elif ext in [".xlsx", ".xls"]:
            df_raw = pd.read_excel(filepath, header=None)
            h = smart_header(df_raw)
            return pd.read_excel(filepath, header=h)
        
        elif ext == ".txt":
            return load_txt_as_df(filepath)
    
    except Exception as e:
        print(f"  [ERREUR LECTURE] {e}")
    
    return None



def preparer_contexte(df, filename, scope):
    """
    Prepare les donnees avant envoi a Groq.
    - Fichiers texte : envoie le texte brut
    - Fichiers structures : filtre les colonnes inutiles et priorise les totaux
    """
    # Cas d'un fichier texte (une seule colonne "contenu")
    if list(df.columns) == ["contenu"]:
        texte = "\n".join(df["contenu"].astype(str).tolist())
        if len(texte) > 5000:
            texte = texte[:5000] + "\n... (tronque)"
        print(f"  [PREP] Fichier texte : {len(texte)} caracteres")
        print(f"  [PREP] Apercu :")
        for line in texte.split('\n')[:5]:
            print(f"    {line[:150]}")
        return "FICHIER TEXTE - Cherche les donnees carbone (kWh, L, kg, km, t) dans le texte. Priorite aux totaux.", texte
    
    # Cas d'un tableau structure
    cols_utiles = [c for c in df.columns if colonne_est_utile(c)]
    print(f"  [PREP] Colonnes totales : {len(df.columns)} | Colonnes utiles : {len(cols_utiles)}")
    
    # Fallback si aucune colonne utile
    if not cols_utiles:
        cols_utiles = [c for c in df.columns 
                       if not any(ig in gd.normalize_text(str(c)).lower() 
                                  for ig in ['€', 'euros', 'tva', 'siren', 'siret', 'adresse'])]
        print(f"  [PREP] Fallback : {len(cols_utiles)} colonnes")
    
    if not cols_utiles:
        cols_utiles = list(df.columns)
        print(f"  [PREP] Fallback ultime : toutes les colonnes ({len(cols_utiles)})")
    
    df_clean = df[cols_utiles].copy()
    
    if df_clean.empty:
        print(f"  [PREP] DataFrame vide apres filtrage")
        return "AUCUNE DONNEE", ""

    # Recherche des lignes de total
    lignes_total = []
    for _, row in df_clean.iterrows():
        if est_ligne_total(row):
            vals_num = [extract_numeric(v) for v in row.values]
            if any(v and v > 0 for v in vals_num):
                lignes_total.append(row)

    if lignes_total:
        df_contexte = pd.DataFrame(lignes_total)
        contexte_type = "TOTAL DETECTE - Retourne uniquement ces valeurs"
        print(f"  [PREP] {len(lignes_total)} total(aux) trouve(s)")
    else:
        nb = len(df_clean.dropna(how='all'))
        limite = nb if nb < 80 else 60
        df_contexte = df_clean.dropna(how='all').head(limite)
        contexte_type = f"PAS DE TOTAL ({nb} lignes) - Additionne les valeurs de meme designation"
        print(f"  [PREP] {nb} lignes -> {limite} envoyees a Groq")

    contexte_str = df_contexte.to_string(index=False)
    
    print(f"  [PREP] Apercu ({len(contexte_str)} caracteres) :")
    for line in contexte_str.split('\n')[:5]:
        print(f"    {line[:150]}")
    
    return contexte_type, contexte_str



def get_annee_cible():
    """Determine l'exercice comptable en cours."""
    annee = datetime.now().year
    mois = datetime.now().month
    if mois >= 10:
        return f"{annee % 100:02d}-{(annee + 1) % 100:02d}"
    return f"{(annee - 1) % 100:02d}-{annee % 100:02d}"

ANNEE_CIBLE = get_annee_cible()


def extraire_avec_groq(filepath, scope):
    """
    Extrait les donnees carbone d'un fichier via l'API Groq.
    Retourne une liste de dictionnaires normalises.
    """
    filename = os.path.basename(filepath)
    
    # Ignorer les fichiers de copies/impressions
    if est_fichier_copies(filename):
        print(f"  [SKIP] Fichier copies/impressions -> ignore")
        return []
    
    # Charger le fichier
    df = load_file(filepath)
    if df is None or df.empty:
        print(f"  [EXTRACT] Fichier vide ou illisible")
        return []

    # Preparer le contexte
    contexte_type, contexte_str = preparer_contexte(df, filename, scope)
    
    if contexte_str == "" or contexte_str.strip() == "":
        print(f"  [EXTRACT] Contexte vide apres preparation")
        return []

    # Construire le prompt
    prompt = f"""FICHIER : {filename}
SCOPE GHG : {scope}
EXERCICE CIBLE : {ANNEE_CIBLE}
INSTRUCTION : {contexte_type}

DONNEES BRUTES :
{contexte_str}

Applique le protocole d'extraction defini dans tes instructions systeme.
Retourne les donnees CO2 de ce fichier au format JSON."""

    # Appeler Groq
    reponse = appel_groq_securise(prompt)
    if not reponse:
        print(f"  [EXTRACT] Pas de reponse de Groq")
        return []

    # Parser la reponse
    try:
        reponse_clean = re.sub(r'```json|```', '', reponse).strip()
        match = re.search(r'\[.*?\]', reponse_clean, re.DOTALL)
        if match:
            reponse_clean = match.group(0)
        elif reponse_clean == '' or reponse_clean == '[]':
            print(f"  [EXTRACT] Reponse vide")
            return []
        
        data = json.loads(reponse_clean)
        
        if not data:
            print(f"  [EXTRACT] JSON vide")
            return []

        # Agreger les doublons
        agregat = {}
        for item in data:
            if not isinstance(item, dict):
                continue
            
            qte = item.get("quantite")
            unite = str(item.get("unite", "")).strip()
            designation_brute = item.get("designation", filename)
            fiabilite = item.get("fiabilite", "faible")
            justification = item.get("justification", "")

            if not qte or float(qte) <= 0:
                continue

            # Filtrer les unites invalides
            unite_lower = unite.lower().strip()
            if unite_lower in UNITES_INVALIDES:
                print(f"  [FILTRE] Unite invalide : {qte} {unite} ({designation_brute})")
                continue
            
            # Filtrer les designations de copies/impressions
            design_lower = gd.normalize_text(designation_brute)
            if any(mot in design_lower for mot in ["copie", "impression", "imprimante", "page"]):
                print(f"  [FILTRE] Designation copies/impressions : {designation_brute}")
                continue

            # Filtrer les valeurs aberrantes
            if float(qte) > 10_000_000:
                print(f"  [FILTRE] Valeur suspecte : {qte} {unite} ({designation_brute})")
                continue

            # Normaliser la designation
            designation = gd.CORRESPONDANCES.get(
                gd.normalize_text(designation_brute), designation_brute)

            # Agregation par cle (designation + unite)
            cle = (designation, unite.lower())
            if cle in agregat:
                agregat[cle]["quantite"] += float(qte)
                if fiabilite == "faible":
                    agregat[cle]["fiabilite"] = "faible"
            else:
                agregat[cle] = {
                    "source": filename,
                    "scope": scope,
                    "designation": designation,
                    "quantite": float(qte),
                    "unite": unite,
                    "fiabilite": fiabilite,
                    "justification": justification,
                    "est_calcule": False
                }

        results = list(agregat.values())
        for r in results:
            r["quantite"] = round(r["quantite"], 4)
        
        print(f"  [EXTRACT] {len(results)} valeur(s) extraite(s)")
        return results

    except (json.JSONDecodeError, KeyError, ValueError) as e:
        print(f"  [GROQ] JSON invalide : {e}")
        print(f"  [GROQ] Reponse brute : {reponse[:500]}")
        return []


def lancer_le_bilan(base_path="."):
    """
    Fonction principale qui parcourt les dossiers, extrait les donnees
    et retourne une liste de resultats normalises.
    """
    final_data = []
    non_traites = []

    print("=" * 70)
    print(f"BILAN CARBONE - Extraction automatique")
    print(f"Exercice cible : {ANNEE_CIBLE}")
    print(f"Limites Groq   : {LIMITE_RPM} req/min | {LIMITE_RPD} req/jour")
    print("=" * 70)

    # Parcourir les dossiers
    for root, dirs, files in os.walk(base_path):
        folder = os.path.basename(root).upper()
        if "SCOPE" not in folder:
            continue
        scope = folder

        print(f"\n{'─'*70}")
        print(f"Dossier : {scope}")
        print(f"{'─'*70}")

        for f in sorted(files):
            # Filtrer les fichiers temporaires
            if f.startswith("~$") or f.startswith(".~"):
                continue
            
            # Filtrer les fichiers exclus
            if any(exclu in f.lower() for exclu in FICHIERS_EXCLUS):
                print(f"  [SKIP] {f} -> Exclu")
                continue
            
            # Filtrer les extensions non supportees
            if f.lower().endswith((".py", ".pkl", ".jpg", ".jpeg", ".png")):
                continue

            filepath = os.path.join(root, f)
            if not os.path.isfile(filepath):
                continue

            print(f"\n  Traitement : {f}")
            results = extraire_avec_groq(filepath, scope)

            if results:
                print(f"  -> {len(results)} valeur(s) extraite(s) :")
                for r in results:
                    print(f"     {r['designation'][:40]:<40} | {r['quantite']:>10.2f} {r['unite']:<8} | {r['fiabilite']}")
                final_data.extend(results)
            else:
                print(f"  -> Aucune donnee extraite")
                non_traites.append(f)

    print(f"\n{'='*70}")
    print(f"POST-TRAITEMENT")
    print(f"{'='*70}")
    print(f"  Donnees brutes : {len(final_data)}")

    # Deduplication
    seen = set()
    dedup = []
    for d in final_data:
        key = (d["source"], d["designation"], round(d["quantite"], 2), d["unite"].lower())
        if key not in seen:
            seen.add(key)
            dedup.append(d)
    
    nb_doublons = len(final_data) - len(dedup)
    if nb_doublons > 0:
        print(f"  {nb_doublons} doublon(s) supprime(s)")


    print(f"\n{'='*70}")
    print(f"RESUME FINAL")
    print(f"{'='*70}")
    print(f"Total valeurs extraites : {len(dedup)}")
    print(f"Requetes Groq utilisees : {_req_jour}/{LIMITE_RPD}")
    
    if non_traites:
        print(f"\nFichiers sans donnees ({len(non_traites)}) :")
        for f in non_traites:
            print(f"  - {f}")

    for scope in ["SCOPE_1", "SCOPE_2", "SCOPE_3"]:
        scope_data = [d for d in dedup if d["scope"] == scope]
        if scope_data:
            print(f"\n  [{scope}] {len(scope_data)} valeur(s) :")
            for d in scope_data:
                print(f"    {d['source'][:30]:<30} | {d['designation'][:30]:<30} | {d['quantite']:>10.2f} {d['unite']:<8} | {d['fiabilite']}")

    print(f"\n{'='*70}")
    return dedup

if __name__ == "__main__":
    data = lancer_le_bilan()
    
    print(f"\n{'FICHIER':<35} | {'DESIGNATION':<28} | {'QUANTITE':>10} | {'UNITE':<10} | {'FIABILITE':<8} | SCOPE")
    print("-" * 115)
    for d in data:
        print(
            f"{d['source'][:35]:<35} | "
            f"{d['designation'][:28]:<28} | "
            f"{d['quantite']:>10.2f} | "
            f"{d['unite']:<10} | "
            f"{d['fiabilite']:<8} | "
            f"{d['scope']}"
        )