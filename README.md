# Outil Bilan Carbone Automatisé

## Présentation

Cet outil permet d’automatiser une partie du traitement des dossiers clients dans le cadre d’un pré-bilan carbone.

Il permet de charger un dossier contenant différents types de fichiers, de nettoyer les données, de classer les documents par scope, d’extraire les informations utiles, puis de calculer les émissions de gaz à effet de serre en kgCO₂e et en tCO₂e.

L’outil a été conçu pour faciliter le travail d’analyse, mais il ne remplace pas une validation humaine. Les résultats doivent être vérifiés à l’aide des fichiers d’audit générés.

## Fonctionnalités principales

* Chargement de fichiers clients hétérogènes
* Nettoyage automatique des fichiers
* Classification des documents par scope
* Extraction des données utiles au bilan carbone
* Recherche des facteurs d’émission
* Calcul des émissions par scope
* Gestion des données physiques et financières
* Génération de fichiers d’audit
* Interface graphique simple d’utilisation

## Formats de fichiers pris en charge

L’outil peut traiter plusieurs types de fichiers :

* Excel : `.xlsx`, `.xls`
* CSV : `.csv`
* Texte : `.txt`
* PDF : `.pdf`
* Word : `.docx`
* Images : `.jpg`, `.jpeg`, `.png` si l’OCR est disponible

Les fichiers Excel, CSV et PDF numériques donnent généralement de meilleurs résultats que les documents scannés ou les photos.

## Structure du dossier de l’application

Le dossier de l’application doit conserver cette organisation :

```text
BilanCarbone
│
├── Bilan_Carbone.exe
├── Bilan_Carbone_V9.01.xlsx
├── fond.jpg
├── .env
└── _internal
```

Le dossier `_internal` ne doit pas être supprimé, renommé ou déplacé. Il contient les fichiers nécessaires au fonctionnement de l’exécutable.

## Ressources importantes

### Fichier des facteurs d’émission

Le fichier suivant est indispensable :

```text
Bilan_Carbone_V9.01.xlsx
```

Il contient les facteurs d’émission utilisés pour les calculs. Si une nouvelle version est fournie, il faut remplacer l’ancien fichier en gardant le même nom.

### Fichier d’environnement

Le fichier suivant contient les variables d’environnement :

```text
.env
```

Il peut contenir des clés API, notamment pour Groq ou l’ADEME. Ce fichier est confidentiel et ne doit pas être partagé publiquement.

### Image de fond

Le fichier suivant est utilisé pour l’affichage de l’interface :

```text
fond.jpg
```

Il n’est pas lié au calcul, mais il est nécessaire pour conserver l’apparence prévue de l’outil.

## Installation

L’outil est fourni sous forme d’un dossier compressé.

Étapes d’installation :

1. Télécharger le fichier `.zip`
2. Extraire le dossier à l’emplacement souhaité
3. Vérifier que les fichiers nécessaires sont présents
4. Lancer l’application avec `Bilan_Carbone.exe`

Lors du premier lancement, l’ouverture peut prendre quelques minutes. Cela est normal.

## Prérequis

Pour utiliser directement l’exécutable, Python n’est normalement pas nécessaire.

Python est uniquement nécessaire si l’on souhaite modifier le code source, lancer les scripts manuellement ou recompiler l’application.

Version recommandée :

```text
Python 3.11 ou supérieur
```

Pour améliorer la lecture des PDF scannés et des images, il est recommandé d’installer Tesseract OCR.

Version recommandée :

```text
Tesseract OCR 5.5.0 ou supérieur
```

## Utilisation

Lancer l’application en double-cliquant sur :

```text
Bilan_Carbone.exe
```

L’utilisation se fait dans cet ordre :

1. Choisir le dossier client
2. Nettoyer les fichiers
3. Classifier les fichiers par scope
4. Calculer les émissions

Il est important de respecter cet ordre, car chaque étape dépend de la précédente.

## Organisation conseillée du dossier client

Il est conseillé de regrouper tous les documents du client dans un même dossier.

Exemple :

```text
Dossier_Client
│
├── Electricite
├── Eau
├── Dechets
├── Vehicules
├── Achats
├── Immobilisations
├── Factures
└── Autres_documents
```

Cette organisation n’est pas obligatoire, mais elle améliore la qualité du traitement automatique.

Les noms de fichiers doivent être les plus clairs possible, par exemple :

```text
Facture EDF 2024.pdf
Tableau consommation eau 2024.xlsx
Reporting déchets 2024.csv
Parc informatique 2024.xlsx
Déplacements domicile-travail.xlsx
Récap achats fournisseurs 2024.xlsx
```

## Résultats générés

Après traitement, l’outil affiche le bilan carbone par scope :

```text
SCOPE_1
SCOPE_2
SCOPE_3
TOTAL
```

Les résultats sont exprimés en kgCO₂e et en tCO₂e.

Certains résultats peuvent contenir la mention :

```text
[€ variable]
```

Cela signifie que le calcul est basé sur une donnée financière estimée. Ces données sont utiles pour ne pas oublier certains postes importants, mais elles doivent être relues avec plus de prudence.

## Fichiers d’audit

L’outil peut générer plusieurs fichiers d’audit, notamment :

```text
audit_extraction_bilan_carbone.csv
audit_rejets_financiers_non_classes.csv
```

Le fichier `audit_extraction_bilan_carbone.csv` contient les données extraites et utilisées dans le calcul.

Le fichier `audit_rejets_financiers_non_classes.csv` contient les montants financiers détectés mais non utilisés, car l’outil n’a pas réussi à les classer avec assez de confiance.

Ces fichiers doivent être consultés lorsque le résultat semble trop faible, trop élevé ou incomplet.

## Points de vérification

Après chaque calcul, il est conseillé de vérifier :

* que les fichiers importants ont bien été pris en compte ;
* que les fichiers sont classés dans le bon scope ;
* que les gros montants financiers ne sont pas oubliés ;
* que les fichiers d’audit ne contiennent pas de données importantes rejetées ;
* qu’il n’y a pas de doublon entre un récapitulatif global et des factures détaillées ;
* que le total obtenu est cohérent avec la taille et l’activité de l’entreprise.

## Limites connues

L’outil dépend fortement de la qualité des fichiers fournis.

Les cas les plus difficiles sont :

* PDF scannés de mauvaise qualité
* Photos floues ou coupées
* Tableaux mal reconnus
* Factures avec plusieurs montants
* Fichiers sans année claire
* Fichiers sans unité
* Libellés comptables trop vagues

Dans ces situations, certaines données peuvent ne pas être extraites automatiquement. Une vérification manuelle reste nécessaire.

## Recommandations

Pour obtenir de meilleurs résultats, il est conseillé de demander aux clients :

* des fichiers Excel ou CSV lorsque c’est possible ;
* des PDF numériques plutôt que des scans ;
* des récapitulatifs annuels ;
* des noms de fichiers explicites ;
* des documents séparés par type de donnée ;
* des fichiers indiquant clairement l’année concernée.

## Compilation avec PyInstaller

Pour recompiler l’application en dossier exécutable, utiliser la commande suivante depuis le dossier du projet :

```powershell
py -m PyInstaller --onedir --windowed --name Bilan_Carbone --add-data "fond.jpg;." --add-data "Bilan_Carbone_V9.01.xlsx;." gui.py
```

L’exécutable sera généré dans :

```text
dist\Bilan_Carbone
```

Le fichier `.env` doit ensuite être copié manuellement dans le même dossier que l’exécutable.

Exemple :

```powershell
copy .env dist\Bilan_Carbone\.env
```

## Nettoyage avant recompilation

Avant une nouvelle compilation, il est possible de supprimer les anciens fichiers générés :

```powershell
Remove-Item -Recurse -Force build, dist, __pycache__
Remove-Item -Force Bilan_Carbone.spec
```
