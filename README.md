# Suivi livraisons transporteurs

Reporting du statut de livraison des commandes à partir des liens de tracking
transporteurs (Heppner, DB Schenker / DSV, Teliway VCV, Speed Distribution, DHL).

> ⚠️ Projet **distinct** du générateur de rapports de stocks. Celui-ci ne traite
> **que** le suivi des livraisons (statuts *en cours / en retard / incident / livré*).

## Ce que fait l'outil

1. Lit un export ERP/TMS (`commandes.csv`) : n° commande, transporteur, lien de tracking, date prévue.
2. Ouvre chaque lien dans un Chromium headless (Playwright) qui exécute le JavaScript des portails.
3. Extrait le statut brut, le **classe automatiquement** et calcule les jours de retard.
4. Écrit `suivi_resultats.csv`, qui alimente le tableau de bord (`templates/Suivi_livraisons_transporteurs.xlsx`).

## Installation

```bash
python -m venv .venv
# Windows : .venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium
```

## Utilisation

```bash
python scraper_tracking.py                      # commandes.csv -> suivi_resultats.csv
python scraper_tracking.py entree.csv sortie.csv
python scraper_tracking.py --headful            # affiche le navigateur (debug)
```

Format d'entrée : voir `commandes.example.csv` (séparateur `;`, dates en `AAAA-MM-JJ`).

## Règles de classification

| Statut normalisé | Condition |
|---|---|
| Incident | libellé contient anomalie / avarie / refus / retour / litige / souffrance… |
| Livré | libellé contient livré / remis / distribué… |
| En retard | non livré **et** date prévue dépassée |
| En cours | non livré, date prévue non dépassée |
| À vérifier | aucun statut lisible (lien vide, page KO) |

Les mots-clés se règlent en haut de `scraper_tracking.py` (`INCIDENT_KEYWORDS`, `DELIVERED_KEYWORDS`).

## À affiner par transporteur

Le dictionnaire `CARRIERS` (haut du script) permet d'indiquer un sélecteur CSS
précis du statut par transporteur. Sans ça, l'extraction reste générique
(recherche de libellés dans le texte de la page). Pour fiabiliser : inspecter
la page rendue (clic droit → Inspecter) et renseigner `selector`.

## Structure

```
suivi-livraisons-transporteurs/
├── scraper_tracking.py          # scraper + classification
├── commandes.example.csv        # exemple d'entrée (données factices)
├── requirements.txt
├── .gitignore
└── templates/
    └── Suivi_livraisons_transporteurs.xlsx   # gabarit Excel + dashboard
```

## Confidentialité

`commandes.csv` et `suivi_resultats.csv` (données réelles) sont **exclus du dépôt**
via `.gitignore`. Ne versionner que `commandes.example.csv`.
