# BRVM Data Extractor v2 📈

Application Streamlit pour l'extraction **automatique et historique** des données actions
depuis les Bulletins Officiels de la Cote (BOC) de la BRVM.

## 🆕 Nouveautés v2

- **Extraction par année antérieure complète** (2015–présent)
- **4 modes d'extraction** : Récents, Mois spécifique, Année complète, Import Excel
- **Détection intelligente des séances** : génère tous les jours Lun–Ven, filtre les fériés fixes UEMOA, détecte automatiquement les PDF manquants (fériés variables, clôtures exceptionnelles)
- **Log d'extraction persistant** : sait exactement quelles dates ont été tentées
- **Barre de progression détaillée** avec log en temps réel
- **Vue historique annuelle** : performances, couverture, calendrier mensuel
- **Suppression d'une année** sans toucher aux autres
- **Graphique Candlestick** quand les données ouverture/clôture sont disponibles

## Installation

```bash
pip install -r requirements.txt
streamlit run brvm_app_v2.py
```

## Modes d'extraction

### 📅 Bulletins récents
Scrape la page officielle `bfin.brvm.org/boc/boc_jour.aspx` et télécharge les N derniers BOC.

### 📆 Mois spécifique
Sélectionner une année + un mois → génère toutes les dates Lun–Ven candidates et tente chaque PDF.

### 📚 Année complète ⭐
```
Sélectionner une année (ex: 2024) → ~257 dates candidates
↓
Pour chaque date :
  • Tente GET https://bfin.brvm.org/boc/BOC_JOUR/BOC_20241205.pdf
  • PDF valide (>30KB, Content-Type=application/pdf) → extraction
  • PDF absent (404 ou HTML) → marqué "not_found" (férié ou pas de séance)
↓
Cache Parquet mis à jour
Log JSON persistant (ne retente pas les dates déjà traitées)
```

### 📂 Import Excel
Compatible avec le fichier `BRVM_20260225.xlsx` généré précédemment.

## Données extraites

| Champ | Description |
|-------|-------------|
| date | JJ/MM/AAAA |
| symbole | Code BRVM (ETIT, SLBC…) |
| titre | Nom complet |
| secteur | Secteur UEMOA |
| ouverture | Cours d'ouverture (FCFA) |
| cloture | Cours de clôture (FCFA) |
| variation_jour_pct | Variation journalière % |
| volume | Titres échangés |
| valeur_fcfa | Valeur transigée (FCFA) |
| variation_annee_pct | Performance depuis 01/01 |
| dividende_net | Dernier dividende net |
| rendement_net_pct | Rendement net % |
| per | Price-Earnings Ratio |

## Structure

```
brvm_data/
├── extracted_data.parquet   # Cache des données (toutes années)
└── extraction_log.json      # Journal des tentatives par date
```

## Estimation des temps d'extraction

| Périmètre | Dates candidates | Temps estimé* |
|-----------|-----------------|---------------|
| 1 mois | ~22 | 1–3 min |
| 1 trimestre | ~65 | 3–8 min |
| 1 année | ~257 | 10–25 min |
| 5 années | ~1280 | 50–120 min |

*Avec délai 0.4s entre requêtes (configurable)

## Exports disponibles

- **CSV Brut** : toutes colonnes, séparateur `;`
- **CSV Pivot** : format `Date,ABJC,BICB,...,UNXC` (cours clôture)
- **Excel** : 4 feuilles (Données, Pivot, Stats, Couverture extraction)
