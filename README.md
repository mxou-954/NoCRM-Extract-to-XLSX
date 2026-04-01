# Export noCRM → Excel

Script Python qui récupère les leads d'une étape spécifique de **noCRM.io** via l'API et les exporte dans un fichier Excel structuré, avec une feuille par type de données (entreprises, contacts, résumé).

---

## Fonctionnement

1. **Récupération paginée** des leads depuis noCRM avec filtre strict sur `step_id`
2. **Parsing de la description** de chaque lead pour extraire les champs entreprise (SIREN, NAF, effectif, CA, etc.) et les blocs contacts (nom, fonction, téléphone, email, LinkedIn)
3. **Export Excel** en trois feuilles :
   - `Entreprises` — données générales du lead + champs extraits de la description
   - `Contacts` — un contact par ligne, rattaché à son lead
   - `Résumé` — total leads, total contacts, date d'export

---

## Prérequis

- Python 3.8+
- Un compte **noCRM.io** avec accès API

### Dépendances

```bash
pip install requests openpyxl python-dotenv
```

---

## Configuration

Créer un fichier `.env` à la racine du projet :

```env
NOCRM_API_KEY=''
NOCRM_SUBDOMAIN=''
```

| Variable          | Description                                                        |
|-------------------|--------------------------------------------------------------------|
| `NOCRM_API_KEY`   | Clé API noCRM (disponible dans les paramètres du compte)           |
| `NOCRM_SUBDOMAIN` | Sous-domaine noCRM (ex : `monentreprise` pour `monentreprise.nocrm.io`) |

Puis dans `main()`, configurer :

```python
STEP_ID  = 267810   # ID de l'étape à exporter
MAX_LEADS = None    # None = tous les leads, ou un entier pour limiter
```

---

## Utilisation

```bash
python main.py
```

Le fichier Excel est généré dans le répertoire courant avec un nom horodaté :

```
export_leads_nocrm_20250401_143022.xlsx
```

---

## Format attendu des descriptions noCRM

Le script s'appuie sur la structure des descriptions pour extraire les données.

### Champs entreprise

```
SIREN : 123 456 789
NAF : 4941A - Transports routiers de fret interurbains
Effectif : 50-99 salariés
Adresse : 12 rue de la Paix, 75001 Paris
Chiffre d'affaires : 8 500 000 €
Résultat net : 320 000 €
Site web : https://www.exemple.fr
Budget transport : 200 000 €/an
Description : Entreprise spécialisée dans...
```

### Blocs contacts (séparés par `----------`)

```
----------
Nom : Frédéric Mignon
Fonction : Chief Financial Officer
Téléphone : +33 3 80 44 71 63
Email : f.mignon@urgo.fr
Source : https://www.linkedin.com/in/frederic-m-01962710/
----------
```

---

## Structure du projet

```
.
├── main.py
├── .env
└── README.md
```

---

## Colonnes exportées

### Feuille `Entreprises`

| Colonne | Source |
|---|---|
| ID Lead, Titre, Étape, Tags | API noCRM |
| Créé le, Mis à jour le, Status | API noCRM |
| Amount, Prochaine action, Rappel | API noCRM |
| SIREN, NAF, Effectif, Adresse | Description parsée |
| Chiffre d'affaires, Résultat net | Description parsée |
| Site web, Budget transport | Description parsée |

### Feuille `Contacts`

ID Lead · Titre Lead · Nom · Fonction · Téléphone · Email · Source LinkedIn

---

## Notes

- Le filtre `step_id` est appliqué côté API **et** côté Python pour garantir la stricte cohérence des résultats.
- Les leads hors étape éventuellement renvoyés par l'API sont ignorés et signalés en console.
- La pagination est gérée automatiquement (100 leads par appel).
