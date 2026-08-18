# Export noCRM → Excel

Script Python qui récupère les leads d'une étape spécifique de **noCRM.io** via l'API et les exporte dans un fichier Excel structuré, avec une feuille par type de données (entreprises, contacts, anomalies, résumé).

---

## Fonctionnement

1. **Validation de la configuration** avant tout appel réseau (fichier `.env`, clé API, sous-domaine, paramètres d'export)
2. **Récupération paginée** des leads depuis noCRM avec filtre strict sur `step_id`, timeout et nouvelles tentatives automatiques
3. **Parsing de la description** de chaque lead pour extraire les champs entreprise (SIREN, NAF, effectif, CA, etc.) et les blocs contacts (nom, fonction, téléphone, email, LinkedIn)
4. **Contrôles de cohérence** sur les données extraites (SIREN, email, téléphone, site web)
5. **Export Excel** en quatre feuilles :
   - `Entreprises` — données générales du lead + champs extraits de la description
   - `Contacts` — un contact par ligne, rattaché à son lead
   - `Anomalies` — tout ce qui mérite une vérification humaine
   - `Résumé` — totaux, compteurs et date d'export

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

Puis dans `executer()`, configurer :

```python
STEP_ID  = 267810       # ID de l'étape à exporter (None = toutes les étapes)
MAX_LEADS = None        # None = tous les leads, ou un entier positif
LIMITE_PAR_PAGE = 100   # nombre de leads demandés par appel API (1 à 500)
```

Ces trois valeurs sont validées au démarrage : une valeur incohérente arrête le
script avec un message explicite, avant le moindre appel réseau.

---

## Utilisation

```bash
python main.py
```

Le fichier Excel est généré à côté du script avec un nom horodaté :

```
export_leads_nocrm_20250401_143022.xlsx
```

Un journal quotidien est écrit au même endroit :

```
export_20250401.log
```

### Codes de sortie

| Code | Signification |
|---|---|
| `0` | Export terminé, fichier Excel généré |
| `1` | Erreur : configuration, API, ou export impossible |
| `2` | Interruption manuelle (Ctrl+C) |

---

## Gestion des erreurs

Le script est conçu pour ne jamais afficher de trace Python brute : chaque
problème est traduit en message lisible, affiché en console et enregistré dans
le journal.

### Au démarrage

- Dépendances manquantes → liste des paquets à installer
- Fichier `.env` absent → chemin attendu et contenu à créer
- `NOCRM_API_KEY` vide ou trop courte → message dédié
- `NOCRM_SUBDOMAIN` vide ou mal formé → message dédié
  (une URL complète ou `monentreprise.nocrm.io` est acceptée et nettoyée automatiquement)
- `STEP_ID`, `MAX_LEADS`, `LIMITE_PAR_PAGE` invalides → arrêt avant tout appel réseau

### Pendant les appels API

| Situation | Comportement |
|---|---|
| Timeout (30 s), coupure réseau, DNS | 3 tentatives espacées de 5 s, puis arrêt |
| `429` quota atteint | attente indiquée par l'en-tête `Retry-After` (max 120 s), puis nouvelle tentative |
| `5xx` erreur serveur noCRM | 3 tentatives, puis arrêt |
| `401` / `403` clé invalide ou sans droits | arrêt immédiat, message dédié |
| `404` sous-domaine ou étape inconnus | arrêt immédiat, message dédié |
| Réponse non JSON (page de maintenance) | arrêt avec un aperçu de la réponse |
| Réponse au mauvais format (objet au lieu d'une liste) | arrêt avec le message renvoyé par l'API |

En cas d'échec définitif, le script **s'arrête** plutôt que de produire un
fichier incomplet sans le signaler. Un garde-fou limite aussi la pagination à
500 pages pour éviter toute boucle infinie.

### Pendant le traitement des leads

- Un lead illisible (mauvais type, champ manquant) est **ignoré** et signalé,
  les autres continuent d'être traités
- Les champs `None`, numériques, listes ou objets sont convertis en texte au
  lieu de provoquer une erreur
- Les leads renvoyés hors de l'étape demandée sont écartés et comptés

### À l'écriture du fichier Excel

- Les caractères de contrôle refusés par Excel sont retirés
- Les cellules de plus de 32 000 caractères sont tronquées proprement
- Si le fichier est déjà ouvert dans Excel, le script enregistre sous un nom de
  repli horodaté au lieu de tout perdre
- Dossier inaccessible ou en lecture seule → message explicite

---

## Feuille `Anomalies`

Elle liste les points à vérifier manuellement, sans bloquer l'export :

| Type | Déclencheur |
|---|---|
| `Description vide` | le lead n'a aucune description exploitable |
| `Aucun contact` | aucun bloc `Nom :` trouvé dans la description |
| `Contact sans coordonnées` | ni téléphone ni email |
| `Email suspect` | format d'email invalide |
| `Téléphone suspect` | moins de 8 chiffres |
| `SIREN suspect` | nombre de chiffres différent de 9 |
| `Site web suspect` | ne ressemble pas à une adresse web |
| `Lead illisible` / `Lead non traité` | élément ignoré pendant la récupération ou l'extraction |

---

## Format attendu des descriptions noCRM

Le script s'appuie sur la structure des descriptions pour extraire les données.
Pour normer vos fiches NoCRM, vous pouvez utiliser le répo suivant : ici

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
├── README.md
├── export_leads_nocrm_AAAAMMJJ_HHMMSS.xlsx   (généré)
└── export_AAAAMMJJ.log                       (généré)
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
- Le bilan affiché en fin d'exécution (appels API, leads ignorés, anomalies, erreurs) est également repris dans la feuille `Résumé`.
- Le journal `export_AAAAMMJJ.log` conserve l'historique complet des exécutions de la journée, avec la trace technique complète en cas d'erreur inattendue.
