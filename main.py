# -*- coding: utf-8 -*-
"""
Export des leads noCRM.io vers un fichier Excel.

"""

import os
import re
import sys
import html
import time
import traceback
from datetime import datetime
from os.path import join, dirname, abspath, exists, isdir


# ──────────────────────────────────────────────
# 0. Vérification des dépendances externes
# ──────────────────────────────────────────────

DEPENDANCES_MANQUANTES = []

try:
    import requests
except ImportError:
    requests = None
    DEPENDANCES_MANQUANTES.append("requests")

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None
    DEPENDANCES_MANQUANTES.append("python-dotenv")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    Workbook = None
    DEPENDANCES_MANQUANTES.append("openpyxl")

if DEPENDANCES_MANQUANTES:
    print("")
    print("ERREUR : des bibliothèques Python sont manquantes.")
    for nom in DEPENDANCES_MANQUANTES:
        print("  - " + nom)
    print("")
    print("Installez-les avec la commande :")
    print("  pip install requests openpyxl python-dotenv")
    print("")
    sys.exit(1)


# ──────────────────────────────────────────────
# 1. Constantes de configuration
# ──────────────────────────────────────────────

DOSSIER_SCRIPT = dirname(abspath(__file__))
CHEMIN_ENV = join(DOSSIER_SCRIPT, ".env")

# Réseau
DELAI_ATTENTE_REQUETE = 30      # secondes avant d'abandonner une requête
NOMBRE_TENTATIVES = 3           # nombre d'essais par appel API
PAUSE_ENTRE_TENTATIVES = 5      # secondes d'attente avant de réessayer

# Garde-fous
LIMITE_PAR_PAGE_MIN = 1
LIMITE_PAR_PAGE_MAX = 500
NOMBRE_MAX_DE_PAGES = 500       # évite une boucle infinie si l'API se comporte mal
LONGUEUR_MAX_CELLULE = 32000    # Excel refuse au-delà de 32767 caractères

# Codes de sortie du programme
CODE_SUCCES = 0
CODE_ERREUR = 1
CODE_INTERRUPTION = 2

# Colonnes de la feuille "Entreprises", dans l'ordre.
# La liste est figée ici pour que le fichier Excel ait toujours les mêmes
# colonnes, même si un lead est incomplet.
COLONNES_ENTREPRISE = [
    "ID Lead",
    "Titre",
    "Étape",
    "Tags",
    "Créé le",
    "Mis à jour le",
    "Status",
    "Amount",
    "Prochaine action",
    "Date de rappel",
    "Heure de rappel",
    "Créé par",
    "SIREN",
    "NAF",
    "Effectif",
    "Adresse",
    "Chiffre d'affaires",
    "Résultat net",
    "Site web",
    "Budget transport",
    "Description activité",
]

COLONNES_CONTACT = [
    "ID Lead",
    "Titre Lead",
    "Nom",
    "Fonction",
    "Téléphone",
    "Email",
    "Source LinkedIn",
]

COLONNES_ANOMALIE = [
    "Type",
    "ID Lead",
    "Titre Lead",
    "Détail",
]


# ──────────────────────────────────────────────
# 2. Journalisation (console + fichier log)
# ──────────────────────────────────────────────

# Compteurs remplis au fil de l'exécution, affichés dans le bilan final.
STATS = {
    "appels_api": 0,
    "tentatives_echouees": 0,
    "leads_recus": 0,
    "leads_hors_etape": 0,
    "leads_ignores": 0,
    "leads_sans_description": 0,
    "leads_sans_contact": 0,
    "avertissements": 0,
    "erreurs": 0,
}

# Liste des anomalies détectées, exportée dans une feuille Excel dédiée.
ANOMALIES = []

# État de la journalisation dans un fichier.
JOURNAL = {
    "chemin": None,
    "actif": False,
    "deja_prevenu": False,
}


def preparer_journal():
    """Prépare le fichier de log. Si c'est impossible, on continue sans."""
    nom = "export_" + datetime.now().strftime("%Y%m%d") + ".log"
    chemin = join(DOSSIER_SCRIPT, nom)
    JOURNAL["chemin"] = chemin
    try:
        fichier = open(chemin, "a", encoding="utf-8")
        fichier.write("\n")
        fichier.write("=" * 70 + "\n")
        fichier.write("Nouvelle exécution : ")
        fichier.write(datetime.now().strftime("%d/%m/%Y %H:%M:%S") + "\n")
        fichier.write("=" * 70 + "\n")
        fichier.close()
        JOURNAL["actif"] = True
    except OSError as detail:
        JOURNAL["actif"] = False
        print("[ATTENTION] Impossible d'écrire le fichier de log (" + str(detail) + ").")
        print("            L'export continue, mais sans journal sur disque.")


def ecrire_dans_journal(ligne):
    """Écrit une ligne dans le journal, sans jamais faire planter le script."""
    if not JOURNAL["actif"]:
        return
    try:
        fichier = open(JOURNAL["chemin"], "a", encoding="utf-8")
        fichier.write(ligne + "\n")
        fichier.close()
    except OSError as detail:
        JOURNAL["actif"] = False
        if not JOURNAL["deja_prevenu"]:
            JOURNAL["deja_prevenu"] = True
            print("[ATTENTION] Écriture dans le journal interrompue (" + str(detail) + ").")


def journaliser(niveau, message):
    """Affiche un message en console et l'enregistre dans le journal."""
    horodatage = datetime.now().strftime("%H:%M:%S")
    print("[" + niveau + "] " + message)
    ecrire_dans_journal("[" + horodatage + "] [" + niveau + "] " + message)


def info(message):
    journaliser("INFO", message)


def succes(message):
    journaliser("OK", message)


def avertissement(message):
    STATS["avertissements"] = STATS["avertissements"] + 1
    journaliser("ATTENTION", message)


def erreur(message):
    STATS["erreurs"] = STATS["erreurs"] + 1
    journaliser("ERREUR", message)


def ajouter_anomalie(type_anomalie, id_lead, titre_lead, detail):
    """Enregistre une anomalie de données pour la feuille Excel dédiée."""
    anomalie = {
        "Type": type_anomalie,
        "ID Lead": id_lead,
        "Titre Lead": titre_lead,
        "Détail": detail,
    }
    ANOMALIES.append(anomalie)


# ──────────────────────────────────────────────
# 3. Erreurs personnalisées
# ──────────────────────────────────────────────

class ErreurConfiguration(Exception):
    """Problème de configuration : fichier .env, clé API, paramètres."""
    pass


class ErreurAPI(Exception):
    """Problème lors d'un appel à l'API noCRM."""
    pass


class ErreurExport(Exception):
    """Problème lors de la création du fichier Excel."""
    pass


# ──────────────────────────────────────────────
# 4. Conversion sûre des valeurs
# ──────────────────────────────────────────────
# L'API peut renvoyer None, un nombre, un dictionnaire ou une liste là où on
# attend du texte. Ces fonctions évitent les plantages du type
# "AttributeError: 'NoneType' object has no attribute 'strip'".

def en_texte(valeur):
    """Transforme n'importe quelle valeur en texte propre."""
    if valeur is None:
        return ""

    if isinstance(valeur, str):
        return valeur.strip()

    if isinstance(valeur, bool):
        if valeur:
            return "Oui"
        return "Non"

    if isinstance(valeur, (int, float)):
        return str(valeur)

    if isinstance(valeur, dict):
        # Les objets noCRM exposent souvent un libellé sous l'une de ces clés.
        for cle in ["name", "title", "label", "value"]:
            if cle in valeur:
                return en_texte(valeur[cle])
        return ""

    if isinstance(valeur, list):
        morceaux = []
        for element in valeur:
            texte = en_texte(element)
            if texte != "":
                morceaux.append(texte)
        return ", ".join(morceaux)

    return str(valeur)


def sous_dictionnaire(source, cle):
    """Retourne source[cle] si c'est bien un dictionnaire, sinon un dict vide."""
    if not isinstance(source, dict):
        return {}
    valeur = source.get(cle)
    if isinstance(valeur, dict):
        return valeur
    return {}


def en_entier_ou_none(valeur):
    """Convertit en entier si c'est possible, sinon retourne None."""
    if valeur is None:
        return None
    if isinstance(valeur, bool):
        return None
    if isinstance(valeur, int):
        return valeur
    if isinstance(valeur, float):
        return int(valeur)
    if isinstance(valeur, str):
        texte = valeur.strip()
        if texte == "":
            return None
        try:
            return int(texte)
        except ValueError:
            return None
    return None


# ──────────────────────────────────────────────
# 5. Chargement et validation de la configuration
# ──────────────────────────────────────────────

def nettoyer_sous_domaine(valeur):
    """Accepte 'monentreprise', 'monentreprise.nocrm.io' ou une URL complète."""
    texte = en_texte(valeur)
    texte = texte.replace("https://", "")
    texte = texte.replace("http://", "")
    texte = texte.strip("/")
    if texte.endswith(".nocrm.io"):
        texte = texte[:-len(".nocrm.io")]
    # Certains fichiers .env gardent des guillemets collés à la valeur.
    texte = texte.strip("'").strip('"')
    return texte.strip()


def charger_configuration():
    """Lit le .env, vérifie les valeurs et retourne (clé API, sous-domaine)."""
    if not exists(CHEMIN_ENV):
        message = (
            "Fichier .env introuvable à l'emplacement suivant :\n"
            "  " + CHEMIN_ENV + "\n"
            "Créez ce fichier à côté de main.py avec le contenu :\n"
            "  NOCRM_API_KEY=votre_cle_api\n"
            "  NOCRM_SUBDOMAIN=votre_sous_domaine"
        )
        raise ErreurConfiguration(message)

    try:
        load_dotenv(CHEMIN_ENV)
    except OSError as detail:
        raise ErreurConfiguration("Impossible de lire le fichier .env : " + str(detail))

    cle_api = en_texte(os.environ.get("NOCRM_API_KEY"))
    cle_api = cle_api.strip("'").strip('"').strip()
    sous_domaine = nettoyer_sous_domaine(os.environ.get("NOCRM_SUBDOMAIN"))

    problemes = []

    if cle_api == "":
        problemes.append("NOCRM_API_KEY est vide ou absente du fichier .env")
    elif len(cle_api) < 10:
        # Simple garde-fou : une vraie clé noCRM est bien plus longue.
        problemes.append(
            "NOCRM_API_KEY semble trop courte (" + str(len(cle_api)) + " caractères). "
            "Vérifiez que la clé a bien été copiée en entier."
        )

    if sous_domaine == "":
        problemes.append("NOCRM_SUBDOMAIN est vide ou absent du fichier .env")
    elif not re.match(r"^[a-zA-Z0-9][a-zA-Z0-9-]*$", sous_domaine):
        problemes.append(
            "NOCRM_SUBDOMAIN contient des caractères invalides : '" + sous_domaine + "'. "
            "Attendu : uniquement des lettres, chiffres et tirets (ex : monentreprise)."
        )

    if problemes:
        message = "Configuration invalide dans " + CHEMIN_ENV + " :"
        for probleme in problemes:
            message = message + "\n  - " + probleme
        raise ErreurConfiguration(message)

    return cle_api, sous_domaine


def valider_parametres_export(step_id, max_leads, limite_par_page):
    """Vérifie les paramètres définis en dur dans executer()."""
    problemes = []

    if step_id is not None:
        entier = en_entier_ou_none(step_id)
        if entier is None or entier <= 0:
            problemes.append(
                "STEP_ID doit être un entier positif ou None (valeur actuelle : "
                + repr(step_id) + ")"
            )

    if max_leads is not None:
        entier = en_entier_ou_none(max_leads)
        if entier is None or entier <= 0:
            problemes.append(
                "MAX_LEADS doit être un entier positif ou None (valeur actuelle : "
                + repr(max_leads) + ")"
            )

    entier = en_entier_ou_none(limite_par_page)
    if entier is None or entier < LIMITE_PAR_PAGE_MIN or entier > LIMITE_PAR_PAGE_MAX:
        problemes.append(
            "LIMITE_PAR_PAGE doit être un entier entre "
            + str(LIMITE_PAR_PAGE_MIN) + " et " + str(LIMITE_PAR_PAGE_MAX)
            + " (valeur actuelle : " + repr(limite_par_page) + ")"
        )

    if problemes:
        message = "Paramètres d'export invalides :"
        for probleme in problemes:
            message = message + "\n  - " + probleme
        raise ErreurConfiguration(message)


# ──────────────────────────────────────────────
# 6. Appel API avec nouvelles tentatives
# ──────────────────────────────────────────────

def lire_delai_avant_reessai(reponse):
    """Lit l'en-tête Retry-After d'une réponse 429, avec une valeur de repli."""
    valeur = reponse.headers.get("Retry-After")
    secondes = en_entier_ou_none(valeur)
    if secondes is None or secondes <= 0:
        return PAUSE_ENTRE_TENTATIVES
    if secondes > 120:
        return 120
    return secondes


def appeler_api(url, entetes, parametres):
    """
    Effectue un GET sur l'API noCRM.

    Réessaie automatiquement en cas de problème temporaire (réseau, 429, 5xx).
    Retourne les données JSON, ou lève ErreurAPI avec un message explicite.
    """
    tentative = 1

    while tentative <= NOMBRE_TENTATIVES:
        pause = PAUSE_ENTRE_TENTATIVES
        raison = ""

        try:
            STATS["appels_api"] = STATS["appels_api"] + 1
            reponse = requests.get(
                url,
                headers=entetes,
                params=parametres,
                timeout=DELAI_ATTENTE_REQUETE,
            )
        except requests.exceptions.Timeout:
            raison = (
                "l'API n'a pas répondu en moins de "
                + str(DELAI_ATTENTE_REQUETE) + " secondes"
            )
        except requests.exceptions.ConnectionError:
            raison = "impossible de joindre l'API (connexion internet ou DNS ?)"
        except requests.exceptions.RequestException as detail:
            raison = "erreur réseau inattendue : " + str(detail)
        else:
            code = reponse.status_code

            if code == 200:
                try:
                    donnees = reponse.json()
                except ValueError:
                    apercu = reponse.text[:200]
                    raise ErreurAPI(
                        "L'API a répondu 200 mais le contenu n'est pas du JSON.\n"
                        "  Début de la réponse : " + apercu
                    )
                return donnees

            # Erreurs définitives : inutile de réessayer.
            if code == 401:
                raise ErreurAPI(
                    "Authentification refusée (401). La clé NOCRM_API_KEY est "
                    "invalide, révoquée ou mal copiée."
                )
            if code == 403:
                raise ErreurAPI(
                    "Accès refusé (403). La clé API n'a pas les droits nécessaires "
                    "pour lire les leads."
                )
            if code == 404:
                raise ErreurAPI(
                    "Ressource introuvable (404). Vérifiez NOCRM_SUBDOMAIN et "
                    "l'existence du STEP_ID demandé.\n  URL appelée : " + url
                )
            if code == 422:
                raise ErreurAPI(
                    "Paramètres refusés par l'API (422). Vérifiez STEP_ID et la "
                    "limite par page.\n  Réponse : " + reponse.text[:200]
                )

            # Erreurs temporaires : on réessaie.
            if code == 429:
                pause = lire_delai_avant_reessai(reponse)
                raison = "quota d'appels API atteint (429)"
            elif code >= 500:
                raison = "erreur serveur noCRM (" + str(code) + ")"
            else:
                raise ErreurAPI(
                    "Réponse HTTP inattendue (" + str(code) + ").\n"
                    "  Réponse : " + reponse.text[:200]
                )

        # Si on arrive ici, la tentative a échoué pour une raison temporaire.
        STATS["tentatives_echouees"] = STATS["tentatives_echouees"] + 1

        if tentative < NOMBRE_TENTATIVES:
            avertissement(
                "Tentative " + str(tentative) + "/" + str(NOMBRE_TENTATIVES)
                + " échouée : " + raison + ". Nouvel essai dans "
                + str(pause) + " s."
            )
            time.sleep(pause)
        else:
            raise ErreurAPI(
                "Échec après " + str(NOMBRE_TENTATIVES) + " tentatives : " + raison
            )

        tentative = tentative + 1

    # Sécurité : cette ligne ne devrait jamais être atteinte.
    raise ErreurAPI("Échec de l'appel API pour une raison inconnue.")


# ──────────────────────────────────────────────
# 7. Récupération des leads (pagination + filtre strict)
# ──────────────────────────────────────────────

def recuperer_tous_les_leads(base_url, entetes, step_id, limite_par_page, max_leads):
    """
    Récupère les leads page par page, avec filtre strict sur step_id.

    En cas d'erreur API définitive, lève ErreurAPI : on préfère s'arrêter
    plutôt que produire un export silencieusement incomplet.
    """
    tous_les_leads = []
    decalage = 0
    numero_page = 0
    url = base_url + "/leads"

    while True:
        numero_page = numero_page + 1

        if numero_page > NOMBRE_MAX_DE_PAGES:
            avertissement(
                "Limite de sécurité atteinte (" + str(NOMBRE_MAX_DE_PAGES)
                + " pages). Arrêt de la pagination."
            )
            break

        parametres = {"limit": limite_par_page, "offset": decalage}
        if step_id is not None:
            parametres["step_id"] = step_id

        donnees = appeler_api(url, entetes, parametres)

        # L'API doit renvoyer une liste. Un dictionnaire signale en général une
        # erreur applicative renvoyée malgré un code 200.
        if isinstance(donnees, dict):
            message_api = en_texte(donnees.get("error"))
            if message_api == "":
                message_api = en_texte(donnees.get("message"))
            if message_api == "":
                message_api = str(donnees)[:200]
            raise ErreurAPI(
                "L'API a renvoyé un message au lieu d'une liste de leads : " + message_api
            )

        if not isinstance(donnees, list):
            raise ErreurAPI(
                "Format de réponse inattendu : "
                + type(donnees).__name__ + " au lieu d'une liste."
            )

        nombre_recu = len(donnees)
        STATS["leads_recus"] = STATS["leads_recus"] + nombre_recu

        if nombre_recu == 0:
            break

        # On ne garde que les éléments exploitables (des dictionnaires).
        leads_valides = []
        for element in donnees:
            if isinstance(element, dict):
                leads_valides.append(element)
            else:
                STATS["leads_ignores"] = STATS["leads_ignores"] + 1
                ajouter_anomalie(
                    "Lead illisible", "", "",
                    "Élément ignoré, type inattendu : " + type(element).__name__
                )

        # Filtre strict : l'API renvoie parfois des leads d'autres étapes.
        if step_id is not None:
            leads_de_l_etape = []
            for lead in leads_valides:
                etape_du_lead = en_entier_ou_none(lead.get("step_id"))
                if etape_du_lead == step_id:
                    leads_de_l_etape.append(lead)
                else:
                    STATS["leads_hors_etape"] = STATS["leads_hors_etape"] + 1
            nombre_ecarte = len(leads_valides) - len(leads_de_l_etape)
            if nombre_ecarte > 0:
                avertissement(
                    str(nombre_ecarte) + " lead(s) hors étape ignoré(s) sur la page "
                    + str(numero_page)
                )
            leads_valides = leads_de_l_etape

        for lead in leads_valides:
            tous_les_leads.append(lead)

        info(
            "Page " + str(numero_page) + " : " + str(nombre_recu)
            + " lead(s) reçu(s), " + str(len(tous_les_leads)) + " conservé(s) au total"
        )

        if max_leads is not None and len(tous_les_leads) >= max_leads:
            tous_les_leads = tous_les_leads[:max_leads]
            info("Limite MAX_LEADS atteinte (" + str(max_leads) + ").")
            break

        # Dernière page : l'API a renvoyé moins d'éléments que demandé.
        if nombre_recu < limite_par_page:
            break

        decalage = decalage + limite_par_page

    return tous_les_leads


# ──────────────────────────────────────────────
# 8. Nettoyage de la description
# ──────────────────────────────────────────────

def nettoyer_description(brut):
    """Convertit une description HTML ou texte brut en texte lisible."""
    texte = en_texte(brut)
    if texte == "":
        return ""

    try:
        texte = re.sub(r"<br\s*/?>|</p>|</div>|</li>", "\n", texte, flags=re.IGNORECASE)
        texte = re.sub(r"<[^>]+>", "", texte)
        texte = html.unescape(texte)
        texte = texte.replace("\r\n", "\n")
        texte = texte.replace("\r", "\n")
        texte = re.sub(r"\n{3,}", "\n\n", texte)
    except (re.error, TypeError) as detail:
        avertissement("Nettoyage de description impossible : " + str(detail))
        return en_texte(brut)

    return texte.strip()


# ──────────────────────────────────────────────
# 9. Extraction des champs entreprise
# ──────────────────────────────────────────────

MOTIFS_ENTREPRISE = {
    "SIREN": r"SIREN\s*:\s*(.+)",
    "NAF": r"NAF\s*:\s*([^-–—\n]+)",
    "Effectif": r"Effectif\s*:\s*(.+)",
    "Adresse": r"Adresse\s*:\s*(.+)",
    "Chiffre d'affaires": r"Chiffre d[''']affaire[s]?[^:]*:\s*(.+)",
    "Résultat net": r"R[ée]sultat\s+net[^:]*:\s*(.+)",
    "Site web": r"Site\s+web\s*:\s*(.+)",
    "Budget transport": r"Budget\s+transport[^:]*:\s*(.+)",
    "Description": r"Description\s*:\s*(.+)",
}


def extraire_champs_entreprise(description_nettoyee):
    """Retourne un dictionnaire des champs trouvés dans la description."""
    champs = {}
    if description_nettoyee == "":
        return champs

    for nom_champ in MOTIFS_ENTREPRISE:
        motif = MOTIFS_ENTREPRISE[nom_champ]
        try:
            trouve = re.search(motif, description_nettoyee, re.IGNORECASE)
        except re.error as detail:
            avertissement("Motif invalide pour le champ '" + nom_champ + "' : " + str(detail))
            continue
        if trouve:
            valeur = trouve.group(1).strip()
            if valeur != "":
                champs[nom_champ] = valeur

    return champs


def verifier_siren(valeur, id_lead, titre_lead):
    """Contrôle simple du SIREN : 9 chiffres une fois les espaces retirés."""
    if valeur == "":
        return
    chiffres = re.sub(r"\D", "", valeur)
    if len(chiffres) != 9:
        ajouter_anomalie(
            "SIREN suspect", id_lead, titre_lead,
            "Valeur lue : '" + valeur + "' (" + str(len(chiffres))
            + " chiffres au lieu de 9)"
        )


def verifier_site_web(valeur, id_lead, titre_lead):
    """Contrôle simple de l'adresse du site web."""
    if valeur == "":
        return
    if not re.match(r"^(https?://)?[\w.-]+\.[a-zA-Z]{2,}", valeur):
        ajouter_anomalie(
            "Site web suspect", id_lead, titre_lead,
            "Valeur lue : '" + valeur + "'"
        )


# ──────────────────────────────────────────────
# 10. Extraction des contacts
# ──────────────────────────────────────────────

def verifier_email(valeur, id_lead, titre_lead, nom_contact):
    """Contrôle simple du format de l'email."""
    if valeur == "":
        return
    if not re.match(r"^[^@\s]+@[^@\s]+\.[a-zA-Z]{2,}$", valeur):
        ajouter_anomalie(
            "Email suspect", id_lead, titre_lead,
            "Contact '" + nom_contact + "' : '" + valeur + "'"
        )


def verifier_telephone(valeur, id_lead, titre_lead, nom_contact):
    """Contrôle simple du téléphone : au moins 8 chiffres."""
    if valeur == "":
        return
    chiffres = re.sub(r"\D", "", valeur)
    if len(chiffres) < 8:
        ajouter_anomalie(
            "Téléphone suspect", id_lead, titre_lead,
            "Contact '" + nom_contact + "' : '" + valeur + "'"
        )


def chercher_champ(bloc, motif):
    """Cherche un champ dans un bloc de texte, retourne '' si absent."""
    try:
        trouve = re.search(motif, bloc, re.IGNORECASE)
    except re.error:
        return ""
    if trouve:
        return trouve.group(1).strip()
    return ""


def extraire_contacts(description_nettoyee, id_lead, titre_lead):
    """Découpe la description en blocs contacts et valide chaque contact."""
    contacts = []
    if description_nettoyee == "":
        return contacts

    blocs = re.split(r"\s*-{5,}\s*", description_nettoyee)

    for bloc in blocs:
        nom = chercher_champ(bloc, r"Nom\s*:\s*(.+)")
        if nom == "":
            continue

        contact = {}
        contact["Nom"] = nom
        contact["Fonction"] = chercher_champ(bloc, r"Fonction\s*:\s*(.+)")
        contact["Téléphone"] = chercher_champ(bloc, r"T[ée]l[ée]phone\s*:\s*(.+)")
        contact["Email"] = chercher_champ(bloc, r"Email\s*:\s*(.+)")
        contact["Source LinkedIn"] = chercher_champ(bloc, r"Source\s*:\s*(https?://\S+|.+)")

        verifier_email(contact["Email"], id_lead, titre_lead, contact["Nom"])
        verifier_telephone(contact["Téléphone"], id_lead, titre_lead, contact["Nom"])

        if contact["Téléphone"] == "" and contact["Email"] == "":
            ajouter_anomalie(
                "Contact sans coordonnées", id_lead, titre_lead,
                "Contact '" + contact["Nom"] + "' : ni téléphone ni email"
            )

        contacts.append(contact)

    return contacts


# ──────────────────────────────────────────────
# 11. Extraction structurée d'un lead
# ──────────────────────────────────────────────

def extraire_donnees_lead(lead):
    """
    Transforme un lead brut de l'API en (entreprise, contacts).

    Toutes les valeurs passent par en_texte() : un champ absent ou d'un type
    inattendu donne une chaîne vide au lieu de faire planter le script.
    """
    id_lead = en_texte(lead.get("id"))
    titre_lead = en_texte(lead.get("title"))

    description = nettoyer_description(lead.get("description"))

    if description == "":
        STATS["leads_sans_description"] = STATS["leads_sans_description"] + 1
        ajouter_anomalie(
            "Description vide", id_lead, titre_lead,
            "Aucune donnée entreprise ni contact ne peut être extraite"
        )

    champs = extraire_champs_entreprise(description)
    contacts = extraire_contacts(description, id_lead, titre_lead)

    if len(contacts) == 0 and description != "":
        STATS["leads_sans_contact"] = STATS["leads_sans_contact"] + 1
        ajouter_anomalie(
            "Aucun contact", id_lead, titre_lead,
            "Aucun bloc 'Nom :' trouvé dans la description"
        )

    infos_etendues = sous_dictionnaire(sous_dictionnaire(lead, "extended_info"), "fields")

    adresse = champs.get("Adresse", "")
    if adresse == "":
        adresse = en_texte(infos_etendues.get("address"))

    site_web = champs.get("Site web", "")
    if site_web == "":
        site_web = en_texte(infos_etendues.get("web"))

    siren = champs.get("SIREN", "")
    verifier_siren(siren, id_lead, titre_lead)
    verifier_site_web(site_web, id_lead, titre_lead)

    if id_lead == "":
        ajouter_anomalie(
            "Lead sans identifiant", "", titre_lead,
            "Le champ 'id' est absent de la réponse API"
        )

    entreprise = {}
    entreprise["ID Lead"] = id_lead
    entreprise["Titre"] = titre_lead
    entreprise["Étape"] = en_texte(lead.get("step"))
    entreprise["Tags"] = en_texte(lead.get("tags"))
    entreprise["Créé le"] = en_texte(lead.get("created_at"))
    entreprise["Mis à jour le"] = en_texte(lead.get("updated_at"))
    entreprise["Status"] = en_texte(lead.get("status"))
    entreprise["Amount"] = en_texte(lead.get("amount"))
    entreprise["Prochaine action"] = en_texte(lead.get("next_action_at"))
    entreprise["Date de rappel"] = en_texte(lead.get("remind_date"))
    entreprise["Heure de rappel"] = en_texte(lead.get("remind_time"))
    entreprise["Créé par"] = en_texte(lead.get("created_from"))
    entreprise["SIREN"] = siren
    entreprise["NAF"] = champs.get("NAF", "")
    entreprise["Effectif"] = champs.get("Effectif", "")
    entreprise["Adresse"] = adresse
    entreprise["Chiffre d'affaires"] = champs.get("Chiffre d'affaires", "")
    entreprise["Résultat net"] = champs.get("Résultat net", "")
    entreprise["Site web"] = site_web
    entreprise["Budget transport"] = champs.get("Budget transport", "")
    entreprise["Description activité"] = champs.get("Description", "")

    return entreprise, contacts


def traiter_les_leads(leads):
    """
    Traite tous les leads.

    Un lead en erreur est signalé et mis de côté, mais n'interrompt pas
    l'export des autres.
    """
    donnees = []

    for lead in leads:
        try:
            entreprise, contacts = extraire_donnees_lead(lead)
        except Exception as detail:
            STATS["leads_ignores"] = STATS["leads_ignores"] + 1
            identifiant = ""
            if isinstance(lead, dict):
                identifiant = en_texte(lead.get("id"))
            erreur(
                "Lead #" + identifiant + " ignoré ("
                + type(detail).__name__ + " : " + str(detail) + ")"
            )
            ecrire_dans_journal(traceback.format_exc())
            ajouter_anomalie(
                "Lead non traité", identifiant, "",
                type(detail).__name__ + " : " + str(detail)
            )
            continue

        donnees.append((entreprise, contacts))

        titre_affiche = entreprise["Titre"]
        if len(titre_affiche) > 40:
            titre_affiche = titre_affiche[:37] + "..."
        info(
            "Lead #" + entreprise["ID Lead"] + " " + titre_affiche
            + " -> " + str(len(contacts)) + " contact(s)"
        )

    return donnees


# ──────────────────────────────────────────────
# 12. Export Excel
# ──────────────────────────────────────────────

# Excel refuse les caractères de contrôle : on les retire avant écriture.
CARACTERES_INTERDITS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def valeur_pour_excel(valeur):
    """Prépare une valeur pour openpyxl : type accepté, longueur bornée, sans caractère interdit."""
    if isinstance(valeur, bool):
        if valeur:
            return "Oui"
        return "Non"

    if isinstance(valeur, (int, float)):
        return valeur

    texte = en_texte(valeur)
    texte = CARACTERES_INTERDITS.sub("", texte)

    if len(texte) > LONGUEUR_MAX_CELLULE:
        texte = texte[:LONGUEUR_MAX_CELLULE] + " [...tronqué]"

    return texte


def ajouter_ligne(feuille, valeurs):
    """Ajoute une ligne après nettoyage de chaque valeur."""
    ligne = []
    for valeur in valeurs:
        ligne.append(valeur_pour_excel(valeur))
    feuille.append(ligne)


def construire_nom_de_fichier(nom_demande):
    """Valide le nom de fichier demandé, ou en génère un horodaté."""
    if nom_demande is None or en_texte(nom_demande) == "":
        horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
        return "export_leads_nocrm_" + horodatage + ".xlsx"

    nom = en_texte(nom_demande)
    # On remplace les caractères interdits par Windows dans un nom de fichier.
    nom = re.sub(r'[<>:"|?*\\/]', "_", nom)
    if not nom.lower().endswith(".xlsx"):
        nom = nom + ".xlsx"
    return nom


def enregistrer_classeur(classeur, chemin):
    """
    Enregistre le classeur.

    Si le fichier est déjà ouvert dans Excel, on tente un nom de repli plutôt
    que de perdre tout le travail effectué.
    """
    try:
        classeur.save(chemin)
        return chemin
    except PermissionError:
        avertissement(
            "Écriture impossible dans '" + chemin + "' : le fichier est "
            "probablement ouvert dans Excel."
        )
    except OSError as detail:
        raise ErreurExport("Impossible d'enregistrer le fichier : " + str(detail))

    # Deuxième essai avec un nom de repli horodaté à la seconde.
    chemin_repli = chemin[:-len(".xlsx")] + "_" + datetime.now().strftime("%H%M%S") + ".xlsx"
    try:
        classeur.save(chemin_repli)
        avertissement("Enregistré sous un nom de repli : " + chemin_repli)
        return chemin_repli
    except (PermissionError, OSError) as detail:
        raise ErreurExport(
            "Impossible d'enregistrer le fichier Excel, même sous un autre nom.\n"
            "  Détail : " + str(detail) + "\n"
            "  Fermez le fichier dans Excel puis relancez le script."
        )


def exporter_vers_excel(donnees_leads, dossier_sortie, nom_fichier=None):
    """Construit et enregistre le fichier Excel. Retourne le chemin réel du fichier."""
    if not isdir(dossier_sortie):
        raise ErreurExport("Dossier de sortie introuvable : " + dossier_sortie)

    if not os.access(dossier_sortie, os.W_OK):
        raise ErreurExport("Écriture interdite dans le dossier de sortie : " + dossier_sortie)

    chemin = join(dossier_sortie, construire_nom_de_fichier(nom_fichier))

    classeur = Workbook()

    police_entete = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    fond_entete = PatternFill("solid", fgColor="2F5496")
    police_cellule = Font(name="Arial", size=10)
    bordure = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    alignement_entete = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alignement_cellule = Alignment(vertical="center", wrap_text=True)

    def styliser_entete(feuille):
        for cellule in feuille[1]:
            cellule.font = police_entete
            cellule.fill = fond_entete
            cellule.alignment = alignement_entete
            cellule.border = bordure

    def styliser_donnees(feuille):
        if feuille.max_row < 2:
            return
        for ligne in feuille.iter_rows(min_row=2, max_row=feuille.max_row,
                                       max_col=feuille.max_column):
            for cellule in ligne:
                cellule.font = police_cellule
                cellule.alignment = alignement_cellule
                cellule.border = bordure

    def ajuster_largeurs(feuille, minimum=10, maximum=50):
        numero_colonne = 1
        while numero_colonne <= feuille.max_column:
            longueur_max = 0
            for ligne in feuille.iter_rows(min_row=1, max_row=feuille.max_row,
                                           min_col=numero_colonne,
                                           max_col=numero_colonne):
                for cellule in ligne:
                    if cellule.value is not None:
                        longueur = len(str(cellule.value))
                        if longueur > longueur_max:
                            longueur_max = longueur
            largeur = longueur_max + 2
            if largeur < minimum:
                largeur = minimum
            if largeur > maximum:
                largeur = maximum
            feuille.column_dimensions[get_column_letter(numero_colonne)].width = largeur
            numero_colonne = numero_colonne + 1

    # Feuille 1 : Entreprises
    feuille_entreprises = classeur.active
    feuille_entreprises.title = "Entreprises"
    feuille_entreprises.append(COLONNES_ENTREPRISE)
    styliser_entete(feuille_entreprises)

    for entreprise, contacts_du_lead in donnees_leads:
        valeurs = []
        for nom_colonne in COLONNES_ENTREPRISE:
            valeurs.append(entreprise.get(nom_colonne, ""))
        ajouter_ligne(feuille_entreprises, valeurs)

    styliser_donnees(feuille_entreprises)
    ajuster_largeurs(feuille_entreprises)
    feuille_entreprises.auto_filter.ref = feuille_entreprises.dimensions
    feuille_entreprises.freeze_panes = "A2"

    # Feuille 2 : Contacts
    feuille_contacts = classeur.create_sheet("Contacts")
    feuille_contacts.append(COLONNES_CONTACT)
    styliser_entete(feuille_contacts)

    nombre_total_contacts = 0
    for entreprise, contacts_du_lead in donnees_leads:
        for contact in contacts_du_lead:
            valeurs = [
                entreprise.get("ID Lead", ""),
                entreprise.get("Titre", ""),
                contact.get("Nom", ""),
                contact.get("Fonction", ""),
                contact.get("Téléphone", ""),
                contact.get("Email", ""),
                contact.get("Source LinkedIn", ""),
            ]
            ajouter_ligne(feuille_contacts, valeurs)
            nombre_total_contacts = nombre_total_contacts + 1

    styliser_donnees(feuille_contacts)
    ajuster_largeurs(feuille_contacts)
    feuille_contacts.auto_filter.ref = feuille_contacts.dimensions
    feuille_contacts.freeze_panes = "A2"

    # Feuille 3 : Anomalies (tout ce qui mérite une vérification humaine)
    feuille_anomalies = classeur.create_sheet("Anomalies")
    feuille_anomalies.append(COLONNES_ANOMALIE)
    styliser_entete(feuille_anomalies)

    for anomalie in ANOMALIES:
        valeurs = []
        for nom_colonne in COLONNES_ANOMALIE:
            valeurs.append(anomalie.get(nom_colonne, ""))
        ajouter_ligne(feuille_anomalies, valeurs)

    styliser_donnees(feuille_anomalies)
    ajuster_largeurs(feuille_anomalies, 10, 80)
    feuille_anomalies.auto_filter.ref = feuille_anomalies.dimensions
    feuille_anomalies.freeze_panes = "A2"

    # Feuille 4 : Résumé
    feuille_resume = classeur.create_sheet("Résumé")
    feuille_resume.append(["Métrique", "Valeur"])
    styliser_entete(feuille_resume)

    lignes_resume = [
        ["Date d'export", datetime.now().strftime("%d/%m/%Y %H:%M")],
        ["Total leads exportés", len(donnees_leads)],
        ["Total contacts extraits", nombre_total_contacts],
        ["Leads reçus depuis l'API", STATS["leads_recus"]],
        ["Leads hors étape ignorés", STATS["leads_hors_etape"]],
        ["Leads ignorés (illisibles)", STATS["leads_ignores"]],
        ["Leads sans description", STATS["leads_sans_description"]],
        ["Leads sans aucun contact", STATS["leads_sans_contact"]],
        ["Anomalies détectées", len(ANOMALIES)],
        ["Avertissements", STATS["avertissements"]],
        ["Erreurs", STATS["erreurs"]],
        ["Appels API effectués", STATS["appels_api"]],
        ["Tentatives API échouées", STATS["tentatives_echouees"]],
    ]
    for ligne in lignes_resume:
        ajouter_ligne(feuille_resume, ligne)

    styliser_donnees(feuille_resume)
    ajuster_largeurs(feuille_resume)

    return enregistrer_classeur(classeur, chemin)


# ──────────────────────────────────────────────
# 13. Programme principal
# ──────────────────────────────────────────────

def afficher_titre():
    print("")
    print("=" * 60)
    print("  Export leads noCRM -> Excel")
    print("=" * 60)
    print("")


def afficher_bilan(chemin_fichier):
    print("")
    print("-" * 60)
    print("  Bilan de l'exécution")
    print("-" * 60)
    print("  Appels API              : " + str(STATS["appels_api"]))
    print("  Leads reçus             : " + str(STATS["leads_recus"]))
    print("  Leads hors étape        : " + str(STATS["leads_hors_etape"]))
    print("  Leads ignorés           : " + str(STATS["leads_ignores"]))
    print("  Leads sans description  : " + str(STATS["leads_sans_description"]))
    print("  Leads sans contact      : " + str(STATS["leads_sans_contact"]))
    print("  Anomalies détectées     : " + str(len(ANOMALIES)))
    print("  Avertissements          : " + str(STATS["avertissements"]))
    print("  Erreurs                 : " + str(STATS["erreurs"]))
    print("")
    print("  Fichier Excel : " + str(chemin_fichier))
    if JOURNAL["actif"]:
        print("  Journal       : " + str(JOURNAL["chemin"]))
    print("-" * 60)
    print("")


def executer():
    """Déroulé complet de l'export. Lève une erreur explicite en cas de problème."""

    # ⚠️ CONFIGURER ICI
    STEP_ID = 267810        # ID de l'étape à exporter (None = toutes les étapes)
    MAX_LEADS = None        # None = tous les leads, ou un entier positif
    LIMITE_PAR_PAGE = 100   # nombre de leads demandés par appel API

    valider_parametres_export(STEP_ID, MAX_LEADS, LIMITE_PAR_PAGE)

    cle_api, sous_domaine = charger_configuration()
    succes("Configuration valide (sous-domaine : " + sous_domaine + ")")

    base_url = "https://" + sous_domaine + ".nocrm.io/api/v2"
    entetes = {
        "X-API-KEY": cle_api,
        "Content-Type": "application/json",
    }

    info("Récupération des leads (step_id=" + str(STEP_ID)
         + ", max=" + str(MAX_LEADS) + ")...")
    leads = recuperer_tous_les_leads(base_url, entetes, STEP_ID,
                                     LIMITE_PAR_PAGE, MAX_LEADS)

    if len(leads) == 0:
        avertissement(
            "Aucun lead récupéré. Vérifiez que l'étape " + str(STEP_ID)
            + " existe et contient bien des leads."
        )
        return None

    succes(str(len(leads)) + " lead(s) récupéré(s).")

    info("Extraction des données...")
    donnees_leads = traiter_les_leads(leads)

    if len(donnees_leads) == 0:
        raise ErreurExport(
            "Aucun lead n'a pu être exploité : le fichier Excel serait vide.\n"
            "  Consultez le journal pour le détail des erreurs."
        )

    nombre_contacts = 0
    for entreprise, contacts in donnees_leads:
        nombre_contacts = nombre_contacts + len(contacts)

    succes(str(len(donnees_leads)) + " entreprise(s), "
           + str(nombre_contacts) + " contact(s).")

    if nombre_contacts == 0:
        avertissement(
            "Aucun contact extrait. Le format des descriptions noCRM ne "
            "correspond peut-être pas à celui attendu (blocs 'Nom :' séparés "
            "par des lignes de tirets)."
        )

    info("Génération du fichier Excel...")
    chemin_fichier = exporter_vers_excel(donnees_leads, DOSSIER_SCRIPT)
    succes("Fichier généré : " + chemin_fichier)

    return chemin_fichier


def main():
    afficher_titre()
    preparer_journal()

    try:
        chemin_fichier = executer()

    except ErreurConfiguration as detail:
        erreur("Problème de configuration.")
        print("")
        print(str(detail))
        print("")
        return CODE_ERREUR

    except ErreurAPI as detail:
        erreur("Problème avec l'API noCRM.")
        print("")
        print(str(detail))
        print("")
        return CODE_ERREUR

    except ErreurExport as detail:
        erreur("Problème lors de la création du fichier Excel.")
        print("")
        print(str(detail))
        print("")
        return CODE_ERREUR

    except KeyboardInterrupt:
        print("")
        avertissement("Exécution interrompue par l'utilisateur (Ctrl+C).")
        return CODE_INTERRUPTION

    except Exception as detail:
        # Filet de sécurité : aucune trace Python brute à l'écran, mais elle
        # est conservée en entier dans le journal.
        erreur("Erreur inattendue : " + type(detail).__name__ + " : " + str(detail))
        ecrire_dans_journal(traceback.format_exc())
        print("")
        print("Le détail technique a été écrit dans le journal :")
        print("  " + str(JOURNAL["chemin"]))
        print("")
        return CODE_ERREUR

    if chemin_fichier is None:
        afficher_bilan("aucun (pas de données à exporter)")
        return CODE_ERREUR

    afficher_bilan(chemin_fichier)

    if STATS["erreurs"] > 0 or STATS["avertissements"] > 0 or len(ANOMALIES) > 0:
        print("Des points à vérifier ont été signalés : voir la feuille "
              "'Anomalies' du fichier Excel et le journal.")
        print("")

    return CODE_SUCCES


if __name__ == "__main__":
    sys.exit(main())
