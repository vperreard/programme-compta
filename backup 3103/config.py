import os
import json
import logging

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Configuration du logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Définir le chemin du fichier de paramètres
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(SCRIPT_DIR, "file_paths.json")

# Chemins par défaut utilisant ~ pour le répertoire utilisateur
default_paths = {
    "pdf_mar": "~/Documents/Contrats/MAR",
    "pdf_iade": "~/Documents/Contrats/IADE",
    "excel_mar": "~/Dropbox/SEL:SPFPL Mathilde/contrat remplacement MAR/MARS SELARL.xlsx",
    "excel_iade": "~/Dropbox/SEL:SPFPL Mathilde/Compta SEL/contrats/CDD à faire signer/IADE remplaçants.xlsx",
    "word_mar": "~/Dropbox/SEL:SPFPL Mathilde/contrat remplacement MAR/CDOM Remplacement Word.docx",
    "word_iade": "~/Dropbox/SEL:SPFPL Mathilde/Compta SEL/contrats/CDD à faire signer/modèle CDD IADE SELMATHILDE.docx",
    "bulletins_salaire": "~/Documents/Bulletins_Salaire",
    "excel_salaries": "~/Dropbox/SEL:SPFPL Mathilde/Compta SEL/salaries.xlsx",
    "dossier_factures": "~/Documents/Frais_Factures",
    "chemin_fichier_virement": "~/Dropbox/SEL:SPFPL Mathilde/Compta SEL/compta SEL -rempla 2025.xlsx",
    "script_docusign": "~/script contrats/envoidocusign12.py"


}

file_paths = {}


# Ajouter une fonction pour s'assurer qu'un chemin existe
def ensure_path_exists(path, create_if_missing=False):
    """Vérifie si un chemin existe et le crée si nécessaire."""
    expanded_path = os.path.expanduser(path)
    if not os.path.exists(expanded_path):
        if create_if_missing:
            try:
                os.makedirs(expanded_path)
                logger.info(f"Dossier créé: {expanded_path}")
                return expanded_path
            except Exception as e:
                logger.error(f"Erreur lors de la création du dossier {expanded_path}: {e}")
                return None
        else:
            logger.warning(f"Chemin inexistant: {expanded_path}")
            return None
    return expanded_path


def charger_file_paths():
    """Charge les chemins depuis le fichier de configuration ou utilise les valeurs par défaut"""
    global file_paths

    try:
        # Charger le fichier s'il existe, sinon initialiser avec les valeurs par défaut
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                file_paths = json.load(f)
                logger.info(f"Fichier de configuration chargé: {SETTINGS_FILE}")
        else:
            file_paths = default_paths.copy()
            logger.info("Utilisation des chemins par défaut (fichier de configuration non trouvé)")

        # Compléter les chemins manquants avec les valeurs par défaut
        for key, value in default_paths.items():
            if key not in file_paths:
                file_paths[key] = value
                logger.info(f"Ajout du chemin par défaut pour: {key}")

        # Sauvegarder les chemins mis à jour
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(file_paths, f, indent=4, ensure_ascii=False)
            logger.info("Configuration sauvegardée")
            
    except Exception as e:
        logger.error(f"Erreur lors du chargement des chemins: {str(e)}")
        # Utiliser les valeurs par défaut en cas d'erreur
        file_paths = default_paths.copy()

def get_file_path(key, verify_exists=False, create_if_missing=False):
    """
    Récupère le chemin d'un fichier à partir de sa clé
    
    Args:
        key: Clé du chemin à récupérer
        verify_exists: Si True, vérifie si le fichier/dossier existe
        create_if_missing: Si True et verify_exists=True, crée le dossier si manquant
        
    Returns:
        Le chemin du fichier/dossier ou None si non trouvé
    """
    if not file_paths:
        charger_file_paths()
        
    if key not in file_paths:
        logger.warning(f"Clé de chemin non trouvée: {key}")
        return None
        
    path = os.path.expanduser(file_paths.get(key, ""))
    
    if verify_exists and path:
        return ensure_path_exists(path, create_if_missing)
    return path

def verify_all_paths():
    """Vérifie l'existence de tous les chemins configurés"""
    results = {}
    for key in file_paths:
        path = os.path.expanduser(file_paths[key])
        exists = os.path.exists(path)
        results[key] = {
            "path": path,
            "exists": exists
        }
        if not exists:
            logger.warning(f"Le chemin {key} n'existe pas: {path}")
    
    return results

def update_path(key, new_path):
    """Met à jour un chemin dans la configuration"""
    if not file_paths:
        charger_file_paths()
        
    file_paths[key] = new_path
    
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(file_paths, f, indent=4, ensure_ascii=False)
        logger.info(f"Chemin mis à jour: {key} -> {new_path}")
        return True
    except Exception as e:
        logger.error(f"Erreur lors de la mise à jour du chemin: {str(e)}")
        return False

# Chargement automatique des chemins lors de l'import du module
charger_file_paths()