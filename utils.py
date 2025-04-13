"""
Module d'utilitaires communs pour l'application de gestion des contrats.
Ce module centralise les fonctions utilitaires, la gestion des erreurs et le logging
pour assurer une cohérence à travers les différents modules de l'application.
"""

import os
import json
import logging
import traceback
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import sys

# Configuration du logging
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

LOG_FILE = os.path.join(LOG_DIR, f"application_{datetime.now().strftime('%Y%m%d')}.log")

# Configuration du logger
logger = logging.getLogger("ContratApp")
logger.setLevel(logging.DEBUG)

# Handler pour fichier
file_handler = logging.FileHandler(LOG_FILE, encoding='utf-8')
file_handler.setLevel(logging.DEBUG)

# Handler pour console
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.INFO)

# Format des logs
log_format = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(log_format)
console_handler.setFormatter(log_format)

# Ajout des handlers au logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)

# Chemins des fichiers de configuration
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(SCRIPT_DIR, "file_paths.json")
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")

# Icônes pour les messages
ICONS = {
    "info": "ℹ️",
    "success": "✅",
    "warning": "⚠️",
    "error": "❌",
    "question": "❓"
}

def load_settings():
    """
    Charge les paramètres depuis le fichier de configuration.
    
    Returns:
        dict: Dictionnaire des paramètres
    """
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                settings = json.load(f)
            logger.info("Paramètres chargés avec succès")
            return settings
        else:
            logger.warning(f"Fichier de paramètres non trouvé: {SETTINGS_FILE}")
            return {}
    except Exception as e:
        logger.error(f"Erreur lors du chargement des paramètres: {str(e)}")
        return {}

def save_settings(settings):
    """
    Sauvegarde les paramètres dans le fichier de configuration.
    
    Args:
        settings (dict): Dictionnaire des paramètres à sauvegarder
    
    Returns:
        bool: True si la sauvegarde a réussi, False sinon
    """
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=4, ensure_ascii=False)
        logger.info("Paramètres sauvegardés avec succès")
        return True
    except Exception as e:
        logger.error(f"Erreur lors de la sauvegarde des paramètres: {str(e)}")
        return False

def load_config():
    """
    Charge la configuration depuis le fichier de configuration.
    
    Returns:
        dict: Dictionnaire de configuration
    """
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)
            logger.info("Configuration chargée avec succès")
            return config
        else:
            logger.warning(f"Fichier de configuration non trouvé: {CONFIG_FILE}")
            return {}
    except Exception as e:
        logger.error(f"Erreur lors du chargement de la configuration: {str(e)}")
        return {}

def save_config(config):
    """
    Sauvegarde la configuration dans le fichier de configuration.
    
    Args:
        config (dict): Dictionnaire de configuration à sauvegarder
    
    Returns:
        bool: True si la sauvegarde a réussi, False sinon
    """
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        logger.info("Configuration sauvegardée avec succès")
        return True
    except Exception as e:
        logger.error(f"Erreur lors de la sauvegarde de la configuration: {str(e)}")
        return False

def get_file_path(key, verify_exists=False, create_if_missing=False):
    """
    Récupère le chemin d'un fichier depuis les paramètres.
    
    Args:
        key (str): Clé du chemin à récupérer
        verify_exists (bool): Vérifie si le chemin existe
        create_if_missing (bool): Crée le dossier s'il n'existe pas
    
    Returns:
        str: Chemin du fichier ou None si non trouvé
    """
    settings = load_settings()
    path = settings.get(key)
    
    if not path:
        logger.warning(f"Chemin non trouvé pour la clé: {key}")
        return None
    
    # Expansion du chemin utilisateur (~)
    path = os.path.expanduser(path)
    
    if verify_exists and not os.path.exists(path):
        logger.warning(f"Le chemin n'existe pas: {path}")
        
        if create_if_missing and key.startswith("dossier_"):
            try:
                os.makedirs(path, exist_ok=True)
                logger.info(f"Dossier créé: {path}")
                return path
            except Exception as e:
                logger.error(f"Erreur lors de la création du dossier {path}: {str(e)}")
                return None
        return None
    
    return path

def update_file_path(key, path):
    """
    Met à jour le chemin d'un fichier dans les paramètres.
    
    Args:
        key (str): Clé du chemin à mettre à jour
        path (str): Nouveau chemin
    
    Returns:
        bool: True si la mise à jour a réussi, False sinon
    """
    settings = load_settings()
    settings[key] = path
    return save_settings(settings)

def ensure_directory_exists(directory_path):
    """
    S'assure qu'un répertoire existe, le crée si nécessaire.
    
    Args:
        directory_path (str): Chemin du répertoire
    
    Returns:
        str: Chemin du répertoire ou None si erreur
    """
    try:
        expanded_path = os.path.expanduser(directory_path)
        if not os.path.exists(expanded_path):
            os.makedirs(expanded_path)
            logger.info(f"Dossier créé: {expanded_path}")
        return expanded_path
    except Exception as e:
        logger.error(f"Erreur lors de la création du dossier {directory_path}: {str(e)}")
        return None

def format_error_message(error, with_traceback=False):
    """
    Formate un message d'erreur pour l'affichage.
    
    Args:
        error (Exception): L'erreur à formater
        with_traceback (bool): Inclure la trace d'appel
    
    Returns:
        str: Message d'erreur formaté
    """
    message = f"{ICONS['error']} Erreur: {str(error)}"
    
    if with_traceback:
        tb = traceback.format_exc()
        message += f"\n\nDétails techniques:\n{tb}"
    
    return message

def show_error(title, error, parent=None, with_traceback=False):
    """
    Affiche une boîte de dialogue d'erreur.
    
    Args:
        title (str): Titre de la boîte de dialogue
        error (Exception or str): L'erreur à afficher
        parent (tk.Widget): Widget parent pour la boîte de dialogue
        with_traceback (bool): Inclure la trace d'appel
    """
    if isinstance(error, Exception):
        message = format_error_message(error, with_traceback)
        logger.error(f"{title}: {str(error)}")
        if with_traceback:
            logger.error(traceback.format_exc())
    else:
        message = f"{ICONS['error']} {error}"
        logger.error(f"{title}: {error}")
    
    messagebox.showerror(title, message, parent=parent)

def show_warning(title, message, parent=None):
    """
    Affiche une boîte de dialogue d'avertissement.
    
    Args:
        title (str): Titre de la boîte de dialogue
        message (str): Message à afficher
        parent (tk.Widget): Widget parent pour la boîte de dialogue
    """
    formatted_message = f"{ICONS['warning']} {message}"
    logger.warning(f"{title}: {message}")
    messagebox.showwarning(title, formatted_message, parent=parent)

def show_info(title, message, parent=None):
    """
    Affiche une boîte de dialogue d'information.
    
    Args:
        title (str): Titre de la boîte de dialogue
        message (str): Message à afficher
        parent (tk.Widget): Widget parent pour la boîte de dialogue
    """
    formatted_message = f"{ICONS['info']} {message}"
    logger.info(f"{title}: {message}")
    messagebox.showinfo(title, formatted_message, parent=parent)

def show_success(title, message, parent=None):
    """
    Affiche une boîte de dialogue de succès.
    
    Args:
        title (str): Titre de la boîte de dialogue
        message (str): Message à afficher
        parent (tk.Widget): Widget parent pour la boîte de dialogue
    """
    formatted_message = f"{ICONS['success']} {message}"
    logger.info(f"{title}: {message}")
    messagebox.showinfo(title, formatted_message, parent=parent)

def ask_question(title, message, parent=None):
    """
    Affiche une boîte de dialogue de question.
    
    Args:
        title (str): Titre de la boîte de dialogue
        message (str): Message à afficher
        parent (tk.Widget): Widget parent pour la boîte de dialogue
    
    Returns:
        bool: True si l'utilisateur a répondu Oui, False sinon
    """
    formatted_message = f"{ICONS['question']} {message}"
    logger.info(f"Question: {title}: {message}")
    return messagebox.askyesno(title, formatted_message, parent=parent)

def save_excel_with_updated_sheet(file_path, sheet_name, updated_data):
    """
    Sauvegarde une feuille spécifique dans un fichier Excel sans supprimer les autres.
    
    Args:
        file_path (str): Chemin du fichier Excel
        sheet_name (str): Nom de la feuille à mettre à jour
        updated_data (pd.DataFrame): Données mises à jour
    
    Returns:
        bool: True si la sauvegarde a réussi, False sinon
    """
    try:
        import pandas as pd
        
        # Charger toutes les feuilles existantes
        with pd.ExcelFile(file_path, engine="openpyxl") as xls:
            sheets = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
        
        # Mettre à jour uniquement la feuille concernée
        sheets[sheet_name] = updated_data
        
        # Réécrire le fichier avec toutes les feuilles
        with pd.ExcelWriter(file_path, engine="openpyxl", mode="w") as writer:
            for sheet, df in sheets.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
        
        logger.info(f"Feuille '{sheet_name}' sauvegardée avec succès dans {file_path}")
        return True
    except Exception as e:
        logger.error(f"Erreur lors de la sauvegarde de la feuille '{sheet_name}' : {str(e)}")
        return False

def create_backup(file_path):
    """
    Crée une sauvegarde d'un fichier.
    
    Args:
        file_path (str): Chemin du fichier à sauvegarder
    
    Returns:
        str: Chemin de la sauvegarde ou None si erreur
    """
    try:
        if not os.path.exists(file_path):
            logger.warning(f"Impossible de créer une sauvegarde, le fichier n'existe pas: {file_path}")
            return None
        
        backup_dir = os.path.join(os.path.dirname(file_path), "backups")
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = os.path.basename(file_path)
        name, ext = os.path.splitext(filename)
        backup_path = os.path.join(backup_dir, f"{name}_{timestamp}{ext}")
        
        import shutil
        shutil.copy2(file_path, backup_path)
        
        logger.info(f"Sauvegarde créée: {backup_path}")
        return backup_path
    except Exception as e:
        logger.error(f"Erreur lors de la création de la sauvegarde de {file_path}: {str(e)}")
        return None

def clean_amount(value):
    """
    Nettoie et convertit une valeur monétaire extraite en float.
    
    Args:
        value (str): Valeur monétaire à nettoyer
    
    Returns:
        float: Valeur convertie ou None si erreur
    """
    if not value:
        logger.warning("Valeur vide détectée pour un montant.")
        return None
    
    try:
        # Remplace les espaces et les virgules avant conversion
        cleaned = value.replace(" ", "").replace(",", ".").replace("€", "").strip()
        return float(cleaned)
    except ValueError:
        logger.error(f"Erreur de conversion : Impossible de convertir '{value}' en float.")
        return None

def format_date(date_str, input_format="%Y-%m-%d", output_format="%d-%m-%Y"):
    """
    Convertit une date d'un format à un autre.
    
    Args:
        date_str (str): Date à convertir
        input_format (str): Format d'entrée
        output_format (str): Format de sortie
    
    Returns:
        str: Date convertie ou la date originale si erreur
    """
    try:
        date_obj = datetime.strptime(date_str, input_format)
        return date_obj.strftime(output_format)
    except ValueError:
        logger.warning(f"Format de date invalide: {date_str}")
        return date_str

def center_window(window, width=None, height=None):
    """
    Centre une fenêtre sur l'écran.
    
    Args:
        window (tk.Toplevel or tk.Tk): Fenêtre à centrer
        width (int): Largeur de la fenêtre (optionnel)
        height (int): Hauteur de la fenêtre (optionnel)
    """
    window.update_idletasks()
    
    if width is None:
        width = window.winfo_width()
    if height is None:
        height = window.winfo_height()
    
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    
    window.geometry(f'{width}x{height}+{x}+{y}')
