#!/usr/bin/env python3
"""
Script de lancement de l'application de gestion des contrats.
Exécutez ce script pour démarrer l'application.
"""

import os
import sys
import subprocess

def main():
    """
    Fonction principale qui lance l'application.
    """
    # Obtenir le chemin du répertoire courant
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Vérifier si les dépendances sont installées
    try:
        # Vérifier si les modules requis sont disponibles
        import tkinter
        import pandas
        import tkcalendar
    except ImportError as e:
        print(f"Erreur: Module manquant - {e}")
        print("Installation des dépendances requises...")
        
        # Installer les dépendances
        requirements_file = os.path.join(current_dir, "requirements.txt")
        if os.path.exists(requirements_file):
            subprocess.run([sys.executable, "-m", "pip", "install", "-r", requirements_file])
        else:
            print("Fichier requirements.txt non trouvé.")
            return
    
    # Lancer l'application
    print("Démarrage de l'application...")
    main_script = os.path.join(current_dir, "main.py")
    
    if os.path.exists(main_script):
        # Exécuter le script principal
        subprocess.run([sys.executable, main_script])
    else:
        print(f"Erreur: Fichier principal {main_script} non trouvé.")

if __name__ == "__main__":
    main()
