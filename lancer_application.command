#!/bin/bash
# Script de lancement pour macOS

# Obtenir le chemin du répertoire du script
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Afficher un message de bienvenue
echo "==================================================="
echo "  Lancement de l'application de gestion des contrats"
echo "  SELARL Anesthésistes Mathilde"
echo "==================================================="
echo ""

# Se déplacer dans le répertoire du script
cd "$DIR"

# Vérifier si Python est installé
if command -v python3 &>/dev/null; then
    # Lancer l'application avec Python 3
    python3 "$DIR/lancer_application.py"
else
    # Afficher un message d'erreur si Python n'est pas installé
    echo "Erreur: Python 3 n'est pas installé sur votre système."
    echo "Veuillez installer Python 3 pour utiliser cette application."
    echo ""
    echo "Appuyez sur une touche pour fermer cette fenêtre..."
    read -n 1
fi
