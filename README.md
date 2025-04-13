# Gestion des contrats - SELARL Anesthésistes Mathilde

Application de gestion intégrée pour la SELARL des Anesthésistes de la Clinique Mathilde.

## Fonctionnalités

L'application centralise la gestion des contrats, factures et bulletins de salaire avec les fonctionnalités suivantes :

### Gestion des contrats

- Création de contrats de remplacement pour les Médecins Anesthésistes Réanimateurs (MAR)
- Création de contrats CDD pour les Infirmiers Anesthésistes Diplômés d'État (IADE)
- Gestion des médecins titulaires et remplaçants
- Gestion des IADE remplaçants
- Génération automatique de documents au format Word et PDF
- Intégration avec DocuSign pour la signature électronique

### Gestion des factures

- Analyse automatique des factures
- Extraction des informations pertinentes
- Calcul des montants à payer
- Préparation des virements

### Gestion des bulletins de salaire

- Stockage et organisation des bulletins de salaire
- Consultation par période et par salarié
- Scan et intégration de nouveaux bulletins

## Installation

### Prérequis

- Python 3.8 ou supérieur
- Bibliothèques requises :
  - tkinter
  - pandas
  - openpyxl
  - python-docx
  - tkcalendar
  - pillow
  - reportlab
  - pytesseract (avec Tesseract OCR installé)

### Installation des dépendances

```bash
pip install pandas openpyxl python-docx tkcalendar pillow reportlab pytesseract
```

### Configuration

1. Lancez l'application pour la première fois
2. Configurez les chemins des fichiers dans l'onglet "Paramètres > Chemins"
3. Configurez les paramètres DocuSign si nécessaire

## Utilisation

### Démarrage de l'application

Plusieurs méthodes sont disponibles pour lancer l'application :

#### Méthode 1 : Script de lancement (recommandé)

Sur macOS :
1. Double-cliquez sur le fichier `lancer_application.command`
2. Si c'est la première utilisation, vous devrez peut-être autoriser l'exécution dans les paramètres de sécurité

Sur tous les systèmes :
```bash
python3 lancer_application.py
```

#### Méthode 2 : Lancement direct

```bash
python3 main.py
```

### Création d'un contrat MAR

1. Accédez à l'onglet "Contrats"
2. Sélectionnez "Contrats MAR"
3. Cliquez sur "Nouveau contrat MAR"
4. Remplissez les informations requises
5. Générez le contrat au format Word et/ou PDF
6. Envoyez le contrat pour signature via DocuSign (optionnel)

### Création d'un contrat IADE

1. Accédez à l'onglet "Contrats"
2. Sélectionnez "Contrats IADE"
3. Cliquez sur "Nouveau contrat IADE"
4. Remplissez les informations requises
5. Générez le contrat au format Word et/ou PDF
6. Envoyez le contrat pour signature via DocuSign (optionnel)

### Analyse des factures

1. Accédez à l'onglet "Factures"
2. Cliquez sur "Analyser les factures"
3. Sélectionnez le dossier contenant les factures à analyser
4. Consultez les résultats de l'analyse
5. Exportez les résultats si nécessaire

### Consultation des bulletins de salaire

1. Accédez à l'onglet "Bulletins"
2. Cliquez sur "Consulter les bulletins"
3. Filtrez par période et/ou par salarié
4. Consultez les bulletins sélectionnés

## Structure des fichiers

- `main.py` : Point d'entrée principal de l'application
- `ui_styles.py` : Styles d'interface utilisateur communs
- `utils.py` : Fonctions utilitaires communes
- `widgets.py` : Widgets personnalisés réutilisables
- `contrat53.py` : Module de gestion des contrats
- `analyse_facture.py` : Module d'analyse des factures
- `bulletins.py` : Module de gestion des bulletins de salaire
- `file_paths.json` : Configuration des chemins de fichiers
- `config.json` : Configuration générale de l'application
- `icons/` : Dossier contenant les icônes de l'application
- `logs/` : Dossier contenant les fichiers de journalisation

## Maintenance

### Sauvegarde

L'application crée automatiquement des sauvegardes des fichiers importants avant modification. Ces sauvegardes sont stockées dans un sous-dossier `backups` à côté du fichier original.

### Journalisation

Les journaux de l'application sont stockés dans le dossier `logs`. Ils contiennent des informations détaillées sur les opérations effectuées et les erreurs rencontrées.

## Support

Pour toute question ou problème, veuillez contacter l'administrateur système.

## Licence

Cette application est destinée à un usage interne uniquement. Tous droits réservés.
