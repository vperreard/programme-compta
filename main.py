"""
Application de gestion des contrats - Point d'entrée principal
Ce module centralise l'accès à toutes les fonctionnalités de l'application
et fournit une interface utilisateur unifiée.
"""

import os
import sys
import tkinter as tk
from tkinter import ttk
import threading

# Import des modules personnalisés
from ui_styles import (
    COLORS, FONTS, configure_ttk_style, create_button, 
    create_label, create_frame
)
from utils import (
    logger, show_error, show_info, show_success, show_warning,
    ask_question, load_settings, save_settings, get_file_path,
    update_file_path, ensure_directory_exists, center_window
)
from widgets import (
    HeaderFrame, FooterFrame, TabView, StatusBar
)

# Import des modules fonctionnels
import analyse_facture
import bulletins
import contrat53_new as contrat53

class MainApplication(tk.Tk):
    """
    Application principale qui intègre tous les modules.
    """
    def __init__(self):
        """
        Initialise l'application principale.
        """
        super().__init__()
        
        # Configuration de la fenêtre principale
        self.title("Gestion des contrats - SELARL Anesthésistes Mathilde")
        self.geometry("1200x800")
        self.minsize(800, 600)
        
        # Configurer le style ttk
        self.style = configure_ttk_style()
        
        # Vérifier les chemins de fichiers
        self.check_file_paths()
        
        # Créer l'interface
        self.create_widgets()
        
        # Centrer la fenêtre
        center_window(self, 1200, 800)
        
        # Journalisation
        logger.info("Application démarrée")
    
    def check_file_paths(self):
        """
        Vérifie que tous les chemins de fichiers nécessaires sont définis.
        """
        settings = load_settings()
        
        # Liste des chemins requis
        required_paths = {
            "dossier_factures": "Dossier des factures",
            "bulletins_salaire": "Dossier des bulletins de salaire",
            "excel_mar": "Fichier Excel MAR",
            "excel_iade": "Fichier Excel IADE",
            "word_mar": "Modèle Word MAR",
            "word_iade": "Modèle Word IADE",
            "pdf_mar": "Dossier PDF Contrats MAR",
            "pdf_iade": "Dossier PDF Contrats IADE",
            "excel_salaries": "Fichier Excel Salariés"
        }
        
        # Vérifier chaque chemin
        missing_paths = []
        for key, description in required_paths.items():
            if key not in settings or not settings[key]:
                missing_paths.append(description)
        
        # Si des chemins sont manquants, afficher un avertissement
        if missing_paths:
            message = "Certains chemins de fichiers ne sont pas définis :\n"
            message += "\n".join([f"- {path}" for path in missing_paths])
            message += "\n\nVeuillez les configurer dans les paramètres."
            
            show_warning("Configuration incomplète", message)
    
    def create_widgets(self):
        """
        Crée les widgets de l'interface principale.
        """
        # Cadre principal
        self.main_frame = tk.Frame(self, bg=COLORS["secondary"])
        self.main_frame.pack(fill="both", expand=True)
        
        # En-tête
        self.header = HeaderFrame(
            self.main_frame, 
            "Gestion des contrats - SELARL Anesthésistes Mathilde",
            "Système de gestion intégré"
        )
        self.header.pack(fill="x")
        
        # Barre de statut (créée avant les onglets pour être disponible)
        self.status_bar = StatusBar(self.main_frame)
        self.status_bar.pack(fill="x", side="bottom")
        self.status_bar.set_message("Prêt")
        
        # Onglets
        self.tabs = TabView(self.main_frame)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Onglet Accueil
        self.home_tab = self.tabs.add_tab("Accueil", bg=COLORS["secondary"])
        self.create_home_tab()
        
        # Onglet Contrats
        self.contracts_tab = self.tabs.add_tab("Contrats", bg=COLORS["secondary"])
        self.create_contracts_tab()
        
        # Onglet Comptabilité
        self.accounting_tab = self.tabs.add_tab("Comptabilité", bg=COLORS["secondary"])
        self.create_accounting_tab()
        
        # Onglet Factures
        self.invoices_tab = self.tabs.add_tab("Factures", bg=COLORS["secondary"])
        self.create_invoices_tab()
        
        # Onglet Bulletins
        self.bulletins_tab = self.tabs.add_tab("Bulletins", bg=COLORS["secondary"])
        self.create_bulletins_tab()
        
        # Onglet Paramètres
        self.settings_tab = self.tabs.add_tab("Paramètres", bg=COLORS["secondary"])
        self.create_settings_tab()
    
    def create_home_tab(self):
        """
        Crée le contenu de l'onglet Accueil.
        """
        # Titre
        create_label(
            self.home_tab, 
            "Bienvenue dans l'application de gestion des contrats",
            style="subtitle"
        ).pack(pady=20)
        
        # Description
        description = (
            "Cette application centralise la gestion des contrats, factures et bulletins de salaire "
            "pour la SELARL des Anesthésistes de la Clinique Mathilde.\n\n"
            "Utilisez les onglets ci-dessus pour accéder aux différentes fonctionnalités."
        )
        
        create_label(
            self.home_tab, 
            description,
            style="body"
        ).pack(pady=10, padx=50)
        
        # Cadre pour les raccourcis
        shortcuts_frame = create_frame(self.home_tab)
        shortcuts_frame.pack(pady=30, fill="x")
        
        # Titre des raccourcis
        create_label(
            shortcuts_frame, 
            "Accès rapide",
            style="subtitle"
        ).pack(pady=10)
        
        # Grille de boutons
        buttons_frame = create_frame(shortcuts_frame)
        buttons_frame.pack(pady=10)
        
        # Boutons de raccourci
        shortcuts = [
            ("Nouveau contrat MAR", lambda: self.tabs.select_tab("Contrats"), "primary"),
            ("Nouveau contrat IADE", lambda: self.tabs.select_tab("Contrats"), "primary"),
            ("Analyser factures", lambda: self.tabs.select_tab("Factures"), "info"),
            ("Consulter bulletins", lambda: self.tabs.select_tab("Bulletins"), "info"),
            ("Paramètres", lambda: self.tabs.select_tab("Paramètres"), "neutral")
        ]
        
        for i, (text, command, style) in enumerate(shortcuts):
            create_button(
                buttons_frame, 
                text=text, 
                command=command,
                style=style,
                width=20,
                height=2
            ).grid(row=i//3, column=i%3, padx=10, pady=10)
    
    def create_contracts_tab(self):
        """
        Crée le contenu de l'onglet Contrats.
        """
        # Sous-onglets pour les différents types de contrats
        contracts_tabs = TabView(self.contracts_tab)
        contracts_tabs.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Sous-onglet MAR
        mar_tab = contracts_tabs.add_tab("Contrats MAR", bg=COLORS["secondary"])
        
        # Boutons pour les contrats MAR
        create_button(
            mar_tab, 
            text="Nouveau contrat MAR", 
            command=self.open_contract_creation_mar,
            style="primary",
            width=25,
            height=2
        ).pack(pady=20)
        
        create_button(
            mar_tab, 
            text="Gérer les contrats MAR existants", 
            command=self.manage_mar_contracts,
            style="info",
            width=25,
            height=2
        ).pack(pady=10)
        
        # Sous-onglet IADE
        iade_tab = contracts_tabs.add_tab("Contrats IADE", bg=COLORS["secondary"])
        
        # Boutons pour les contrats IADE
        create_button(
            iade_tab, 
            text="Nouveau contrat IADE", 
            command=self.open_contract_creation_iade,
            style="primary",
            width=25,
            height=2
        ).pack(pady=20)
        
        create_button(
            iade_tab, 
            text="Gérer les contrats IADE existants", 
            command=self.manage_iade_contracts,
            style="info",
            width=25,
            height=2
        ).pack(pady=10)
    
    def create_invoices_tab(self):
        """
        Crée le contenu de l'onglet Factures.
        """
        # Cadre pour l'analyse des factures (occupe tout l'espace)
        self.invoices_frame = create_frame(self.invoices_tab)
        self.invoices_frame.pack(fill="both", expand=True, padx=0, pady=0)
        
        # Ouvrir directement l'interface d'analyse des factures
        self.open_invoice_analysis()
    
    def create_bulletins_tab(self):
        """
        Crée le contenu de l'onglet Bulletins.
        """
        # Cadre pour les boutons
        buttons_frame = create_frame(self.bulletins_tab)
        buttons_frame.pack(pady=20)
        
        # Bouton pour consulter les bulletins
        create_button(
            buttons_frame, 
            text="Consulter les bulletins", 
            command=self.open_bulletins,
            style="primary",
            width=25,
            height=2
        ).pack(pady=10)
        
        # Bouton pour scanner un nouveau bulletin
        create_button(
            buttons_frame, 
            text="Scanner un nouveau bulletin", 
            command=self.scan_new_bulletin,
            style="info",
            width=25,
            height=2
        ).pack(pady=10)
        
        # Cadre pour les bulletins (initialement vide)
        self.bulletins_frame = create_frame(self.bulletins_tab)
        self.bulletins_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    def create_accounting_tab(self):
        """
        Crée le contenu de l'onglet Comptabilité.
        """
        # Titre
        create_label(
            self.accounting_tab, 
            "Gestion comptable",
            style="subtitle"
        ).pack(pady=20)
        
        # Cadre pour les boutons
        buttons_frame = create_frame(self.accounting_tab)
        buttons_frame.pack(pady=20)
        
        # Boutons pour les différentes fonctionnalités
        accounting_buttons = [
            ("Bulletins de salaire", self.open_bulletins_accounting, "primary"),
            ("Frais et factures", self.open_invoice_analysis, "primary"),
            ("Effectuer un virement", self.open_transfer, "info"),
            ("Virement rémunération MARS", self.open_virement_mar, "info")
        ]
        
        for text, command, style in accounting_buttons:
            create_button(
                buttons_frame, 
                text=text, 
                command=command,
                style=style,
                width=25,
                height=2
            ).pack(pady=10)
        
        # Cadre pour le contenu (initialement vide)
        self.accounting_content_frame = create_frame(self.accounting_tab)
        self.accounting_content_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    def create_settings_tab(self):
        """
        Crée le contenu de l'onglet Paramètres.
        """
        # Sous-onglets pour les différents types de paramètres
        settings_tabs = TabView(self.settings_tab)
        settings_tabs.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Sous-onglet Chemins
        paths_tab = settings_tabs.add_tab("Chemins", bg=COLORS["secondary"])
        self.create_paths_tab(paths_tab)
        
        # Sous-onglet Médecins
        doctors_tab = settings_tabs.add_tab("Médecins", bg=COLORS["secondary"])
        self.create_doctors_tab(doctors_tab)
        
        # Sous-onglet IADE
        iade_tab = settings_tabs.add_tab("IADE", bg=COLORS["secondary"])
        self.create_iade_tab(iade_tab)
        
        # Sous-onglet Salariés
        employees_tab = settings_tabs.add_tab("Salariés", bg=COLORS["secondary"])
        self.create_employees_tab(employees_tab)
        
        # Sous-onglet DocuSign
        docusign_tab = settings_tabs.add_tab("DocuSign", bg=COLORS["secondary"])
        self.create_docusign_tab(docusign_tab)
    
    def create_paths_tab(self, parent):
        """
        Crée le contenu du sous-onglet Chemins.
        """
        from widgets import FileSelector
        
        # Titre
        create_label(
            parent, 
            "Configuration des chemins de fichiers",
            style="subtitle"
        ).pack(pady=10)
        
        # Cadre avec scrollbar
        canvas = tk.Canvas(parent, bg=COLORS["secondary"])
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = create_frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y")
        
        # Charger les paramètres
        settings = load_settings()
        
        # Liste des chemins à configurer
        paths = [
            ("Chemin Excel MAR", "excel_mar", "file"),
            ("Chemin Excel IADE", "excel_iade", "file"),
            ("Chemin Word MAR", "word_mar", "file"),
            ("Chemin Word IADE", "word_iade", "file"),
            ("Dossier PDF Contrats MAR", "pdf_mar", "directory"),
            ("Dossier PDF Contrats IADE", "pdf_iade", "directory"),
            ("Dossier Bulletins de salaire", "bulletins_salaire", "directory"),
            ("Dossier Frais/Factures", "dossier_factures", "directory"),
            ("Chemin Excel Salariés", "excel_salaries", "file")
        ]
        
        # Variables pour les chemins
        self.path_vars = {}
        
        # Créer les sélecteurs de fichiers
        for i, (label, key, file_type) in enumerate(paths):
            var = tk.StringVar(value=settings.get(key, ""))
            self.path_vars[key] = var
            
            selector = FileSelector(
                scrollable_frame, 
                label, 
                variable=var,
                file_type=file_type
            )
            selector.pack(fill="x", pady=5)
        
        # Bouton de sauvegarde
        create_button(
            parent, 
            text="Enregistrer", 
            command=self.save_paths,
            style="accent",
            width=15
        ).pack(pady=10)
    
    def create_doctors_tab(self, parent):
        """
        Crée le contenu du sous-onglet Médecins.
        """
        # Sous-onglets pour les différents types de médecins
        doctors_tabs = TabView(parent)
        doctors_tabs.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Sous-onglet MAR titulaires
        mar_tab = doctors_tabs.add_tab("MAR titulaires", bg=COLORS["secondary"])
        
        # Bouton pour gérer les MAR titulaires
        create_button(
            mar_tab, 
            text="Gérer les MAR titulaires", 
            command=self.manage_mar_titulaires,
            style="primary",
            width=25
        ).pack(pady=20)
        
        # Sous-onglet MAR remplaçants
        mar_remp_tab = doctors_tabs.add_tab("MAR remplaçants", bg=COLORS["secondary"])
        
        # Bouton pour gérer les MAR remplaçants
        create_button(
            mar_remp_tab, 
            text="Gérer les MAR remplaçants", 
            command=self.manage_mar_remplacants,
            style="primary",
            width=25
        ).pack(pady=20)
    
    def create_iade_tab(self, parent):
        """
        Crée le contenu du sous-onglet IADE.
        """
        # Bouton pour gérer les IADE remplaçants
        create_button(
            parent, 
            text="Gérer les IADE remplaçants", 
            command=self.manage_iade_remplacants,
            style="primary",
            width=25
        ).pack(pady=20)
    
    def create_employees_tab(self, parent):
        """
        Crée le contenu du sous-onglet Salariés.
        """
        # Bouton pour gérer les salariés
        create_button(
            parent, 
            text="Gérer les salariés", 
            command=self.manage_salaries,
            style="primary",
            width=25
        ).pack(pady=20)
    
    def create_docusign_tab(self, parent):
        """
        Crée le contenu du sous-onglet DocuSign.
        """
        from widgets import FormField
        import json
        
        # Titre
        create_label(
            parent, 
            "Configuration DocuSign",
            style="subtitle"
        ).pack(pady=10)
        
        # Charger la configuration
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
        if os.path.exists(config_file):
            with open(config_file, "r") as f:
                config = json.load(f)
        else:
            config = {}
        
        # Variables pour les champs
        self.docusign_vars = {
            "login_page": tk.StringVar(value=config.get("docusign_login_page", "https://account.docusign.com")),
            "email": tk.StringVar(value=config.get("docusign_email", "")),
            "password": tk.StringVar(value=config.get("docusign_password", ""))
        }
        
        # Champs de formulaire
        FormField(
            parent, 
            "Page de login DocuSign :", 
            variable=self.docusign_vars["login_page"]
        ).pack(fill="x", padx=50, pady=5)
        
        FormField(
            parent, 
            "Email DocuSign :", 
            variable=self.docusign_vars["email"]
        ).pack(fill="x", padx=50, pady=5)
        
        FormField(
            parent, 
            "Mot de passe (laisser vide si non stocké) :", 
            variable=self.docusign_vars["password"]
        ).pack(fill="x", padx=50, pady=5)
        
        # Bouton de sauvegarde
        create_button(
            parent, 
            text="Enregistrer", 
            command=self.save_docusign,
            style="accent",
            width=15
        ).pack(pady=20)
    
    def save_paths(self):
        """
        Sauvegarde les chemins de fichiers.
        """
        settings = load_settings()
        
        # Mettre à jour les chemins
        for key, var in self.path_vars.items():
            settings[key] = var.get()
        
        # Sauvegarder les paramètres
        if save_settings(settings):
            show_success("Paramètres", "Les chemins ont été enregistrés avec succès.")
        else:
            show_error("Paramètres", "Erreur lors de l'enregistrement des chemins.")
    
    def save_docusign(self):
        """
        Sauvegarde les paramètres DocuSign.
        """
        import json
        
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
        
        # Charger la configuration existante
        if os.path.exists(config_file):
            with open(config_file, "r") as f:
                config = json.load(f)
        else:
            config = {}
        
        # Mettre à jour les paramètres DocuSign
        config["docusign_login_page"] = self.docusign_vars["login_page"].get()
        config["docusign_email"] = self.docusign_vars["email"].get()
        config["docusign_password"] = self.docusign_vars["password"].get()
        
        # Sauvegarder la configuration
        try:
            with open(config_file, "w") as f:
                json.dump(config, f, indent=4)
            
            show_success("Paramètres", "Les paramètres DocuSign ont été enregistrés avec succès.")
        except Exception as e:
            show_error("Paramètres", f"Erreur lors de l'enregistrement des paramètres DocuSign : {str(e)}")
    
    def open_contract_creation_mar(self):
        """
        Ouvre l'interface de création de contrat MAR.
        """
        self.status_bar.set_message("Ouverture de l'interface de création de contrat MAR...")
        
        # Nettoyer l'interface existante
        for widget in self.contracts_tab.winfo_children():
            if isinstance(widget, TabView):
                continue
            widget.destroy()
        
        # Créer un cadre pour l'interface de création de contrat
        contract_frame = create_frame(self.contracts_tab)
        contract_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        try:
            # Appeler la fonction de création de contrat MAR avec le frame créé
            contrat53.create_contract_interface_mar(contract_frame)
            self.status_bar.set_message("Interface de création de contrat MAR ouverte.")
        except Exception as e:
            show_error("Erreur", f"Impossible d'ouvrir l'interface de création de contrat MAR: {str(e)}")
            self.status_bar.set_message("Erreur lors de l'ouverture de l'interface de création de contrat MAR.")
    
    def open_contract_creation_iade(self):
        """
        Ouvre l'interface de création de contrat IADE.
        """
        self.status_bar.set_message("Ouverture de l'interface de création de contrat IADE...")
        
        # Nettoyer l'interface existante
        for widget in self.contracts_tab.winfo_children():
            if isinstance(widget, TabView):
                continue
            widget.destroy()
        
        # Créer un cadre pour l'interface de création de contrat
        contract_frame = create_frame(self.contracts_tab)
        contract_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        try:
            # Appeler la fonction de création de contrat IADE avec le frame créé
            contrat53.create_contract_interface_iade(contract_frame)
            self.status_bar.set_message("Interface de création de contrat IADE ouverte.")
        except Exception as e:
            show_error("Erreur", f"Impossible d'ouvrir l'interface de création de contrat IADE: {str(e)}")
            self.status_bar.set_message("Erreur lors de l'ouverture de l'interface de création de contrat IADE.")
    
    def manage_mar_contracts(self):
        """
        Ouvre l'interface de gestion des contrats MAR existants.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des contrats MAR...")
        
        # À implémenter
        show_info("Fonctionnalité à venir", "La gestion des contrats MAR existants sera disponible prochainement.")
        
        self.status_bar.set_message("Prêt")
    
    def manage_iade_contracts(self):
        """
        Ouvre l'interface de gestion des contrats IADE existants.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des contrats IADE...")
        
        # À implémenter
        show_info("Fonctionnalité à venir", "La gestion des contrats IADE existants sera disponible prochainement.")
        
        self.status_bar.set_message("Prêt")
    
    def open_invoice_analysis(self):
        """
        Ouvre l'interface d'analyse des factures.
        """
        self.status_bar.set_message("Ouverture de l'interface d'analyse des factures...")
        
        # Nettoyer l'interface existante
        for widget in self.invoices_frame.winfo_children():
            widget.destroy()
        
        # Créer un cadre pour l'interface d'analyse des factures
        analysis_frame = create_frame(self.invoices_frame)
        analysis_frame.pack(fill="both", expand=True)
        
        # Créer l'instance de l'analyseur
        factures_path = get_file_path("dossier_factures", verify_exists=True)
        analyseur = analyse_facture.AnalyseFactures(factures_path)
        
        # Créer l'interface de l'analyseur dans le cadre
        analyseur.creer_interface(analysis_frame)
        
        self.status_bar.set_message("Interface d'analyse des factures ouverte.")
    
    def open_transfer(self):
        """
        Ouvre l'interface de virement.
        """
        self.status_bar.set_message("Ouverture de l'interface de virement...")
        
        # À implémenter
        show_info("Fonctionnalité à venir", "L'interface de virement sera disponible prochainement.")
        
        self.status_bar.set_message("Prêt")
    
    def open_bulletins(self):
        """
        Ouvre l'interface de consultation des bulletins.
        """
        self.status_bar.set_message("Ouverture de l'interface de consultation des bulletins...")
        
        # Nettoyer l'interface existante
        for widget in self.bulletins_frame.winfo_children():
            widget.destroy()
        
        # Créer un cadre pour l'interface de consultation des bulletins
        bulletins_frame = create_frame(self.bulletins_frame)
        bulletins_frame.pack(fill="both", expand=True)
        
        # Appeler la fonction de consultation des bulletins
        bulletins.show_bulletins_in_frame(bulletins_frame)
        
        self.status_bar.set_message("Interface de consultation des bulletins ouverte.")
    
    def scan_new_bulletin(self):
        """
        Ouvre l'interface de scan d'un nouveau bulletin.
        """
        self.status_bar.set_message("Ouverture de l'interface de scan d'un nouveau bulletin...")
        
        # Appeler la fonction de scan d'un nouveau bulletin
        bulletins.scan_new_pdf()
        
        self.status_bar.set_message("Prêt")
    
    def manage_mar_titulaires(self):
        """
        Ouvre l'interface de gestion des MAR titulaires.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des MAR titulaires...")
        
        # Appeler la fonction de gestion des MAR titulaires
        contrat53.manage_mar_titulaires()
        
        self.status_bar.set_message("Interface de gestion des MAR titulaires ouverte.")
    
    def manage_mar_remplacants(self):
        """
        Ouvre l'interface de gestion des MAR remplaçants.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des MAR remplaçants...")
        
        # Appeler la fonction de gestion des MAR remplaçants
        contrat53.manage_mar_remplacants()
        
        self.status_bar.set_message("Interface de gestion des MAR remplaçants ouverte.")
    
    def manage_iade_remplacants(self):
        """
        Ouvre l'interface de gestion des IADE remplaçants.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des IADE remplaçants...")
        
        # Appeler la fonction de gestion des IADE remplaçants
        contrat53.manage_iade_remplacants()
        
        self.status_bar.set_message("Interface de gestion des IADE remplaçants ouverte.")
    
    def manage_salaries(self):
        """
        Ouvre l'interface de gestion des salariés.
        """
        self.status_bar.set_message("Ouverture de l'interface de gestion des salariés...")
        
        # Appeler la fonction de gestion des salariés
        contrat53.manage_salaries()
        
        self.status_bar.set_message("Interface de gestion des salariés ouverte.")
    
    def open_bulletins_accounting(self):
        """
        Ouvre l'interface de consultation des bulletins dans l'onglet comptabilité.
        """
        self.status_bar.set_message("Ouverture de l'interface de consultation des bulletins...")
        
        # Nettoyer l'interface existante
        for widget in self.accounting_content_frame.winfo_children():
            widget.destroy()
        
        # Créer un cadre pour l'interface de consultation des bulletins
        bulletins_frame = create_frame(self.accounting_content_frame)
        bulletins_frame.pack(fill="both", expand=True)
        
        # Appeler la fonction de consultation des bulletins
        bulletins.show_bulletins_in_frame(bulletins_frame)
        
        self.status_bar.set_message("Interface de consultation des bulletins ouverte.")
    
    def open_virement_mar(self):
        """
        Ouvre l'interface de virement pour les MARS.
        """
        self.status_bar.set_message("Ouverture de l'interface de virement pour les MARS...")
        
        # Nettoyer l'interface existante
        for widget in self.accounting_content_frame.winfo_children():
            widget.destroy()
        
        # Créer un cadre pour l'interface de virement
        virement_frame = create_frame(self.accounting_content_frame)
        virement_frame.pack(fill="both", expand=True)
        
        # Titre
        create_label(
            virement_frame, 
            "Virement rémunération MARS",
            style="subtitle"
        ).pack(pady=20)
        
        # Description
        description = (
            "Cette fonctionnalité permet de générer et d'envoyer des virements pour la rémunération des MARS.\n"
            "Vous pouvez sélectionner les médecins et les montants à virer."
        )
        
        create_label(
            virement_frame, 
            description,
            style="body"
        ).pack(pady=10, padx=50)
        
        # Boutons
        buttons_frame = create_frame(virement_frame)
        buttons_frame.pack(pady=20)
        
        # Bouton pour ouvrir le fichier Excel des virements
        excel_file = get_file_path("chemin_fichier_virement", "")
        
        create_button(
            buttons_frame, 
            text="Ouvrir le fichier Excel des virements", 
            command=lambda: self.open_excel_file(excel_file),
            style="primary",
            width=30
        ).pack(pady=10)
        
        # Bouton pour générer les virements
        create_button(
            buttons_frame, 
            text="Générer et envoyer les virements", 
            command=self.generate_mar_virements,
            style="accent",
            width=30
        ).pack(pady=10)
        
        self.status_bar.set_message("Interface de virement pour les MARS ouverte.")
    
    def open_excel_file(self, file_path):
        """
        Ouvre un fichier Excel avec l'application par défaut.
        """
        if not file_path or not os.path.exists(file_path):
            show_error("Erreur", f"Le fichier {file_path} est introuvable.")
            return
        
        try:
            if sys.platform == "darwin":  # macOS
                os.system(f"open -a 'Microsoft Excel' '{file_path}'")
            elif sys.platform == "win32":  # Windows
                os.startfile(file_path)
            else:  # Linux
                os.system(f"xdg-open '{file_path}'")
        except Exception as e:
            show_error("Erreur", f"Impossible d'ouvrir le fichier Excel : {str(e)}")
    
    def generate_mar_virements(self):
        """
        Génère et envoie les virements pour les MARS.
        """
        try:
            # Appeler la fonction de génération des virements
            from generer_virement import open_virement_selection_window
            open_virement_selection_window()
            
            self.status_bar.set_message("Interface de sélection des virements ouverte.")
        except Exception as e:
            show_error("Erreur", f"Impossible de générer les virements : {str(e)}")
            self.status_bar.set_message("Erreur lors de la génération des virements.")

def main():
    """
    Point d'entrée principal de l'application.
    """
    app = MainApplication()
    app.mainloop()

if __name__ == "__main__":
    main()
