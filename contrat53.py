import pandas as pd
from tkinter import Tk, Label, Button, Canvas, Entry, Frame, StringVar, OptionMenu, Toplevel, Listbox, END
from tkinter.filedialog import askopenfilename, askdirectory
from PIL import Image, ImageTk, ImageEnhance  
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkcalendar import Calendar
from functools import partial
from datetime import datetime
import json
import os
import subprocess
import pyautogui
import threading
import time
from threading import Thread  # Import ajouté pour gérer les threads
from datetime import datetime
import gestion_contrats1
from generer_virement import generer_xml_virements  # Import de la fonction
from config import charger_file_paths, get_file_path, ensure_path_exists, update_path, file_paths, SETTINGS_FILE


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Charger les chemins au démarrage
charger_file_paths()

factures_path = get_file_path("dossier_factures", verify_exists=True, create_if_missing=True)
bulletins_path = get_file_path("bulletins_salaire", verify_exists=True, create_if_missing=True)



# Vérifier que les chemins sont bien disponibles
if not factures_path:
    print("⚠️ Erreur : Le chemin des factures n'est pas défini.")
else:
    print(f"📂 Dossier Factures trouvé : {factures_path}")

if not bulletins_path:
    print("⚠️ Erreur : Le chemin des bulletins de salaire n'est pas défini.")
else:
    print(f"📂 Dossier Bulletins de salaire trouvé : {bulletins_path}")

# Remplacer les chemins hardcodés par les références à get_file_path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

pyautogui.FAILSAFE = False
current_frame = None  # Cette variable sera utilisée pour communiquer entre les fonctions


# Fichier pour sauvegarder les chemins de fichiers
chemin_script_virement = "/Users/vincentperreard/script contrats/generer_virement.py"



# Coordonnées des boutons (à ajuster si besoin)
x_oui, y_oui = 936, 455  # Coordonnées du bouton "Oui"
x_ok, y_ok = 946, 420    # Coordonnées du bouton "OK"



from bulletins import show_bulletins, show_bulletins_in_frame, show_details_in_frame


if not bulletins_path:
    print("⚠️ Erreur : Le chemin des bulletins de salaire n'est pas défini.")
else:
    print(f"📂 Dossier Bulletins de salaire trouvé : {bulletins_path}")

# Assurer que toutes les clés existent dans file_paths avant utilisation
keys_to_check = ["pdf_mar", "pdf_iade", "excel_mar", "excel_iade", "word_mar", "word_iade", 
                 "bulletins_salaire", "dossier_factures", "excel_salaries"]
for key in keys_to_check:
    if key not in file_paths:
        file_paths[key] = ""

def ensure_directory_exists(directory_path):
    """Crée le dossier s'il n'existe pas."""
    expanded_path = os.path.expanduser(directory_path)
    if not os.path.exists(expanded_path):
        os.makedirs(expanded_path)
        print(f"📂 Dossier créé à : {expanded_path}")
    return expanded_path



if factures_path:
    print(f"📂 Dossier 'Frais_Factures' disponible à : {factures_path}")


"""Vérifie et crée le dossier Bulletins de salaire si nécessaire."""
bulletins_path = file_paths.get("bulletins_salaire", "~/Documents/Bulletins_Salaire")
bulletins_path = os.path.expanduser(bulletins_path)  # Gère les chemins avec ~

if not os.path.exists(bulletins_path):
    os.makedirs(bulletins_path)
    print(f"📂 Dossier 'Bulletins de salaire' créé à : {bulletins_path}")
else:
    print(f"📂 Le dossier 'Bulletins de salaire' existe déjà à : {bulletins_path}")


# Style pour les boutons
button_style = {
    "font": ("Arial", 14, "bold"),
    "bg": "#007ACC",
    "fg": "black",
    "width": 30,
    "height": 2,
    "borderwidth": 3,
    "relief": "raised"
}

def open_parameters():
    """Fenêtre pour modifier les chemins de fichiers et gérer les listes."""
    # Nettoyer le panneau droit
    clear_right_frame()
    
    # Conteneur principal - utilisation d'un PanedWindow pour diviser l'espace
    main_container = tk.PanedWindow(right_frame, orient=tk.HORIZONTAL)
    main_container.pack(fill="both", expand=True)
    
    # Cadre gauche pour le menu des paramètres
    params_menu_frame = tk.Frame(main_container, bg="#f0f4f7", padx=20, pady=20)
    main_container.add(params_menu_frame, width=300)  # Largeur fixe pour le menu
    
    # Cadre droit pour l'affichage des fonctionnalités (initialement vide)
    content_container = tk.Frame(main_container, bg="#f5f5f5", padx=20, pady=20)
    main_container.add(content_container, width=900)  # Plus de place pour le contenu
    
    # Titre du menu des paramètres
    tk.Label(params_menu_frame, text="⚙️ Paramètres", font=("Arial", 14, "bold"), 
             bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    # Boutons du menu
    tk.Button(params_menu_frame, text="🔄 Chemins des fichiers", 
              command=lambda: display_file_paths_in_container(content_container),
              width=25, bg="#DDDDDD", fg="black", 
              font=("Arial", 10, "bold")).pack(pady=10)
    
    tk.Button(params_menu_frame, text="👨‍⚕️ Gestion MARS titulaires", 
              command=lambda: display_mars_titulaires_in_container(content_container),
              width=25, bg="#DDDDDD", fg="black", 
              font=("Arial", 10, "bold")).pack(pady=10)
    
    tk.Button(params_menu_frame, text="🩺 Gestion MARS remplaçants", 
              command=lambda: display_mars_remplacants_in_container(content_container),
              width=25, bg="#DDDDDD", fg="black", 
              font=("Arial", 10, "bold")).pack(pady=10)
    
    tk.Button(params_menu_frame, text="💉 Gestion IADE remplaçants", 
              command=lambda: display_iade_remplacants_in_container(content_container),
              width=25, bg="#DDDDDD", fg="black", 
              font=("Arial", 10, "bold")).pack(pady=10)
    
    tk.Button(params_menu_frame, text="👥 Gestion des salariés", 
              command=lambda: display_salaries_in_container(content_container),
              width=25, bg="#DDDDDD", fg="black", 
              font=("Arial", 10, "bold")).pack(pady=10)
    
    tk.Button(params_menu_frame, text="📝 Paramètres DocuSign", 
              command=lambda: display_docusign_in_container(content_container),
              width=25, bg="#DDDDDD", fg="black", 
              font=("Arial", 10, "bold")).pack(pady=10)
    
    # Retour au menu principal
    tk.Button(params_menu_frame, text="🔙 Retour au menu principal", 
              command=lambda: [clear_right_frame(), show_welcome_image()], 
              width=25, bg="#BBBBBB", fg="black", 
              font=("Arial", 10, "bold")).pack(pady=20)
    
    # Afficher un message d'accueil initial dans le conteneur de droite
    tk.Label(content_container, text="⚙️ Bienvenue dans les paramètres", 
             font=("Arial", 16, "bold"), fg="#4a90e2", bg="#f5f5f5").pack(pady=20)
    tk.Label(content_container, text="Sélectionnez une option dans le menu de gauche pour configurer l'application.", 
             font=("Arial", 12), bg="#f5f5f5").pack(pady=10)

def display_file_paths_in_container(container):
    """Affiche l'interface de modification des chemins dans le conteneur spécifié."""
    # Vider le conteneur
    for widget in container.winfo_children():
        widget.destroy()
    
    # Charger les paramètres depuis config.json
    config_file = "config.json"
    if not os.path.exists(config_file):
        config = {"bank_url": "https://espacepro.secure.lcl.fr/", "bank_id": ""}
    else:
        with open(config_file, "r") as f:
            config = json.load(f)

    # Variables pour stocker les entrées
    bank_url_var = StringVar(value=config.get("bank_url", "https://espacepro.secure.lcl.fr/"))
    bank_id_var = StringVar(value=config.get("bank_id", ""))
    
    # Titre
    tk.Label(container, text="🔄 Configuration des chemins de fichiers", 
           font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    # Cadre principal avec scrollbar pour gérer beaucoup de chemins
    main_frame = tk.Frame(container, bg="#f5f5f5")
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    # Ajouter une scrollbar si nécessaire
    canvas = tk.Canvas(main_frame, bg="#f5f5f5")
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg="#f5f5f5")
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    paths = [
        ("Chemin Excel MAR :", "excel_mar"),
        ("Chemin Excel IADE :", "excel_iade"),
        ("Chemin Word MAR :", "word_mar"),
        ("Chemin Word IADE :", "word_iade"),
        ("Dossier PDF Contrats MAR :", "pdf_mar"),
        ("Dossier PDF Contrats IADE :", "pdf_iade"),
        ("Dossier Bulletins de salaire :", "bulletins_salaire"),
        ("Dossier Frais/Factures :", "dossier_factures"),
        ("Chemin Excel Salariés :", "excel_salaries"),
    ]
    
    # Dictionnaire pour stocker les variables de chaque chemin
    path_vars = {}
    
    def select_path(var, key):
        """Sélectionne un fichier ou un dossier selon le type de chemin."""
        print(f"📂 DEBUG : select_path() appelé pour {key}")  # ✅ Ajout d'un print pour debug

        if key in ["pdf_mar", "pdf_iade", "bulletins_salaire", "dossier_factures"]:
            # ✅ Pour les dossiers, utiliser askdirectory()
            path = filedialog.askdirectory(title="Sélectionnez un dossier")
        else:
            # ✅ Pour les fichiers, utiliser askopenfilename()
            path = filedialog.askopenfilename(title="Sélectionnez un fichier")

        if path:  # ✅ Vérifier qu'un chemin a été sélectionné
            print(f"📂 DEBUG : Nouveau chemin sélectionné : {path}")  # ✅ Vérification du retour
            var.set(path)
    
    row_index = 0
    for label_text, key in paths:
        path_vars[key] = StringVar(value=file_paths.get(key, ""))
        ttk.Label(scrollable_frame, text=label_text, font=("Arial", 11)).grid(row=row_index, column=0, sticky="w", pady=3)
        ttk.Entry(scrollable_frame, textvariable=path_vars[key], width=60).grid(row=row_index, column=1, sticky="w", pady=3)
        ttk.Button(scrollable_frame, text="…", command=partial(select_path, path_vars[key], key)).grid(row=row_index, column=2, padx=5, pady=3)
        row_index += 1
    
    # Section pour les paramètres bancaires
    row_index += 1
    ttk.Label(scrollable_frame, text="Paramètres bancaires", font=("Arial", 12, "bold")).grid(row=row_index, column=0, columnspan=3, pady=10)
    
    row_index += 1
    ttk.Label(scrollable_frame, text="URL de la banque :").grid(row=row_index, column=0, sticky="w", pady=5)
    ttk.Entry(scrollable_frame, textvariable=bank_url_var, width=60).grid(row=row_index, column=1, pady=5)
    
    row_index += 1
    ttk.Label(scrollable_frame, text="Identifiant bancaire :").grid(row=row_index, column=0, sticky="w", pady=5)
    ttk.Entry(scrollable_frame, textvariable=bank_id_var, width=60).grid(row=row_index, column=1, pady=5)
    
    # Fonction pour sauvegarder les paramètres
    def save_parameters():
        """Sauvegarde les paramètres dans file_paths.json et config.json"""
        
        # Mise à jour des chemins de fichiers
        for key, var in path_vars.items():
            file_paths[key] = var.get()

        # Mise à jour des paramètres bancaires
        config["bank_url"] = bank_url_var.get()
        config["bank_id"] = bank_id_var.get()

        try:
            # Sauvegarde des chemins de fichiers
            with open(SETTINGS_FILE, "w") as f:
                json.dump(file_paths, f, indent=4)
            print("✅ Paramètres des fichiers sauvegardés avec succès.")
        except Exception as e:
            print(f"❌ Erreur lors de la sauvegarde des fichiers : {e}")
            messagebox.showerror("Erreur", f"Impossible de sauvegarder les chemins des fichiers : {e}")

        try:
            # Sauvegarde des paramètres bancaires
            with open(config_file, "w") as f:
                json.dump(config, f, indent=4)
            print("✅ Paramètres bancaires sauvegardés avec succès.")
            
            # Afficher un message de succès
            messagebox.showinfo("Succès", "Les paramètres ont été sauvegardés avec succès.")
        except Exception as e:
            print(f"❌ Erreur lors de la sauvegarde des paramètres bancaires : {e}")
            messagebox.showerror("Erreur", f"Impossible de sauvegarder les paramètres bancaires : {e}")

    # Boutons de sauvegarde
    buttons_frame = tk.Frame(container, bg="#f5f5f5")
    buttons_frame.pack(fill="x", pady=10)
    
    tk.Button(buttons_frame, text="💾 Enregistrer", command=save_parameters, 
             bg="#4caf50", fg="black", font=("Arial", 12, "bold")).pack(side="right", padx=10)


# Fonctions pour afficher chaque type de paramètre dans le conteneur
def display_mars_titulaires_in_container(container):
    """Affiche la gestion des MARS titulaires dans le conteneur."""
    # Vider le conteneur
    for widget in container.winfo_children():
        widget.destroy()
    
    # Créer un cadre pour la gestion des MARS titulaires
    frame = tk.Frame(container, bg="#f5f5f5", padx=20, pady=20)
    frame.pack(fill="both", expand=True)
    
    # Titre
    tk.Label(frame, text="👨‍⚕️ Gestion des MARS titulaires", 
             font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    try:
        # Chargement des données depuis la bonne feuille
        mars_titulaires = pd.read_excel(file_paths["excel_mar"], sheet_name="MARS SELARL")
        
        # Cadre principal
        main_frame = tk.Frame(frame, bg="#f5f5f5")
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Créer une listbox pour afficher les MARS titulaires
        listbox = tk.Listbox(main_frame, width=50, height=15, font=("Arial", 12))
        listbox.pack(side="left", fill="both", expand=True, padx=5, pady=10)
        
        # Ajouter une scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=listbox.yview)
        scrollbar.pack(side="right", fill="y")
        listbox.config(yscrollcommand=scrollbar.set)
        
        # Variable pour stocker les données
        mars_data = []
        
        def refresh_listbox():
            """Met à jour la liste des MARS affichée dans la Listbox."""
            listbox.delete(0, tk.END)
            mars_data.clear()
            
            for _, row in mars_titulaires.iterrows():
                nom = row["NOM"] if not pd.isna(row["NOM"]) else ""
                prenom = row["PRENOM"] if not pd.isna(row["PRENOM"]) else ""
                full_name = f"{nom} {prenom}".strip()
                mars_data.append(row)
                listbox.insert(tk.END, full_name)
        
        # Remplir la listbox initialement
        refresh_listbox()
        
        # Fonctions pour les boutons
        def on_modify():
            """Modifier un MAR titulaire."""
            selected_index = listbox.curselection()
            if not selected_index:
                messagebox.showwarning("Avertissement", "Veuillez sélectionner un médecin.")
                return
            
            selected_index = selected_index[0]
            selected_row = mars_data[selected_index]
            
            # Créer une fenêtre de modification
            modify_window = tk.Toplevel(container)
            modify_window.title("Modifier un MAR titulaire")
            modify_window.geometry("400x400")
            modify_window.grab_set()  # Rendre la fenêtre modale
            
            # Variables pour les champs
            nom_var = StringVar(value=selected_row["NOM"] if not pd.isna(selected_row["NOM"]) else "")
            prenom_var = StringVar(value=selected_row["PRENOM"] if not pd.isna(selected_row["PRENOM"]) else "")
            ordre_var = StringVar(value=selected_row["N ORDRE"] if not pd.isna(selected_row["N ORDRE"]) else "")
            email_var = StringVar(value=selected_row["EMAIL"] if not pd.isna(selected_row["EMAIL"]) else "")
            iban_var = StringVar(value=selected_row.get("IBAN", "") if not pd.isna(selected_row.get("IBAN", "")) else "")
            
            # Création des champs
            padx, pady = 10, 5
            row = 0
            
            ttk.Label(modify_window, text="Nom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
            ttk.Entry(modify_window, textvariable=nom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
            row += 1
            
            ttk.Label(modify_window, text="Prénom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
            ttk.Entry(modify_window, textvariable=prenom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
            row += 1
            
            ttk.Label(modify_window, text="N° Ordre:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
            ttk.Entry(modify_window, textvariable=ordre_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
            row += 1
            
            ttk.Label(modify_window, text="Email:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
            ttk.Entry(modify_window, textvariable=email_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
            row += 1
            
            ttk.Label(modify_window, text="IBAN:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
            ttk.Entry(modify_window, textvariable=iban_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
            row += 1
            
            def save_changes():
                """Enregistre les modifications et met à jour Excel."""
                # Mise à jour des données
                mars_titulaires.loc[selected_index, "NOM"] = nom_var.get()
                mars_titulaires.loc[selected_index, "PRENOM"] = prenom_var.get()
                mars_titulaires.loc[selected_index, "N ORDRE"] = ordre_var.get()
                mars_titulaires.loc[selected_index, "EMAIL"] = email_var.get()
                
                # Ajouter IBAN s'il n'existe pas
                if "IBAN" not in mars_titulaires.columns:
                    mars_titulaires["IBAN"] = ""
                mars_titulaires.loc[selected_index, "IBAN"] = iban_var.get()
                
                try:
                    # Sauvegarde dans Excel
                    save_excel_with_updated_sheet(file_paths["excel_mar"], "MARS SELARL", mars_titulaires)
                    messagebox.showinfo("Succès", "Modifications enregistrées avec succès.")
                    
                    # Mise à jour de l'affichage
                    refresh_listbox()
                    modify_window.destroy()
                except Exception as e:
                    messagebox.showerror("Erreur", f"Erreur lors de la sauvegarde : {e}")
            
            # Boutons de validation
            buttons_frame = tk.Frame(modify_window)
            buttons_frame.grid(row=row, column=0, columnspan=2, pady=15)
            
            ttk.Button(buttons_frame, text="Enregistrer", command=save_changes).pack(side="left", padx=10)
            ttk.Button(buttons_frame, text="Annuler", command=modify_window.destroy).pack(side="left", padx=10)
        
        def on_add():
            """Ajouter un nouveau MAR titulaire."""
            # Créer une fenêtre d'ajout
            add_window = tk.Toplevel(container)
            add_window.title("Ajouter un MAR titulaire")
            add_window.geometry("400x400")
            add_window.grab_set()  # Rendre la fenêtre modale
            
            # Variables pour les champs
            nom_var = StringVar()
            prenom_var = StringVar()
            ordre_var = StringVar()
            email_var = StringVar()
            iban_var = StringVar()
            
            # Création des champs
            padx, pady = 10, 5
            row = 0
            
            ttk.Label(add_window, text="Nom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
            ttk.Entry(add_window, textvariable=nom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
            row += 1
            
            ttk.Label(add_window, text="Prénom:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
            ttk.Entry(add_window, textvariable=prenom_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
            row += 1
            
            ttk.Label(add_window, text="N° Ordre:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
            ttk.Entry(add_window, textvariable=ordre_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
            row += 1
            
            ttk.Label(add_window, text="Email:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
            ttk.Entry(add_window, textvariable=email_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
            row += 1
            
            ttk.Label(add_window, text="IBAN:").grid(row=row, column=0, sticky="w", padx=padx, pady=pady)
            ttk.Entry(add_window, textvariable=iban_var, width=30).grid(row=row, column=1, padx=padx, pady=pady)
            row += 1
            
            def save_new():
                """Enregistre le nouveau MAR et met à jour Excel."""
                # Vérification des champs obligatoires
                if not nom_var.get().strip() or not prenom_var.get().strip():
                    messagebox.showwarning("Attention", "Veuillez renseigner au moins le nom et le prénom.")
                    return
                
                # Préparation des nouvelles données
                new_row = {
                    "NOM": nom_var.get(),
                    "PRENOM": prenom_var.get(),
                    "N ORDRE": ordre_var.get(),
                    "EMAIL": email_var.get()
                }
                
                # Ajouter IBAN s'il n'existe pas
                if "IBAN" not in mars_titulaires.columns:
                    mars_titulaires["IBAN"] = ""
                new_row["IBAN"] = iban_var.get()
                
                try:
                    # Ajout au DataFrame
                    mars_titulaires.loc[len(mars_titulaires)] = new_row
                    
                    # Sauvegarde dans Excel
                    save_excel_with_updated_sheet(file_paths["excel_mar"], "MARS SELARL", mars_titulaires)
                    messagebox.showinfo("Succès", "MAR titulaire ajouté avec succès.")
                    
                    # Mise à jour de l'affichage
                    refresh_listbox()
                    add_window.destroy()
                except Exception as e:
                    messagebox.showerror("Erreur", f"Erreur lors de l'ajout : {e}")
            
            # Boutons de validation
            buttons_frame = tk.Frame(add_window)
            buttons_frame.grid(row=row, column=0, columnspan=2, pady=15)
            
            ttk.Button(buttons_frame, text="Ajouter", command=save_new).pack(side="left", padx=10)
            ttk.Button(buttons_frame, text="Annuler", command=add_window.destroy).pack(side="left", padx=10)
        
        def on_delete():
            """Supprimer un MAR titulaire."""
            selected_index = listbox.curselection()
            if not selected_index:
                messagebox.showwarning("Avertissement", "Veuillez sélectionner un médecin.")
                return
            
            selected_index = selected_index[0]
            selected_row = mars_data[selected_index]
            
            # Confirmation de suppression
            nom = selected_row["NOM"] if not pd.isna(selected_row["NOM"]) else ""
            prenom = selected_row["PRENOM"] if not pd.isna(selected_row["PRENOM"]) else ""
            full_name = f"{nom} {prenom}".strip()
            
            if messagebox.askyesno("Confirmation", f"Voulez-vous vraiment supprimer {full_name} ?"):
                try:
                    # Suppression de la ligne
                    mars_titulaires.drop(mars_titulaires.index[selected_index], inplace=True)
                    mars_titulaires.reset_index(drop=True, inplace=True)
                    
                    # Sauvegarde dans Excel
                    save_excel_with_updated_sheet(file_paths["excel_mar"], "MARS SELARL", mars_titulaires)
                    messagebox.showinfo("Succès", f"{full_name} supprimé avec succès.")
                    
                    # Mise à jour de l'affichage
                    refresh_listbox()
                except Exception as e:
                    messagebox.showerror("Erreur", f"Erreur lors de la suppression : {e}")
        
        # Boutons d'action
        buttons_frame = tk.Frame(frame, bg="#f5f5f5")
        buttons_frame.pack(pady=10)
        
        tk.Button(buttons_frame, text="➕ Ajouter", command=on_add, 
                 bg="#4caf50", fg="black", font=("Arial", 10, "bold")).pack(side="left", padx=10)
        tk.Button(buttons_frame, text="✏️ Modifier", command=on_modify, 
                 bg="#2196f3", fg="black", font=("Arial", 10, "bold")).pack(side="left", padx=10)
        tk.Button(buttons_frame, text="🗑️ Supprimer", command=on_delete, 
                 bg="#f44336", fg="black", font=("Arial", 10, "bold")).pack(side="left", padx=10)
    
    except Exception as e:
        tk.Label(frame, text=f"Erreur lors du chargement des données : {e}", 
                fg="red", bg="#f5f5f5", wraplength=600).pack(pady=20)


def display_mars_remplacants_in_container(container):
    """Affiche la gestion des MARS remplaçants dans le conteneur."""
    # Vider le conteneur
    for widget in container.winfo_children():
        widget.destroy()
    
    # Créer un cadre pour la gestion des MARS remplaçants
    frame = tk.Frame(container, bg="#f5f5f5")
    frame.pack(fill="both", expand=True)
    
    # Titre
    tk.Label(frame, text="🩺 Gestion des MARS remplaçants", 
             font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    # Adaptez ici le contenu de votre fonction manage_mar_remplacants


def display_iade_remplacants_in_container(container):
    """Affiche la gestion des IADE remplaçants dans le conteneur."""
    # Vider le conteneur
    for widget in container.winfo_children():
        widget.destroy()
    
    # Créer un cadre pour la gestion des IADE remplaçants
    frame = tk.Frame(container, bg="#f5f5f5")
    frame.pack(fill="both", expand=True)
    
    # Titre
    tk.Label(frame, text="💉 Gestion des IADE remplaçants", 
             font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    # Adaptez ici le contenu de votre fonction manage_iade_remplacants


def display_salaries_in_container(container):
    """Affiche la gestion des salariés dans le conteneur."""
    # Vider le conteneur
    for widget in container.winfo_children():
        widget.destroy()
    
    # Créer un cadre pour la gestion des salariés
    frame = tk.Frame(container, bg="#f5f5f5")
    frame.pack(fill="both", expand=True)
    
    # Titre
    tk.Label(frame, text="👥 Gestion des salariés", 
             font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    # Adaptez ici le contenu de votre fonction manage_salaries


def display_docusign_in_container(container):
    """Affiche les paramètres DocuSign dans le conteneur."""
    # Vider le conteneur
    for widget in container.winfo_children():
        widget.destroy()
    
    # Créer un cadre pour les paramètres DocuSign
    frame = tk.Frame(container, bg="#f5f5f5")
    frame.pack(fill="both", expand=True)
    
    # Titre
    tk.Label(frame, text="📝 Paramètres DocuSign", 
             font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    # Charger les identifiants depuis config.json
    config_file = "config.json"
    if not os.path.exists(config_file):
        config = {"docusign_login_page": "https://account.docusign.com", "docusign_email": "", "docusign_password": ""}
    else:
        with open(config_file, "r") as f:
            config = json.load(f)   
    
    # Variables pour stocker les entrées
    login_page_var = StringVar(value=config.get("docusign_login_page", "https://account.docusign.com"))
    email_var = StringVar(value=config.get("docusign_email", ""))
    password_var = StringVar(value=config.get("docusign_password", ""))
    
    # Création des champs
    tk.Label(frame, text="Page de login DocuSign :").pack(pady=5)
    tk.Entry(frame, textvariable=login_page_var, width=50).pack(pady=5)
    
    tk.Label(frame, text="Email DocuSign :").pack(pady=5)
    tk.Entry(frame, textvariable=email_var, width=50).pack(pady=5)
    
    tk.Label(frame, text="Mot de passe (laisser vide si non stocké) :").pack(pady=5)
    tk.Entry(frame, textvariable=password_var, width=50, show="*").pack(pady=5)
    
    # Fonction pour sauvegarder
    def save_docusign_settings():
        config["docusign_login_page"] = login_page_var.get()
        config["docusign_email"] = email_var.get()
        config["docusign_password"] = password_var.get()
        
        with open(config_file, "w") as f:
            json.dump(config, f, indent=4)
        
        messagebox.showinfo("Succès", "Paramètres DocuSign enregistrés.")
    
    # Bouton de sauvegarde
    tk.Button(frame, text="💾 Enregistrer", command=save_docusign_settings,
             bg="#4caf50", fg="black", font=("Arial", 12, "bold")).pack(pady=20)    
    
        
def open_docusign_parameters():
    """Fenêtre pour modifier les identifiants DocuSign et les enregistrer dans un fichier JSON."""
    param_window = Toplevel()
    param_window.title("Paramètres DocuSign")
    param_window.geometry("500x300")

    # Charger les identifiants depuis config.json
    config_file = "config.json"
    if not os.path.exists(config_file):
        config = {"docusign_login_page": "https://account.docusign.com", "docusign_email": "", "docusign_password": ""}
    else:
        with open(config_file, "r") as f:
            config = json.load(f)    
    try:
        with open(config_file, "r") as f:
            config = json.load(f)
    except FileNotFoundError:
        config = {
            "docusign_login_page": "https://account.docusign.com",
            "docusign_email": "",
            "docusign_password": ""
        }

    # Variables pour stocker les entrées
    login_page_var = StringVar(value=config.get("docusign_login_page", "https://account.docusign.com"))
    email_var = StringVar(value=config.get("docusign_email", ""))
    password_var = StringVar(value=config.get("docusign_password", ""))

    # Création des champs
    Label(param_window, text="Page de login DocuSign :").pack(pady=5)
    Entry(param_window, textvariable=login_page_var, width=50).pack(pady=5)

    Label(param_window, text="Email DocuSign :").pack(pady=5)
    Entry(param_window, textvariable=email_var, width=50).pack(pady=5)

    Label(param_window, text="Mot de passe (laisser vide si non stocké) :").pack(pady=5)
    Entry(param_window, textvariable=password_var, width=50, show="*").pack(pady=5)

    # Fonction pour sauvegarder
    def save_docusign_settings():
        config["docusign_login_page"] = login_page_var.get()
        config["docusign_email"] = email_var.get()
        config["docusign_password"] = password_var.get()

        with open(config_file, "w") as f:
            json.dump(config, f, indent=4)

        messagebox.showinfo("Succès", "Paramètres DocuSign enregistrés.")
        param_window.destroy()

    # Boutons d'action
    Button(param_window, text="Enregistrer", command=save_docusign_settings).pack(pady=10)
    Button(param_window, text="Annuler", command=param_window.destroy).pack(pady=5)


def show_post_contract_actions(container, pdf_path, replaced_name, replacing_name, 
                            replaced_email, replacing_email, start_date, end_date, contract_type):
    """
    Affiche les options post-création de contrat dans le conteneur spécifié.
    """
    # Nettoyer le conteneur d'abord
    for widget in container.winfo_children():
        widget.destroy()
    
    # Titre
    tk.Label(container, text="Actions disponibles", font=("Arial", 14, "bold"), 
           bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    # Description
    tk.Label(container, text=f"Contrat généré pour:\n{replacing_name}\nremplaçant {replaced_name}\n{start_date} - {end_date}", 
           font=("Arial", 10), justify=tk.LEFT, bg="#f5f5f5").pack(pady=10, anchor="w")
    
    
    def send_to_docusign_callback(pdf_path, contract_type, start_date, end_date, replacing_name, replacing_email, replaced_name=None, replaced_email=None):
        print("🔹 Début de la fonction send_to_docusign_callback()")

        # Vérification et correction des valeurs None
        replaced_name = replaced_name if replaced_name else "Non précisé"
        replaced_email = replaced_email if replaced_email else "email_inconnu@exemple.com"
        replacing_email = replacing_email if replacing_email and replacing_email not in ["", "nan", "None", None] else "email_inconnu@exemple.com"

        print(f"📤 DEBUG Avant subprocess : replaced_name={replaced_name}, replaced_email={replaced_email}, replacing_name={replacing_name}, replacing_email={replacing_email}")

        # Vérification finale avant exécution
        if None in [pdf_path, replacing_name, replacing_email, start_date, end_date, contract_type]:
            print("❌ ERREUR : L'un des paramètres essentiels est None, envoi annulé.")
            return

        print(f"📤 DEBUG Final avant envoi DocuSign :")
        print(f"   🔹 pdf_path = {pdf_path}")
        print(f"   🔹 contract_type = {contract_type}")
        print(f"   🔹 start_date = {start_date}")
        print(f"   🔹 end_date = {end_date}")
        print(f"   🔹 replacing_name = {replacing_name}")
        print(f"   🔹 replacing_email = {replacing_email}")
        print(f"   🔹 replaced_name = {replaced_name}")
        print(f"   🔹 replaced_email = {replaced_email}")

        # Récupérer le chemin du script DocuSign depuis config
        script_docusign = get_file_path("script_docusign")
        if not script_docusign:
            print("❌ ERREUR : Impossible de trouver le script DocuSign")
            return

        # Gestion spécifique pour IADE (sans remplacé)
        if contract_type == "IADE":
            args = [
                "python3", script_docusign,
                pdf_path, contract_type, start_date, end_date, replacing_name, replacing_email
            ]
        else:  # MAR et autres types
            args = [
                "python3", script_docusign,
                pdf_path, contract_type, start_date, end_date, replacing_name, replacing_email, replaced_name, replaced_email
            ]

        try:
            print("🚀 Lancement de l'envoi à DocuSign via subprocess...")
            subprocess.run(args, check=True)
            print("✅ Envoi à DocuSign terminé avec succès.")
        except subprocess.CalledProcessError as e:
            print(f"❌ Erreur lors de l'envoi à DocuSign : {e}")


    # Titre
    tk.Label(container, text="Actions disponibles", font=("Arial", 14, "bold"), 
           bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    # Description
    tk.Label(container, text=f"Contrat généré pour:\n{replacing_name}\nremplaçant {replaced_name}\n{start_date} - {end_date}", 
           font=("Arial", 10), justify=tk.LEFT, bg="#f5f5f5").pack(pady=10, anchor="w")
    
    # Bouton pour ouvrir le PDF
    def open_pdf():
        """Ouvre le contrat avec PDF Expert."""
        subprocess.run(["open", "-a", "PDF Expert", pdf_path])
    

    tk.Button(container, text="📄 Ouvrir avec PDF Expert", command=open_pdf, 
            width=30).pack(pady=5)
    

    # Bouton pour envoyer à DocuSign
    tk.Button(container, text="📩 Envoyer en DocuSign", 
            command=send_to_docusign_callback, width=30).pack(pady=5)
    

    # Bouton pour effectuer le règlement (désactivé comme dans votre code original)
    tk.Button(container, text="💰 Effectuer le règlement (à venir)", 
            state="disabled", width=30).pack(pady=5)
    
    # Bouton pour revenir à l'écran d'accueil
    tk.Button(container, text="🏠 Retour à l'accueil", 
            command=lambda: [clear_right_frame(), show_welcome_image()], 
            width=30).pack(pady=20)

def send_to_docusign(pdf_path, contract_type, start_date, end_date, replacing_name, replacing_email, replaced_email=None, replaced_name=None):
    """
    Envoie le fichier PDF généré à DocuSign en automatisant l'importation.

    Args:
        pdf_path (str): Chemin du fichier PDF à envoyer.
        contract_type (str): Type du contrat ("IADE" ou "MAR").
        start_date (str): Date de début du contrat.
        end_date (str): Date de fin du contrat.
        replacing_name (str): Nom du remplaçant.
        replacing_email (str): Email du remplaçant.
        replaced_name (str, optional): Nom du médecin remplacé (pour MAR).
        replaced_email (str, optional): Email du médecin remplacé (pour MAR).
    """

    # 🔹 Vérification des paramètres obligatoires
    if not pdf_path or not os.path.exists(pdf_path):
        print(f"❌ Erreur : Le fichier PDF {pdf_path} n'existe pas.")
        return
    
    if not replacing_name or not replacing_email:
        print("❌ Erreur : Le nom ou l'email du remplaçant est manquant.")
        return
    
    # 🔹 Vérification des dates
    if not start_date or not end_date:
        print("❌ Erreur : Les dates du contrat ne sont pas définies correctement.")
        return

    replacing_email = replacing_email if replacing_email and replacing_email not in ["", "nan", None] else "email_inconnu@exemple.com"
    replaced_email = replaced_email if replaced_email and replaced_email not in ["", "nan", None] else "email_inconnu@exemple.com"

    # 📌 Debugging - Vérification des paramètres envoyés
    print(f"📤 Envoi à DocuSign...")
    print(f"   📄 Fichier : {pdf_path}")
    print(f"   🏷️ Type de contrat : {contract_type}")
    print(f"   📆 Période : {start_date} - {end_date}")
    print(f"   👨‍⚕️ Remplaçant : {replacing_name} ({replacing_email})")
    print(f"📤 DEBUG Avant envoi DocuSign :")
    print(f"   🔹 replaced_name = {replaced_name}")
    print(f"   🔹 replaced_email = {replaced_email}")
    print(f"   🔹 replacing_name = {replacing_name}")
    print(f"   🔹 replacing_email = {replacing_email}")

    # Récupérer le chemin du script DocuSign depuis config
    script_docusign = get_file_path("script_docusign")
    if not script_docusign:
        print("❌ ERREUR : Impossible de trouver le script DocuSign")
        return

    try:
        if contract_type == "MAR":
            print(f"📤 DEBUG Avant subprocess : replaced_name={replaced_name}, replaced_email={replaced_email}, replacing_name={replacing_name}, replacing_email={replacing_email}")
            subprocess.run([
                "python3", script_docusign,
                pdf_path, contract_type, start_date, end_date, replacing_name, replacing_email, replaced_name, replaced_email
            ], check=True)

        
        elif contract_type == "IADE":
            subprocess.run([
                "python3", script_docusign,
                pdf_path, "IADE", start_date, end_date, replacing_name, replacing_email
            ], check=True)

        print("✅ Contrat envoyé avec succès à DocuSign.")

    except subprocess.CalledProcessError as e:
        print(f"❌ Erreur lors de l'envoi à DocuSign : {e}")
    except Exception as e:
        print(f"❌ Erreur inattendue : {e}")


def manage_salaries():
    """Fenêtre pour gérer la liste des salariés."""
    try:
        salaries_data = pd.read_excel(file_paths["excel_salaries"], sheet_name="Salariés")
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible de charger la liste des salariés : {e}")
        return

    def refresh_listbox():
        """Met à jour la liste des salariés dans la Listbox."""
        listbox.delete(0, "end")
        for _, row in salaries_data.iterrows():
            listbox.insert("end", f"{row['NOM']} {row['PRENOM']}")

    def on_modify():
        """Modifier les informations du salarié sélectionné."""
        try:
            selected_index = listbox.curselection()[0]
        except IndexError:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un salarié.")
            return

        selected_row = salaries_data.iloc[selected_index]

        modify_window = Toplevel(window)
        modify_window.title("Modifier un salarié")

        # Variables pour les champs
        nom_var = StringVar(value=selected_row["NOM"])
        prenom_var = StringVar(value=selected_row["PRENOM"])
        email_var = StringVar(value=selected_row["EMAIL"])
        poste_var = StringVar(value=selected_row.get("POSTE", ""))
        adresse_var = StringVar(value=selected_row.get("ADRESSE", ""))
        iban_var = StringVar(value=selected_row.get("IBAN", ""))

        fields = [
            ("Nom", nom_var),
            ("Prénom", prenom_var),
            ("Email", email_var),
            ("Poste", poste_var),
            ("Adresse", adresse_var),
            ("IBAN", iban_var)
        ]

        for i, (label_text, var) in enumerate(fields):
            Label(modify_window, text=label_text + " :").grid(row=i, column=0, padx=10, pady=5)
            Entry(modify_window, textvariable=var).grid(row=i, column=1, padx=10, pady=5)

        def save_changes():
            salaries_data.loc[selected_index, "NOM"] = nom_var.get()
            salaries_data.loc[selected_index, "PRENOM"] = prenom_var.get()
            salaries_data.loc[selected_index, "EMAIL"] = email_var.get()
            salaries_data.loc[selected_index, "POSTE"] = poste_var.get()
            salaries_data.loc[selected_index, "ADRESSE"] = adresse_var.get()
            salaries_data.loc[selected_index, "IBAN"] = iban_var.get()

            try:
                save_excel_with_updated_sheet(file_paths["excel_salaries"], "Salariés", salaries_data)
                refresh_listbox()
                modify_window.destroy()
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible d'enregistrer les modifications : {e}")

        Button(modify_window, text="Enregistrer", command=save_changes).grid(row=len(fields), column=0, columnspan=2, pady=10)
        
    def on_delete():
        """Supprime un salarié sélectionné."""
        try:
            selected_index = listbox.curselection()[0]
        except IndexError:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un salarié.")
            return

        if not messagebox.askyesno("Confirmation", "Êtes-vous sûr de vouloir supprimer ce salarié ?"):
            return

        salaries_data.drop(index=selected_index, inplace=True)
        salaries_data.reset_index(drop=True, inplace=True)

        try:
            save_excel_with_updated_sheet(file_paths["excel_salaries"], "Salariés", salaries_data)
            messagebox.showinfo("Succès", "Salarié supprimé.")
            refresh_listbox()
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de supprimer le salarié : {e}")

    def on_add():
        """Ajoute un nouveau salarié."""
        add_window = Toplevel(window)
        add_window.title("Ajouter un salarié")

        # Variables pour les champs
        nom_var = StringVar()
        prenom_var = StringVar()
        email_var = StringVar()
        poste_var = StringVar()
        adresse_var = StringVar()
        iban_var = StringVar()

        fields = [
            ("Nom", nom_var),
            ("Prénom", prenom_var),
            ("Email", email_var),
            ("Poste", poste_var),
            ("Adresse", adresse_var),
            ("IBAN", iban_var)
        ]

        for i, (label_text, var) in enumerate(fields):
            Label(add_window, text=label_text + " :").grid(row=i, column=0, padx=10, pady=5)
            Entry(add_window, textvariable=var).grid(row=i, column=1, padx=10, pady=5)

        def save_new():
            new_row = {
                "NOM": nom_var.get(),
                "PRENOM": prenom_var.get(),
                "EMAIL": email_var.get(),
                "POSTE": poste_var.get(),
                "ADRESSE": adresse_var.get(),
                "IBAN": iban_var.get()
            }
            salaries_data.loc[len(salaries_data)] = new_row

            try:
                save_excel_with_updated_sheet(file_paths["excel_salaries"], "Salariés", salaries_data)
                refresh_listbox()
                add_window.destroy()
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible d'ajouter le salarié : {e}")

        Button(add_window, text="Enregistrer", command=save_new).grid(row=len(fields), column=0, columnspan=2, pady=10)

    # Fenêtre principale
    window = Toplevel()
    window.title("Liste des Salariés")
    window.geometry("550x500")

    # Titre
    Label(window, text="Liste des Salariés", font=("Arial", 16, "bold"), bg="#007ACC", fg="white").pack(fill="x", pady=10)

    # Liste des salariés
    listbox = Listbox(window, width=50, height=15, font=("Arial", 12))
    listbox.pack(pady=10)
    refresh_listbox()

    # Boutons d'actions
    Button(window, text="Modifier", command=on_modify, width=20, bg="#4caf50").pack(pady=5)
    Button(window, text="Ajouter", command=on_add, width=20, bg="#2196f3").pack(pady=5)
    Button(window, text="Supprimer", command=on_delete, width=20, bg="#f44336").pack(pady=5)
    Button(window, text="Retour", command=window.destroy, width=20, bg="#607d8b").pack(pady=10)

def manage_mar_titulaires():
    """Fenêtre pour gérer la liste des MARS titulaires."""
    try:
        # Chargement des données depuis la bonne feuille
        mars_titulaires = pd.read_excel(file_paths["excel_mar"], sheet_name="MARS SELARL")
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible de charger la liste des MARS titulaires : {e}")
        return


    def refresh_listbox():
        """Met à jour la liste des MARS titulaires dans la Listbox."""
        listbox.delete(0, "end")
        for _, row in mars_titulaires.iterrows():
            listbox.insert("end", f"{row['NOM']} {row['PRENOM']}")

    def on_modify():
        """Modifier les informations du médecin sélectionné."""
        try:
            selected_index = listbox.curselection()[0]
        except IndexError:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un médecin.")
            return

        selected_row = mars_titulaires.iloc[selected_index]

        modify_window = Toplevel(window)
        modify_window.title("Modifier un MAR titulaire")

        # Variables pour les champs
        nom_var = StringVar(value=selected_row["NOM"])
        prenom_var = StringVar(value=selected_row["PRENOM"])
        ordre_var = StringVar(value=selected_row["N ORDRE"])
        email_var = StringVar(value=selected_row["EMAIL"])
        iban_var = StringVar(value=selected_row.get("IBAN", ""))  # Ajout de l'IBAN

        Label(modify_window, text="Nom :").grid(row=0, column=0, padx=10, pady=5)
        Entry(modify_window, textvariable=nom_var).grid(row=0, column=1, padx=10, pady=5)

        Label(modify_window, text="Prénom :").grid(row=1, column=0, padx=10, pady=5)
        Entry(modify_window, textvariable=prenom_var).grid(row=1, column=1, padx=10, pady=5)

        Label(modify_window, text="N° Ordre :").grid(row=2, column=0, padx=10, pady=5)
        Entry(modify_window, textvariable=ordre_var).grid(row=2, column=1, padx=10, pady=5)

        Label(modify_window, text="Email :").grid(row=3, column=0, padx=10, pady=5)
        Entry(modify_window, textvariable=email_var).grid(row=3, column=1, padx=10, pady=5)

        Label(modify_window, text="IBAN :").grid(row=4, column=0, padx=10, pady=5)
        Entry(modify_window, textvariable=iban_var).grid(row=4, column=1, padx=10, pady=5)

        def save_changes():
            mars_titulaires.loc[selected_index, "NOM"] = nom_var.get()
            mars_titulaires.loc[selected_index, "PRENOM"] = prenom_var.get()
            mars_titulaires.loc[selected_index, "N ORDRE"] = ordre_var.get()
            mars_titulaires.loc[selected_index, "EMAIL"] = email_var.get()
            mars_titulaires.loc[selected_index, "IBAN"] = iban_var.get()  # Sauvegarde de l'IBAN

            try:
                save_excel_with_updated_sheet(file_paths["excel_mar"], "MARS SELARL", mars_titulaires)
                refresh_listbox()
                modify_window.destroy()
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible d'enregistrer les modifications : {e}")

        Button(modify_window, text="Enregistrer", command=save_changes).grid(row=5, column=0, columnspan=2, pady=10)

    def on_delete():
        try:
            selected_index = listbox.curselection()[0]
        except IndexError:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un médecin.")
            return

        if not messagebox.askyesno("Confirmation", "Êtes-vous sûr de vouloir supprimer ce médecin ?"):
            return

        mars_titulaires.drop(index=selected_index, inplace=True)
        mars_titulaires.reset_index(drop=True, inplace=True)

        try:
            save_excel_with_updated_sheet(file_paths["excel_mar"], "MARS SELARL", mars_titulaires)

            messagebox.showinfo("Succès", "Médecin supprimé.")
            refresh_listbox()
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de supprimer le médecin : {e}")

    def on_add():
        """Ajoute un nouveau médecin."""
        add_window = Toplevel(window)
        add_window.title("Ajouter un MAR titulaire")
        add_window.geometry("400x400")  # Ajuste la taille (largeur x hauteur)

        # Variables pour les champs
        nom_var = StringVar()
        prenom_var = StringVar()
        ordre_var = StringVar()
        email_var = StringVar()
        iban_var = StringVar()


        Label(add_window, text="Nom :").grid(row=0, column=0, padx=10, pady=5)
        Entry(add_window, textvariable=nom_var).grid(row=0, column=1, padx=10, pady=5)

        Label(add_window, text="Prénom :").grid(row=1, column=0, padx=10, pady=5)
        Entry(add_window, textvariable=prenom_var).grid(row=1, column=1, padx=10, pady=5)

        Label(add_window, text="N° Ordre :").grid(row=2, column=0, padx=10, pady=5)
        Entry(add_window, textvariable=ordre_var).grid(row=2, column=1, padx=10, pady=5)

        Label(add_window, text="Email :").grid(row=3, column=0, padx=10, pady=5)
        Entry(add_window, textvariable=email_var).grid(row=3, column=1, padx=10, pady=5)
        
        Label(add_window, text="IBAN :").grid(row=4, column=0, padx=10, pady=5)
        Entry(add_window, textvariable=iban_var).grid(row=4, column=1, padx=10, pady=5)


        def save_new():
            new_row = {
                "NOM": nom_var.get(),
                "PRENOM": prenom_var.get(),
                "N ORDRE": ordre_var.get(),
                "EMAIL": email_var.get(),
                "IBAN": iban_var.get()
            }
            mars_titulaires.loc[len(mars_titulaires)] = new_row

            try:
                save_excel_with_updated_sheet(file_paths["excel_mar"], "MARS SELARL", mars_titulaires)
                refresh_listbox()
                add_window.destroy()
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible d'ajouter le médecin : {e}")

    def on_add():
        """Ajoute un nouveau médecin."""
        add_window = Toplevel(window)
        add_window.title("Ajouter un MAR titulaire")
        add_window.geometry("400x400")  # Ajuste la taille (largeur x hauteur)

        # Variables pour les champs
        nom_var = StringVar()
        prenom_var = StringVar()
        ordre_var = StringVar()
        email_var = StringVar()
        iban_var = StringVar()

        Label(add_window, text="Nom :").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        Entry(add_window, textvariable=nom_var).grid(row=0, column=1, padx=10, pady=5)

        Label(add_window, text="Prénom :").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        Entry(add_window, textvariable=prenom_var).grid(row=1, column=1, padx=10, pady=5)

        Label(add_window, text="N° Ordre :").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        Entry(add_window, textvariable=ordre_var).grid(row=2, column=1, padx=10, pady=5)

        Label(add_window, text="Email :").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        Entry(add_window, textvariable=email_var).grid(row=3, column=1, padx=10, pady=5)

        Label(add_window, text="IBAN :").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        Entry(add_window, textvariable=iban_var).grid(row=4, column=1, padx=10, pady=5)

    def save_new():
        """Sauvegarde le nouveau MAR titulaire dans le fichier Excel."""
        new_row = {
            "NOM": nom_var.get(),
            "PRENOM": prenom_var.get(),
            "N ORDRE": ordre_var.get(),
            "EMAIL": email_var.get(),
            "IBAN": iban_var.get()
        }
        try:
            mars_titulaires.loc[len(mars_titulaires)] = new_row
            save_excel_with_updated_sheet(file_paths["excel_mar"], "MARS SELARL", mars_titulaires)
            messagebox.showinfo("Succès", "Médecin ajouté.")
            refresh_listbox()
            add_window.destroy()
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ajouter le médecin : {e}")

        # Ajout du bouton Enregistrer avec la bonne fonction
        Button(add_window, text="Enregistrer", command=save_new, bg="#4caf50", fg="black").grid(row=5, column=0, columnspan=2, pady=10)

    # Fenêtre principale
    window = tk.Toplevel()
    window.title("Liste des MARS titulaires")
    window.geometry("550x500")
    window.configure(bg="#f2f7ff")

    # Titre
    title_frame = tk.Frame(window, bg="#4a90e2")
    title_frame.pack(fill="x")
    tk.Label(
        title_frame, text="Liste des MARS titulaires", font=("Arial", 16, "bold"), fg="white", bg="#4a90e2", pady=10
    ).pack()

    # Liste des titulaires
    listbox_frame = tk.Frame(window, bg="#f2f7ff")
    listbox_frame.pack(pady=20, padx=20)
    listbox = tk.Listbox(
        listbox_frame, width=50, height=15, font=("Arial", 12), selectbackground="#d1e7ff", selectforeground="black"
    )
    listbox.pack()

    # Boutons
    button_frame = tk.Frame(window, bg="#f2f7ff")
    button_frame.pack(pady=20)

    tk.Button(
        button_frame, text="Modifier", command=on_modify, width=15, bg="#4caf50", fg="black", font=("Arial", 10, "bold")
    ).grid(row=0, column=0, padx=10, pady=10)
    tk.Button(
        button_frame, text="Ajouter", command=on_add, width=15, bg="#2196f3", fg="black", font=("Arial", 10, "bold")
    ).grid(row=0, column=1, padx=10, pady=10)
    tk.Button(
        button_frame, text="Supprimer", command=on_delete, width=15, bg="#f44336", fg="black", font=("Arial", 10, "bold")
    ).grid(row=0, column=2, padx=10, pady=10)

    # Bouton retour
    tk.Button(
        window, text="Retour", command=window.destroy, width=20, bg="#607d8b", fg="black", font=("Arial", 12, "bold")
    ).pack(pady=10)

    # Remplissage initial de la liste
    refresh_listbox()


def manage_iade_remplacants():
    """Gérer les IADE remplaçants dans le fichier Excel."""
    try:
        iade_data = pd.read_excel(file_paths["excel_iade"], sheet_name="Coordonnées IADEs", dtype={"DPTN": str, "NOSSR": str, "IBAN": str})
    except FileNotFoundError:
        print("Erreur : Le fichier Excel pour les IADEs est introuvable.")
        return
    except Exception as e:
        print(f"Erreur lors de l'ouverture du fichier Excel : {e}")
        return

    def update_listbox():
        """Met à jour la liste des IADEs affichée dans la Listbox."""
        listbox.delete(0, tk.END)
        for index, row in iade_data.iterrows():
            listbox.insert(tk.END, f"{row['NOMR']} {row['PRENOMR']}")

    def confirm_action(message):
        """Affiche une boîte de dialogue de confirmation."""
        return messagebox.askyesno("Confirmation", message)

    def delete_entry():
        """Supprime un IADE remplaçant."""
        selected_index = listbox.curselection()
        if not selected_index:
            messagebox.showerror("Erreur", "Aucun IADE sélectionné.")
            return

        selected_index = selected_index[0]
        selected_entry = iade_data.iloc[selected_index]

        if not confirm_action(f"Êtes-vous sûr de vouloir supprimer {selected_entry['NOMR']} {selected_entry['PRENOMR']} ?"):
            return

        try:
            iade_data.drop(index=selected_index, inplace=True)
            iade_data.reset_index(drop=True, inplace=True)
            save_to_excel(file_paths["excel_iade"], "Coordonnées IADEs", iade_data)
            update_listbox()
            print(f"IADE {selected_entry['NOMR']} {selected_entry['PRENOMR']} supprimé avec succès.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de supprimer l'IADE : {e}")


    def add_entry():
        """Ajoute un nouvel IADE."""
        def update_sexe_dependent_fields(*args):
            if sexe_var.get() == "Madame":
                er_var.set("e")
                ilr_var.set("elle")
                salarier_var.set("à la salariée")
            else:
                er_var.set("")
                ilr_var.set("il")
                salarier_var.set("au salarié")

    
        def save_new_entry():
            """Enregistre les données du nouvel IADE dans le fichier Excel."""
            new_data = {
                "NOMR": nomr_var.get(),
                "PRENOMR": prenomr_var.get(),
                "EMAIL": email_var.get(),
                "LIEUNR": lieunr_var.get(),
                "DPTN": dptn_var.get(),
                "ADRESSER": adresser_var.get(),
                "NOSSR": nossr_var.get(),
                "NATR": natr_var.get(),
                "IBAN": iban_var.get(),
                "SEXE": sexe_var.get(),
                "ER": "e" if sexe_var.get() == "Madame" else "",
                "ILR": "elle" if sexe_var.get() == "Madame" else "il",
                "SALARIER": "à la salariée" if sexe_var.get() == "Madame" else "au salarié",
                "DDNR": ddnr_var.get()
            }

            # Vérifier si tous les champs sont remplis
            if any((value is None or str(value).strip() == "") and key != "ER" for key, value in new_data.items()):
                empty_fields = [key for key, value in new_data.items() if (value is None or str(value).strip() == "") and key != "ER"]
                print(f"❌ Erreur : Champs vides mais requis : {empty_fields}")
                messagebox.showerror("Erreur", f"Tous les champs doivent être remplis. Manquants : {', '.join(empty_fields)}")
                return

            # Ajouter les données au DataFrame
            try:
                iade_data.loc[len(iade_data)] = new_data
                save_to_excel(file_paths["excel_iade"], "Coordonnées IADEs", iade_data)
                update_listbox()
                add_window.destroy()
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible d'ajouter l'IADE : {e}")

        add_window = tk.Toplevel(window)
        add_window.title("Ajouter IADE")
        window_width = 480  # 400px + 20%
        window_height = 540
        add_window.geometry(f"{window_width}x{window_height}")

        nomr_var = tk.StringVar()
        prenomr_var = tk.StringVar()
        email_var = tk.StringVar()
        lieunr_var = tk.StringVar()
        dptn_var = tk.StringVar()
        adresser_var = tk.StringVar()
        nossr_var = tk.StringVar()
        natr_var = tk.StringVar()
        iban_var = tk.StringVar()  # Ajout du champ IBAN/RIB
        sexe_var = tk.StringVar(value="Monsieur")
        ddnr_var = tk.StringVar()
        er_var = tk.StringVar()
        ilr_var = tk.StringVar()
        salarier_var = tk.StringVar()
        sexe_var.trace_add("write", update_sexe_dependent_fields)
        
        update_sexe_dependent_fields()  

        fields = [
            ("Nom", nomr_var), ("Prénom", prenomr_var), ("Date de naissance", ddnr_var),
            ("Lieu de naissance", lieunr_var), ("Département de naissance", dptn_var), ("Nationalité", natr_var),
            ("Adresse", adresser_var), ("Email", email_var), ("Numéro de sécurité sociale", nossr_var),
            ("IBAN / RIB", iban_var), 
        ]

        for i, (label, var) in enumerate(fields):
            tk.Label(add_window, text=label).grid(row=i, column=0, sticky="w")
            tk.Entry(add_window, textvariable=var).grid(row=i, column=1)
            
        sexe_index = len(fields)
        tk.Label(add_window, text="Sexe").grid(row=sexe_index, column=0, sticky="w")
        tk.OptionMenu(add_window, sexe_var, "Monsieur", "Madame").grid(row=sexe_index, column=1)

        # Affichage des champs ER, ILR et SALARIER
        extra_fields = [
            ("ER", er_var),
            ("ILR", ilr_var),
            ("SALARIER", salarier_var)
        ]

        for i, (label, var) in enumerate(extra_fields, start=sexe_index + 1):
            tk.Label(add_window, text=label).grid(row=i, column=0, sticky="w")
            tk.Entry(add_window, textvariable=var, state="readonly").grid(row=i, column=1)



        tk.Label(add_window, text="Sexe").grid(row=len(fields), column=0, sticky="w")
        tk.OptionMenu(add_window, sexe_var, "Monsieur", "Madame").grid(row=len(fields), column=1)

        tk.Button(add_window, text="Enregistrer", command=save_new_entry, bg="#4caf50", fg="black").grid(row=len(fields) + 5, column=0, columnspan=3, pady=15)
        tk.Button(add_window, text="Annuler", command=add_window.destroy, bg="#d32f2f", fg="black").grid(row=len(fields) + 7, column=0, columnspan=3, pady=15)

        add_window.transient(window)
        add_window.grab_set()


    def modify_entry():
        """Modifier une entrée existante."""
        selected_index = listbox.curselection()
        if not selected_index:
            messagebox.showerror("Erreur", "Aucun IADE sélectionné.")
            return

        selected_index = selected_index[0]
        selected_entry = iade_data.iloc[selected_index]

        modify_window = tk.Toplevel(window)
        modify_window.title(f"Modifier IADE - {selected_entry['NOMR']} {selected_entry['PRENOMR']}")
        modify_window.geometry("600x600")

        fields = {
            "NOMR": "Nom",
            "PRENOMR": "Prénom",
            "EMAIL": "Email",
            "LIEUNR": "Lieu de naissance",
            "DPTN": "Département de naissance",
            "ADRESSER": "Adresse",
            "NOSSR": "Numéro de sécurité sociale",
            "NATR": "Nationalité",
            "IBAN": "IBAN / RIB"  # Ajout du champ IBAN
        }

        entries = {}
        for i, (key, label) in enumerate(fields.items()):
            tk.Label(modify_window, text=label).grid(row=i, column=0, padx=10, pady=5, sticky="w")
            var = tk.StringVar(value=str(selected_entry.get(key, "")))
            tk.Entry(modify_window, textvariable=var, width=40).grid(row=i, column=1, padx=10, pady=5)
            entries[key] = var

        # Sélection du sexe
        tk.Label(modify_window, text="Sexe").grid(row=len(fields), column=0, padx=10, pady=5, sticky="w")
        sexe_var = tk.StringVar(value=selected_entry.get("SEXE", "Monsieur"))
        sexe_menu = tk.OptionMenu(modify_window, sexe_var, "Monsieur", "Madame")
        sexe_menu.grid(row=len(fields), column=1, padx=10, pady=5)

        # Variables dépendantes du sexe
        er_var = tk.StringVar(value=selected_entry.get("ER", ""))
        ilr_var = tk.StringVar(value=selected_entry.get("ILR", ""))
        salarier_var = tk.StringVar(value=selected_entry.get("SALARIER", ""))
        sexe_var.trace("w", update_gender_fields)

        def update_gender_fields(*args):
            """Met à jour les champs ER, ILR et SALARIER en fonction du sexe."""
            if sexe_var.get() == "Monsieur":
                er_var.set("")
                ilr_var.set("il")
                salarier_var.set("au salarié")
            else:
                er_var.set("e")
                ilr_var.set("elle")
                salarier_var.set("à la salariée")


        # Champs dépendants du sexe
        tk.Label(modify_window, text="ER").grid(row=len(fields) + 1, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(modify_window, textvariable=er_var, state="readonly", width=40).grid(row=len(fields) + 1, column=1, padx=10, pady=5)

        tk.Label(modify_window, text="ILR").grid(row=len(fields) + 2, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(modify_window, textvariable=ilr_var, state="readonly", width=40).grid(row=len(fields) + 2, column=1, padx=10, pady=5)

        tk.Label(modify_window, text="SALARIER").grid(row=len(fields) + 3, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(modify_window, textvariable=salarier_var, state="readonly", width=40).grid(row=len(fields) + 3, column=1, padx=10, pady=5)

        def save_changes():
            """Enregistre les modifications dans le fichier Excel."""
            for key, var in entries.items():
                iade_data.at[selected_index, key] = var.get()
            iade_data.at[selected_index, "SEXE"] = sexe_var.get()
            iade_data.at[selected_index, "ER"] = er_var.get()
            iade_data.at[selected_index, "ILR"] = ilr_var.get()
            iade_data.at[selected_index, "SALARIER"] = salarier_var.get()
            save_to_excel(file_paths["excel_iade"], "Coordonnées IADEs", iade_data)
            modify_window.destroy()
            update_listbox()



        tk.Button(modify_window, text="Enregistrer", command=save_changes, bg="#4caf50", fg="black").grid(row=len(fields) + 4, column=0, columnspan=2, pady=10)
        tk.Button(modify_window, text="Retour", command=modify_window.destroy, bg="#d32f2f", fg="black").grid(row=len(fields) + 5, column=0, columnspan=2, pady=10)

    window = tk.Toplevel()
    window.title("Modifier liste IADE remplaçants")
    window.geometry("500x500")
    window.configure(bg="#f2f7ff")

    listbox = tk.Listbox(window, width=50, height=15, font=("Arial", 12))
    listbox.pack(pady=10)
    update_listbox()

    tk.Button(window, text="Ajouter", command=add_entry, width=20, bg="#2196f3", fg="black").pack(pady=5)
    tk.Button(window, text="Supprimer", command=delete_entry, width=20, bg="#f44336", fg="black").pack(pady=5)
    tk.Button(window, text="Modifier", command=modify_entry, width=20, bg="#4caf50", fg="black").pack(pady=5)
    tk.Button(window, text="Retour", command=window.destroy, width=20, bg="#607d8b", fg="black").pack(pady=5)

    def save_to_excel(file_path, sheet_name, updated_data):
        """Sauvegarde les modifications d'une seule feuille sans écraser les autres."""
        try:
            # Charger toutes les feuilles existantes
            with pd.ExcelFile(file_path, engine="openpyxl") as xls:
                sheets = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}

            # Mettre à jour uniquement la feuille concernée
            sheets[sheet_name] = updated_data

            # Réécrire le fichier avec toutes les feuilles
            with pd.ExcelWriter(file_path, engine="openpyxl", mode="w") as writer:
                for sheet, df in sheets.items():
                    df.to_excel(writer, sheet_name=sheet, index=False)

        except Exception as e:
            print(f"Erreur lors de la sauvegarde de la feuille {sheet_name} : {e}")


    update_listbox()





def manage_mar_remplacants():
    """Gérer les MARS remplaçants dans le fichier Excel."""
    global mar_data, listbox

    try:
        mar_data = pd.read_excel(file_paths["excel_mar"], sheet_name="MARS Remplaçants", dtype={"URSSAF": str, "secu": str, "IBAN": str, "N ORDRER": str})
    except FileNotFoundError:
        print("Erreur : Le fichier Excel pour les MARS est introuvable.")
        return
    except Exception as e:
        print(f"Erreur lors de l'ouverture du fichier Excel : {e}")
        return

    def update_listbox():
        """Met à jour la liste des MARS affichée dans la Listbox."""
        listbox.delete(0, tk.END)
        for _, row in mar_data.iterrows():
            listbox.insert(tk.END, f"{row['NOMR']} {row['PRENOMR']}")

    def save_to_excel(file_path, sheet_name, updated_data):
        """Sauvegarde les modifications d'une seule feuille sans écraser les autres."""
        try:
            with pd.ExcelFile(file_path, engine="openpyxl") as xls:
                sheets = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}

            sheets[sheet_name] = updated_data

            with pd.ExcelWriter(file_path, engine="openpyxl", mode="w") as writer:
                for sheet, df in sheets.items():
                    df.to_excel(writer, sheet_name=sheet, index=False)

        except Exception as e:
            print(f"Erreur lors de la sauvegarde : {e}")

    def delete_entry():
        """Supprime un MAR remplaçant."""
        global mar_data
        selected_index = listbox.curselection()
        if not selected_index:
            messagebox.showerror("Erreur", "Aucun MAR sélectionné.")
            return

        selected_index = selected_index[0]
        selected_entry = mar_data.iloc[selected_index]

        if not messagebox.askyesno("Confirmation", f"Supprimer {selected_entry['NOMR']} {selected_entry['PRENOMR']} ?"):
            return

        try:
            mar_data.drop(index=selected_index, inplace=True)
            mar_data.reset_index(drop=True, inplace=True)
            save_to_excel(file_paths["excel_mar"], "MARS Remplaçants", mar_data)
            update_listbox()
            print(f"MAR {selected_entry['NOMR']} {selected_entry['PRENOMR']} supprimé.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de supprimer : {e}")

    def add_entry():
        """Ajouter un nouveau MARS remplaçant."""
        add_window = tk.Toplevel(window)
        add_window.title("Ajouter un MARS Remplaçant")
        add_window.geometry("400x700")

        fields = {label: tk.StringVar() for label in ["Nom", "Prénom", "Email", "Adresse", "Numéro URSSAF", "Numéro Sécurité sociale", "IBAN", "Numéro de Conseil de l'Ordre"]}

        tk.Label(add_window, text="Ajouter un MARS Remplaçant", font=("Arial", 14, "bold")).pack(pady=10)
        for label, var in fields.items():
            tk.Label(add_window, text=label).pack(pady=2)
            tk.Entry(add_window, textvariable=var).pack(pady=2)

        def save_new_entry():
            """Enregistre le nouveau MARS remplaçant."""
            global mar_data
            new_data = {key: var.get() for key, var in fields.items()}
            new_data = dict(zip(["NOMR", "PRENOMR", "EMAILR", "AdresseR", "URSSAF", "secu", "IBAN", "N ORDRER"], new_data.values()))

            if not all(new_data.values()):
                messagebox.showerror("Erreur", "Tous les champs doivent être remplis.")
                return

            mar_data = mar_data.append(new_data, ignore_index=True)
            save_to_excel(file_paths["excel_mar"], "MARS Remplaçants", mar_data)
            update_listbox()
            add_window.destroy()

        tk.Button(add_window, text="Enregistrer", command=save_new_entry, bg="green", fg="black").pack(pady=10)
        tk.Button(add_window, text="Annuler", command=add_window.destroy, bg="red", fg="black").pack(pady=10)

    def modify_entry():
        """Modifier un MARS remplaçant existant."""
        global mar_data
        selected_index = listbox.curselection()
        if not selected_index:
            messagebox.showerror("Erreur", "Aucun MARS sélectionné.")
            return

        selected_index = selected_index[0]
        selected_entry = mar_data.iloc[selected_index]

        modify_window = tk.Toplevel(window)
        modify_window.title("Modifier un MARS Remplaçant")
        modify_window.geometry("400x700")

        fields = {
            "Nom": "NOMR",
            "Prénom": "PRENOMR",
            "Email": "EMAILR",
            "Adresse": "AdresseR",
            "Numéro URSSAF": "URSSAF",
            "Numéro Sécurité sociale": "secu",
            "IBAN": "IBAN",
            "Numéro de Conseil de l'Ordre": "N ORDRER"
        }

        entries = {key: tk.StringVar(value=selected_entry[val]) for key, val in fields.items()}

        for label, var in entries.items():
            tk.Label(modify_window, text=label).pack(pady=2)
            tk.Entry(modify_window, textvariable=var).pack(pady=2)

        def save_changes():
            """Enregistre les modifications d'un MAR remplaçant."""
            global mar_data
            for key, val in entries.items():
                mar_data.at[selected_index, fields[key]] = val.get()

            save_to_excel(file_paths["excel_mar"], "MARS Remplaçants", mar_data)
            modify_window.destroy()
            update_listbox()

        tk.Button(modify_window, text="Enregistrer", command=save_changes, bg="green", fg="black").pack(pady=10)
        tk.Button(modify_window, text="Annuler", command=modify_window.destroy, bg="red", fg="black").pack(pady=10)

    window = tk.Toplevel()
    window.title("Liste des MARS Remplaçants")
    window.geometry("500x400")

    listbox = tk.Listbox(window, width=50, height=15)
    listbox.pack(pady=10)
    update_listbox()

    tk.Button(window, text="Ajouter", command=add_entry).pack()
    tk.Button(window, text="Modifier", command=modify_entry).pack()
    tk.Button(window, text="Supprimer", command=delete_entry).pack()
    tk.Button(window, text="Fermer", command=window.destroy).pack()




def save_excel_with_updated_sheet(file_path, sheet_name, updated_data):
    """Sauvegarde une feuille spécifique sans supprimer les autres."""
    try:
        # Charger toutes les feuilles existantes
        with pd.ExcelFile(file_path, engine="openpyxl") as xls:
            sheets = {sheet: xls.parse(sheet) for sheet in xls.sheet_names}

        # Mettre à jour uniquement la feuille concernée
        sheets[sheet_name] = updated_data
        if "Remplaçants" in sheet_name:
            print("DEBUG: Sauvegarde des informations des remplaçants:", updated_data.head())

        # Réécrire le fichier avec toutes les feuilles
        with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            for sheet, df in sheets.items():
                df.to_excel(writer, sheet_name=sheet, index=False)

        print(f"✅ Feuille '{sheet_name}' sauvegardée avec succès dans {file_path}")

    except Exception as e:
        print(f"❌ Erreur lors de la sauvegarde de la feuille '{sheet_name}' : {e}")





def activate_word():
    """
    Active Microsoft Word en utilisant AppleScript avant de chercher les boutons.
    """
    applescript = """
    tell application "Microsoft Word"
        activate
    end tell
    """
    try:
        subprocess.run(["osascript", "-e", applescript], check=True)
        print("Microsoft Word activé.")
    except Exception as e:
        print(f"Erreur lors de l'activation de Microsoft Word : {e}")

def quit_word():
    """
    Ferme Microsoft Word à la fin du script.
    """
    applescript = """
    tell application "Microsoft Word"
        quit saving no
    end tell
    """
    try:
        subprocess.run(["osascript", "-e", applescript], check=True)
        print("Microsoft Word fermé.")
    except Exception as e:
        print(f"Erreur lors de la fermeture de Microsoft Word : {e}")

def handle_popups():
    """
    Gère les popups de Word avec des clics automatiques basés sur des coordonnées fixes.
    """
    try:
        print("Attente des fenêtres popups...")
        time.sleep(1)  # Attendre que les popups apparaissent

        # Clic sur "Oui"
        print(f"Clic sur 'Oui' aux coordonnées : ({x_oui}, {y_oui})")
        pyautogui.moveTo(x_oui, y_oui, duration=0.2)
        pyautogui.click()

        # Pause avant de gérer "OK"
        time.sleep(1)

        # Clic sur "OK"
        print(f"Clic sur 'OK' aux coordonnées : ({x_ok}, {y_ok})")
        pyautogui.moveTo(x_ok, y_ok, duration=0.2)
        pyautogui.click()

        print("Les boutons ont été cliqués avec succès.")
    except Exception as e:
        print(f"Erreur lors de la gestion des popups : {e}")



def save_to_new_excel(dataframe, new_excel_path, sheet_name="CONTRAT"):
    """Crée un nouveau fichier Excel avec la feuille spécifiée sans boîtes de dialogue."""
    try:
        with pd.ExcelWriter(new_excel_path, engine="openpyxl") as writer:
            dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as e:
        print(f"Erreur lors de la sauvegarde du fichier Excel : {e}")


def perform_mail_merge(word_file, replaced_name, replacing_name, start_date, pdf_folder, contract_type="MAR"):
    print(f"📄 Dossier PDF : {pdf_folder}")
    print(f"📄 Modèle Word : {word_file}")
    print(f"👨‍⚕️ Remplacé : {replaced_name}")
    print(f"👨‍⚕️ Remplaçant : {replacing_name}")
    print(f"📅 Date début : {start_date}")
    print(f"📄 Type contrat : {contract_type}")

    # Assurer que le dossier cible existe
    pdf_folder = ensure_directory_exists(pdf_folder)

    # Formatage de la date pour le nom du fichier
    formatted_date = datetime.strptime(start_date, "%Y-%m-%d").strftime("%Y%m%d")

    # Génération du nom de fichier PDF
    if contract_type == "MAR":
        if not replaced_name or not replacing_name:
            print("❌ Erreur : Noms des médecins incorrects.")
            return None
        pdf_filename = f"{formatted_date}_Contrat_{replacing_name}_{replaced_name}.pdf"
    elif contract_type == "IADE":
        if not replacing_name:
            print("❌ Erreur : Nom du remplaçant manquant.")
            return None
        pdf_filename = f"{formatted_date}_CDD_{replacing_name.strip()}.pdf"
    else:
        print(f"❌ Erreur : Type de contrat inconnu {contract_type}")
        return None

    output_pdf_path = os.path.join(pdf_folder, pdf_filename)
    print(f"📄 Chemin prévu pour le PDF : {output_pdf_path}")

    # 📌 Script AppleScript mis à jour
    applescript = f"""
    tell application "Microsoft Word"
        activate
        open POSIX file "{word_file}"
        delay 3 -- Attendre que Word charge le document

        -- Sauvegarde en PDF
        set output_pdf_path to POSIX file "{output_pdf_path}"
        save as active document file name output_pdf_path file format format PDF

        -- Fermer le document sans enregistrer
        close active document saving no
    end tell
    """

    try:
        # 📌 Lancer Word en arrière-plan pour éviter que Python ne soit bloqué
        word_process = subprocess.Popen(["osascript", "-e", applescript])

        # 📌 Attendre un peu que Word ouvre sa fenêtre
        time.sleep(2)

        # 📌 Lancer handle_popups() en **parallèle** pour cliquer sur "Oui" et "OK"
        popup_thread = threading.Thread(target=handle_popups)
        popup_thread.start()

        # 📌 Attendre que Word termine (optionnel, si besoin)
        word_process.wait()

        print(f"✅ PDF généré : {output_pdf_path}")

        if os.path.exists(output_pdf_path):
            return output_pdf_path  
        else:
            print("❌ Erreur : Fichier PDF non trouvé après génération.")
            return None  

    except Exception as e:
        print(f"❌ Erreur lors du publipostage : {e}")
        return None  
    
def open_contract_creation_iade():
    """Affiche le formulaire de création de contrat IADE dans le panneau droit."""
    clear_right_frame()  # Nettoie le panneau droit avant d'y ajouter de nouveaux éléments
    
    try:
        iade_data = pd.read_excel(file_paths["excel_iade"], sheet_name="Coordonnées IADEs")
    except FileNotFoundError:
        print("Erreur : Le fichier Excel pour les contrats IADE est introuvable.")
        return
    except Exception as e:
        print(f"Erreur lors de l'ouverture du fichier Excel IADE : {e}")
        return
    
    # Conteneur principal - utilisation d'un PanedWindow pour diviser l'espace
    main_container = tk.PanedWindow(right_frame, orient=tk.HORIZONTAL)
    main_container.pack(fill="both", expand=True)
    
    # Cadre gauche pour le formulaire
    form_container = tk.Frame(main_container, bg="#f0f4f7", padx=20, pady=20)
    main_container.add(form_container, width=420)  # Largeur fixe pour le formulaire
    
    # Cadre droit pour les actions post-contrat (initialement vide)
    actions_container = tk.Frame(main_container, bg="#f5f5f5", padx=20, pady=20)
    main_container.add(actions_container, width=400)  # Largeur pour les actions
    
    # Titre du formulaire
    tk.Label(form_container, text="🩺 Nouveau contrat IADE", 
            font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    # Fonction pour sélectionner les dates
    def select_dates():
        """Ouvre un calendrier pour sélectionner les dates de début et de fin."""
        selected_dates = []

        def on_date_select():
            selected_date = calendar.get_date()
            selected_dates.append(selected_date)
            if len(selected_dates) == 1:
                start_date_var.set(selected_date)
                message_var.set("Sélectionnez la date de fin.")
            elif len(selected_dates) == 2:
                end_date_var.set(selected_date)
                date_picker.destroy()

        date_picker = Toplevel(root)
        date_picker.title("Sélectionner les dates")

        message_var = StringVar(value="Sélectionnez la date de début.")
        Label(date_picker, textvariable=message_var, font=("Arial", 12)).pack(pady=5)

        calendar = Calendar(
            date_picker,
            date_pattern="yyyy-mm-dd",
            background="white",
            foreground="black",
            selectbackground="blue",
            selectforeground="white",
        )
        calendar.pack(pady=10)

        Button(date_picker, text="Valider", command=on_date_select).pack(pady=5)
        Button(date_picker, text="Fermer", command=date_picker.destroy).pack(pady=5)
    
    # Fonction pour enregistrer le contrat
    def save_contract_iade():
        """Génère un fichier Excel et un contrat IADE, puis affiche les options post-contrat."""
        # Récupération des valeurs depuis l'interface
        replacing_name = replacing_var.get()
        start_date = start_date_var.get()
        end_date = end_date_var.get()
        sign_date = sign_date_var.get()
        nb_hours = nb_hours_var.get()

        # Vérification que tous les champs sont bien remplis
        if not replacing_name or not start_date or not end_date or not sign_date or not nb_hours:
            print("❌ Erreur : Tous les champs doivent être remplis.")
            return

        try:
            # Récupérer les données de l'IADE depuis la base
            replacing_data = iade_data[iade_data["NOMR"] == replacing_name].iloc[0]
            replacing_email = replacing_data["EMAIL"]
            
            # Correction : Récupération du prénom et nom complet
            replacing_full_name = f"{replacing_data['PRENOMR']} {replacing_data['NOMR']}".strip()

        except IndexError:
            print(f"❌ Erreur : Impossible de trouver l'IADE '{replacing_name}' dans la base.")
            return

        # Correction de l'email pour éviter 'nan'
        if pd.isna(replacing_email) or replacing_email in ["", "nan"]:
            replacing_email = "email_inconnu@exemple.com"

        # Définir les colonnes du fichier Excel temporaire
        columns = [
            "SEXE", "NOMR", "PRENOMR", "DDNR", "ER", "ILR", "SALARIER",
            "LIEUNR", "DPTN", "ADRESSER", "NOSSR", "NATR", "EMAIL",
            "DATEDEBUT", "DATEFIN", "DATESIGN", "NBHEURES"
        ]

        # Récupérer les informations de l'IADE sélectionné
        row_data = {
            "SEXE": replacing_data["SEXE"],
            "NOMR": replacing_data["NOMR"],
            "PRENOMR": replacing_data["PRENOMR"],
            "DDNR": replacing_data["DDNR"],
            "ER": replacing_data["ER"],
            "ILR": replacing_data["ILR"],
            "SALARIER": replacing_data["SALARIER"],
            "LIEUNR": replacing_data["LIEUNR"],
            "DPTN": replacing_data["DPTN"],
            "ADRESSER": replacing_data["ADRESSER"],
            "NOSSR": replacing_data["NOSSR"],
            "NATR": replacing_data["NATR"],
            "EMAIL": replacing_email,
            "DATEDEBUT": start_date,
            "DATEFIN": end_date,
            "DATESIGN": sign_date,
            "NBHEURES": nb_hours
        }

        # Créer un DataFrame avec les données
        contrat_iade = pd.DataFrame([row_data])

        # Définir le chemin du fichier temporaire
        excel_temp_path = "/Users/vincentperreard/Documents/Contrats/IADE/temp_contrat_iade.xlsx"

        # Sauvegarder dans un fichier Excel temporaire
        with pd.ExcelWriter(excel_temp_path, engine="openpyxl") as writer:
            contrat_iade.to_excel(writer, index=False, sheet_name="CONTRAT")

        print(f"✅ Données du contrat IADE enregistrées dans {excel_temp_path}")

        # Effectuer le publipostage et générer le PDF
        pdf_path = perform_mail_merge(
            file_paths["word_iade"],  # Modèle Word pour IADE
            None,  # IADE n'a pas de remplacé
            replacing_full_name,
            start_date,
            os.path.expanduser(file_paths["pdf_iade"]),
            contract_type="IADE"
        )

        if not pdf_path:
            print("❌ Erreur : Le fichier PDF n'a pas pu être généré.")
            return
            
        print(f"✅ Contrat IADE enregistré sous : {pdf_path}")
        
        # Désactiver les éléments du formulaire
        for widget in form_frame.winfo_children():
            if isinstance(widget, (tk.Entry, tk.Button, tk.OptionMenu)):
                widget.configure(state="disabled")
                
        # Vider le conteneur d'actions
        for widget in actions_container.winfo_children():
            widget.destroy()
            
        # Titre
        tk.Label(actions_container, text="Actions sur le contrat", 
                font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
        
        # Information sur le contrat généré
        info_text = f"Contrat généré pour:\n{replacing_full_name}\ndu {start_date} au {end_date}"
        tk.Label(actions_container, text=info_text, justify=tk.LEFT, 
                bg="#f5f5f5", padx=10, pady=5).pack(fill="x", pady=10)
        
        # Bouton pour ouvrir le PDF
        def open_pdf():
            subprocess.run(["open", "-a", "PDF Expert", pdf_path])
            
        tk.Button(actions_container, text="📄 Ouvrir avec PDF Expert", 
                command=open_pdf, width=30).pack(pady=5)
        
        # Bouton pour envoyer à DocuSign
        tk.Button(actions_container, text="📩 Envoyer en DocuSign", 
                command=lambda: send_to_docusign(pdf_path, "IADE", start_date, end_date, 
                                                replacing_full_name, replacing_email, 
                                                None, None), 
                width=30).pack(pady=5)
        
        # Bouton pour le règlement (désactivé)
        tk.Button(actions_container, text="💰 Effectuer le règlement (à venir)", 
                state="disabled", width=30).pack(pady=5)
        
        # Bouton pour revenir à l'accueil
        tk.Button(actions_container, text="🏠 Retour à l'accueil", 
                command=lambda: [clear_right_frame(), show_welcome_image()], 
                width=30).pack(pady=20)
        
    # Formulaire principal
    form_frame = tk.Frame(form_container, bg="#f0f4f7")
    form_frame.pack(pady=10, fill="x")
    
    # IADE remplaçant
    tk.Label(form_frame, text="IADE remplaçant :", bg="#f0f4f7").grid(row=0, column=0, sticky="w", pady=5)
    replacing_var = StringVar()
    replacing_menu = tk.OptionMenu(form_frame, replacing_var, *iade_data["NOMR"].tolist())
    replacing_menu.grid(row=0, column=1, sticky="w", padx=5, pady=5)
    
    # Bouton pour sélectionner les dates
    tk.Label(form_frame, text="Dates de début et de fin :", bg="#f0f4f7").grid(row=1, column=0, sticky="w", pady=5)
    date_btn = Button(form_frame, text="📅 Sélectionner les dates", command=select_dates)
    date_btn.grid(row=1, column=1, sticky="w", padx=5, pady=5)
    
    start_date_var = StringVar()
    end_date_var = StringVar()
    
    # Date de début
    tk.Label(form_frame, text="Date de début :", bg="#f0f4f7").grid(row=2, column=0, sticky="w", pady=5)
    start_entry = Entry(form_frame, textvariable=start_date_var, state="readonly")
    start_entry.grid(row=2, column=1, sticky="w", padx=5, pady=5)
    
    # Date de fin
    tk.Label(form_frame, text="Date de fin :", bg="#f0f4f7").grid(row=3, column=0, sticky="w", pady=5)
    end_entry = Entry(form_frame, textvariable=end_date_var, state="readonly")
    end_entry.grid(row=3, column=1, sticky="w", padx=5, pady=5)
    
    # Date de signature
    tk.Label(form_frame, text="Date de signature :", bg="#f0f4f7").grid(row=4, column=0, sticky="w", pady=5)
    sign_date_var = StringVar(value=datetime.today().strftime("%Y-%m-%d"))
    sign_entry = Entry(form_frame, textvariable=sign_date_var)
    sign_entry.grid(row=4, column=1, sticky="w", padx=5, pady=5)
    
    # Nombre d'heures par jour
    tk.Label(form_frame, text="Nombre d'heures par jour :", bg="#f0f4f7").grid(row=5, column=0, sticky="w", pady=5)
    nb_hours_var = StringVar(value="9")
    hours_entry = Entry(form_frame, textvariable=nb_hours_var)
    hours_entry.grid(row=5, column=1, sticky="w", padx=5, pady=5)
    
    # Bouton de création et annulation
    create_btn = Button(form_frame, text="Créer le contrat", command=save_contract_iade, 
                       font=("Arial", 12, "bold"), bg="#007ACC", fg="black")
    create_btn.grid(row=6, column=0, pady=10, padx=5, sticky="w")
    
    cancel_btn = Button(form_frame, text="Annuler", command=lambda: [clear_right_frame(), show_welcome_image()], 
                       font=("Arial", 10), bg="#f44336", fg="black")
    cancel_btn.grid(row=6, column=1, pady=10, padx=5, sticky="w")



def open_contract_creation_mar():
    """Affiche le formulaire de création de contrat MAR dans le panneau droit."""
    clear_right_frame()  # Nettoie le panneau droit avant d'y ajouter de nouveaux éléments
    
    # Chargement des données
    mars_selarl = pd.read_excel(file_paths["excel_mar"], sheet_name="MARS SELARL")
    mars_rempla = pd.read_excel(file_paths["excel_mar"], sheet_name="MARS Remplaçants")
    
    # Concaténation PRENOM + NOM en une seule colonne normalisée
    mars_selarl["FULL_NAME"] = mars_selarl["PRENOM"].fillna("").str.strip() + " " + mars_selarl["NOM"].fillna("").str.strip()
    mars_rempla["FULL_NAME"] = mars_rempla["PRENOMR"].fillna("").str.strip() + " " + mars_rempla["NOMR"].fillna("").str.strip()
    
    if mars_selarl is None or mars_rempla is None:
        return
    
    # Conteneur principal - utilisation d'un PanedWindow pour diviser l'espace
    main_container = tk.PanedWindow(right_frame, orient=tk.HORIZONTAL)
    main_container.pack(fill="both", expand=True)
    
    # Cadre gauche pour le formulaire
    form_container = tk.Frame(main_container, bg="#f0f4f7", padx=20, pady=20)
    main_container.add(form_container, width=420)  # Largeur fixe pour le formulaire
    
    # Cadre droit pour les actions post-contrat (initialement vide)
    actions_container = tk.Frame(main_container, bg="#f5f5f5", padx=20, pady=20)
    main_container.add(actions_container, width=400)  # Largeur pour les actions
    
    # Titre du formulaire
    tk.Label(form_container, text="Nouveau contrat remplacement MAR", 
            font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    def select_dates():
        """Ouvre un calendrier pour sélectionner les dates de début et de fin."""
        selected_dates = []

        def update_message():
            """Mise à jour du message selon le nombre de dates sélectionnées."""
            if len(selected_dates) == 0:
                message_var.set("Sélectionnez la date de début.")
            elif len(selected_dates) == 1:
                message_var.set("Sélectionnez la date de fin.")

        def on_date_select():
            """Capture la date sélectionnée."""
            selected_date = calendar.get_date()
            selected_dates.append(selected_date)
            update_message()

            if len(selected_dates) == 2:  # Deux dates sélectionnées
                start_date_var.set(min(selected_dates))
                end_date_var.set(max(selected_dates))
                date_picker.destroy()
            elif len(selected_dates) == 1:  # Une seule date sélectionnée
                start_date_var.set(selected_date)
                end_date_var.set(selected_date)

        def close_calendar():
            """Ferme le calendrier en ne retenant qu'une seule date si nécessaire."""
            if len(selected_dates) == 1:  # Si une seule date est sélectionnée
                start_date_var.set(selected_dates[0])
                end_date_var.set(selected_dates[0])
            date_picker.destroy()

        # Crée la fenêtre de calendrier (reste un Toplevel car c'est une fenêtre modale)
        date_picker = Toplevel(root)
        date_picker.title("Sélectionner les dates")

        message_var = StringVar()
        message_var.set("Sélectionnez la date de début.")
        Label(date_picker, text="Sélectionner les dates", font=("Arial", 12, "bold")).pack(pady=5)
        Label(date_picker, textvariable=message_var, font=("Arial", 10)).pack(pady=5)

        calendar = Calendar(
            date_picker,
            date_pattern="yyyy-mm-dd",
            background="white",
            foreground="black",
            selectbackground="blue",
            selectforeground="white",
        )
        calendar.pack(pady=10)

        Button(date_picker, text="Valider", command=on_date_select).pack(pady=5, side=tk.LEFT, padx=10)
        Button(date_picker, text="Fermer", command=close_calendar).pack(pady=5, side=tk.RIGHT, padx=10)

        pass
    
    def save_contract():
        """Sauvegarde les informations du contrat et génère le PDF."""
        print("file_paths keys:", file_paths.keys())
        
        replaced_name = replaced_var.get().strip()
        replacing_name = replacing_var.get().strip()
        start_date = start_date_var.get()
        end_date = end_date_var.get()
        sign_date = sign_date_var.get()
        daily_fee = daily_fee_var.get()

        # Vérification insensible à la casse et aux espaces
        replaced_matches = mars_selarl[mars_selarl["FULL_NAME"].str.strip().str.upper() == replaced_name.strip().upper()]
        if replaced_matches.empty:
            print(f"❌ Erreur : Médecin remplacé '{replaced_name}' introuvable dans la base.")
            print("📋 Liste des noms disponibles :", mars_selarl["FULL_NAME"].tolist())
            return
        replaced_data = replaced_matches.iloc[0]

        replacing_matches = mars_rempla[mars_rempla["FULL_NAME"].str.strip().str.upper() == replacing_name.strip().upper()]
        if replacing_matches.empty:
            print(f"❌ Erreur : Médecin remplaçant '{replacing_name}' introuvable dans la base.")
            print("📋 Liste des noms disponibles :", mars_rempla["FULL_NAME"].tolist())
            return
        replacing_data = replacing_matches.iloc[0]

        print(f"🛠️ DEBUG : Données du remplacé : {replaced_data.to_dict()}")
        print(f"🛠️ DEBUG : Données du remplaçant : {replacing_data.to_dict()}")

        # Vérification des dates
        if not start_date or not end_date:
            print("❌ Erreur : Les dates ne sont pas définies correctement.")
            return

        # Récupération et nettoyage des emails
        replaced_email = replaced_data.get("EMAIL", "").strip()
        if pd.isna(replaced_email) or not replaced_email:
            replaced_email = "email_inconnu@exemple.com"
            
        replacing_email = replacing_data.get("EMAILR", "").strip()
        if pd.isna(replacing_email) or not replacing_email:
            replacing_email = "email_inconnu@exemple.com"

        # Création d'une ligne pour la feuille CONTRAT
        new_row = [
            replaced_data["NOM"],
            replaced_data["PRENOM"],
            str(replaced_data["N ORDRE"]),
            replaced_email,
            replacing_data["NOMR"],
            replacing_data["PRENOMR"],
            str(replacing_data["N ORDRER"]),
            replacing_email,
            replacing_data["AdresseR"],
            str(replacing_data["URSSAF"]),
            start_date,
            end_date,
            sign_date,
            int(daily_fee)
        ]

        # Création d'un DataFrame avec la structure de la feuille CONTRAT
        contrat_columns = [
            "NOM", "PRENOM", "N ORDRE", "EMAIL",
            "NOMR", "PRENOMR", "N ORDRER", "EMAILR",
            "AdresseR", "URSSAF", "DATEDEBUT", "DATEFIN", "DATESIGN", "forfait"
        ]
        contrat = pd.DataFrame([new_row], columns=contrat_columns)

        # Préparation pour le publipostage
        word_file = file_paths["word_mar"]
        base_dir = os.path.dirname(word_file)
        excel_path = os.path.join(base_dir, "contrat_mars_selarl.xlsx")
        
        # Sauvegarde du fichier Excel temporaire
        save_to_new_excel(contrat, excel_path)

        # Génération du contrat PDF
        pdf_path = perform_mail_merge(
            file_paths["word_mar"],
            replaced_name,
            replacing_name,
            start_date,
            os.path.expanduser(file_paths["pdf_mar"]),
            contract_type="MAR"
        )

        if not pdf_path or not os.path.exists(pdf_path):
            print(f"❌ Erreur : Le fichier PDF n'a pas pu être généré.")
            return

        print(f"✅ Contrat généré : {pdf_path}")

        # Désactivation des widgets du formulaire
        for widget in form_frame.winfo_children():
            if isinstance(widget, (tk.Entry, tk.Button, tk.OptionMenu)):
                widget.configure(state="disabled")
        
        # Vider le conteneur d'actions
        for widget in actions_container.winfo_children():
            widget.destroy()
        
        # Création du conteneur pour les actions post-contrat
        title_label = tk.Label(actions_container, text="Actions sur le contrat", 
                            font=("Arial", 14, "bold"), bg="#4a90e2", fg="white")
        title_label.pack(fill="x", pady=10)
        
        # Information sur le contrat généré
        info_text = f"Contrat généré pour:\n{replacing_name}\nremplaçant {replaced_name}\ndu {start_date} au {end_date}"
        info_label = tk.Label(actions_container, text=info_text, justify=tk.LEFT, 
                            bg="#f5f5f5", padx=10, pady=5)
        info_label.pack(fill="x", pady=10)
        
        # Fonction pour ouvrir le PDF
        def open_pdf():
            subprocess.run(["open", "-a", "PDF Expert", pdf_path])
        
        # Bouton pour ouvrir le PDF
        open_button = tk.Button(actions_container, text="📄 Ouvrir avec PDF Expert", 
                            command=open_pdf, width=30)
        open_button.pack(pady=5)
        
        # Bouton pour envoyer à DocuSign
        docusign_button = tk.Button(actions_container, text="📩 Envoyer en DocuSign", 
                            command=lambda: send_to_docusign(pdf_path, "MAR", start_date, end_date, 
                                                            replacing_name, replacing_email, 
                                                            replaced_name, replaced_email), 
                            width=30)
        docusign_button.pack(pady=5)
        
        # Bouton pour le règlement (désactivé)
        pay_button = tk.Button(actions_container, text="💰 Effectuer le règlement (à venir)", 
                            state="disabled", width=30)
        pay_button.pack(pady=5)
        
        # Bouton pour revenir à l'écran d'accueil
        home_button = tk.Button(actions_container, text="🏠 Retour à l'accueil", 
                            command=lambda: [clear_right_frame(), show_welcome_image()], 
                            width=30)
        home_button.pack(pady=20)


   # Formulaire principal
    form_frame = tk.Frame(form_container, bg="#f0f4f7")
    form_frame.pack(pady=10, fill="x")
    
    # Médecin remplacé
    tk.Label(form_frame, text="Médecin remplacé :", bg="#f0f4f7").grid(row=0, column=0, sticky="w", pady=5)
    replaced_var = StringVar()
    tk.OptionMenu(form_frame, replaced_var, *mars_selarl["FULL_NAME"].tolist()).grid(row=0, column=1, sticky="w", padx=5, pady=5)
    
    # Médecin remplaçant
    tk.Label(form_frame, text="Médecin remplaçant :", bg="#f0f4f7").grid(row=1, column=0, sticky="w", pady=5)
    replacing_var = StringVar()
    tk.OptionMenu(form_frame, replacing_var, *mars_rempla["FULL_NAME"].tolist()).grid(row=1, column=1, sticky="w", padx=5, pady=5)
    
    # Bouton pour sélectionner les dates
    tk.Label(form_frame, text="Dates de début et de fin :", bg="#f0f4f7").grid(row=2, column=0, sticky="w", pady=5)
    Button(form_frame, text="📅 Sélectionner les dates", command=select_dates).grid(row=2, column=1, sticky="w", padx=5, pady=5)
    
    start_date_var = StringVar()
    end_date_var = StringVar()
    
    # Champ pour afficher la date de début
    tk.Label(form_frame, text="Date de début :", bg="#f0f4f7").grid(row=3, column=0, sticky="w", pady=5)
    Entry(form_frame, textvariable=start_date_var, state="readonly").grid(row=3, column=1, sticky="w", padx=5, pady=5)
    
    # Champ pour afficher la date de fin
    tk.Label(form_frame, text="Date de fin :", bg="#f0f4f7").grid(row=4, column=0, sticky="w", pady=5)
    Entry(form_frame, textvariable=end_date_var, state="readonly").grid(row=4, column=1, sticky="w", padx=5, pady=5)
    
    # Date de signature
    tk.Label(form_frame, text="Date de signature :", bg="#f0f4f7").grid(row=5, column=0, sticky="w", pady=5)
    sign_date_var = StringVar(value=datetime.today().strftime("%Y-%m-%d"))
    Entry(form_frame, textvariable=sign_date_var).grid(row=5, column=1, sticky="w", padx=5, pady=5)
    
    # Forfait journalier
    tk.Label(form_frame, text="Forfait journalier (€) :", bg="#f0f4f7").grid(row=6, column=0, sticky="w", pady=5)
    daily_fee_var = StringVar(value="1000")
    Entry(form_frame, textvariable=daily_fee_var).grid(row=6, column=1, sticky="w", padx=5, pady=5)
    
    # Bouton de création
    Button(form_frame, text="Créer le contrat", command=save_contract, 
           font=("Arial", 12, "bold"), bg="#007ACC", fg="black").grid(row=7, column=0, pady=10, padx=5, sticky="w")
    
    # Bouton de retour
    Button(form_frame, text="Annuler", command=lambda: [clear_right_frame(), show_welcome_image()], 
           font=("Arial", 10), bg="#f44336", fg="black").grid(row=7, column=1, pady=10, padx=5, sticky="w")
    
def open_accounting_menu():
    """Affiche le menu de gestion comptable dans le panneau droit."""
    clear_right_frame()  # Nettoie le panneau droit avant d'y ajouter de nouveaux éléments
    
    # Conteneur principal - utilisation d'un PanedWindow pour diviser l'espace
    main_container = tk.PanedWindow(right_frame, orient=tk.HORIZONTAL)
    main_container.pack(fill="both", expand=True)
    
    # Cadre gauche pour le menu comptabilité (même largeur que le menu principal)
    compta_frame = tk.Frame(main_container, bg="#f0f4f7", padx=20, pady=20)
    main_container.add(compta_frame, width=300)  # Même largeur que left_frame
    
    # Cadre droit pour l'affichage des fonctionnalités (initialement vide)
    content_container = tk.Frame(main_container, bg="#f5f5f5", padx=20, pady=20)
    main_container.add(content_container, width=900)  # Plus de place pour le contenu
    
    # Titre du menu
    tk.Label(compta_frame, text="📊 Menu Comptabilité", font=("Arial", 14, "bold"), 
             bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    # Boutons du menu comptabilité
    tk.Button(compta_frame, text="📄 Bulletins de salaire", 
              command=lambda: display_bulletins_in_container(content_container),
              width=25, bg="#DDDDDD", fg="black", 
              font=("Arial", 10, "bold")).pack(pady=10)
    
    tk.Button(compta_frame, text="📂 Frais et factures", 
              command=lambda: display_factures_in_container(content_container), 
              width=25, bg="#DDDDDD", fg="black", 
              font=("Arial", 10, "bold")).pack(pady=10)
    
    tk.Button(compta_frame, text="💰 Effectuer un virement", 
              command=lambda: display_transfer_in_container(content_container), 
              width=25, bg="#DDDDDD", fg="black", 
              font=("Arial", 10, "bold")).pack(pady=10)
    
    tk.Button(compta_frame, text="💰 Virement rémunération MARS", 
              command=lambda: display_virement_mar_in_container(content_container),
              width=25, bg="#DDDDDD", fg="black", 
              font=("Arial", 10, "bold")).pack(pady=10)
    
    # Retour au menu principal
    tk.Button(compta_frame, text="🔙 Retour au menu principal", 
              command=lambda: [clear_right_frame(), show_welcome_image()], 
              width=25, bg="#BBBBBB", fg="black", 
              font=("Arial", 10, "bold")).pack(pady=20)
    
    # Afficher un message d'accueil initial dans le conteneur de droite
    tk.Label(content_container, text="Bienvenue dans le module comptabilité", 
             font=("Arial", 16, "bold"), fg="#4a90e2", bg="#f5f5f5").pack(pady=20)
    tk.Label(content_container, text="Choisissez une option dans le menu de gauche.", 
             font=("Arial", 12), bg="#f5f5f5").pack(pady=10)


def display_bulletins_in_container(container):
    """Affiche les bulletins de salaire dans le conteneur spécifié."""
    # Vider le conteneur
    for widget in container.winfo_children():
        widget.destroy()
    
    # Créer un cadre pour les bulletins
    bulletins_frame = tk.Frame(container, bg="#f0f0f0")
    bulletins_frame.pack(fill="both", expand=True)
    
    # Importer la fonction du module bulletins
    from bulletins import show_bulletins_in_frame
    
    # Appeler la fonction avec le cadre créé
    show_bulletins_in_frame(bulletins_frame)


def display_factures_in_container(container):
    """Affiche une interface pour lancer l'analyse des factures."""
    # Vider le conteneur
    for widget in container.winfo_children():
        widget.destroy()
    
    # Créer un frame pour l'analyse des factures
    factures_frame = tk.Frame(container, bg="#f5f5f5", padx=20, pady=20)
    factures_frame.pack(fill="both", expand=True)
    
    # Titre
    tk.Label(factures_frame, text="📂 Analyse des factures", 
             font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    # Description
    tk.Label(factures_frame, text="Cette fonction va lancer l'outil d'analyse des factures dans une fenêtre séparée.", 
             font=("Arial", 12), bg="#f5f5f5").pack(pady=20)
    
    # Bouton pour lancer l'analyse
    tk.Button(factures_frame, text="🚀 Lancer l'analyse des factures", 
              command=launch_facture_analysis,
              width=30, height=2, bg="#FFA500", fg="black", 
              font=("Arial", 12, "bold")).pack(pady=20)


def display_transfer_in_container(container):
    """Affiche le menu de virement dans le conteneur spécifié."""
    # Vider le conteneur
    for widget in container.winfo_children():
        widget.destroy()
    
    # Créer un frame pour le menu de virement
    transfer_frame = tk.Frame(container, bg="#f5f5f5", padx=20, pady=20)
    transfer_frame.pack(fill="both", expand=True)
    
    # Titre
    tk.Label(transfer_frame, text="💰 Effectuer un virement", 
             font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    # Message d'instruction
    tk.Label(transfer_frame, text="Cette fonctionnalité permet d'effectuer des virements aux fournisseurs.",
             font=("Arial", 12), bg="#f5f5f5").pack(pady=20)
    
    # Bouton pour lancer la génération d'un virement
    tk.Button(transfer_frame, text="🔄 Générer un virement",
              command=lambda: open_virement_selection_window(),
              width=30, height=2, bg="#4CAF50", fg="black",
              font=("Arial", 12, "bold")).pack(pady=20)


def display_virement_mar_in_container(container):
    """Affiche le menu de virement pour les MARS dans le conteneur spécifié."""
    # Vider le conteneur
    for widget in container.winfo_children():
        widget.destroy()
    
    # Créer un frame pour le menu de virement MAR
    mar_frame = tk.Frame(container, bg="#f5f5f5", padx=20, pady=20)
    mar_frame.pack(fill="both", expand=True)
    
    # Titre
    tk.Label(mar_frame, text="💰 Virement rémunération MARS", 
             font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    
    # Récupération du fichier depuis le chemin configuré
    excel_file = file_paths.get("chemin_fichier_virement", "")
    
    # Message d'instruction
    instruction_text = (
        "1️⃣ Vérifiez attentivement chaque ligne des virements prévus.\n"
        "2️⃣ Faites les modifications nécessaires.\n"
        "3️⃣ Cliquez sur 'Sauvegarder et Valider les virements'."
    )
    
    tk.Label(mar_frame, text=instruction_text, font=("Arial", 12), 
             bg="#f5f5f5", fg="black", justify="left").pack(pady=20, padx=20)
    
    # Bouton pour ouvrir le fichier Excel
    tk.Button(mar_frame, text="📊 Ouvrir le fichier Excel des virements", 
              command=lambda: subprocess.run(["open", "-a", "Microsoft Excel", excel_file], check=True) if os.path.exists(excel_file) else messagebox.showerror("Erreur", f"Le fichier {excel_file} est introuvable."),
              bg="#4682B4", fg="black", font=("Arial", 12, "bold"), width=30).pack(pady=10)
    
    # Bouton pour sauvegarder et valider les virements
    tk.Button(mar_frame, text="💾 Sauvegarder et Valider les virements", 
              command=lambda: open_virement_selection_window(),
              bg="#4caf50", fg="black", font=("Arial", 12, "bold"), width=30).pack(pady=10)



def display_in_right_frame(function_to_display):
    """Fonction helper pour afficher une fonction dans le panneau droit"""
    global right_frame
    
    # On sauvegarde la fonction originale
    original_function = function_to_display
    
    # On définit une nouvelle fonction qui va intercepter la fenêtre toplevel
    def wrapped_function(*args, **kwargs):
        # Nettoyer le panneau droit
        for widget in right_frame.winfo_children():
            widget.destroy()
        
        # Créer un cadre spécial qui va servir de "fausse fenêtre Toplevel"
        # pour la fonction d'origine
        fake_toplevel_frame = tk.Frame(right_frame, bg="#f0f0f0")
        fake_toplevel_frame.pack(fill="both", expand=True)
        
        # Stocke le cadre actuel pour les fonctions qui supposent être dans une fenêtre Toplevel
        original_function.__globals__['current_frame'] = fake_toplevel_frame
        
        # Appeler la fonction d'origine mais rediriger ses widgets dans notre cadre
        original_function(*args, **kwargs)
        
        # Ajouter un bouton de retour au menu comptabilité
        tk.Button(fake_toplevel_frame, text="🔙 Retour au menu comptabilité", 
                  command=open_accounting_menu, 
                  bg="#B0C4DE", fg="black", 
                  font=("Arial", 10, "bold")).pack(pady=10)
    
    # Exécuter notre fonction wrapper
    wrapped_function()

def launch_facture_analysis():
    """Lance le script analyse_facture.py dans un processus séparé."""
    script_path = "/Users/vincentperreard/script contrats/analyse_facture.py"
    factures_path = get_file_path("dossier_factures", verify_exists=True)
    
    # DEBUGGING: Afficher le chemin utilisé
    print("\n" + "="*50)
    print("🔍 DEBUG - LANCEMENT ANALYSE FACTURES")
    print(f"📂 Chemin du script: {script_path}")
    print(f"📂 Chemin des factures trouvé: {factures_path}")
    print("="*50 + "\n")
    
    try:
        if factures_path:
            # Passer le chemin comme argument avec --dossier
            cmd = ["python3", script_path, "--dossier", factures_path]
            print(f"🚀 Exécution de la commande: {' '.join(cmd)}")
            # Utiliser Popen pour ne pas bloquer l'interface
            subprocess.Popen(cmd)
        else:
            print("⚠️ Chemin des factures non trouvé, utilisation du chemin par défaut.")
            subprocess.Popen(["python3", script_path])
            print("✅ Analyse des factures lancée avec succès (chemin par défaut).")
    except Exception as e:
        print(f"❌ Erreur lors de l'exécution du script : {e}")
        messagebox.showerror("Erreur", f"Impossible de lancer l'analyse des factures : {e}")
        
def open_transfer_menu():
    """Ouvre le menu pour effectuer un virement."""
    messagebox.showinfo("Effectuer un virement", "Fonction à implémenter : Générer un virement.")


def open_virement_mar():
    """
    Ouvre le fichier Excel des virements et affiche une fenêtre d'instruction.
    """
    # 📌 Récupération du fichier depuis le chemin configuré
    excel_file = file_paths["chemin_fichier_virement"]    
    if not os.path.exists(excel_file):
        messagebox.showerror("Erreur", f"Le fichier des virements '{excel_file}' est introuvable.")
        return

    # 📌 Ouvrir le fichier Excel sur la feuille "Virement"
    try:
        subprocess.run(["open", "-a", "Microsoft Excel", excel_file], check=True)
    except Exception as e:
        messagebox.showerror("Erreur", f"Impossible d'ouvrir Excel : {e}")
        return

    # 📌 Création de la fenêtre d'instruction
    instruction_window = tk.Toplevel()
    instruction_window.title("🔹 Instructions pour les Virements MAR")
    instruction_window.geometry("500x300")
    instruction_window.configure(bg="#f2f7ff")

    # 📌 Message d'instruction
    instruction_text = (
        "1️⃣ Vérifiez attentivement chaque ligne des virements prévus.\n"
        "2️⃣ Faites les modifications nécessaires.\n"
        "3️⃣ Cliquez sur 'Sauvegarder et Valider les virements'."
    )

    tk.Label(instruction_window, text="✅ Instructions", font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    tk.Label(instruction_window, text=instruction_text, font=("Arial", 12), bg="#f2f7ff", fg="black", justify="left").pack(pady=20, padx=20)

    def save_and_validate():
        """
        Sauvegarde le fichier Excel, ferme Excel et ouvre la sélection des virements.
        """
        try:
            # 📌 Sauvegarde du fichier
            df = pd.read_excel(excel_file, sheet_name="Virement")
            with pd.ExcelWriter(excel_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name="Virement", index=False)

            # 📌 Fermer Excel proprement via AppleScript
            applescript = """
            tell application "Microsoft Excel"
                if (count of windows) > 0 then
                    save active workbook
                    close active workbook
                end if
                quit
            end tell
            """
            subprocess.run(["osascript", "-e", applescript], check=True)

            # 📌 Ouvrir la sélection des virements
            instruction_window.destroy()
            open_virement_selection_window()

        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de sauvegarder le fichier : {e}")

    # 📌 Bouton "Sauvegarder et Valider"
    tk.Button(
        instruction_window, text="💾 Sauvegarder et Valider les virements", command=save_and_validate,
        bg="#4caf50", fg="black", font=("Arial", 12, "bold"), width=30
    ).pack(pady=20)

    # 📌 Bouton "Annuler"
    tk.Button(
        instruction_window, text="❌ Annuler", command=instruction_window.destroy,
        bg="#f44336", fg="black", font=("Arial", 12, "bold"), width=30
    ).pack(pady=10)

def open_virement_selection_window():
    """Ouvre la fenêtre pour sélectionner les virements MAR à traiter."""

    # 📌 Vérification du fichier des virements
    excel_file = file_paths.get("chemin_fichier_virement", "/Users/vincentperreard/Dropbox/SEL:SPFPL Mathilde/Compta SEL/compta SEL -rempla 2025.xlsx")
    
    if not os.path.exists(excel_file):
        messagebox.showerror("Erreur", f"Le fichier {excel_file} est introuvable.")
        return

    # 📌 Chargement sécurisé des données
    try:
        df = pd.read_excel(excel_file, sheet_name="Virement")
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de la lecture du fichier Excel : {e}")
        return

    # 📌 Vérifier les colonnes requises
    required_columns = ["Beneficiaire", "IBAN", "Date", "Reference", "Montant"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        messagebox.showerror("Erreur", f"Colonnes manquantes dans la feuille Virement : {', '.join(missing_columns)}")
        return

    # 📌 Convertir les dates en datetime sans filtrage spécifique à mars 2025
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    
    # 📌 Utilisation de tout le DataFrame sans filtre de date
    df_filtered = df
    
    # 📌 Vérifier si des virements sont présents
    if df_filtered.empty:
        messagebox.showwarning("Aucun virement", "Aucun virement trouvé dans le fichier.")
        return
 
    # 📌 Créer la fenêtre de sélection des virements
    virement_window = tk.Toplevel()
    virement_window.title("Sélection des virements")
    virement_window.geometry("800x500")

    tk.Label(virement_window, text="Sélectionnez les virements à traiter", font=("Arial", 12, "bold")).pack(pady=10)

    # 📌 Création de la Treeview
    columns = ("Bénéficiaire", "IBAN", "Date", "Référence", "Montant")
    tree = ttk.Treeview(virement_window, columns=columns, show="headings", selectmode="extended")

    # 📌 Définition des en-têtes
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=150)

    # 📌 Ajout des données
    for _, row in df_filtered.iterrows():
        date_str = row["Date"].strftime("%Y-%m-%d") if pd.notna(row["Date"]) else "Date inconnue"
        tree.insert("", "end", values=(row["Beneficiaire"], row["IBAN"], date_str, row["Reference"], f"{row['Montant']:.2f}"))

    tree.pack(padx=10, pady=10, fill="both", expand=True)

    def generate_selected_virements():
        """Génère un fichier XML et exécute le virement avec les données sélectionnées."""
        selected_items = tree.selection()
        if not selected_items:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner au moins un virement.")
            return

        selected_virements = []
        execution_date = None
        
        for item in selected_items:
            values = tree.item(item, "values")
            # Récupérer la date d'exécution du premier virement sélectionné
            if not execution_date and values[2] != "Date inconnue":
                try:
                    # Convertir au format YYYY-MM-DD
                    execution_date = values[2]
                except:
                    pass
                    
            selected_virements.append({
                "beneficiaire": values[0],
                "iban": values[1],
                "objet": f"{values[2]} - {values[3]}",  # Date + Référence
                "montant": float(values[4])
            })

        # 📌 Générer le fichier XML avec la date d'exécution spécifiée
        fichier_xml = generer_xml_virements(selected_virements, date_execution=execution_date)
        if fichier_xml:
            # 🔥 Envoi du fichier XML pour exécution du virement
            try:
                import generer_virement
                generer_virement.envoyer_virement_vers_lcl(fichier_xml)
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'envoi du virement : {e}")

            virement_window.destroy()
        else:
            messagebox.showerror("Erreur", "Échec de la génération du fichier XML.")

    # 📌 Bouton de validation
    tk.Button(virement_window, text="Générer XML et Réaliser Virement", command=generate_selected_virements, bg="#4caf50", fg="black", font=("Arial", 10, "bold")).pack(pady=10)
    
def planning():
    print ("Fonction en implémentation")

def clear_right_frame():
    """Supprime tous les widgets du cadre droit avant d'en ajouter de nouveaux."""
    for widget in right_frame.winfo_children():
        widget.destroy()    



def main_menu():
    """Initialise l'interface et affiche le menu principal."""
    global root, left_frame, right_frame  # Définir ces variables globalement

    root = tk.Tk()
    root.title("Gestion des contrats")
    root.geometry("1500x1000")  # Définit une taille initiale
    root.configure(bg="#f2f7ff")

    # Cadre de gauche : Menu
    left_frame = Frame(root, bg="lightgray", width=200, height=500)
    left_frame.pack(side="left", fill="y")

    # Cadre de droite : Zone d'affichage dynamique
    right_frame = Frame(root, width=600, height=500)
    right_frame.pack(side="right", fill="both", expand=True)
    right_frame.pack_propagate(False)

    show_main_menu()  # Affiche le menu principal

    root.mainloop()


def show_main_menu():
    """Affiche le menu principal dans le cadre de gauche et affiche l’image dans le cadre de droite."""
    # Effacer le contenu du menu de gauche
    for widget in left_frame.winfo_children():
        widget.destroy()

    # Création des boutons du menu
    Label(left_frame, text="Menu principal", font=("Arial", 16, "bold"), bg="#4a90e2", fg="white").pack(fill="x", pady=10)
    Button(left_frame, text="📋 Nouveau contrat MAR", command=open_contract_creation_mar).pack(fill="x", pady=10)
    Button(left_frame, text="🩺 Nouveau contrat IADE", command=open_contract_creation_iade).pack(fill="x", pady=10)  # Ajout de l'emoji 🩺
    Button(left_frame, text="📊 Comptabilité", command=open_accounting_menu).pack(fill="x", pady=10)
    Button(left_frame, text="📅 Plannings opératoires", command=planning).pack(fill="x", pady=10)
    Button(left_frame, text="⚙️ Paramètres", command=open_parameters).pack(fill="x", pady=10)
    Button(left_frame, text="🚪 Quitter", command=root.destroy).pack(fill="x", pady=10)

    # Affichage de l'image dans le cadre droit
    show_welcome_image()



def show_welcome_image():
    """Affiche une image d'accueil avec un voile blanc semi-transparent."""
    global right_frame  # Accéder à right_frame défini dans main_menu()

    # Effacer le contenu existant
    for widget in right_frame.winfo_children():
        widget.destroy()

    image_path = "/Users/vincentperreard/script contrats/clinique_mathilde.png"

    if not os.path.exists(image_path):
        print(f"Image non trouvée: {image_path}")
        return
    
    try:
        original_image = Image.open(image_path).convert("RGBA")

        # Vérifier la taille du cadre
        right_frame.update_idletasks()
        frame_width = right_frame.winfo_width()
        frame_height = right_frame.winfo_height()

        if frame_width <= 1 or frame_height <= 1:
            print("Le cadre right_frame n'a pas encore de dimensions valides.")
            return

        # Redimensionner l'image tout en conservant les proportions
        resized_image = original_image.resize((frame_width, frame_height), Image.Resampling.LANCZOS)

        # Créer un calque blanc semi-transparent (30% d'opacité)
        overlay = Image.new("RGBA", resized_image.size, (255, 255, 255, 100))  # 100 = transparence légère
        blended_image = Image.alpha_composite(resized_image, overlay)

        # Convertir pour affichage Tkinter
        image_tk = ImageTk.PhotoImage(blended_image)

        # Utilisation d'un Canvas pour placer l'image et futurs éléments
        canvas = Canvas(right_frame, width=frame_width, height=frame_height, bg="white", highlightthickness=0)
        canvas.pack(fill="both", expand=True)

        # Affichage de l'image avec le voile blanc
        canvas.create_image(frame_width // 2, frame_height // 2, image=image_tk, anchor="center")
        canvas.image = image_tk  # Conserver la référence

        # Ajout d'un titre exemple (modifiable)
        canvas.create_text(frame_width // 2, frame_height // 4, text="Bienvenue", font=("Arial", 24, "bold"), fill="black")

    except Exception as e:
        print(f"Erreur d'affichage de l'image : {e}")

if __name__ == "__main__":
    main_menu()