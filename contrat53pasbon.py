# Nouvelles fonctions pour l'interface modernisée

def clear_right_frame():
    """
    Nettoie le panneau droit avant d'y ajouter de nouveaux éléments.
    Cette fonction est appelée par d'autres fonctions pour préparer l'espace d'affichage.
    """
    # Cette fonction est maintenant implémentée dans main.py
    pass

def convert_date_format(date_str, from_format="yyyy-mm-dd", to_format="dd-mm-yyyy"):
    """Convertit une date d'un format à un autre."""
    if not date_str:
        return date_str
        
    try:
        if from_format == "yyyy-mm-dd" and to_format == "dd-mm-yyyy":
            year, month, day = date_str.split('-')
            return f"{day}-{month}-{year}"
        elif from_format == "dd-mm-yyyy" and to_format == "yyyy-mm-dd":
            day, month, year = date_str.split('-')
            return f"{year}-{month}-{day}"
        return date_str
    except:
        return date_str  # En cas d'erreur, renvoyer la date d'origine

def ensure_correct_date_format(date_str, required_format="yyyy-mm-dd"):
    """
    S'assure que la date est dans le format requis pour traitement par la base de données.
    Convertit une date du format DD-MM-YYYY vers YYYY-MM-DD si nécessaire.
    """
    if not date_str:
        return date_str
        
    # Si la date est déjà au format YYYY-MM-DD
    if len(date_str.split('-')) == 3 and date_str.split('-')[0].isdigit() and len(date_str.split('-')[0]) == 4:
        return date_str
        
    # Convertir de DD-MM-YYYY vers YYYY-MM-DD
    try:
        day, month, year = date_str.split('-')
        if len(year) == 4 and len(day) <= 2 and len(month) <= 2:  # Vérifier si c'est bien DD-MM-YYYY
            return f"{year}-{month}-{day}"
    except:
        pass  # Si la séparation échoue, on continue
        
    # Sinon, renvoyer la chaîne d'origine
    return date_str

def create_contract_interface_mar(parent_frame):
    """
    Crée l'interface de création de contrat MAR dans le frame parent fourni.
    
    Args:
        parent_frame: Le frame parent dans lequel créer l'interface
    """
    from ui_styles import create_frame, create_label, create_button, COLORS
    from tkinter import StringVar, Entry, OptionMenu
    from tkcalendar import Calendar
    from datetime import datetime
    import pandas as pd
    import os
    import subprocess
    
    try:
        # Chargement des données
        mars_selarl = pd.read_excel(file_paths["excel_mar"], sheet_name="MARS SELARL")
        mars_rempla = pd.read_excel(file_paths["excel_mar"], sheet_name="MARS Remplaçants")
        
        # Concaténation PRENOM + NOM en une seule colonne normalisée
        mars_selarl["FULL_NAME"] = mars_selarl["PRENOM"].fillna("").str.strip() + " " + mars_selarl["NOM"].fillna("").str.strip()
        mars_rempla["FULL_NAME"] = mars_rempla["PRENOMR"].fillna("").str.strip() + " " + mars_rempla["NOMR"].fillna("").str.strip()
        
        # Conteneur principal - utilisation d'un PanedWindow pour diviser l'espace
        main_container = tk.PanedWindow(parent_frame, orient=tk.HORIZONTAL)
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

            def on_date_select():
                """Capture la date sélectionnée et met à jour les variables."""
                selected_date = calendar.get_date()  # Format YYYY-MM-DD
                selected_dates.append(selected_date)
                
                if len(selected_dates) == 1:
                    # Mise à jour de la variable uniquement (pas directement le widget)
                    start_date_var.set(selected_date)
                    message_var.set("Sélectionnez la date de fin.")
                elif len(selected_dates) == 2:
                    # Trier les dates si nécessaire
                    start, end = sorted(selected_dates)
                    start_date_var.set(start)
                    end_date_var.set(end)
                    date_picker.destroy()

            def close_calendar():
                """Ferme le calendrier en ne retenant qu'une seule date si nécessaire."""
                if len(selected_dates) == 1:  # Si une seule date est sélectionnée
                    # Une seule date = même début et fin
                    start_date_var.set(selected_dates[0])
                    end_date_var.set(selected_dates[0])
                date_picker.destroy()

            # Crée la fenêtre de calendrier
            date_picker = tk.Toplevel(parent_frame.winfo_toplevel())
            date_picker.title("Sélectionner les dates")
            date_picker.geometry("400x400")  # Fenêtre plus grande
            date_picker.configure(bg="#f0f4f7")  # Fond assorti

            message_var = tk.StringVar()
            message_var.set("Sélectionnez la date de début.")
            
            # Titre principal stylisé
            tk.Label(date_picker, 
                    text="Sélection des dates", 
                    font=("Arial", 14, "bold"),
                    bg="#4a90e2", 
                    fg="white",
                    padx=10,
                    pady=5).pack(fill="x", pady=(0, 10))
                    
            # Message d'instruction
            instruction_label = tk.Label(date_picker, 
                                    textvariable=message_var, 
                                    font=("Arial", 11),
                                    bg="#f0f4f7",
                                    pady=5)
            instruction_label.pack(pady=5)

            # Calendrier amélioré
            calendar = Calendar(
                date_picker,
                date_pattern="dd-mm-yyyy",  # Format jj-mm-aaaa
                background="#f0f4f7",       # Fond assorti
                foreground="#333333",       # Texte foncé
                selectbackground="#4a90e2", # Bleu pour sélection
                selectforeground="white",   # Texte blanc sur sélection
                normalbackground="white",   # Fond blanc jours normaux
                weekendbackground="#f5f5f5",# Gris clair weekends
                weekendforeground="#e74c3c",# Rouge pour weekends
                othermonthforeground="#aaaaaa", # Gris pour jours autres mois
                font=("Arial", 10),        # Police lisible
                bordercolor="#4a90e2",     # Bordure bleue
                headersbackground="#4a90e2", # En-tête bleu
                headersforeground="white",   # Texte blanc pour en-tête
            )
            calendar.pack(pady=10, padx=20, fill="both", expand=True)

            # Cadre pour les boutons avec style
            button_frame = tk.Frame(date_picker, bg="#f0f4f7")
            button_frame.pack(pady=10, fill="x")

            # Boutons stylisés
            select_button = tk.Button(
                button_frame, 
                text="Valider", 
                command=on_date_select,
                bg="#4a90e2", 
                fg="white",
                font=("Arial", 10, "bold"),
                width=10,
                padx=5,
                pady=3,
                relief="flat",
                cursor="hand2"
            )
            select_button.pack(side="left", padx=20)

            close_button = tk.Button(
                button_frame, 
                text="Fermer", 
                command=close_calendar,
                bg="#f44336", 
                fg="white",
                font=("Arial", 10, "bold"),
                width=10,
                padx=5,
                pady=3,
                relief="flat",
                cursor="hand2"
            )
            close_button.pack(side="right", padx=20)

            # Centrer la fenêtre
            date_picker.update_idletasks()
            width = date_picker.winfo_width()
            height = date_picker.winfo_height()
            x = (date_picker.winfo_screenwidth() // 2) - (width // 2)
            y = (date_picker.winfo_screenheight() // 2) - (height // 2)
            date_picker.geometry(f'{width}x{height}+{x}+{y}')

            # Définir cette fenêtre comme modale
            date_picker.transient(parent_frame.winfo_toplevel())
            date_picker.grab_set()
            date_picker.focus_set()
        
        def save_contract():
            """Sauvegarde les informations du contrat et génère le PDF."""
            print("file_paths keys:", file_paths.keys())
            
            replaced_name = replaced_var.get().strip()
            replacing_name = replacing_var.get().strip()
            start_date = ensure_correct_date_format(start_date_var.get())
            end_date = ensure_correct_date_format(end_date_var.get())
            sign_date = ensure_correct_date_format(sign_date_var.get())
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
            show_post_contract_actions(
                actions_container, 
                pdf_path, 
                replaced_name, 
                replacing_name, 
                replaced_email, 
                replacing_email, 
                start_date, 
                end_date, 
                "MAR"
            )
        
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
        tk.Button(form_frame, text="📅 Sélectionner les dates", command=select_dates).grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
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
        tk.Button(form_frame, text="Créer le contrat", command=save_contract, 
               font=("Arial", 12, "bold"), bg="#007ACC", fg="black").grid(row=7, column=0, pady=10, padx=5, sticky="w")
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        tk.Label(parent_frame, text=f"Erreur lors de la création de l'interface : {str(e)}", 
                 font=("Arial", 12), fg="red").pack(pady=20)

def create_contract_interface_iade(parent_frame):
    """
    Crée l'interface de création de contrat IADE dans le frame parent fourni.
    
    Args:
        parent_frame: Le frame parent dans lequel créer l'interface
    """
    from ui_styles import create_frame, create_label, create_button, COLORS
    from tkinter import StringVar, Entry, OptionMenu
    from tkcalendar import Calendar
    from datetime import datetime
    import pandas as pd
    import os
    import subprocess
    
    try:
        # Chargement des données
        iade_data = pd.read_excel(file_paths["excel_iade"], sheet_name="Coordonnées IADEs")
        
        # Conteneur principal - utilisation d'un PanedWindow pour diviser l'espace
        main_container = tk.PanedWindow(parent_frame, orient=tk.HORIZONTAL)
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

            def update_message():
                """Mise à jour du message selon le nombre de dates sélectionnées."""
                if len(selected_dates) == 0:
                    message_var.set("Sélectionnez la date de début.")
                elif len(selected_dates) == 1:
                    message_var.set("Sélectionnez la date de fin.")

            def on_date_select():
                """Capture la date sélectionnée."""
                selected_date = calendar.get_date()  # Format YYYY-MM-DD
                
                # Convertir au format DD-MM-YYYY pour l'affichage
                display_date = convert_date_format(selected_date, "yyyy-mm-dd", "dd-mm-yyyy")
                
                selected_dates.append(selected_date)
                update_message()

                if len(selected_dates) == 2:  # Deux dates sélectionnées
                    # Trier les dates si nécessaire
                    start_date, end_date = sorted(selected_dates)
                    
                    # Stocker les dates originales (YYYY-MM-DD) dans les variables pour traitement
                    start_date_var.set(start_date)
                    end_date_var.set(end_date)
                    
                    # Convertir pour affichage
                    display_start = convert_date_format(start_date, "yyyy-mm-dd", "dd-mm-yyyy")
                    display_end = convert_date_format(end_date, "yyyy-mm-dd", "dd-mm-yyyy")
                    
                    # Mettre à jour les champs d'affichage
                    start_entry.config(state="normal")
                    start_entry.delete(0, tk.END)
                    start_entry.insert(0, display_start)
                    start_entry.config(state="readonly")
                    
                    end_entry.config(state="normal")
                    end_entry.delete(0, tk.END)
                    end_entry.insert(0, display_end)
                    end_entry.config(state="readonly")
                    
                    date_picker.destroy()
                elif len(selected_dates) == 1:  # Une seule date sélectionnée
                    # Stocker la date originale (YYYY-MM-DD) dans la variable pour traitement
                    start_date_var.set(selected_date)
                    
                    # Convertir pour affichage
                    display_start = convert_date_format(selected_date, "yyyy-mm-dd", "dd-mm-yyyy")
                    
                    # Mettre à jour le champ d'affichage
                    start_entry.config(state="normal")
                    start_entry.delete(0, tk.END)
                    start_entry.insert(0, display_start)
                    start_entry.config(state="readonly")

            def close_calendar():
                """Ferme le calendrier en ne retenant qu'une seule date si nécessaire."""
                if len(selected_dates) == 1:  # Si une seule date est sélectionnée
                    # Une seule date = même début et fin
                    start_date_var.set(selected_dates[0])
                    end_date_var.set(selected_dates[0])
                    
                    # Convertir pour affichage
                    display_date = convert_date_format(selected_dates[0], "yyyy-mm-dd", "dd-mm-yyyy")
                    
                    # Mettre à jour les champs d'affichage
                    start_entry.config(state="normal")
                    start_entry.delete(0, tk.END)
                    start_entry.insert(0, display_date)
                    start_entry.config(state="readonly")
                    
                    end_entry.config(state="normal")
                    end_entry.delete(0, tk.END)
                    end_entry.insert(0, display_date)
                    end_entry.config(state="readonly")
                    
                date_picker.destroy()

            # Crée la fenêtre de calendrier
            date_picker = tk.Toplevel(parent_frame.winfo_toplevel())
            date_picker.title("Sélectionner les dates")
            date_picker.geometry("400x400")  # Fenêtre plus grande
            date_picker.configure(bg="#f0f4f7")  # Fond assorti

            message_var = tk.StringVar()
            message_var.set("Sélectionnez la date de début.")
            
            # Titre principal stylisé
            tk.Label(date_picker, 
                    text="Sélection des dates", 
                    font=("Arial", 14, "bold"),
                    bg="#4a90e2", 
                    fg="white",
                    padx=10,
                    pady=5).pack(fill="x", pady=(0, 10))
                    
            # Message d'instruction
            instruction_label = tk.Label(date_picker, 
                                    textvariable=message_var, 
                                    font=("Arial", 11),
                                    bg="#f0f4f7",
                                    pady=5)
            instruction_label.pack(pady=5)

            # Calendrier amélioré
            calendar = Calendar(
                date_picker,
                date_pattern="dd-mm-yyyy",  # Format jj-mm-aaaa
                background="#f0f4f7",       # Fond assorti
                foreground="#333333",       # Texte foncé
                selectbackground="#4a90e2", # Bleu pour sélection
                selectforeground="white",   # Texte blanc sur sélection
                normalbackground="white",   # Fond blanc jours normaux
                weekendbackground="#f5f5f5",# Gris clair weekends
                weekendforeground="#e74c3c",# Rouge pour weekends
                othermonthforeground="#aaaaaa", # Gris pour jours autres mois
                font=("Arial", 10),        # Police lisible
                bordercolor="#4a90e2",     # Bordure bleue
                headersbackground="#4a90e2", # En-tête bleu
                headersforeground="white",   # Texte blanc pour en-tête
            )
            calendar.pack(pady=10, padx=20, fill="both", expand=True)

            # Cadre pour les boutons avec style
            button_frame = tk.Frame(date_picker, bg="#f0f4f7")
            button_frame.pack(pady=10, fill="x")

            # Boutons stylisés
            select_button = tk.Button(
                button_frame, 
                text="Valider", 
                command=on_date_select,
                bg="#4a90e2", 
                fg="white",
                font=("Arial", 10, "bold"),
                width=10,
                padx=5,
                pady=3,
                relief="flat",
                cursor="hand2"
            )
            select_button.pack(side="left", padx=20)

            close_button = tk.Button(
                button_frame, 
                text="Fermer", 
                command=close_calendar,
                bg="#f44336", 
                fg="white",
                font=("Arial", 10, "bold"),
                width=10,
                padx=5,
                pady=3,
                relief="flat",
                cursor="hand2"
            )
            close_button.pack(side="right", padx=20)

            # Centrer la fenêtre
            date_picker.update_idletasks()
            width = date_picker.winfo_width()
            height = date_picker.winfo_height()
            x = (date_picker.winfo_screenwidth() // 2) - (width // 2)
            y = (date_picker.winfo_screenheight() // 2) - (height // 2)
            date_picker.geometry(f'{width}x{height}+{x}+{y}')

            # Définir cette fenêtre comme modale
            date_picker.transient(parent_frame.winfo_toplevel())
            date_picker.grab_set()
            date_picker.focus_set()
        
        # Fonction pour enregistrer le contrat
        def save_contract_iade():
            """Génère un fichier Excel et un contrat IADE, puis affiche les options post-contrat."""
            # Récupération des valeurs depuis l'interface
            replacing_name = replacing_var.get()
            start_date = ensure_correct_date_format(start_date_var.get())
            end_date = ensure_correct_date_format(end_date_var.get())
            sign_date = ensure_correct_date_format(sign_date_var.get())
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
                replacing_full_name = f"{replacing_data['PRENOMR']} {replacing_data['
