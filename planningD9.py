import gspread
import datetime as dt
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from analyse_excel import analyse_depuis_script
from gestion_mars import gestion_mars_gui
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import time
import os
import glob


# 📌 Connexion à Google Sheets
def get_chirurgiens_presences(sheet_name, worksheet_name):
    """
    Récupère les présences des chirurgiens depuis Google Sheets.
    """
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("/Users/vincentperreard/script contrats/credentials.json", scope)    
    client = gspread.authorize(creds)

    sheet = client.open(sheet_name)
    worksheet = sheet.worksheet(worksheet_name)
    
    data = worksheet.get_all_values()
    df = pd.DataFrame(data)

    # Extraction des noms des chirurgiens (ligne 4, une colonne sur deux à partir de E)
    colonnes_chirurgiens = list(range(4, len(df.columns), 2))
    noms_chirurgiens = {col: df.iloc[3, col] for col in colonnes_chirurgiens if pd.notna(df.iloc[3, col]) and df.iloc[3, col] not in ["BLOC", "---------"]}

    return df, noms_chirurgiens

# 📌 Détection des chirurgiens non répondants
def get_chirurgiens_non_repondus(df, noms_chirurgiens):
    """
    Identifie les chirurgiens qui n'ont pas renseigné leur présence.
    """
    non_repondants_par_chirurgien = {}  # Dictionnaire pour regrouper les réponses manquantes par chirurgien
    colonnes_chirurgiens = list(noms_chirurgiens.keys())

    for index, row in df.iterrows():  # Parcourir toutes les lignes du DataFrame
        # Ignorer les lignes sans date valide
        if pd.isna(row[2]):
            continue

        date = row[2]  # La date est déjà propagée dans le DataFrame filtré
        moment = row[3]  # Matin ou Après-midi

        for col in colonnes_chirurgiens:
            chirurgien = noms_chirurgiens.get(col, None)
            if chirurgien:
                bloc_status = str(row[col]).strip().upper() if pd.notna(row[col]) else ""
                presence_status = str(row[col + 1]).strip().upper() if pd.notna(row[col + 1]) else ""

                # Debug: Afficher les informations pour chaque chirurgien et chaque ligne
                print(f"Chirurgien: {chirurgien}, Date: {date}, Moment: {moment}")
                print(f"Bloc status: '{bloc_status}', Presence status: '{presence_status}'")

                if "BLOC" in bloc_status.upper() and presence_status.strip() in ["", "NAN", " "]:
                    # Ajouter la réponse manquante au dictionnaire
                    if chirurgien not in non_repondants_par_chirurgien:
                        non_repondants_par_chirurgien[chirurgien] = []
                    non_repondants_par_chirurgien[chirurgien].append((date, moment))

    return non_repondants_par_chirurgien



def demander_periode_via_gui():
    def on_date_click(event):
        nonlocal choix_date_debut, date_debut, date_fin
        selected_date = cal.selection_get()

        if choix_date_debut:
            # Réinitialiser la sélection précédente
            cal.calevent_remove('all')  # Supprimer tous les événements de mise en évidence
            date_debut = selected_date
            choix_date_debut = False
            label_periode.config(
                text=f"Sélectionnez la date de fin.",
                font=('Helvetica', 14, 'bold'),
                foreground='darkblue'
            )
        else:
            date_fin = selected_date
            choix_date_debut = True

            # Mettre en surbrillance les dates intermédiaires
            current_date = date_debut
            while current_date <= date_fin:
                if current_date == date_debut or current_date == date_fin:
                    # Style pour les bornes (début et fin)
                    cal.calevent_create(current_date, 'Boundary', 'selected_boundaries')
                else:
                    # Style pour les dates intermédiaires
                    cal.calevent_create(current_date, 'Range', 'selected_range')
                current_date += dt.timedelta(days=1)

            # Afficher un résumé des dates sélectionnées
            label_periode.config(
                text=f"Période sélectionnée : {date_debut.strftime('%d/%m/%Y')} - {date_fin.strftime('%d/%m/%Y')}",
                font=('Helvetica', 14, 'bold'),
                foreground='darkblue'
            )

    def valider_periode():
        root.destroy()

    def recommencer_selection():
        nonlocal choix_date_debut, date_debut, date_fin
        # Réinitialiser les variables et l'affichage
        choix_date_debut = True
        date_debut = None
        date_fin = None
        cal.calevent_remove('all')  # Supprimer tous les événements de mise en évidence
        label_periode.config(
            text="Cliquez sur les dates pour choisir la période.",
            font=('Helvetica', 14),
            foreground='#2c3e50'
        )

    root = tk.Tk()
    root.title("Sélectionner une période")
    root.geometry("600x500")  # Taille de la fenêtre

    choix_date_debut = True
    date_debut = None
    date_fin = None

    # Style global
    style = ttk.Style()
    style.configure('TButton', font=('Helvetica', 12), padding=10)
    style.configure('TLabel', font=('Helvetica', 12))

    # Création du calendrier
    cal = Calendar(
        root,
        selectmode='day',
        date_pattern='dd/MM/yyyy',
        background='white',  # Fond du calendrier
        foreground='black',  # Texte par défaut en noir
        borderwidth=2,
        headersbackground='#2c3e50',  # Fond du mois et de l'année en haut (bleu foncé)
        normalbackground='white',  # Fond des jours normaux
        weekendbackground='#f0f0f0',  # Fond des week-ends (gris clair)
        headersforeground='white',  # Texte du mois et de l'année en blanc
        selectbackground='#3498db',  # Fond bleu clair pour la sélection
        selectforeground='white',  # Texte blanc pour les dates sélectionnées
        font=('Helvetica', 14),  # Police plus grande pour les jours
        weekendforeground='red',  # Texte des week-ends en rouge
        othermonthbackground='#f0f0f0',  # Fond des jours des autres mois (gris clair)
        othermonthforeground='gray',  # Texte des jours des autres mois en gris
        bordercolor='#2c3e50',  # Bordure du calendrier (bleu foncé)
    )

    cal.bind("<<CalendarSelected>>", on_date_click)
    cal.pack(padx=20, pady=20, fill='both', expand=True)

    # Ajout des styles pour les plages de dates sélectionnées
    cal.tag_config('selected_range', background='#3498db', foreground='white')  # Fond bleu clair pour la période
    cal.tag_config('selected_boundaries', background='#2c3e50', foreground='white', font=('Helvetica', 14, 'bold'))  # Début & fin en gras et plus foncé

    # Indication des dates sélectionnées
    label_periode = ttk.Label(
        root,
        text="Cliquez sur les dates pour choisir la période.",
        font=('Helvetica', 14),
        foreground='#2c3e50'
    )
    label_periode.pack(padx=10, pady=10)

    # Boutons
    button_frame = ttk.Frame(root)
    button_frame.pack(pady=10)

    valider_btn = ttk.Button(button_frame, text="Valider", command=valider_periode, style='TButton')
    valider_btn.pack(side='left', padx=10)

    recommencer_btn = ttk.Button(button_frame, text="Recommencer", command=recommencer_selection, style='TButton')
    recommencer_btn.pack(side='left', padx=10)

    root.mainloop()

    return date_debut, date_fin


def afficher_resultats_par_chirurgien(non_repondants_par_chirurgien):
    """
    Affiche la liste des chirurgiens non répondants regroupés par chirurgien.
    """
    result_window = tk.Tk()
    result_window.title("Chirurgiens non répondants")

    # Créer un widget Text pour afficher les résultats
    text_widget = tk.Text(result_window, wrap="word", height=20, width=80)
    text_widget.pack(expand=True, fill="both", padx=10, pady=10)

    # Parcourir le dictionnaire et afficher les résultats
    for chirurgien, reponses_manquantes in non_repondants_par_chirurgien.items():
        text_widget.insert("end", f"{chirurgien} :\n")
        for date, moment in reponses_manquantes:
            text_widget.insert("end", f"  {date} : {moment}\n")
        text_widget.insert("end", "\n")  # Ajouter une ligne vide entre les chirurgiens

    # Ajouter un bouton pour fermer la fenêtre
    ttk.Button(result_window, text="Fermer", command=result_window.destroy).pack(pady=10)

    result_window.mainloop()

def afficher_resultats_anesthesie(absents, presences):
    fenetre = tk.Tk()
    fenetre.title("Analyse IADE : Absents et Présents")
    fenetre.geometry("800x600")

    texte = tk.Text(fenetre, wrap="word", font=("Helvetica", 12))
    texte.pack(expand=True, fill="both", padx=10, pady=10)

    texte.insert("end", "📌 Absents par date :\n")
    for date, data in absents.items():
        mars = data.get("mars", [])
        iades = data.get("iade", [])
        texte.insert("end", f"  - {date} : {len(mars)} MARS ({', '.join(mars) if mars else '-'}) | {len(iades)} IADEs ({', '.join(iades) if iades else '-'})\n")

    texte.insert("end", "\n📌 Présences par date :\n")
    for date, data in presences.items():
        gf = data.get("GF", [])
        g  = data.get("G", [])
        p  = data.get("P", [])
        journee_complete = gf + g
        texte.insert("end", f"  - {date} : Journée complète = {len(journee_complete)} ({', '.join(journee_complete) if journee_complete else '-'}) | Petite journée = {len(p)} ({', '.join(p) if p else '-'})\n")

    # Bouton de fermeture
    fermer_btn = ttk.Button(fenetre, text="Fermer", command=fenetre.destroy)
    fermer_btn.pack(pady=10)

    fenetre.mainloop()

def filtrer_par_periode(df, date_debut, date_fin):
    # Convertir date_debut et date_fin en datetime
    date_debut = pd.Timestamp(date_debut)
    date_fin = pd.Timestamp(date_fin)
    
    # Convertir la colonne des dates en datetime
    df[2] = pd.to_datetime(df[2], errors='coerce', dayfirst=True)
    print("Dates après conversion :")
    print(df[2].head(20))  # Afficher les 20 premières dates pour vérifier
    
    # Remplir les dates manquantes
    df[2] = df[2].fillna(method='ffill')
    print("Dates après propagation :")
    print(df[2].head(20))  # Afficher les 20 premières dates pour vérifier
    
    print("Date de début :", date_debut)
    print("Date de fin :", date_fin)
    
    # Filtrer les données selon la période sélectionnée
    df_filtre = df[(df[2] >= date_debut) & (df[2] <= date_fin)]
    print("Dates après filtrage :")
    print(df_filtre[[2, 3]].head(20))  # Afficher les dates et moments après filtrage
    
    return df_filtre

def lancer_programme_anesthesie():
    """
    Connexion à Momentum, sélection de la période et téléchargement du fichier Excel des anesthésistes.
    """
    # 📌 Demander la période comme pour les chirurgiens
    date_debut, date_fin = demander_periode_via_gui()

    # 📌 Définition du dossier Téléchargements de l'utilisateur
    dossier_telechargements = os.path.expanduser("~/Downloads")

    # 📌 Configuration de Selenium
    chrome_options = webdriver.ChromeOptions()

    # 📌 Activer le mode headless (exécution en arrière-plan)
    #chrome_options.add_argument("--headless=new")  # Mode headless moderne
    chrome_options.add_argument("--disable-gpu")  # Nécessaire sous Windows
    chrome_options.add_argument("--window-size=1920x1080")  # Simule une grande fenêtre
    chrome_options.add_argument("--no-sandbox")  # Évite certains blocages
    chrome_options.add_argument("--disable-dev-shm-usage")  # Empêche les erreurs sous Linux

    # 📌 Gérer le téléchargement sans interaction utilisateur
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": dossier_telechargements,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })

    # 📌 Lancer Selenium en mode invisible
    driver = webdriver.Chrome(options=chrome_options)




    driver.get("https://planningopmathilde.biosked.com/Momentum/Forms/frmHome.aspx")

    # 📌 Attendre que le champ utilisateur soit visible et interagissable
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "rtxtUserID"))).send_keys("vperreard")
    driver.find_element(By.ID, "rtxtPassword").send_keys("#Pervipeobe1")
    driver.find_element(By.ID, "cmdLogin_input").click()

    # 📌 Attendre la page d'accueil de Momentum
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ctl00_ctl00_mn_bd_rlvQuickLinks_ctrl1_btnImageLink")))

    # 📌 Aller dans "Planning Admin"
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_mn_bd_rlvQuickLinks_ctrl1_btnImageLink"))).click()

    # 📌 Passer en mode lecture seule immédiatement après l'ouverture de Planning Admin
    try:
        bouton_lecture_seule = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_mn_bd_cmdGoToReadOnlyView")))
        bouton_lecture_seule.click()
        time.sleep(1)
    except:
        print("✅ Déjà en mode lecture seule.")

    # 📌 Ensuite seulement : Sélectionner le menu déroulant de filtres
    menu_filtres = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_mn_bd_uclSavedFilter_rcbFilters_Input")))
    menu_filtres.click()

    # 📌 Puis choisir le bon filtre
    option_filtres = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, "//li[contains(text(),'VP ABSENCES MARS IADES G GF P')]")))
    option_filtres.click()

    # 📌 Attendre que les champs de date soient disponibles
    champ_date_debut = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_mn_bd_ucSchedControl_urDateRange_rdpStartDate_dateInput")))
    champ_date_debut.clear()
    champ_date_debut.send_keys(date_debut.strftime("%d/%m/%Y"))

    champ_date_fin = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_mn_bd_ucSchedControl_urDateRange_rdpEndDate_dateInput")))
    champ_date_fin.clear()
    champ_date_fin.send_keys(date_fin.strftime("%d/%m/%Y"))

    # 📌 Cliquer sur "Actualiser" immédiatement après que le bouton soit prêt
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_mn_bd_cmdRefresh"))).click()

    # 📌 Attendre l’apparition du bouton d’exportation avant de cliquer
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_mn_bd_cmdExcel"))).click()
    time.sleep(1)  # Attendre que le téléchargement se termine

    driver.quit()

    # 📌 Attendre que le fichier apparaisse dans Téléchargements
    timeout = 30  # Maximum 30 secondes d'attente
    fichier_excel = None

    while timeout > 0:
        # 📌 Lister tous les fichiers téléchargés récemment
        fichiers_xls = [f for f in os.listdir(dossier_telechargements) if f.endswith(".xls")]

        # 📌 Sélectionner uniquement le dernier fichier .xls
        if fichiers_xls:
            fichier_excel = max(
                fichiers_xls, 
                key=lambda f: os.path.getmtime(os.path.join(dossier_telechargements, f))
            )
            chemin_complet = os.path.join(dossier_telechargements, fichier_excel)

            # 📌 Renommer le fichier
            nouveau_nom = "VP_ABSENCES_MOMENTUM.xls"
            nouveau_chemin = os.path.join(dossier_telechargements, nouveau_nom)

            try:
                os.rename(chemin_complet, nouveau_chemin)
                print(f"📥 Fichier renommé : {nouveau_chemin}")
            except Exception as e:
                print(f"⚠️ Erreur lors du renommage : {e}")
                nouveau_chemin = chemin_complet  # fallback

            break
        
        time.sleep(1)
        timeout -= 1

    if os.path.exists(nouveau_chemin):
        print("🔍 Lancement de l'analyse des absents et présences IADE...")

        absents, presences = analyse_depuis_script(nouveau_chemin)
        afficher_resultats_anesthesie(absents, presences)

        print("\n📌 Absents par date :")
        for date, data in absents.items():
            mars = data.get("mars", [])
            iades = data.get("iade", [])
            print(f"  - {date} : {len(mars)} MARS ({', '.join(mars) if mars else '-'}) | {len(iades)} IADEs ({', '.join(iades) if iades else '-'})")

        print("\n📌 Présences par date :")
        for date, data in presences.items():
            gf = data.get("GF", [])
            g  = data.get("G", [])
            p  = data.get("P", [])
            journee_complete = gf + g
            print(f"  - {date} : Journée complète = {len(journee_complete)} ({', '.join(journee_complete) if journee_complete else '-'}) | Petite journée = {len(p)} ({', '.join(p) if p else '-'})")

        # 🔁 Demande unique après avoir tout affiché
        fenetre_confirmation = tk.Tk()
        fenetre_confirmation.withdraw()
        reponse = tk.messagebox.askyesno("Intégration Google Sheets", "Souhaitez-vous intégrer ces données dans Google Sheets ?")
        fenetre_confirmation.destroy()

        if reponse:
            integrer_donnees_google_sheets(absents, presences, date_debut, date_fin)
        else:
            print("❌ Données non intégrées à Google Sheets.")

        return  # Ajout explicite du retour pour éviter toute récursion de menu
    
    else:
        print("⚠️ Le fichier renommé n'a pas été trouvé.")
        return  # Évite la boucle infinie


    def afficher_fichier_excel(fichier_excel):
        df = pd.read_excel(fichier_excel, engine='xlrd')  # Supporte les fichiers .xls
        print(df.head())  # Affiche un aperçu du fichier dans la console
        

        
        

# 📌 Interface principale du programme
def menu_principal():
    import webbrowser

    def ouvrir_google_sheet(url):
        chrome_path = 'open -a "Google Chrome" %s'
        webbrowser.get(chrome_path).open(url)

    def lancer_chirurgiens():
        root.destroy()
        lancer_programme_chirurgiens()

    def lancer_anesthesie():
        root.destroy()
        lancer_programme_anesthesie()

    def quitter_programme():
        root.destroy()

    # ⚙️ Paramètres
    def ouvrir_parametres():
        fenetre_params = tk.Toplevel(root)
        fenetre_params.title("⚙️ Paramètres")
        fenetre_params.geometry("400x350")
        fenetre_params.configure(bg="#f0f4f7")

        label = ttk.Label(
            fenetre_params,
            text="🛠️ Menu des paramètres",
            font=('Helvetica', 16, 'bold'),
            background="#f0f4f7",
            foreground="#2c3e50"
        )
        label.pack(pady=20)

        # 👨‍⚕️ Gestion opérateurs
        btn_operateurs = ttk.Button(fenetre_params, text="👨‍⚕️ Gestion opérateurs", command=gestion_operateurs_gui)
        btn_operateurs.pack(pady=10, fill='x', padx=40)

        # 🩺 Gestion MARS
        btn_mars = ttk.Button(fenetre_params, text="🩺 Gestion MARS", command=gestion_mars_gui)
        btn_mars.pack(pady=10, fill='x', padx=40)

        # Fermer
        btn_fermer = ttk.Button(fenetre_params, text="❌ Fermer", command=fenetre_params.destroy)
        btn_fermer.pack(pady=20, fill='x', padx=40)

    root = tk.Tk()
    root.title("🏥 Menu Principal")
    root.geometry("500x500")
    root.configure(bg="#f0f4f7")

    style = ttk.Style()
    style.configure("TButton", font=("Helvetica", 13), padding=10)

    label_titre = ttk.Label(
        root,
        text="📋 Sélectionnez une action",
        font=('Helvetica', 18, 'bold'),
        background="#f0f4f7",
        foreground="#2c3e50"
    )
    label_titre.pack(pady=20)

    # 🩺 Chirurgiens
    bouton_chirurgiens = ttk.Button(root, text="🩺 Programme Chirurgiens", command=lancer_chirurgiens)
    bouton_chirurgiens.pack(pady=10, fill='x', padx=50)

    # 💉 Anesthésie
    bouton_anesthesie = ttk.Button(root, text="💉 Programme Anesthésie", command=lancer_anesthesie)
    bouton_anesthesie.pack(pady=10, fill='x', padx=50)

    # 📊 Accès Google Sheets Opérateurs
    bouton_gsheet_operateurs = ttk.Button(
        root,
        text="📊 Accès Google Sheets Opérateurs",
        command=lambda: ouvrir_google_sheet("https://docs.google.com/spreadsheets/d/1JyGxKKlvVtAGY57wM7JDWrclML-E2WUx1yLrMS49u_Q/edit?usp=sharing")
    )
    bouton_gsheet_operateurs.pack(pady=10, fill='x', padx=50)

    # 📈 Accès Google Sheets Synthèse
    bouton_gsheet_synthese = ttk.Button(
        root,
        text="📈 Accès Google Sheets Synthèse",
        command=lambda: ouvrir_google_sheet("https://docs.google.com/spreadsheets/d/1I_rbDbbVPy414u8-qlEbQj8H1vL8jgCOUNQyVVCo8fQ/edit?usp=sharing")
    )
    bouton_gsheet_synthese.pack(pady=10, fill='x', padx=50)

    bouton_parametres = ttk.Button(root, text="⚙️ Paramètres", command=ouvrir_parametres)
    bouton_parametres.pack(pady=10, fill='x', padx=50)

    # ❌ Quitter
    bouton_quitter = ttk.Button(root, text="❌ Quitter", command=quitter_programme)
    bouton_quitter.pack(pady=20, fill='x', padx=50)

    root.mainloop()
    
    
    
    
# 📌 Lancer le programme Chirurgiens
def lancer_programme_chirurgiens():
    sheet_name = "PLANNING AUTO"
    worksheet_name = "planning"

    df_chirurgiens, noms_chirurgiens = get_chirurgiens_presences(sheet_name, worksheet_name)
    date_debut, date_fin = demander_periode_via_gui()
    df_filtre = filtrer_par_periode(df_chirurgiens, date_debut, date_fin)
    
    # 📌 Récupération des chirurgiens qui n'ont pas répondu dans cette période
    chirurgiens_non_repondus = get_chirurgiens_non_repondus(df_filtre, noms_chirurgiens)
    
    # Afficher les résultats regroupés par chirurgien
    afficher_resultats_par_chirurgien(chirurgiens_non_repondus)
    
    print("\n📊 Aperçu des données après filtrage :")
    print(df_filtre.head(10))

    menu_principal()
    
def integrer_donnees_google_sheets(absents, presences, date_debut, date_fin):

    print("📤 Connexion à Google Sheets...")

    # Connexion à Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("/Users/vincentperreard/script contrats/credentials.json", scope)
    client = gspread.authorize(creds)

    sheet = client.open("PLANNING AUTO")
    worksheet = sheet.worksheet("planning")

    all_data = worksheet.get_all_values()
    maj_cellules = []  # Regroupe les mises à jour

    for idx, row in enumerate(all_data):
        if idx < 5 or len(row) < 166:
            continue  # ignorer les entêtes et lignes incomplètes

        try:
            cell_date = row[2]
            if not cell_date:
                continue

            date_cell = pd.to_datetime(cell_date, dayfirst=True).date()

            if date_cell < date_debut or date_cell > date_fin:
                continue

            ligne_matin = idx if idx % 2 == 0 else idx - 1
            ligne_apm   = idx if idx % 2 != 0 else idx + 1

            date_str = date_cell.strftime("%d/%m/%Y")

            absents_du_jour = absents.get(date_str, {}).get("mars", [])
            nb_absents = len(absents_du_jour)
            print(f"{date_str} matin : ABS MARS = ({', '.join(absents_du_jour)}) - valeur intégrée {nb_absents}")
            print(f"{date_str} après-midi : ABS MARS = ({', '.join(absents_du_jour)}) - valeur intégrée {nb_absents}")

            # Ajouter les cellules pour absents MARS
            maj_cellules.append({'range': f'FJ{ligne_matin + 1}', 'values': [[nb_absents]]})
            maj_cellules.append({'range': f'FJ{ligne_apm + 1}', 'values': [[nb_absents]]})

            pres_jour = presences.get(date_str, {})
            print(f"[DEBUG] Données présence pour {date_str} : {pres_jour}")

            nb_matin = len(pres_jour.get("GF", [])) + len(pres_jour.get("G", [])) + len(pres_jour.get("P", []))
            nb_apm   = len(pres_jour.get("GF", [])) + len(pres_jour.get("G", []))

            print(f"{date_str} : {len(pres_jour.get('GF', []))} GF, {len(pres_jour.get('G', []))} G, {len(pres_jour.get('P', []))} P")
            print(f"{date_str} matin : valeur intégrée {nb_matin}")
            print(f"{date_str} après-midi : valeur intégrée {nb_apm}")

            # Ajouter les cellules pour présents IADE
            maj_cellules.append({'range': f'FI{ligne_matin + 1}', 'values': [[nb_matin]]})
            maj_cellules.append({'range': f'FI{ligne_apm + 1}', 'values': [[nb_apm]]})

            print(f"✅ Données préparées pour {date_str}")
            print(f"type(nb_matin) = {type(nb_matin)}, type(nb_apm) = {type(nb_apm)}")
        except Exception as e:
            print(f"⚠️ Erreur à la ligne {idx} : {e}")

    # 📤 Envoi en un seul batch
    try:
        worksheet.batch_update([{
            'range': maj['range'],
            'values': maj['values']
        } for maj in maj_cellules])
        print(f"📤 {len(maj_cellules)} cellules mises à jour avec batch_update.")
    except Exception as e:
        print(f"❌ Erreur lors de l'envoi batch_update : {e}")


    
def afficher_confirmation_integration():
    fenetre = tk.Tk()
    fenetre.title("Confirmation")
    fenetre.geometry("400x150")
    fenetre.resizable(False, False)

    label = ttk.Label(
        fenetre,
        text="✅ Les données ont été intégrées avec succès dans Google Sheets.",
        font=("Helvetica", 12),
        wraplength=360,
        justify="center"
    )
    label.pack(padx=20, pady=20)

    bouton_fermer = ttk.Button(fenetre, text="Fermer", command=fenetre.destroy)
    bouton_fermer.pack(pady=5)

    fenetre.mainloop()    

def recuperer_operateurs():
    sheet_name = "PLANNING AUTO"
    worksheet_name = "planning"

    df, noms_chirurgiens = get_chirurgiens_presences(sheet_name, worksheet_name)

    # Chargement du fichier d'e-mails si déjà existant
    chemin_fichier = "emails_operateurs.json"
    if os.path.exists(chemin_fichier):
        with open(chemin_fichier, "r") as f:
            emails = json.load(f)
    else:
        emails = {}

    return noms_chirurgiens, emails   
   
def gestion_operateurs_gui(depuis_cache=True):
    noms_chirurgiens, emails = recuperer_operateurs(depuis_cache)
    fenetre = tk.Toplevel()
    fenetre.title("👨‍⚕️ Gestion des opérateurs")
    fenetre.geometry("700x600")
    fenetre.configure(bg="#f0f4f7")

    def recharger_depuis_gsheets():
        fenetre.destroy()
        gestion_operateurs_gui(depuis_cache=False)

    label = ttk.Label(fenetre, text="✏️ Associez un ou plusieurs emails à chaque opérateur (séparés par des virgules) :", font=('Helvetica', 14, 'bold'))
    label.pack(pady=10)



    # ➕ Canvas avec scrollbar
    canvas = tk.Canvas(fenetre, borderwidth=0, background="#f0f4f7", highlightthickness=0)
    scrollbar = ttk.Scrollbar(fenetre, orient="vertical", command=canvas.yview)
    cadre_scrollable = ttk.Frame(canvas)

    cadre_scrollable.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=cadre_scrollable, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True, padx=10)
    scrollbar.pack(side="right", fill="y")

    champs = {}

    for idx, (col, nom) in enumerate(noms_chirurgiens.items()):
        ttk.Label(cadre_scrollable, text=nom, width=25).grid(row=idx, column=0, sticky="w", pady=3)
        champ = ttk.Entry(cadre_scrollable, width=50)
        champ.insert(0, emails.get(nom, ""))  # supporte déjà plusieurs emails séparés par ,
        champ.grid(row=idx, column=1, pady=3)
        champs[nom] = champ

    def sauvegarder():
        new_emails = {nom: champ.get().strip() for nom, champ in champs.items()}
        with open("emails_operateurs.json", "w") as f:
            json.dump(new_emails, f, indent=2)
        messagebox.showinfo("Sauvegarde", "✅ Emails enregistrés avec succès !")

    bouton_reload = ttk.Button(fenetre, text="🔄 Recharger depuis Google Sheets", command=recharger_depuis_gsheets)
    bouton_reload.pack(pady=5)

    bouton_sauvegarder = ttk.Button(fenetre, text="💾 Sauvegarder", command=sauvegarder)
    bouton_sauvegarder.pack(pady=20)

    bouton_fermer = ttk.Button(fenetre, text="❌ Fermer", command=fenetre.destroy)
    bouton_fermer.pack()
    
def recuperer_operateurs(depuis_cache=True):
    # Lire les e-mails déjà enregistrés
    chemin_emails = "emails_operateurs.json"
    if os.path.exists(chemin_emails):
        with open(chemin_emails, "r") as f:
            emails = json.load(f)
    else:
        emails = {}

    # Charger depuis cache si demandé
    chemin_noms = "noms_operateurs.json"
    if depuis_cache and os.path.exists(chemin_noms):
        with open(chemin_noms, "r") as f:
            noms_chirurgiens = json.load(f)
        return noms_chirurgiens, emails

    # Sinon : recharger depuis Google Sheets
    df, noms_chirurgiens = get_chirurgiens_presences("PLANNING AUTO", "planning")

    # Limiter aux colonnes avant ou jusqu’à AL HAMEEDI inclus
    limite_index = None
    for col, nom in noms_chirurgiens.items():
        if nom.strip().upper() == "AL HAMEEDI":
            limite_index = col
            break

    if limite_index is not None:
        noms_chirurgiens = {col: nom for col, nom in noms_chirurgiens.items() if col <= limite_index}

    # Sauvegarder la liste localement
    with open(chemin_noms, "w") as f:
        json.dump(noms_chirurgiens, f, indent=2)

    return noms_chirurgiens, emails
    
# 📌 Lancer le menu principal
if __name__ == "__main__":
    menu_principal()