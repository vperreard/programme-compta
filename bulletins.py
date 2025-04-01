import os
import json
import pdfplumber
import re
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
import time
import datetime
import pandas as pd
import subprocess
import calendar
from generer_virement import generer_xml_virements  # Import de la fonction


# 🔹 Configuration des fichiers et cache
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(SCRIPT_DIR, "file_paths.json")
CACHE_FILE = os.path.expanduser("~/Documents/Contrats/bulletins_cache.json")

# 🔹 Dictionnaire des mois
MOIS_MAPPING = {
    "janvier": "Janvier", "février": "Février", "mars": "Mars", "avril": "Avril",
    "mai": "Mai", "juin": "Juin", "juillet": "Juillet", "août": "Août",
    "septembre": "Septembre", "octobre": "Octobre", "novembre": "Novembre", "décembre": "Décembre"
}

# 🔹 Expressions régulières pour extraire les données des bulletins PDF
DATE_REGEXES = [
    re.compile(r"Période\s*:\s*([A-Za-zéûîôèê]+\s*\d{4})"),  # Ex: Période : Janvier 2024
    re.compile(r"Mois de paie\s*:\s*([A-Za-zéûîôèê]+\s*\d{4})"),  # Ex: Mois de paie : Février 2023
    re.compile(r"Bulletin du\s*(\d{2}/\d{2}/\d{4})"),  # Ex: Bulletin du 05/11/2023
]

BULLETIN_START_REGEX = re.compile(
    r"3ANESTR##BULLETIN##(\d{2}-\d{4})##(\d{5})##([\w\-\s']+)##([\w\-\s']+)##\d+"
)
BRUT_REGEX = re.compile(r"Salaire brut\s+([\d\s,\.]+)")
NET_AVANT_IMPOT_REGEX = re.compile(r"Net à payer avant impôt sur le revenu\s+([\d\s,\.]+)")
NET_APRES_IMPOT_REGEX = re.compile(r"Net payé\s+([\d\s,\.]+)")

def clean_amount(value):
    """ Nettoie et convertit une valeur monétaire extraite en float """
    if not value:
        print("⚠️ Erreur : Valeur vide détectée pour un montant.")
        return None
    
    try:
        # Remplace les espaces et les virgules avant conversion
        return float(value.replace(" ", "").replace(",", "."))
    except ValueError:
        print(f"⚠️ Erreur de conversion : Impossible de convertir '{value}' en float.")
        return None  # Retourne None si la conversion échoue

# 📂 Charger le cache des bulletins
try:
    with open(CACHE_FILE, "r") as f:
        bulletins_cache = json.load(f)
except FileNotFoundError:
    bulletins_cache = {}

# Charger le chemin des bulletins depuis le fichier JSON (si défini)
try:
    with open(SETTINGS_FILE, "r") as f:
        file_paths = json.load(f)
except FileNotFoundError:
    file_paths = {}

# S'assurer que bulletins_path est bien défini après chargement
bulletins_path = file_paths.get("bulletins_salaire", os.path.expanduser("~/Documents/Bulletins_Salaire"))

print(f"📂 Chemin des bulletins utilisé après mise à jour : {bulletins_path}")


def extract_date(text, filename):
    """
    Extrait la date d'un texte et, en dernier recours, tente de l'extraire du nom du fichier.
    """
    for regex in DATE_REGEXES:
        match = regex.search(text)
        if match:
            raw_date = match.group(1).strip()
            print(f"🔍 Debug : Date brute trouvée dans {filename} -> {raw_date}")  # Debug

            # Vérifier si c'est un format JJ/MM/AAAA
            if re.match(r"\d{2}/\d{2}/\d{4}", raw_date) or re.match(r"\d{2}-\d{2}-\d{4}", raw_date):
                jour, mois, annee = re.split(r"[-/]", raw_date)
                mois_nom = calendar.month_name[int(mois)].capitalize()  # Convertir 01 -> Janvier
                print(f"✅ Date extraite (Format JJ/MM/AAAA) : {mois_nom} {annee}")
                return mois_nom, annee

            # Vérifier si c'est un format "Mois Année"
            parts = raw_date.split()
            if len(parts) == 2:
                mois, annee = parts
                mois = mois.lower()
                if mois in MOIS_MAPPING:
                    mois_nom = MOIS_MAPPING[mois]
                    print(f"✅ Date extraite (Format Mois Année) : {mois_nom} {annee}")
                    return mois_nom, annee
                else:
                    print(f"⚠️ Mois inconnu dans {filename}: {mois}")
                    return None, None

    print(f"⚠️ Aucune date trouvée dans {filename}, tentative d'extraction depuis le nom du fichier...")
    return extract_date_from_filename(filename)

def extract_date_from_filename(filename):
    """
    Tente d'extraire la date à partir du nom du fichier.
    Ex: "Bulletin alacroix 141223.pdf" → Décembre 2023
    """
    match = re.search(r"(\d{2})(\d{2})(\d{2,4})", filename)
    if match:
        jour, mois, annee = match.groups()
        mois_nom = calendar.month_name[int(mois)].capitalize()
        annee = "20" + annee if len(annee) == 2 else annee  # Conversion si année abrégée
        print(f"✅ Date extraite depuis le nom du fichier {filename}: {mois_nom} {annee}")
        return mois_nom, annee
    return None, None

def extract_name(text):
    """
    Extrait le nom et prénom du bulletin de salaire en repérant la ligne de structure connue.
    Capture aussi les noms composés.
    """
    match = re.search(r"##BULLETIN##\d{2}-\d{4}##\d{5}##([\w\-\s]+)##([\w\-\s]+)##\d+", text)
    if match:
        nom = match.group(1).strip().title()  # Conserve les majuscules/minuscules correctes
        prenom = match.group(2).strip().title()
        return nom, prenom
    return "?", "?"

# 📌 Fonction pour extraire les données des PDF
def extract_info_from_pdf(pdf_path):
    """Extrait les bulletins des PDF en fusionnant correctement les pages tout en gardant les bulletins distincts si besoin."""
    bulletins_temp = {}  # Dictionnaire temporaire pour fusionner les pages
    bulletins = []  # Liste finale des bulletins extraits

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                continue

            match = BULLETIN_START_REGEX.search(text)
            if match:
                matricule = match.group(2)
                nom = match.group(3).title()
                prenom = match.group(4).title()
                
                # 🔹 Extraction des dates d'entrée et sortie (si disponibles)
                entree_match = re.search(r"Entrée\s*:\s*(\d{2}/\d{2}/\d{4})", text)
                sortie_match = re.search(r"Sortie\s*:\s*(\d{2}/\d{2}/\d{4})", text)

                date_entree = entree_match.group(1) if entree_match else "INCONNUE"
                date_sortie = sortie_match.group(1) if sortie_match else "INCONNUE"

                # 🔑 Clé unique pour identifier un bulletin
                key = (matricule, nom, prenom, date_entree, date_sortie)

                if key in bulletins_temp:
                    # 🟢 Même bulletin : on fusionne les pages
                    bulletins_temp[key]["text"] += "\n" + text  
                else:
                    # 🟠 Nouveau bulletin : on l'ajoute
                    bulletins_temp[key] = {
                        "text": text,
                        "pdf_path": pdf_path,
                        "page_num": page_num,
                    }

    # 🛠 Traitement des bulletins fusionnés
    for key, data in bulletins_temp.items():
        matricule, nom, prenom, date_entree, date_sortie = key
        bulletin = extract_bulletin_info(data["text"], data["pdf_path"], data["page_num"])
        if bulletin:
            bulletin["DateEntree"] = date_entree
            bulletin["DateSortie"] = date_sortie
            bulletins.append(bulletin)

    return bulletins  # 🔥 Retourne des bulletins propres, sans doublon parasite


def extract_bulletin_info(text, pdf_path, page_num):
    """Extrait les informations d'un bulletin à partir du texte fusionné sur plusieurs pages, sans recalculer les valeurs absentes."""
    
    # 🔹 Vérification de la présence d'un bulletin
    match = BULLETIN_START_REGEX.search(text)
    if not match:
        print(f"⚠️ Aucun en-tête de bulletin détecté dans {pdf_path} (page {page_num})")
        return None

    periode, matricule, nom, prenom = match.groups()

    # 🔹 Normalisation du nom et prénom (mise en majuscule de la première lettre de chaque mot)
    nom = " ".join(word.capitalize() for word in nom.split())
    prenom = " ".join(word.capitalize() for word in prenom.split())

    # 🔹 Extraction de la date
    mois, annee = extract_date(text, pdf_path)
    if not mois or not annee:
        print(f"⚠️ Impossible d'extraire la date pour {nom} {prenom} - {pdf_path} (page {page_num})")
        return None

    # 🔹 Extraction des montants (dernier trouvé dans le texte complet)
    brut_values = BRUT_REGEX.findall(text)
    net_avant_values = NET_AVANT_IMPOT_REGEX.findall(text)
    net_apres_values = NET_APRES_IMPOT_REGEX.findall(text)

    # Sélectionner le dernier montant extrait SANS recalcul
    brut_value = brut_values[-1] if brut_values else None
    net_avant_value = net_avant_values[-1] if net_avant_values else None
    net_apres_value = net_apres_values[-1] if net_apres_values else None

    print(f"🔍 Debug Montants trouvés : Brut='{brut_value}', NetAvant='{net_avant_value}', NetAprès='{net_apres_value}'")  # Log avant conversion

    brut = clean_amount(brut_value)
    net_avant = clean_amount(net_avant_value)
    net_apres = clean_amount(net_apres_value)

    # 🔍 Vérification si les montants sont trouvés
    if brut is None and net_avant is None and net_apres is None:
        print(f"⚠️ Montants totalement manquants pour {nom} {prenom} - {mois} {annee} (page {page_num})")
        return None  # Ignorer ce bulletin si les montants sont introuvables

    # 🔹 Construction de l'objet bulletin
    bulletin = {
        "Nom": nom,
        "Prénom": prenom,
        "Date": f"{mois} {annee}",
        "Brut": brut,
        "NetAvantImpot": net_avant,
        "NetApresImpot": net_apres,
        "Matricule": matricule,
        "Fichier": pdf_path,
        "Page": page_num,
    }

    print(f"✅ Bulletin extrait et ajouté : {bulletin}")  # 🔥 Debugging

    return bulletin

def open_virement_window(tree):
    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("Avertissement", "Veuillez sélectionner au moins un bulletin pour le virement.")
        return

    # Charger le fichier des salariés pour récupérer les IBAN
    salaries_df = None
    salaries_file = file_paths.get("excel_salaries")
    if salaries_file and os.path.exists(salaries_file):
        try:
            salaries_df = pd.read_excel(salaries_file, sheet_name="Salariés")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de charger le fichier des salariés : {e}")
            return

    virements_data = []
    for item in selected_items:
        values = tree.item(item, "values")
        nom = values[0]
        prenom = values[1]
        # On utilise le NetApresImpot (colonne 5) pour le montant, format "XXXX €"
        montant_str = values[5].replace("€", "").replace(" ", "")
        try:
            montant = float(montant_str)
        except:
            montant = 0.0

        # Recherche de l'IBAN dans le fichier des salariés (Excel)
        iban_lookup = ""
        if salaries_df is not None:
            result = salaries_df[
                (salaries_df["PRENOM"].str.lower() == prenom.lower()) &
                (salaries_df["NOM"].str.lower() == nom.lower())
            ]
            if not result.empty:
                iban_lookup = str(result.iloc[0]["IBAN"])

        # Si aucun IBAN n'est trouvé, chercher dans le fichier des IADE remplaçants
        if not iban_lookup:
            excel_iade_file = file_paths.get("excel_iade")
            if excel_iade_file and os.path.exists(excel_iade_file):
                try:
                    iade_df = pd.read_excel(excel_iade_file, sheet_name="Coordonnées IADEs")
                    result_iade = iade_df[
                        (iade_df["PRENOMR"].str.lower() == prenom.lower()) &
                        (iade_df["NOMR"].str.lower() == nom.lower())
                    ]
                    if not result_iade.empty:
                        iban_lookup = str(result_iade.iloc[0]["IBAN"])
                except Exception as e:
                    messagebox.showerror("Erreur", f"Impossible de charger le fichier des IADE remplaçants : {e}")

        virements_data.append({
            "beneficiaire": f"{prenom} {nom}",
            "iban": iban_lookup,
            "net_pay": montant,
        })

    # Création de la fenêtre de saisie des virements
    virement_window = tk.Toplevel()
    virement_window.title("Détails des virements")
    
    # En-têtes des colonnes
    headers = ["Bénéficiaire", "IBAN", "Date (YYYY-MM-DD)", "Référence", "Montant (€)"]
    for col, header in enumerate(headers):
        tk.Label(virement_window, text=header, font=("Arial", 10, "bold")).grid(row=0, column=col, padx=5, pady=5)
    
    # Date par défaut = aujourd'hui
    default_date = time.strftime("%Y-%m-%d")
    virement_entries = []
    for i, v_data in enumerate(virements_data):
        row_entries = {}
        # Bénéficiaire (affiché)
        e_benef = tk.Entry(virement_window, width=20)
        e_benef.insert(0, v_data["beneficiaire"])
        e_benef.grid(row=i+1, column=0, padx=5, pady=5)
        row_entries["beneficiaire"] = e_benef

        # IBAN (pré-rempli si trouvé)
        e_iban = tk.Entry(virement_window, width=25)
        e_iban.insert(0, v_data["iban"])
        e_iban.grid(row=i+1, column=1, padx=5, pady=5)
        row_entries["iban"] = e_iban

        # Date
        e_date = tk.Entry(virement_window, width=15)
        e_date.insert(0, default_date)
        e_date.grid(row=i+1, column=2, padx=5, pady=5)
        row_entries["date"] = e_date

        # Génération de la référence selon le format VIR-SELANE-<initialPrenom><NOM>-<MOISabbrev><YY>
        # Exemple : Pour "Vincent Perreard" et la date par défaut "2025-02-28", la référence sera VIR-SELANE-VPERREARD-FEV25
        parts = v_data["beneficiaire"].split()
        if len(parts) >= 2:
            ref_benef = parts[0][0].upper() + parts[-1].upper()
        else:
            ref_benef = v_data["beneficiaire"].upper()
        try:
            date_obj = datetime.datetime.strptime(default_date, "%Y-%m-%d")
            french_abbr = {
                1: "JAN", 2: "FEV", 3: "MAR", 4: "AVR", 5: "MAI", 6: "JUN",
                7: "JUL", 8: "AOU", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"
            }
            month_abbr = french_abbr.get(date_obj.month, "")
            year_last2 = str(date_obj.year)[-2:]
        except Exception as e:
            month_abbr = ""
            year_last2 = ""
        default_ref = f"VIR-SELANE-{ref_benef}-{month_abbr}{year_last2}"
        
        e_ref = tk.Entry(virement_window, width=20)
        e_ref.insert(0, default_ref)
        e_ref.grid(row=i+1, column=3, padx=5, pady=5)
        row_entries["reference"] = e_ref

        # Montant : pré-rempli avec le net payé
        e_montant = tk.Entry(virement_window, width=15)
        e_montant.insert(0, str(v_data["net_pay"]))
        e_montant.grid(row=i+1, column=4, padx=5, pady=5)
        row_entries["montant"] = e_montant

        virement_entries.append(row_entries)
    
    def generate_xml():
        virements_list = []
        try:
            for row in virement_entries:
                virements_list.append({
                    "beneficiaire": row["beneficiaire"].get(),
                    "iban": row["iban"].get(),
                    "objet": f"{row['date'].get()} - {row['reference'].get()}",
                    "montant": float(row["montant"].get()),
                })
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur dans la saisie des virements : {e}")
            return

        # Génération du fichier XML
        fichier_xml = generer_xml_virements(virements_list)
        if fichier_xml:
            messagebox.showinfo("Succès", f"Fichier XML généré : {fichier_xml}")

            # 🔥 AJOUT : Envoi du fichier XML sur le site de la banque LCL
            import generer_virement
            generer_virement.envoyer_virement_vers_lcl(fichier_xml)

            virement_window.destroy()
        else:
            messagebox.showerror("Erreur", "Échec de la génération du fichier XML.")
    
    tk.Button(virement_window, text="Générer XML et Réaliser Virement", command=generate_xml).grid(row=len(virement_entries)+1, column=0, columnspan=5, pady=10)


# 📌 Fonction pour recharger le cache en forçant la lecture des fichiers PDF
def update_cache(force_reload=False):
    """Recharge les bulletins de salaire et met à jour le cache, uniquement si nécessaire."""
    global bulletins_cache, bulletins_path
    try:
        with open(SETTINGS_FILE, "r") as f:
            file_paths = json.load(f)
        bulletins_path = file_paths.get("bulletins_salaire", os.path.expanduser("~/Documents/Bulletins_Salaire"))
        print(f"📂 Chemin des bulletins utilisé : {bulletins_path}")
    except Exception as e:
        print(f"⚠️ Erreur lors du chargement du chemin des bulletins: {e}")
    
    
    print(f"📂 DEBUG : Chemin utilisé pour recharger le cache = {bulletins_path}")
    print(f"📂 Chemin des bulletins scanné : {bulletins_path}")
    print(f"🔄 Rechargement du cache... Force reload : {force_reload}")

    # Si le cache est déjà rempli et qu’on ne force pas le rechargement, on ne fait rien
    if not force_reload and bulletins_cache:
        print("✅ Cache existant utilisé, aucun rechargement nécessaire.")
        return

    print("🔄 Rechargement du cache à partir des fichiers PDF...")
    bulletins_cache = {}  # Réinitialisation du cache

    nb_total_bulletins = 0  # Compteur du nombre total de bulletins extraits

    for root, _, files in os.walk(bulletins_path):
        for file in files:
            if file.lower().endswith(".pdf"):
                full_path = os.path.join(root, file)
                print(f"📂 Analyse du fichier : {full_path}")  # Debug : Afficher le fichier analysé

                # Extraction des informations du PDF
                bulletins = extract_info_from_pdf(full_path)
                if not bulletins:
                    print(f"⚠️ Aucun bulletin trouvé dans {full_path}")
                    continue  # On passe au fichier suivant si aucun bulletin trouvé

                # Ajout des bulletins au cache
                for bulletin in bulletins:
                    mois, annee = bulletin["Date"].split()
                    
                    # Création des clés si elles n'existent pas encore
                    if annee not in bulletins_cache:
                        bulletins_cache[annee] = {}
                    if mois not in bulletins_cache[annee]:
                        bulletins_cache[annee][mois] = []
                    
                    bulletins_cache[annee][mois].append(bulletin)
                    nb_total_bulletins += 1  # Incrémentation du compteur
                    print(f"📌 Enregistré dans le cache : {bulletin}")  # 🔥 Debugging


    print(f"📊 Nombre total de bulletins ajoutés au cache : {nb_total_bulletins}")

    # 🔥 Vérification finale du cache
    if bulletins_cache:
        print(f"✅ Cache mis à jour avec {nb_total_bulletins} bulletins.")
    else:
        print("⚠️ Aucun bulletin n'a été ajouté au cache. Vérifie les regex et le contenu des PDF.")

    # Sauvegarde du cache mis à jour
    with open(CACHE_FILE, "w") as f:
        json.dump(bulletins_cache, f, indent=4)

    print("💾 Cache sauvegardé avec succès !")  
    
def reset_cache():
    """Efface le cache et recharge tous les bulletins de salaire."""
    global bulletins_cache

    print("🗑️ Suppression du cache en mémoire et sur disque...")

    bulletins_cache.clear()  # Efface le cache en mémoire

    # Supprimer le fichier de cache s'il existe
    if os.path.exists(CACHE_FILE):
        os.remove(CACHE_FILE)
        print("🗑️ Cache supprimé du disque.")

    # Recharger tous les bulletins
    update_cache(force_reload=True)  # On force le rechargement du cache
    refresh_table()  # Rafraîchir l'affichage de la table après rechargement
    
    
def open_pdf_at_page(file_path, page_num):
    """Ouvre un fichier PDF avec PDF Expert, puis utilise AppleScript pour aller à la page spécifiée."""
    
    try:
        if not os.path.exists(file_path):
            messagebox.showerror("Erreur", "Le fichier sélectionné n'existe plus.")
            return

        # 🔹 Étape 1 : Ouvrir le fichier avec PDF Expert (sans AppleScript)
        subprocess.run(["open", "-a", "PDF Expert", file_path], check=True)

        # 🔹 Attendre un court instant que le fichier s'ouvre
        time.sleep(1)  # Ajustable si nécessaire

        # 🔹 Étape 2 : Activer PDF Expert et aller à la page avec AppleScript
        apple_script = f'''
        tell application "PDF Expert"
            activate
        end tell

        delay 0.5

        tell application "System Events"
            keystroke "g" using {{option down, command down}} -- ⌥ + ⌘ + G pour "Aller à la page"
            delay 0.3
            keystroke "{page_num}"
            delay 0.3
            keystroke return
        end tell
        '''

        # 🔥 Exécuter AppleScript pour aller à la page
        subprocess.run(["osascript", "-e", apple_script], check=True)

        print(f"📂 PDF ouvert rapidement dans PDF Expert : {file_path}, page {page_num}")

    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erreur", f"Impossible d'ouvrir le fichier PDF : {e}")
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur inconnue : {e}")

def show_details_in_frame(annee, mois, parent_frame):
    """Affiche les détails des bulletins dans un cadre existant."""
    # Nettoyer le cadre parent
    for widget in parent_frame.winfo_children():
        widget.destroy()

    # Titre
    tk.Label(parent_frame, text=f"📄 Bulletins de {mois} {annee}", 
             font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x")

    columns = ("Nom", "Prénom", "Date", "Brut", "NetAvantImpot", "NetApresImpot", "Fichier", "Page")
    tree = ttk.Treeview(parent_frame, columns=columns, show="headings", selectmode="extended")

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, anchor="center", width=120)

    # Ajouter les bulletins du mois sélectionné
    for bulletin in bulletins_cache[annee][mois]:
        tree.insert("", "end", values=(
            bulletin.get("Nom", "?"),
            bulletin.get("Prénom", "?"),
            bulletin["Date"],
            f"{bulletin['Brut']} €",
            f"{bulletin['NetAvantImpot']} €",
            f"{bulletin['NetApresImpot']} €",
            bulletin["Fichier"],
            bulletin["Page"]
        ))
    tree.pack(expand=True, fill="both", padx=10, pady=10)

    # Boutons
    buttons_frame = tk.Frame(parent_frame, bg="#f0f0f0")
    buttons_frame.pack(fill="x", pady=10)

    # Fonction pour ouvrir le PDF
    def open_selected():
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un bulletin.")
            return

        values = tree.item(selected_item[0], "values")
        file_path = values[6]
        page_num = int(values[7])

        if not os.path.exists(file_path):
            messagebox.showerror("Erreur", "Le fichier sélectionné n'existe plus.")
            return
        
        open_pdf_at_page(file_path, page_num)

    # Bouton pour afficher le bulletin
    tk.Button(buttons_frame, text="📂 Afficher le bulletin", command=open_selected,
              bg="#4682B4", fg="black", font=("Arial", 12, "bold")).pack(side="left", padx=10)
    
    # Bouton pour faire un virement
    tk.Button(buttons_frame, text="💰 Faire virement", command=lambda: open_virement_window(tree),
              bg="#32CD32", fg="black", font=("Arial", 12, "bold")).pack(side="left", padx=10)
    
    # Bouton retour aux bulletins
    tk.Button(buttons_frame, text="🔙 Retour aux bulletins", 
              command=lambda: show_bulletins_in_frame(parent_frame),
              bg="#B0C4DE", fg="black", font=("Arial", 12, "bold")).pack(side="left", padx=10)

def show_bulletins_in_frame(frame):
    """Version de show_bulletins qui affiche dans un cadre existant."""
    # Recharger les chemins depuis la configuration
    global bulletins_path
    try:
        with open(SETTINGS_FILE, "r") as f:
            file_paths = json.load(f)
        bulletins_path = file_paths.get("bulletins_salaire", os.path.expanduser("~/Documents/Bulletins_Salaire"))
        print(f"📂 Chemin des bulletins chargé: {bulletins_path}")
    except Exception as e:
        print(f"⚠️ Erreur lors du chargement du chemin des bulletins: {e}")
    
  
    # Configuration du cadre
    frame.config(bg="#f0f0f0")

    # Bouton pour recharger les bulletins
    reload_button = tk.Button(
        frame, text="🔄 Recharger les bulletins",
        command=lambda: [update_cache(force_reload=True), 
                        refresh_same_frame(frame)],  # Utiliser refresh_same_frame au lieu de display_bulletins_in_container
        bg="#FFA500", fg="black", font=("Arial", 12, "bold")
    )
    reload_button.pack(pady=5)

    # Fonction auxiliaire pour rafraîchir le même cadre
    def refresh_same_frame(frame_to_refresh):
        """Vide et remplit à nouveau le même cadre."""
        for widget in frame_to_refresh.winfo_children():
            widget.destroy()
        show_bulletins_in_frame(frame_to_refresh)

    # Bouton pour scanner un fichier PDF
    scan_pdf_button = tk.Button(
        frame, text="📂 Scanner un PDF",
        command=scan_new_pdf,
        bg="#D3D3D3", fg="black", font=("Arial", 12, "bold")
    )
    scan_pdf_button.pack(pady=5)

    # Titre
    title_label = tk.Label(frame, text="📄 Sélectionnez un mois", font=("Arial", 14, "bold"), bg="#f0f0f0")
    title_label.pack(pady=10)

    # Cadre pour la grille des mois/années
    frame_grille = tk.Frame(frame, bg="#f0f0f0")
    frame_grille.pack()

    # Assurez-vous que le cache est chargé
    if not bulletins_cache:
        update_cache()

    # Affichage des années en haut
    annees = sorted(bulletins_cache.keys()) if bulletins_cache else []
    for col_idx, annee in enumerate(annees):
        tk.Label(frame_grille, text=annee, font=("Arial", 12, "bold"), bg="#f0f0f0").grid(row=0, column=col_idx + 1, padx=20, pady=5)

    # Affichage des mois sous chaque année
    mois_list = list(MOIS_MAPPING.values())
    for row_idx, mois in enumerate(mois_list, start=1):
        # Afficher le nom du mois dans la première colonne
        tk.Label(frame_grille, text=mois, font=("Arial", 10), bg="#f0f0f0").grid(row=row_idx, column=0, padx=10, pady=5)

        for col_idx, annee in enumerate(annees):
            if mois in bulletins_cache.get(annee, {}):
                tk.Button(frame_grille, text=mois, width=12, height=2,
                          command=lambda a=annee, m=mois, f=frame: show_details_in_frame(a, m, f),
                          bg="#87CEEB", fg="black", font=("Arial", 10, "bold")).grid(row=row_idx, column=col_idx + 1, padx=5, pady=5)



# 📌 Interface principale pour afficher les bulletins
def show_bulletins():
    """Affiche l'interface des bulletins triés par année/mois dans le frame fourni"""
    # On utilise une variable globale fournie par display_in_right_frame
    # ou on crée une fenêtre Toplevel si appelé directement
    try:
        # Si on est appelé via display_bulletins_in_container
        window = current_frame
    except NameError:
        # Si on est appelé directement (comportement original)
        window = tk.Toplevel()
        window.title("📄 Bulletins de Salaire")
        window.geometry("700x850")
        window.focus_force()  # Met la fenêtre au premier plan
        window.transient()  # La rattache à la fenêtre principale

    window.config(bg="#f0f0f0")

    # ✅ Bouton pour recharger les bulletins
    reload_button = tk.Button(
        window, text="🔄 Recharger les bulletins",
        command=lambda: [update_cache(force_reload=True), 
                        # Si on est dans une fenêtre Toplevel, on la détruit et en recrée une
                        window.destroy() if isinstance(window, tk.Toplevel) else None,
                        show_bulletins()],
        bg="#FFA500", fg="black", font=("Arial", 12, "bold")
    )
    reload_button.pack(pady=5)

    # ✅ Bouton pour scanner un fichier PDF
    scan_pdf_button = tk.Button(
        window, text="📂 Scanner un PDF",
        command=scan_new_pdf,
        bg="#D3D3D3", fg="black", font=("Arial", 12, "bold")
    )
    scan_pdf_button.pack(pady=5)

    # ✅ Bouton Retour - seulement si on est dans une Toplevel
    if isinstance(window, tk.Toplevel):
        retour_button = tk.Button(
            window, text="🔙 Retour",
            command=window.destroy,
            bg="#B0C4DE", fg="black", font=("Arial", 12, "bold")
        )
        retour_button.pack(pady=10)

    # ✅ Titre
    title_label = tk.Label(window, text="📄 Sélectionnez un mois", font=("Arial", 14, "bold"), bg="#f0f0f0")
    title_label.pack(pady=10)

    # ✅ Cadre pour la grille des mois/années
    frame_grille = tk.Frame(window, bg="#f0f0f0")
    frame_grille.pack()

    # Assurez-vous que le cache est chargé
    if not bulletins_cache:
        update_cache()

    # 🔹 Affichage des années en haut
    annees = sorted(bulletins_cache.keys()) if bulletins_cache else []
    for col_idx, annee in enumerate(annees):
        tk.Label(frame_grille, text=annee, font=("Arial", 12, "bold"), bg="#f0f0f0").grid(row=0, column=col_idx + 1, padx=20, pady=5)

    # 🔹 Affichage des mois sous chaque année
    mois_list = list(MOIS_MAPPING.values())
    for row_idx, mois in enumerate(mois_list, start=1):
        # Afficher le nom du mois dans la première colonne
        tk.Label(frame_grille, text=mois, font=("Arial", 10), bg="#f0f0f0").grid(row=row_idx, column=0, padx=10, pady=5)

        for col_idx, annee in enumerate(annees):
            if mois in bulletins_cache.get(annee, {}):
                tk.Button(frame_grille, text=mois, width=12, height=2,
                          command=lambda a=annee, m=mois: show_details(a, m, window),
                          bg="#87CEEB", fg="black", font=("Arial", 10, "bold")).grid(row=row_idx, column=col_idx + 1, padx=5, pady=5)

    # Si c'est une fenêtre Toplevel, on lance la boucle mainloop
    if isinstance(window, tk.Toplevel):
        window.mainloop()

def scan_new_pdf():
    """Ouvre l'explorateur de fichiers pour sélectionner un PDF à scanner sans rescanner tout le cache."""
    file_path = filedialog.askopenfilename(
        title="Sélectionnez un fichier PDF",
        filetypes=[("Fichiers PDF", "*.pdf")]
    )

    if not file_path:
        return  # L'utilisateur a annulé la sélection

    print(f"📂 Fichier sélectionné : {file_path}")

    # Extraction des bulletins du fichier sélectionné
    bulletins = extract_info_from_pdf(file_path)
    if not bulletins:
        messagebox.showinfo("Résultat", "Aucun bulletin trouvé dans ce fichier.")
        return

    # Ajouter les nouveaux bulletins au cache sans toucher aux autres
    nb_ajoutes = 0
    for bulletin in bulletins:
        mois, annee = bulletin["Date"].split()

        # Création des clés si elles n'existent pas encore
        if annee not in bulletins_cache:
            bulletins_cache[annee] = {}
        if mois not in bulletins_cache[annee]:
            bulletins_cache[annee][mois] = []

        bulletins_cache[annee][mois].append(bulletin)
        nb_ajoutes += 1

    # Sauvegarde du cache mis à jour
    with open(CACHE_FILE, "w") as f:
        json.dump(bulletins_cache, f, indent=4)

    messagebox.showinfo("Succès", f"{nb_ajoutes} bulletins ajoutés !")

    # Rafraîchir l'affichage de la fenêtre des bulletins
    show_bulletins()

def show_details(annee, mois, parent_window=None):
    """Affiche les détails des bulletins d'un mois sélectionné"""
    # Nettoyer la fenêtre parent si elle est fournie
    if parent_window and not isinstance(parent_window, tk.Toplevel):
        for widget in parent_window.winfo_children():
            widget.destroy()
        window = parent_window
    else:
        window = tk.Toplevel()
        window.title(f"📄 Bulletins de {mois} {annee}")
        window.geometry("900x500")

    # Titre
    tk.Label(window, text=f"📄 Bulletins de {mois} {annee}", 
             font=("Arial", 14, "bold"), bg="#4a90e2", fg="white").pack(fill="x")

    columns = ("Nom", "Prénom", "Date", "Brut", "NetAvantImpot", "NetApresImpot", "Fichier", "Page")
    tree = ttk.Treeview(window, columns=columns, show="headings", selectmode="extended")

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, anchor="center", width=120)

    # Ajouter les bulletins du mois sélectionné
    for bulletin in bulletins_cache[annee][mois]:
        tree.insert("", "end", values=(
            bulletin.get("Nom", "?"),
            bulletin.get("Prénom", "?"),
            bulletin["Date"],
            f"{bulletin['Brut']} €",
            f"{bulletin['NetAvantImpot']} €",
            f"{bulletin['NetApresImpot']} €",
            bulletin["Fichier"],
            bulletin["Page"]
        ))
    tree.pack(expand=True, fill="both", padx=10, pady=10)

    def open_selected():
        """Ouvre le fichier PDF du bulletin sélectionné à la page correcte"""
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Avertissement", "Veuillez sélectionner un bulletin.")
            return

        values = tree.item(selected_item[0], "values")

        # Vérification et récupération du chemin du fichier et du numéro de page
        try:
            file_path = values[6]  # Index 6 -> Chemin du fichier PDF
            page_num = int(values[7])  # Index 7 -> Numéro de page
        except IndexError:
            messagebox.showerror("Erreur", "Impossible de récupérer le fichier et la page (index hors limites).")
            return
        except ValueError:
            messagebox.showerror("Erreur", "Le numéro de page n'est pas un entier valide.")
            return

        # Vérifier si le fichier existe avant de l'ouvrir
        if not os.path.exists(file_path):
            messagebox.showerror("Erreur", "Le fichier sélectionné n'existe plus.")
            return
        
        # Ouvrir le PDF à la bonne page avec PDF Expert
        open_pdf_at_page(file_path, page_num)

    # Ajouter un bouton pour afficher le bulletin en PDF Expert
    show_contract_button = tk.Button(
        window, text="📂 Afficher le bulletin",
        command=open_selected,
        bg="#4682B4", fg="black", font=("Arial", 12, "bold")
    )
    show_contract_button.pack(pady=10)
    
    # Bouton "Faire virement" : il utilisera les bulletins sélectionnés dans le Treeview
    tk.Button(window, text="Faire virement", command=lambda: open_virement_window(tree),
              bg="#32CD32", fg="black", font=("Arial", 12, "bold")).pack(pady=10)
    
    # 🔙 Bouton Retour - comportement différent selon le contexte
    if isinstance(window, tk.Toplevel):
        # Si c'est une fenêtre Toplevel, retour = fermer la fenêtre
        retour_button = tk.Button(
            window, text="🔙 Retour",
            command=window.destroy,
            bg="#B0C4DE", fg="black", font=("Arial", 12, "bold")
        )
    else:
        # Si c'est dans le right_frame, retour = afficher les bulletins
        retour_button = tk.Button(
            window, text="🔙 Retour aux bulletins",
            command=lambda: show_bulletins(),
            bg="#B0C4DE", fg="black", font=("Arial", 12, "bold")
        )
    retour_button.pack(pady=10)

    # Si c'est une fenêtre Toplevel, lancer la boucle mainloop
    if isinstance(window, tk.Toplevel):
        window.mainloop()

# 🔹 Exécuter l'interface avec mise à jour du cache
if __name__ == "__main__":
    update_cache()
    show_bulletins()