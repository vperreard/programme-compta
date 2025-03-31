import os
import pdfplumber
import re
import pandas as pd

import pytesseract
from PIL import Image, ImageTk  # Ajout pour gérer les icônes
from datetime import datetime
from pdf2image import convert_from_path
import schwifty
import unicodedata
import tkinter as tk
from tkinter import messagebox
from tkinter import simpledialog
from tkinter import ttk
import subprocess
from tkinter import scrolledtext
import openai
import json
import argparse
import sys
import threading
# Charger les credentials depuis le fichier JSON
with open('credentials.json', 'r') as file:
    credentials = json.load(file)




from factures_db_utils import (
    charger_base_donnees, 
    sauvegarder_base_donnees, 
    ajouter_ou_maj_facture, 
    marquer_factures_manquantes,
    calculer_hash_fichier
)

# DEBUGGING - Afficher les arguments reçus
print("\n" + "="*50)
print("🔍 DEBUG - ANALYSE FACTURES - DÉMARRAGE")
print(f"🔄 Arguments système reçus: {sys.argv}")
print("="*50 + "\n")

# Essayer de récupérer le chemin des arguments en premier
parser = argparse.ArgumentParser(description="Analyse de factures PDF", add_help=False)
parser.add_argument("--dossier", help="Chemin du dossier contenant les factures")
args, _ = parser.parse_known_args()

if args.dossier:
    DOSSIER_FACTURES = args.dossier
    print(f"✅ DOSSIER_FACTURES défini par argument: {DOSSIER_FACTURES}")
else:
    # Sinon, essayer depuis config
    try:
        from config import get_file_path
        config_path = get_file_path("dossier_factures", verify_exists=True)
        if config_path:
            DOSSIER_FACTURES = config_path
            print(f"✅ DOSSIER_FACTURES défini par config: {DOSSIER_FACTURES}")
        else:
            DOSSIER_FACTURES = os.path.expanduser("~/Documents/Frais_Factures")
            print(f"⚠️ Config path vide, utilisation du chemin par défaut: {DOSSIER_FACTURES}")
    except Exception as e:
        DOSSIER_FACTURES = os.path.expanduser("~/Documents/Frais_Factures")
        print(f"❌ Erreur lors de l'import de config: {e}")
        print(f"⚠️ Utilisation du chemin par défaut: {DOSSIER_FACTURES}")

# Vérifier que le dossier existe
if not os.path.exists(DOSSIER_FACTURES):
    print(f"❌ ERREUR: Le dossier {DOSSIER_FACTURES} n'existe pas!")
else:
    print(f"✅ Le dossier {DOSSIER_FACTURES} existe.")
    # Lister les fichiers présents
    all_files = os.listdir(DOSSIER_FACTURES)
    pdf_files = [f for f in all_files if f.lower().endswith(('.pdf', '.jpg', '.jpeg', '.png'))]
    print(f"📄 {len(pdf_files)} fichiers PDF/images trouvés:")
    for i, file in enumerate(pdf_files[:10]):  # Montrer les 10 premiers fichiers
        print(f"   {i+1}. {file}")
    if len(pdf_files) > 10:
        print(f"   ... et {len(pdf_files)-10} autres fichiers.")
    print(f"⚠️ Utilisation du chemin par défaut: {DOSSIER_FACTURES}")

def initialiser_systeme():
    """Charge les données sans réanalyser les factures sauf si nécessaire."""
    global df_factures
    
    # Vérifier si le fichier CSV existe et le charger
    if os.path.exists(CSV_FACTURES):
        try:
            df_factures = pd.read_csv(CSV_FACTURES)
            print(f"✅ Données chargées depuis {CSV_FACTURES}")
            # Ne pas lancer verifier_nouvelles_factures() ici
            return True
        except Exception as e:
            print(f"❌ Erreur lors du chargement des données: {e}")
    
    # Si le CSV n'existe pas ou est corrompu, créer un DataFrame vide
    print("⚠️ Aucune donnée existante, création d'un DataFrame vide")
    df_factures = pd.DataFrame(columns=REQUIRED_COLUMNS)
    return False


def choisir_fournisseur_iban(event, col):
    """Gère le double-clic sur une cellule fournisseur ou IBAN."""
    global df_factures, df_ibans, CSV_FACTURES
    
    row_id = tree.focus()
    if not row_id:
        return
    
    row_data = tree.item(row_id, "values")
    
    selection_window = tk.Toplevel(root)
    selection_window.title("Modifier fournisseur / IBAN")
    selection_window.geometry("500x350")
    selection_window.transient(root)
    selection_window.grab_set()
    
    # Cadre principal
    main_frame = ttk.Frame(selection_window, padding=10)
    main_frame.pack(fill="both", expand=True)
    
    # Valeurs actuelles
    current_fournisseur = row_data[1] if len(row_data) > 1 else ""
    current_iban = row_data[3] if len(row_data) > 3 else ""
    
    # 1. Section pour l'édition manuelle
    edit_frame = ttk.LabelFrame(main_frame, text="Édition manuelle du fournisseur")
    edit_frame.pack(fill="x", pady=10, padx=5)
    
    ttk.Label(edit_frame, text="Fournisseur:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    fournisseur_var = tk.StringVar(value=current_fournisseur)
    fournisseur_entry = ttk.Entry(edit_frame, textvariable=fournisseur_var, width=40)
    fournisseur_entry.grid(row=0, column=1, padx=5, pady=5)
    
    ttk.Label(edit_frame, text="IBAN:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    iban_var = tk.StringVar(value=current_iban)
    iban_entry = ttk.Entry(edit_frame, textvariable=iban_var, width=40)
    iban_entry.grid(row=1, column=1, padx=5, pady=5)
    
    # 2. Section pour choisir un fournisseur existant dans la liste
    list_frame = ttk.LabelFrame(main_frame, text="Ou choisir un fournisseur existant")
    list_frame.pack(fill="x", pady=10, padx=5)
    
    # Récupérer et trier la liste des fournisseurs existants
    fournisseurs_list = df_ibans["fournisseur"].dropna().unique().tolist()
    fournisseurs_list = sorted([f for f in fournisseurs_list if f and str(f).strip()])
    
    ttk.Label(list_frame, text="Fournisseur:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    
    # Combobox pour sélectionner un fournisseur existant
    fournisseur_combo_var = tk.StringVar()
    fournisseur_combo = ttk.Combobox(list_frame, textvariable=fournisseur_combo_var, values=fournisseurs_list, width=40)
    fournisseur_combo.grid(row=0, column=1, padx=5, pady=5)
    
    # Mettre à jour les champs quand un fournisseur est sélectionné dans la liste
    def on_combo_select(event):
        selected_fournisseur = fournisseur_combo_var.get()
        if selected_fournisseur:
            fournisseur_var.set(selected_fournisseur)
            # Trouver l'IBAN correspondant
            matching_rows = df_ibans[df_ibans["fournisseur"] == selected_fournisseur]
            if not matching_rows.empty and pd.notna(matching_rows.iloc[0]["IBAN"]):
                iban_var.set(matching_rows.iloc[0]["IBAN"])
    
    fournisseur_combo.bind("<<ComboboxSelected>>", on_combo_select)
    
    # 3. Boutons d'action
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill="x", pady=20)
    
    def valider():
        global df_ibans
   
        """Valide les changements et met à jour la facture."""
        fournisseur_selectionne = fournisseur_var.get().strip()
        iban_selectionne = iban_var.get().strip()
        
        if not fournisseur_selectionne:
            messagebox.showwarning("Avertissement", "Veuillez entrer un nom de fournisseur.")
            return
        
        # Obtenir l'index de la ligne sélectionnée
        index = tree.index(row_id)
        
        # Mettre à jour le tableau visuel
        new_values = list(row_data)
        new_values[1] = fournisseur_selectionne
        if iban_selectionne:
            new_values[3] = iban_selectionne
        tree.item(row_id, values=new_values)
        
        # Mettre à jour le DataFrame
        fichier = row_data[5] if len(row_data) > 5 else None
        if fichier:
            mask = df_factures["fichier"] == fichier
            df_factures.loc[mask, "fournisseur"] = fournisseur_selectionne
            if iban_selectionne:
                df_factures.loc[mask, "iban"] = iban_selectionne
            
            # Sauvegarder dans le CSV
            df_factures.to_csv(CSV_FACTURES, index=False)
            
            # Vérifier si c'est un nouveau fournisseur à ajouter à la liste
            if fournisseur_selectionne not in df_ibans["fournisseur"].values and iban_selectionne:
                new_row = pd.DataFrame({"fournisseur": [fournisseur_selectionne], "IBAN": [iban_selectionne]})
                df_ibans = pd.concat([df_ibans, new_row], ignore_index=True)
                df_ibans.to_csv(IBAN_LISTE_CSV, index=False)
                print(f"✅ Nouveau fournisseur ajouté: {fournisseur_selectionne}")
        
        text_area.insert("end", f"\n✅ Fournisseur mis à jour: {fournisseur_selectionne}\n")
        selection_window.destroy()
    
    ttk.Button(button_frame, text="Valider", command=valider).pack(side="left", padx=5)
    ttk.Button(button_frame, text="Annuler", command=selection_window.destroy).pack(side="left", padx=5)



from generer_virement import generer_xml_virements
from generer_virement import envoyer_virement_vers_lcl


loupe_window = None
loupe_canvas = None

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Chemin du dossier contenant les factures
IBAN_LISTE_CSV = os.path.join(SCRIPT_DIR, "liste_ibans.csv")
CSV_FACTURES = os.path.join(SCRIPT_DIR, "resultats_factures.csv")

# Liste des colonnes nécessaires à partir de l'analyse du script
REQUIRED_COLUMNS = [
    "iban", "fournisseur", "paiement", "date_facture", "montant", "fichier",  "reference_facture"
]

# Vérifier si le fichier existe, sinon le créer
if not os.path.exists(IBAN_LISTE_CSV):
    pd.DataFrame(columns=["fournisseur", "IBAN"]).to_csv(IBAN_LISTE_CSV, index=False) 

# Chargement unique des fichiers CSV
def charger_dataframe(fichier):
    """Charge un fichier CSV s'il existe, sinon retourne un DataFrame vide."""
    if os.path.exists(fichier):
        return pd.read_csv(fichier, dtype=str)
    return pd.DataFrame()

def validate_dataframe(df, required_columns):
    """Vérifie la présence des colonnes requises, supprime les doublons et ajuste le DataFrame."""
    # Supprimer les colonnes en doublon
    df = df.loc[:, ~df.columns.duplicated()].copy()

    # Vérifier et ajouter les colonnes manquantes
    for col in required_columns:
        if col not in df.columns:
            print(f"⚠️ Colonne manquante détectée : {col}, création d'une colonne vide.")
            df[col] = None

    # Afficher les colonnes après correction
    print(f"✅ Colonnes après correction : {df.columns.tolist()}")
    return df


df_factures = charger_dataframe(CSV_FACTURES)
df_factures.rename(columns={"Fournisseur": "fournisseur"}, inplace=True)
df_factures = validate_dataframe(df_factures, REQUIRED_COLUMNS)

print("Colonnes disponibles dans df_factures :", df_factures.columns)
print("Aperçu des données :", df_factures.head())

df_factures.columns = df_factures.columns.str.strip().str.lower()
print("Colonnes après normalisation :", df_factures.columns)


# Charger les IBANs et fournisseurs une seule fois
df_ibans = pd.read_csv(IBAN_LISTE_CSV, sep=",", dtype=str)
print("Colonnes disponibles dans df_ibans :", df_ibans.columns)
print(df_ibans.head())

if "fournisseur" in df_ibans.columns:
    df_ibans["fournisseur"] = df_ibans["fournisseur"].astype(str).str.strip()
else:
    print("⚠️ La colonne 'fournisseur' est absente de df_ibans.")

# Vérifications et nettoyage
df_ibans.rename(columns={"Fournisseur": "fournisseur", "iban": "IBAN"}, inplace=True)
df_ibans["fournisseur"] = df_ibans["fournisseur"].astype(str).str.strip()
df_ibans["IBAN"] = df_ibans["IBAN"].astype(str).str.replace(" ", "")

# Ajout de la colonne "paiement" si absente
if "paiement" not in df_factures.columns or df_factures["paiement"].isnull().all():
    df_factures["paiement"] = "❌"
else:
    df_factures["paiement"] = df_factures["paiement"].fillna("❌").replace("", "❌")


# Sauvegarde unique après toutes les modifications
df_factures.to_csv(CSV_FACTURES, index=False)



# Fonction pour charger la liste des IBANs une seule fois
def charger_liste_ibans():
    """Charge la liste des IBANs enregistrés."""
    return charger_dataframe(IBAN_LISTE_CSV).to_dict(orient="records")

def charger_en_arriere_plan():
    """Charge les nouvelles factures en arrière-plan."""
    thread = threading.Thread(target=verifier_nouvelles_factures)
    thread.daemon = True  # Le thread se terminera quand le programme principal se termine
    thread.start()



# Chemin du dossier des icônes
ICONS_PATH = "/Users/vincentperreard/script contrats/icons/"

def charger_icone(nom_fichier, taille=(18, 18)):  # Taille ajustable
    """Charge une icône PNG et la redimensionne pour l'affichage."""
    chemin_complet = os.path.join(ICONS_PATH, nom_fichier)
    try:
        image = Image.open(chemin_complet)
        image = image.resize(taille, Image.LANCZOS)  # Redimensionne proprement
        return ImageTk.PhotoImage(image)
    except Exception as e:
        print(f"⚠️ Erreur lors du chargement de l'icône {nom_fichier}: {e}")
        return None

def lister_fichiers_factures():
    """Liste tous les fichiers PDF et images, y compris ceux dans les sous-dossiers."""
    fichiers = []
    print(f"\n🔍 DEBUG - LISTAGE DES FICHIERS")
    print(f"📂 Recherche dans le dossier: {DOSSIER_FACTURES}")
    
    if not os.path.exists(DOSSIER_FACTURES):
        print(f"❌ ERREUR: Le dossier {DOSSIER_FACTURES} n'existe pas!")
        return fichiers
        
    for root, dirs, files in os.walk(DOSSIER_FACTURES):
        print(f"📁 Sous-dossier: {os.path.relpath(root, DOSSIER_FACTURES)}")
        for file in files:
            if file.lower().endswith((".pdf", ".jpg", ".jpeg", ".png")):
                rel_path = os.path.relpath(os.path.join(root, file), DOSSIER_FACTURES)
                fichiers.append(rel_path)
                print(f"   📄 Trouvé: {rel_path}")
    
    print(f"✅ Total: {len(fichiers)} fichiers trouvés.")
    return fichiers

def lire_facture(chemin_fichier):
    """Lit une facture (PDF ou image)."""
    if chemin_fichier.lower().endswith((".jpg", ".jpeg", ".png")):
        return lire_facture_image(chemin_fichier)
    else:
        return lire_facture_pdf(chemin_fichier)


def lire_facture_pdf(chemin_fichier):
    """
    Extrait le texte d'un fichier PDF à l'aide de pdfplumber.
    """
    try:
        with pdfplumber.open(chemin_fichier) as pdf:
            texte = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        return texte
    except Exception as e:
        return f"Erreur lors de la lecture du fichier {chemin_fichier} : {str(e)}"

def lire_facture_image(chemin_fichier):
    """Extrait le texte d'une image avec OCR (Tesseract)."""
    try:
        with Image.open(chemin_fichier) as img:
            img.verify()  # Vérifie si l’image est valide
    except (IOError, SyntaxError) as e:
        return f"Erreur : fichier corrompu ou format inconnu ({chemin_fichier})"
    try:
        image = Image.open(chemin_fichier).convert("L")  # Convertir en niveaux de gris
        image = image.point(lambda x: 0 if x < 200 else 255)  # Améliorer le contraste

        texte = pytesseract.image_to_string(image, lang="fra")  # OCR en français
        print(f"📝 Texte extrait (brut) : {repr(texte[:200])}")  # Vérifier le format du texte brut

        if not texte.strip():
            return "⚠️ Aucun texte détecté (OCR faible)"
        return texte
    except Exception as e:
        return f"Erreur OCR {chemin_fichier} : {str(e)}"


def loupe_on(event):
    """Active la loupe en créant une fenêtre flottante qui s'affiche où la souris clique."""
    global loupe_window, loupe_canvas

    if loupe_window is None:
        loupe_window = tk.Toplevel(root)
        loupe_window.overrideredirect(True)  # Supprime la barre de titre
        loupe_window.resizable(False, False)
        loupe_canvas = tk.Canvas(loupe_window, width=200, height=200, bg="black")
        loupe_canvas.pack()

    update_loupe(event)
    
def loupe_off(event):
    """Désactive la loupe et ferme la fenêtre flottante."""
    global loupe_window
    if loupe_window:
        loupe_window.destroy()
        loupe_window = None

def update_loupe(event):
    """Met à jour la loupe pour qu'elle s'affiche dans une fenêtre flottante en suivant la souris."""
    if apercu_img and loupe_window:
        zoom_size = 100  # Taille de la zone capturée
        zoom_factor = 2  # Facteur de zoom
        x, y = event.x, event.y

        img_width, img_height = apercu_img.size
        canvas_width, canvas_height = apercu_canvas.winfo_width(), apercu_canvas.winfo_height()

        # Calcul des coordonnées réelles sur l'image
        img_x = int(x * img_width / canvas_width)
        img_y = int(y * img_height / canvas_height)

        # Zone à capturer
        x1, y1, x2, y2 = img_x - zoom_size // 2, img_y - zoom_size // 2, img_x + zoom_size // 2, img_y + zoom_size // 2

        # Création d'une image zoomée avec fond noir pour éviter les artefacts
        zoom_area = Image.new("RGB", (zoom_size, zoom_size), "black")

        # Ajuster les bords si la zone capturée dépasse l'image
        crop_x1, crop_y1, crop_x2, crop_y2 = max(0, x1), max(0, y1), min(img_width, x2), min(img_height, y2)
        cropped_part = apercu_img.crop((crop_x1, crop_y1, crop_x2, crop_y2))

        paste_x, paste_y = max(0, -x1), max(0, -y1)
        zoom_area.paste(cropped_part, (paste_x, paste_y))

        # Appliquer le zoom
        zoom_area = zoom_area.resize((zoom_size * zoom_factor, zoom_size * zoom_factor), Image.LANCZOS)
        zoom_tk = ImageTk.PhotoImage(zoom_area)

        # Positionner la fenêtre de la loupe près de la souris
        loupe_window.geometry(f"200x200+{event.x_root-100}+{event.y_root-100}")
        loupe_canvas.delete("all")
        loupe_canvas.create_image(100, 100, image=zoom_tk, anchor="center")
        loupe_canvas.image = zoom_tk  # Empêche la suppression par le garbage collector

def afficher_apercu_fichier(chemin_fichier):
    """Affiche un aperçu d'un fichier PDF ou image."""
    global apercu_img, img_tk  # Garde en mémoire l'image pour éviter la suppression par le garbage collector

    try:
        if chemin_fichier.lower().endswith((".png", ".jpg", ".jpeg")):
            img = Image.open(chemin_fichier)
        elif chemin_fichier.lower().endswith(".pdf"):
            images = convert_from_path(chemin_fichier, first_page=1, last_page=1, dpi=150)
            img = images[0]

        canvas_width = apercu_canvas.winfo_width()
        canvas_height = apercu_canvas.winfo_height()
        img = img.resize((canvas_width, canvas_height), Image.LANCZOS)
        apercu_img = img  # Garde l'image originale pour le zoom
        img_tk = ImageTk.PhotoImage(img)
        
        apercu_canvas.create_image(canvas_width//2, canvas_height//2, anchor="center", image=img_tk)

    except Exception as e:
        print(f"Erreur lors de l'aperçu du fichier : {e}")

def update_apercu(event=None):
    """Met à jour l'aperçu du fichier sélectionné lors du redimensionnement."""
    selected_item = tree.selection()
    if selected_item:
        fichier = tree.item(selected_item, "values")[5]  # Colonne "Fichier"
        chemin_fichier = os.path.join(DOSSIER_FACTURES, fichier)
        afficher_apercu_fichier(chemin_fichier)

def on_facture_selection(event):
    """Affiche l’aperçu du fichier sélectionné lors de la sélection dans le tableau."""
    selected_item = tree.selection()
    if selected_item:
        fichier = tree.item(selected_item, "values")[5]  # Index 5 = colonne "Fichier"
        chemin_fichier = os.path.join(DOSSIER_FACTURES, fichier)
        afficher_apercu_fichier(chemin_fichier)



# Configure ta clé API (idéalement via une variable d'environnement)
openai.api_key = credentials['api_key']


def analyser_facture_api(texte):
    """
    Envoie le texte d'une facture à l'API OpenAI pour extraire les informations clés.
    Inclut des instructions pour différencier clients et fournisseurs.
    """
    # Debug : Vérifier le texte extrait
    print(f"🔍 Texte envoyé (longueur {len(texte)}): {texte[:200]}...")
    
    prompt = (
        "Tu es un expert en extraction d'informations sur des factures.\n"
        "Réponds uniquement par un objet JSON sans texte supplémentaire.\n"
        "Extrait les informations suivantes sous format JSON :\n"
        "- date_facture (format JJ-MM-AAAA)\n"
        "- fournisseur\n"
        "- montant (nombre en euros)\n"
        "- IBAN\n"
        "- reference_facture\n\n"
        "IMPORTANT: GROUPEMENT DES ANESTHÉSISTES DE LA CLINIQUE MATHILDE, GROUPEMENT ANESTHESISTES, SELARL DES ANESTHÉSISTES DE LA CLINIQUE MATHILDE, "
        "SPFPL HOLDING ANESTHÉSIE MATHILDE, PERREARD Vincent, NAFEH Samer ou d'autres médecins anesthésistes "
        "ne sont PAS des fournisseurs, mais des clients. "
        "Si tu identifies ces entités comme expéditeurs de la facture, cherche le véritable fournisseur. "
        "Si tu ne trouves pas de fournisseur clair, indique 'N/A'.\n\n"
        "Texte de la facture :\n" + texte
    )
    
    # Debug : Afficher le prompt complet
    print("🔎 Prompt envoyé à l'API :\n", prompt)
    
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "Tu es un assistant expert dans l'analyse de factures."},
            {"role": "user", "content": prompt}
        ],
        temperature=0,
        max_tokens=300
    )
    
    # Traitement de la réponse
    try:
        # Obtenir la réponse brute
        raw_response = response.choices[0].message.content.strip()
        print("📝 Réponse brute de l'API :", repr(raw_response))
        
        # Nettoyage 1: Supprimer les backticks et les identifiants de code
        if raw_response.startswith("```json"):
            raw_response = raw_response.replace("```json", "", 1)
        if raw_response.endswith("```"):
            raw_response = raw_response.replace("```", "", 1)
        
        # Nettoyage 2: Remplacer les caractères de contrôle par des espaces
        cleaned_response = ""
        for char in raw_response:
            if ord(char) < 32 and char not in '\n\r\t':
                cleaned_response += ' '  # Remplacer par un espace
            else:
                cleaned_response += char
        
        # Parsing du JSON nettoyé
        result = json.loads(cleaned_response)
        print(f"📋 Résultat JSON parsé: {result}")
        
        return result
    except Exception as e:
        print("❌ Erreur de parsing JSON :", e)
        return {
            "date_facture": "N/A",
            "fournisseur": "N/A", 
            "montant": "N/A",
            "iban": "N/A",
            "reference_facture": "N/A"
        }  # Retourner des valeurs par défaut en cas d'erreur
    
    # Vérification supplémentaire pour les fournisseurs incorrects
    fournisseurs_incorrects = [
        "GROUPEMENT DES ANESTHÉSISTES", "SELARL DES ANESTHÉSISTES", 
        "ANESTHÉSISTES DE LA CLINIQUE MATHILDE", "SPFPL HOLDING ANESTHÉSIE", 
        "PERREARD", "NAFEH", "BEGHIN", "DELPECH", "BRILLE", "ELIE", "SACUTO", "ANESTHESISTE", "ANESTHESISTES", "GROUPEMENT ANESTHESISTES"
    ]
    
    if "fournisseur" in result:
        fournisseur = result["fournisseur"]
        if fournisseur and any(nom.lower() in fournisseur.lower() for nom in fournisseurs_incorrects):
            print(f"⚠️ Fournisseur incorrectement identifié: {fournisseur}. Changement pour 'N/A'")
            result["fournisseur"] = "N/A"
    
    return result


def analyser_facture_complete(chemin_fichier, utiliser_cache=True):
    """
    Analyse une facture en extrayant son texte depuis un PDF et en l'envoyant à l'API OpenAI.
    """
    global df_ibans
    
    print(f"🔍 Fichier analysé : {chemin_fichier}")    
    
    # Vérifier si la facture est dans la base de données
    if utiliser_cache:
        db = charger_base_donnees()
        try:
            hash_fichier = calculer_hash_fichier(chemin_fichier)
            if hash_fichier:
                # Créer un dictionnaire inversé pour recherche par hash
                hash_to_id = {info["hash"]: id_facture for id_facture, info in db["factures"].items()}
                
                if hash_fichier in hash_to_id:
                    # Récupérer les infos déjà extraites
                    id_facture = hash_to_id[hash_fichier]
                    info_extraite = db["factures"][id_facture]
                    
                    # Si les infos nécessaires sont présentes, les utiliser
                    cache_info = {}
                    for k in ["date_facture", "fournisseur", "montant", "iban", "reference_facture"]:
                        # Vérifier chaque clé individuellement
                        if k in info_extraite and info_extraite[k] not in [None, "", "N/A"]:
                            cache_info[k] = info_extraite[k]
                        else:
                            # Si manquante ou N/A, l'inclure quand même pour complétion
                            cache_info[k] = "N/A"
                    
                    print(f"✅ Utilisation du cache pour {os.path.basename(chemin_fichier)}")
                    # Ajouter un debug pour voir ce qu'on récupère du cache
                    print(f"📋 Données du cache: {cache_info}")
                    
                    return cache_info
        except Exception as e:
            print(f"⚠️ Erreur lors de la vérification du cache: {e}")
    
    # Si la facture n'est pas dans le cache ou si les infos sont incomplètes, l'analyser
    texte = lire_facture(chemin_fichier)
    result_api = analyser_facture_api(texte)
    result_api = {k.lower(): v for k, v in result_api.items()}
    
    # Débogage: voir ce que l'API a renvoyé
    print(f"📋 Résultat API: {result_api}")
    
    
    # Vérifier si le fournisseur existe déjà dans notre base
    if "fournisseur" in result_api and result_api["fournisseur"] and result_api["fournisseur"] != "N/A":
        detected_fournisseur = result_api["fournisseur"]
        detected_iban = result_api.get("iban", "N/A")
        
        # Chercher une correspondance approximative dans la liste des fournisseurs
        found_match = False
        for index, row in df_ibans.iterrows():
            existing_fournisseur = str(row["fournisseur"]).strip()
            if (existing_fournisseur.lower() in detected_fournisseur.lower() or 
                detected_fournisseur.lower() in existing_fournisseur.lower()):
                result_api["fournisseur"] = existing_fournisseur
                # Récupérer aussi l'IBAN si disponible
                if pd.notna(row["IBAN"]) and row["IBAN"] != "nan":
                    result_api["iban"] = row["IBAN"]
                found_match = True
                break
        
        # Si c'est un nouveau fournisseur valide avec un IBAN, l'ajouter à la liste
        if not found_match and detected_iban != "N/A" and detected_fournisseur != "N/A":
            print(f"🆕 Nouveau fournisseur détecté: {detected_fournisseur} avec IBAN: {detected_iban}")
            
            # Vérifier si l'IBAN n'existe pas déjà pour un autre fournisseur
            if not df_ibans[df_ibans["IBAN"] == detected_iban].empty:
                print(f"⚠️ IBAN {detected_iban} existe déjà, pas d'ajout automatique.")
            else:
                # Ajouter aux ibans connus
                new_row = pd.DataFrame({"fournisseur": [detected_fournisseur], "IBAN": [detected_iban]})
                df_ibans = pd.concat([df_ibans, new_row], ignore_index=True)
                df_ibans.to_csv(IBAN_LISTE_CSV, index=False)
                print(f"✅ Fournisseur automatiquement ajouté à la liste")
    
    # S'assurer que toutes les clés attendues sont présentes
    for key in ["date_facture", "fournisseur", "montant", "iban", "reference_facture"]:
        if key not in result_api or not result_api[key]:
            result_api[key] = "N/A"
    
    return result_api

def normaliser_fournisseur(nom):
    """Normalise le nom d'un fournisseur pour faciliter la comparaison."""
    if pd.isna(nom) or not nom:
        return ""
    
    # Convertir en minuscules et supprimer les espaces superflus
    nom = str(nom).lower().strip()
    
    # Supprimer les articles et mots communs
    mots_a_ignorer = ["le", "la", "les", "l'", "de", "des", "du", "et", "sarl", "sa", "sas", "eurl", "selarl"]
    for mot in mots_a_ignorer:
        nom = nom.replace(f" {mot} ", " ")
    
    # Suppression des caractères spéciaux
    import re
    nom = re.sub(r'[^\w\s]', '', nom)
    
    # Supprimer les espaces multiples
    nom = re.sub(r'\s+', ' ', nom)
    
    return nom.strip()


def recharger_fichiers(forcer_reanalyse=False):

    global df_factures
    print("🔄 Rechargement des données depuis le CSV...")
    print(df_factures.head())  # Afficher les premières lignes pour vérification
    if forcer_reanalyse:
        print("🔄 Rechargement complet de toutes les factures...")
        df_factures = analyser_toutes_les_factures()  # 🔄 Recharger toutes les factures
    else:
        print("🔍 Vérification des nouvelles factures uniquement...")
        verifier_nouvelles_factures()  # 🔍 Vérifie seulement les nouvelles factures

    # Charger la base de données pour avoir accès aux informations sur les fichiers manquants
    db = charger_base_donnees()
    
    # Créer un dict des statuts de disponibilité
    disponibilite = {db["factures"][id_facture]["chemin"].split("/")[-1]: db["factures"][id_facture]["disponible"] 
                     for id_facture in db["factures"] if "chemin" in db["factures"][id_facture]}

    df_factures["date_facture"] = df_factures["date_facture"].astype(str)  # Convertir en string
    
    # Ajouter la colonne "disponible" si elle n'existe pas encore
    if "disponible" not in df_factures.columns:
        df_factures["disponible"] = True
    
    # Mettre à jour le statut de disponibilité
    for idx, row in df_factures.iterrows():
        fichier = row["fichier"]
        if fichier in disponibilite:
            df_factures.at[idx, "disponible"] = disponibilite[fichier]    
    

    # Ajouter la colonne "fournisseur" en fusionnant avec la liste des fournisseurs si elle n'existe pas
    if "fournisseur" not in df_factures.columns:
        df_ibans_local = pd.read_csv(IBAN_LISTE_CSV)
        df_factures = pd.merge(df_factures, df_ibans_local, left_on="iban", right_on="IBAN", how="left")
        df_factures["fournisseur"] = df_factures["fournisseur"].fillna("N/A")

    dates_probleme = df_factures[~df_factures["date_facture"].str.match(r"^\d{2}-\d{2}-\d{4}$", na=False)]

 
    # Appliquer le filtre année/mois
    annee_filtre = annee_var.get()
    mois_filtre = mois_var.get()
    
    # Extraire l'année et le mois des dates
    if "annee" not in df_factures.columns or "mois" not in df_factures.columns:
        df_factures["annee"] = df_factures["date_facture"].apply(
            lambda x: x.split("-")[2] if isinstance(x, str) and len(x.split("-")) == 3 else "N/A"
        )
        df_factures["mois"] = df_factures["date_facture"].apply(
            lambda x: x.split("-")[1] if isinstance(x, str) and len(x.split("-")) == 3 else "N/A"
        )

    if "paiement" not in df_factures.columns:
        df_factures["paiement"] = "❌"
    else:
        df_factures["paiement"] = df_factures["paiement"].fillna("❌").replace("", "❌")
        
    df_filtree = df_factures.copy()
    if annee_filtre and annee_filtre != "---":
        df_filtree = df_filtree[df_filtree["annee"] == annee_filtre]
    if mois_filtre and mois_filtre != "---":
        df_filtree = df_filtree[df_filtree["mois"] == mois_filtre]

    # Tri des factures par date (décroissant)
    df_filtree = df_filtree.sort_values(by=["date_facture"], ascending=False)

    # Supprimer toutes les entrées du tableau
    for item in tree.get_children():
        tree.delete(item)    

    # Réinsérer les nouvelles données dans le tableau
    for index, row in df_filtree.iterrows():
        paiement = row["paiement"] if pd.notna(row["paiement"]) and row["paiement"].strip() != "" else "❌"
        item_id = tree.insert("", "end", values=(
            row["date_facture"], 
            row["fournisseur"],  # Ajout de la colonne fournisseur
            row["montant"], 
            row["iban"], 
            row["reference_facture"], 
            row["fichier"], 
            paiement  # 🔥 On force la présence de la croix rouge si vide
        ))
        
        # Appliquer le style "disparu" si le fichier n'est plus disponible
        if not row.get("disponible", True):
            tree.item(item_id, tags=("disparu",))

    text_area.delete("1.0", "end")
    text_area.insert("end", f"🔄 Factures mises à jour pour {mois_filtre}/{annee_filtre}\n")
    root.update_idletasks()

    # Enregistrer toutes les modifications
    df_factures.to_csv(CSV_FACTURES, index=False)
    print("✅ Toutes les modifications ont été enregistrées.")
    
    try:
        mettre_a_jour_listes_filtres()  # Met à jour les options de filtres
    except NameError:
        pass  # Ignore si la fonction n'existe pas encore

def supprimer_accents(texte):
    return ''.join(c for c in unicodedata.normalize('NFD', texte) if unicodedata.category(c) != 'Mn')

def envoyer_virement(date, reference, iban, montant):
    """Génère un fichier XML de virement avec les données correctes."""
    try:
        # Extraction du nom du fournisseur associé à l'IBAN depuis le fichier liste_ibans.csv
        df_ibans_local = pd.read_csv(IBAN_LISTE_CSV)
        fournisseur_match = df_ibans_local[df_ibans_local["IBAN"] == iban]
        if fournisseur_match.empty:
            raise ValueError(f"IBAN {iban} non trouvé dans la liste des fournisseurs.")
        nom_fournisseur = fournisseur_match.iloc[0]["fournisseur"]
        nom_fournisseur = supprimer_accents(nom_fournisseur)

        virements = [{
            "beneficiaire": nom_fournisseur,  # Utilisation du nom du fournisseur après suppression des accents
            "iban": iban,
            "montant": float(montant.replace("€", "").replace(",", ".").strip()),
            "objet": reference  # Correction : ne mettre que la référence de la facture
        }]

        # Génération du fichier XML
        fichier_xml = generer_xml_virements(virements)
        print(f"Fichier XML généré : {fichier_xml}")  # Vérifier le chemin du fichier

        if fichier_xml:
            print("Lancement de Selenium pour envoyer le fichier XML...")  # Vérifier si cette ligne s'affiche
            envoyer_virement_vers_lcl(fichier_xml)
        else:
            print("Erreur : Le fichier XML n'a pas été généré.")
    except Exception as e:
        print(f"❌ Erreur lors de la génération du virement : {e}")

def analyser_toutes_les_factures():
    """
    Analyse toutes les factures du dossier et retourne un DataFrame.
    Utilise le pipeline complet (API OpenAI avec fallback local).
    """
    fichiers = lister_fichiers_factures()
    resultats = []

    for fichier in fichiers:
        chemin_fichier = os.path.join(DOSSIER_FACTURES, fichier)
        infos = analyser_facture_complete(chemin_fichier)
        infos["fichier"] = fichier
        resultats.append(infos)

    df_resultats = pd.DataFrame(resultats)
    return df_resultats


def afficher_contenu_factures():
    """Affiche le contenu brut des factures PDF pour analyse."""
    fichiers = lister_fichiers_factures()

    for fichier in fichiers:
        chemin_fichier = os.path.join(DOSSIER_FACTURES, fichier)
        texte = lire_facture_pdf(chemin_fichier)
        print(f"\n📄 Contenu de {fichier} :\n")
        print(texte[:1000])  # Afficher seulement les 1000 premiers caractères pour éviter trop de texte
        print("\n" + "-" * 80 + "\n")

def fermer():
    """Ferme proprement toutes les fenêtres Tkinter."""
    global loupe_window
    if loupe_window:
        loupe_window.destroy()
        loupe_window = None

    for window in root.winfo_children():
        window.destroy()  # Détruit toutes les fenêtres enfants

    root.destroy()  # Ferme complètement l'application

def modifier_paiement(event):
    """Ouvre une fenêtre pour modifier manuellement le statut du paiement (✔️ ou ❌)."""
    selected_item = tree.selection()
    if not selected_item:
        return

    item_id = selected_item[0]
    values = tree.item(item_id, "values")

    if not values:
        return

    fichier = values[5]  # Récupération du fichier
    paiement_actuel = values[5] if len(values) > 5 else "❌"  # Valeur actuelle du paiement

    # Création de la fenêtre pop-up
    popup = tk.Toplevel(root)
    popup.title("Modifier le paiement")
    popup.geometry("300x150")
    popup.transient(root)
    popup.grab_set()

    ttk.Label(popup, text="Modifier le statut du paiement :").pack(pady=10)

    def valider():
        """Valide le paiement et met à jour l'interface et le fichier CSV."""
        date_validation = datetime.now().strftime("%d-%m-%Y")  # Date actuelle
        statut_paiement = f"✔️ {date_validation}"

        tree.item(item_id, values=(values[0], values[1], values[2], values[3], values[4], values[5], statut_paiement))
        df_factures.loc[df_factures["fichier"] == fichier, "paiement"] = statut_paiement
        df_factures.to_csv(CSV_FACTURES, index=False)
        popup.destroy()

    def invalider():
        """Invalide le paiement et met à jour l'interface et le fichier CSV."""
        tree.item(item_id, values=(values[0], values[1], values[2], values[3], values[4], "❌"))
        df_factures.loc[df_factures["fichier"] == fichier, "paiement"] = "❌"
        df_factures.to_csv(CSV_FACTURES, index=False)
        popup.destroy()

    btn_valider = ttk.Button(popup, text="✔️ Valider", command=valider)
    btn_valider.pack(pady=5)

    btn_invalider = ttk.Button(popup, text="❌ Invalider", command=invalider)
    btn_invalider.pack(pady=5)

def choisir_fournisseur_iban(event, col):
    """Gère le double-clic sur une cellule fournisseur ou IBAN."""
    global df_factures, df_ibans, CSV_FACTURES
    
    # DEBUG: Afficher l'état initial des DataFrames
    print("\n==== DÉBUT DEBUG: ÉTAT INITIAL ====")
    print(f"Nombre de lignes df_factures: {len(df_factures)}")
    print(f"Colonnes df_factures: {df_factures.columns.tolist()}")
    print("==== FIN ÉTAT INITIAL ====\n")
    
    row_id = tree.focus()
    if not row_id:
        print("❌ DEBUG: Aucune ligne sélectionnée")
        return
    
    # DEBUG: Identifier clairement la ligne sélectionnée
    row_index = tree.index(row_id)
    print(f"🔍 DEBUG: Ligne sélectionnée ID={row_id}, Index={row_index}")
    
    row_data = tree.item(row_id, "values")
    print(f"🔍 DEBUG: Données de la ligne: {row_data}")
    
    # Vérifier que le fichier existe dans df_factures
    fichier = row_data[5] if len(row_data) > 5 else None
    if fichier:
        matching_rows = df_factures[df_factures["fichier"] == fichier]
        print(f"🔍 DEBUG: Lignes correspondant au fichier '{fichier}': {len(matching_rows)}")
        if matching_rows.empty:
            print(f"⚠️ DEBUG: Fichier '{fichier}' non trouvé dans df_factures!")
    
    selection_window = tk.Toplevel(root)
    selection_window.title("Modifier fournisseur / IBAN")
    selection_window.geometry("500x350")
    selection_window.transient(root)
    selection_window.grab_set()
    
    # Cadre principal
    main_frame = ttk.Frame(selection_window, padding=10)
    main_frame.pack(fill="both", expand=True)
    
    # Valeurs actuelles
    current_fournisseur = row_data[1] if len(row_data) > 1 else ""
    current_iban = row_data[3] if len(row_data) > 3 else ""
    
    print(f"🔍 DEBUG: Valeurs actuelles - Fournisseur: '{current_fournisseur}', IBAN: '{current_iban}'")
    
    # 1. Section pour l'édition manuelle
    edit_frame = ttk.LabelFrame(main_frame, text="Édition manuelle du fournisseur")
    edit_frame.pack(fill="x", pady=10, padx=5)
    
    ttk.Label(edit_frame, text="Fournisseur:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    fournisseur_var = tk.StringVar(value=current_fournisseur)
    fournisseur_entry = ttk.Entry(edit_frame, textvariable=fournisseur_var, width=40)
    fournisseur_entry.grid(row=0, column=1, padx=5, pady=5)
    
    ttk.Label(edit_frame, text="IBAN:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    iban_var = tk.StringVar(value=current_iban)
    iban_entry = ttk.Entry(edit_frame, textvariable=iban_var, width=40)
    iban_entry.grid(row=1, column=1, padx=5, pady=5)
    
    # 2. Section pour choisir un fournisseur existant dans la liste
    list_frame = ttk.LabelFrame(main_frame, text="Ou choisir un fournisseur existant")
    list_frame.pack(fill="x", pady=10, padx=5)
    
    # Récupérer et trier la liste des fournisseurs existants
    fournisseurs_list = df_ibans["fournisseur"].dropna().unique().tolist()
    fournisseurs_list = sorted([f for f in fournisseurs_list if f and str(f).strip()])
    
    print(f"🔍 DEBUG: Liste des fournisseurs disponibles: {fournisseurs_list}")
    
    ttk.Label(list_frame, text="Fournisseur:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    
    # Combobox pour sélectionner un fournisseur existant
    fournisseur_combo_var = tk.StringVar()
    fournisseur_combo = ttk.Combobox(list_frame, textvariable=fournisseur_combo_var, values=fournisseurs_list, width=40)
    fournisseur_combo.grid(row=0, column=1, padx=5, pady=5)
    
    # Mettre à jour les champs quand un fournisseur est sélectionné dans la liste
    def on_combo_select(event):
        selected_fournisseur = fournisseur_combo_var.get()
        if selected_fournisseur:
            print(f"🔍 DEBUG: Fournisseur sélectionné dans la liste: '{selected_fournisseur}'")
            fournisseur_var.set(selected_fournisseur)
            # Trouver l'IBAN correspondant
            matching_rows = df_ibans[df_ibans["fournisseur"] == selected_fournisseur]
            if not matching_rows.empty and pd.notna(matching_rows.iloc[0]["IBAN"]):
                iban_var.set(matching_rows.iloc[0]["IBAN"])
                print(f"🔍 DEBUG: IBAN trouvé: '{matching_rows.iloc[0]['IBAN']}'")
    
    fournisseur_combo.bind("<<ComboboxSelected>>", on_combo_select)
    
    # 3. Boutons d'action
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill="x", pady=20)
    
    def valider():
        """Valide les changements et met à jour la facture."""
        global df_ibans
        
        fournisseur_selectionne = fournisseur_var.get().strip()
        iban_selectionne = iban_var.get().strip()
        
        print(f"\n==== DÉBUT DEBUG: VALIDATION ====")
        print(f"🔍 Fournisseur sélectionné: '{fournisseur_selectionne}'")
        print(f"🔍 IBAN sélectionné: '{iban_selectionne}'")
        
        if not fournisseur_selectionne:
            messagebox.showwarning("Avertissement", "Veuillez entrer un nom de fournisseur.")
            return
        
        # APPROCHE 1: Mise à jour par ID exact
        fichier = row_data[5] if len(row_data) > 5 else None
        if not fichier:
            print("❌ DEBUG: Pas de fichier identifié, impossible de mettre à jour")
            return
            
        print(f"🔍 DEBUG: Recherche du fichier '{fichier}' dans df_factures")
        
        # Créer une copie du DataFrame avant modification
        df_before = df_factures.copy()
        
        # Vérifier combien de lignes seraient affectées
        matching_rows = df_factures[df_factures["fichier"] == fichier]
        print(f"🔍 DEBUG: Nombre de lignes correspondant au fichier: {len(matching_rows)}")
        
        # CRITIQUE: Utiliser une condition très restrictive pour ne modifier que cette ligne
        # Trouvons un ID unique pour cette ligne
        unique_condition = (df_factures["fichier"] == fichier)
        
        # Si d'autres colonnes ont des valeurs qui peuvent aider à identifier la ligne
        if current_fournisseur:
            # On ne fait pas ça car cela peut changer
            pass
            
        # Compte combien de lignes correspondent à la condition
        rows_to_update = df_factures[unique_condition]
        print(f"🔍 DEBUG: Nombre de lignes à mettre à jour: {len(rows_to_update)}")
        
        # Mise à jour SEULEMENT si on a identifié UNE SEULE ligne
        if len(rows_to_update) == 1:
            df_factures.loc[unique_condition, "fournisseur"] = fournisseur_selectionne
            if iban_selectionne:
                df_factures.loc[unique_condition, "iban"] = iban_selectionne
                
            # Vérifier combien de lignes ont été modifiées
            changed_rows = df_factures[(df_factures["fournisseur"] != df_before["fournisseur"]) | 
                                       (df_factures["iban"] != df_before["iban"])]
            print(f"🔍 DEBUG: {len(changed_rows)} lignes ont été modifiées")
            
            # Sauvegarder dans le CSV
            df_factures.to_csv(CSV_FACTURES, index=False)
            print("✅ DEBUG: DataFrame sauvegardé dans CSV")
        else:
            print(f"❌ DEBUG: Condition trop ambiguë, trouverait {len(rows_to_update)} lignes")
            messagebox.showerror("Erreur", f"Impossible d'identifier de façon unique la facture à mettre à jour.")
            return
            
        # Mettre à jour le tableau visuel
        new_values = list(row_data)
        new_values[1] = fournisseur_selectionne
        if iban_selectionne:
            new_values[3] = iban_selectionne
        tree.item(row_id, values=new_values)
        print(f"✅ DEBUG: Affichage mis à jour avec fournisseur='{fournisseur_selectionne}', IBAN='{iban_selectionne}'")
            
        # Vérifier si c'est un nouveau fournisseur à ajouter à la liste
        if fournisseur_selectionne not in df_ibans["fournisseur"].values and iban_selectionne:
            new_row = pd.DataFrame({"fournisseur": [fournisseur_selectionne], "IBAN": [iban_selectionne]})
            df_ibans = pd.concat([df_ibans, new_row], ignore_index=True)
            df_ibans.to_csv(IBAN_LISTE_CSV, index=False)
            print(f"✅ DEBUG: Nouveau fournisseur ajouté: {fournisseur_selectionne}")
        
        print("==== FIN DEBUG: VALIDATION ====\n")
        text_area.insert("end", f"\n✅ Fournisseur mis à jour: {fournisseur_selectionne}\n")
        selection_window.destroy()
    
    ttk.Button(button_frame, text="Valider", command=valider).pack(side="left", padx=5)
    ttk.Button(button_frame, text="Annuler", command=selection_window.destroy).pack(side="left", padx=5)

def gerer_fournisseurs():
    """Ouvre une fenêtre pour gérer la liste des fournisseurs et IBANs."""
    global df_ibans  # S'assure d'utiliser la liste actuelle des IBANs

    # Charger la liste des IBANs depuis le fichier
    df_ibans = pd.read_csv(IBAN_LISTE_CSV)

    # Création de la fenêtre
    fournisseurs_window = tk.Toplevel(root)
    fournisseurs_window.title("Gestion des fournisseurs")
    fournisseurs_window.geometry("500x350")  # Ajuste la taille
    fournisseurs_window.transient(root)
    fournisseurs_window.grab_set()

    # Créer un cadre pour la liste des fournisseurs
    frame = ttk.Frame(fournisseurs_window, padding=10)
    frame.pack(fill="both", expand=True)

    # Création de la table pour afficher les fournisseurs et IBANs
    columns = ("fournisseur", "IBAN")
    tree_fournisseurs = ttk.Treeview(frame, columns=columns, show="headings")
    tree_fournisseurs.heading("fournisseur", text="fournisseur")
    tree_fournisseurs.heading("IBAN", text="IBAN")
    tree_fournisseurs.column("fournisseur", width=200, anchor="w")
    tree_fournisseurs.column("IBAN", width=250, anchor="w")

    # Insérer les fournisseurs actuels dans le tableau
    for _, row in df_ibans.iterrows():
        if "fournisseur" not in df_factures.columns:
            print("Erreur : colonne 'fournisseur' non trouvée dans df_factures")
            print("Colonnes actuelles :", df_factures.columns)
            return 
        tree_fournisseurs.insert("", "end", values=(row["fournisseur"], row["IBAN"]))

    tree_fournisseurs.pack(fill="both", expand=True, padx=5, pady=5)

    # Fonction pour modifier un fournisseur/IBAN
    def modifier_fournisseur():
        """Ouvre une fenêtre pour modifier un fournisseur ou son IBAN."""
        selected_item = tree_fournisseurs.selection()
        if not selected_item:
            messagebox.showwarning("Sélection requise", "Veuillez sélectionner un fournisseur à modifier.")
            return

        item_id = selected_item[0]
        values = tree_fournisseurs.item(item_id, "values")

        fournisseur_initial = values[0]
        iban_initial = values[1]

        # Fenêtre de modification
        modif_window = tk.Toplevel(fournisseurs_window)
        modif_window.title("Modifier fournisseur")
        modif_window.geometry("400x150")
        modif_window.transient(fournisseurs_window)
        modif_window.grab_set()

        frame_modif = ttk.Frame(modif_window, padding=10)
        frame_modif.pack(fill="both", expand=True)

        ttk.Label(frame_modif, text="fournisseur :").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        fournisseur_var = tk.StringVar(value=fournisseur_initial)
        fournisseur_entry = ttk.Entry(frame_modif, textvariable=fournisseur_var, width=40)
        fournisseur_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frame_modif, text="IBAN :").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        iban_var = tk.StringVar(value=iban_initial)
        iban_entry = ttk.Entry(frame_modif, textvariable=iban_var, width=40)
        iban_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        def valider_modif():
            """Enregistre les modifications du fournisseur/IBAN."""
            new_fournisseur = fournisseur_var.get().strip()
            new_iban = iban_var.get().strip()

            if not new_fournisseur or not new_iban:
                messagebox.showwarning("Champs vides", "Veuillez renseigner le fournisseur et l'IBAN.")
                return

            # Modifier dans le DataFrame
            df_ibans.loc[df_ibans["IBAN"] == iban_initial, ["fournisseur", "IBAN"]] = [new_fournisseur, new_iban]
            df_ibans.to_csv(IBAN_LISTE_CSV, index=False)

            # Modifier dans l'affichage
            tree_fournisseurs.item(item_id, values=(new_fournisseur, new_iban))
            modif_window.destroy()

        ttk.Button(frame_modif, text="Enregistrer", command=valider_modif).grid(row=2, column=0, padx=5, pady=10)
        ttk.Button(frame_modif, text="Annuler", command=modif_window.destroy).grid(row=2, column=1, padx=5, pady=10)

    # Fonction pour supprimer un fournisseur/IBAN
    def supprimer_fournisseur():
        """Supprime un fournisseur et son IBAN."""
        selected_item = tree_fournisseurs.selection()
        if not selected_item:
            messagebox.showwarning("Sélection requise", "Veuillez sélectionner un fournisseur à supprimer.")
            return

        item_id = selected_item[0]
        values = tree_fournisseurs.item(item_id, "values")

        confirmation = messagebox.askyesno("Confirmation", f"Voulez-vous vraiment supprimer l'IBAN de {values[0]} ?")

        if confirmation:
            # Supprimer du DataFrame
            df_ibans.drop(df_ibans[df_ibans["IBAN"] == values[1]].index, inplace=True)
            df_ibans.to_csv(IBAN_LISTE_CSV, index=False)

            # Supprimer de l'affichage
            tree_fournisseurs.delete(item_id)

    # Fonction pour ajouter un nouveau fournisseur/IBAN
    def ajouter_fournisseur():
        """Ouvre une fenêtre pour ajouter un nouveau fournisseur et son IBAN."""
        global df_ibans
        df_ibans = pd.read_csv(IBAN_LISTE_CSV)
        ajout_window = tk.Toplevel(fournisseurs_window)
        ajout_window.title("Ajouter fournisseur")
        ajout_window.geometry("400x150")
        ajout_window.transient(fournisseurs_window)
        ajout_window.grab_set()

        frame_ajout = ttk.Frame(ajout_window, padding=10)
        frame_ajout.pack(fill="both", expand=True)

        ttk.Label(frame_ajout, text="fournisseur :").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        fournisseur_var = tk.StringVar()
        fournisseur_entry = ttk.Entry(frame_ajout, textvariable=fournisseur_var, width=40)
        fournisseur_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frame_ajout, text="IBAN :").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        iban_var = tk.StringVar()
        iban_entry = ttk.Entry(frame_ajout, textvariable=iban_var, width=40)
        iban_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        def valider_ajout():
            """Ajoute un fournisseur avec son IBAN."""
            fournisseur = fournisseur_var.get().strip()
            iban = iban_var.get().strip()

            if not fournisseur or not iban:
                messagebox.showwarning("Champs vides", "Veuillez renseigner le fournisseur et l'IBAN.")
                return

            # Ajouter au DataFrame
            global df_ibans
            new_row = pd.DataFrame([{"fournisseur": fournisseur, "IBAN": iban}])
            df_ibans = pd.concat([df_ibans, new_row], ignore_index=True)
            df_ibans.to_csv(IBAN_LISTE_CSV, index=False)

            # Ajouter à l'affichage
            tree_fournisseurs.insert("", "end", values=(fournisseur, iban))
            ajout_window.destroy()

        ttk.Button(frame_ajout, text="Enregistrer", command=valider_ajout).grid(row=2, column=0, padx=5, pady=10)
        ttk.Button(frame_ajout, text="Annuler", command=ajout_window.destroy).grid(row=2, column=1, padx=5, pady=10)

    # Ajouter les boutons d'action
    ttk.Button(frame, text="Modifier", command=modifier_fournisseur).pack(side="left", padx=5, pady=5)
    ttk.Button(frame, text="Supprimer", command=supprimer_fournisseur).pack(side="left", padx=5, pady=5)
    ttk.Button(frame, text="Ajouter", command=ajouter_fournisseur).pack(side="left", padx=5, pady=5)
    ttk.Button(frame, text="Retour", command=fournisseurs_window.destroy).pack(side="left", padx=5,pady=5)

def ouvrir_dossier_fichier(event):
    """Ouvre le dossier contenant le fichier sélectionné dans le Finder."""
    selected_item = tree.selection()
    if not selected_item:
        return
        
    item_id = selected_item[0]
    fichier = tree.item(item_id, "values")[5]  # Colonne "Fichier"
    chemin_complet = os.path.join(DOSSIER_FACTURES, fichier)
    
    # Vérifier si le fichier existe
    if not os.path.exists(chemin_complet):
        messagebox.showwarning("Fichier introuvable", f"Le fichier {fichier} n'existe pas dans le dossier spécifié.")
        return
        
    # Révéler le fichier dans le Finder (macOS)
    subprocess.run(["open", "-R", chemin_complet])

def on_tree_double_click(event):
    """Gère les doubles-clics sur le tableau en fonction de la colonne cliquée."""
    column_id = tree.identify_column(event.x)  # Identifie la colonne cliquée
    column_index = int(column_id[1:]) - 1  # Convertit '#6' en 5 (index 0-based)
    
    # Récupérer la ligne sélectionnée
    selected_item = tree.selection()
    if not selected_item:
        return
    
    # Colonne "Fichier" (vérifiez l'index correct)
    if column_index == 5:  # L'index 5 correspond à la 6ème colonne (Fichier)
        fichier = tree.item(selected_item[0], "values")[5]
        chemin_complet = os.path.join(DOSSIER_FACTURES, fichier)
        
        # Vérifier si le fichier existe
        if os.path.exists(chemin_complet):
            print(f"Ouverture du fichier dans le Finder: {chemin_complet}")
            subprocess.run(["open", "-R", chemin_complet])
        else:
            messagebox.showwarning("Fichier introuvable", f"Le fichier {fichier} n'existe pas.")
        return
    
    # Colonne "Paiement"
    elif column_index == 6:  # L'index 6 correspond à la 7ème colonne (Paiement)
        modifier_paiement(event)
        return
    
    # Pour toutes les autres colonnes, utilisez la fonction de modification générale
    else:
        modifier_valeur(event)


# Interface graphique avec Tkinter
root = tk.Tk()
root.title("Analyse des Factures")

# Charge les données existantes sans lancer d'analyse
initialiser_systeme()

# Configuration de la zone de gauche (tableau, zone de texte et boutons)
frame_left = ttk.Frame(root)
frame_left.pack(side="left", fill="both", expand=True, padx=10, pady=10)

# Définition des colonnes
tree = ttk.Treeview(frame_left, columns=("Date","fournisseur", "Montant", "IBAN", "Référence", "Fichier", "paiement"), show="headings")
tree.heading("Date", text="Date")
tree.heading("fournisseur", text="fournisseur")
tree.column("fournisseur", width=150, anchor="w")
tree.heading("Montant", text="Montant")
tree.heading("IBAN", text="IBAN")
tree.heading("Référence", text="Référence")
tree.heading("Fichier", text="Fichier")
tree.heading("paiement", text="paiement")
tree.column("paiement", width=100, anchor="center")
# **Ajout de la liaison après la création de `tree`**
tree.bind("<<TreeviewSelect>>", on_facture_selection)
tree.bind("<Double-1>", on_tree_double_click)


# Ajustement des largeurs
tree.column("Date", width=100, anchor="center")
tree.column("Montant", width=100, anchor="center")
tree.column("IBAN", width=240, anchor="center")
tree.column("Référence", width=150, anchor="w")
tree.column("Fichier", width=200, anchor="w")
tree.pack(fill="both", expand=True)

# Configuration du style pour les fichiers disparus
style = ttk.Style()
tree.tag_configure("disparu", foreground="orange")

# Zone de texte pour afficher les erreurs et infos extraites
text_area = scrolledtext.ScrolledText(frame_left, height=10, wrap="word")
text_area.pack(side="top", fill="x", padx=10, pady=5)

frame_apercu = ttk.Frame(root)
frame_apercu.pack(side="right", fill="y", padx=10, pady=10)

apercu_canvas = tk.Canvas(frame_apercu, width=600, height=800, bg="white")
apercu_canvas.pack()

apercu_canvas.bind("<ButtonPress-1>", loupe_on)
apercu_canvas.bind("<B1-Motion>", update_loupe)
apercu_canvas.bind("<ButtonRelease-1>", loupe_off)
apercu_canvas.bind("<Configure>", update_apercu)

# Création du menu contextuel (clic droit)
menu_contextuel = tk.Menu(root, tearoff=0)
menu_contextuel.add_command(label="Appliquer à Date", command=lambda: appliquer_selection("Date"))
menu_contextuel.add_command(label="Appliquer à Montant", command=lambda: appliquer_selection("Montant"))
menu_contextuel.add_command(label="Appliquer à IBAN", command=lambda: appliquer_selection("IBAN"))
menu_contextuel.add_command(label="Appliquer à Référence", command=lambda: appliquer_selection("Référence"))


# Fonction pour afficher le menu contextuel au clic droit
def afficher_menu_contextuel(event):
    """Affiche le menu contextuel avec les options d'application du texte sélectionné."""
    try:
        selection = text_area.selection_get()  # Vérifie si une sélection existe
        print(f"🕵️‍♂️ Texte sélectionné : {selection}")  # Vérifie si la sélection est bien récupérée

        if selection:
            menu_contextuel.post(event.x_root, event.y_root)  # Affiche le menu au bon endroit
    except tk.TclError:
        print("⚠️ Aucune sélection détectée")  # Message en cas d'erreur
        
# Associer le menu contextuel au clic droit
text_area.bind("<Button-3>", afficher_menu_contextuel)  # Clic droit sur la zone d'affichage


def appliquer_selection(champ):
    """Applique le texte sélectionné au champ choisi (Date, Montant, IBAN, Référence)."""
    try:
        selection = text_area.selection_get().strip()
        if not selection:
            return  # Ne rien faire si aucune sélection

        # Récupérer la ligne sélectionnée dans le tableau
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Aucune sélection", "Veuillez sélectionner une facture dans la liste.")
            return

        item_id = selected_item[0]
        values = tree.item(item_id, "values")
        if not values:
            return

        # Déterminer l'index correspondant au champ
        champs_mapping = {"Date": 0, "Montant": 1, "IBAN": 2, "Référence": 3}
        col_index = champs_mapping[champ]

        # Mise à jour dans l'affichage
        new_values = list(values)
        new_values[col_index] = selection
        tree.item(item_id, values=new_values)

        # Mise à jour dans le DataFrame
        fichier = values[4]  # Nom du fichier
        df_factures.loc[df_factures["fichier"] == fichier, champ.lower()] = selection
        df_factures.to_csv(CSV_FACTURES, index=False)

        messagebox.showinfo("Mise à jour réussie", f"{champ} mis à jour avec : {selection}")

    except tk.TclError:
        messagebox.showwarning("Aucune sélection", "Veuillez sélectionner un texte dans la zone d'affichage.")

def modifier_date_facture(event=None):
    """Ouvre une fenêtre pour modifier la date de facture sans calendrier, avec saisie jour + combobox mois/année."""
    selected_item = tree.selection()
    if not selected_item:
        return

    item_id = tree.focus()
    values = tree.item(item_id, "values")
    if not values:
        return

    fichier = values[4]  # Récupération du nom de fichier
    ancienne_date = values[0] if values[0] != "N/A" else ""

    # Création de la fenêtre modale
    date_window = tk.Toplevel(root)
    date_window.title("Modifier la date de facture")
    date_window.geometry("350x120")
    date_window.transient(root)
    date_window.grab_set()

    # ---- Analyser l'ancienne date (JJ-MM-YYYY) ----
    day_init, month_init, year_init = "01", "01", str(datetime.now().year)
    try:
        if ancienne_date and ancienne_date != "N/A":
            d, m, y = ancienne_date.split("-")
            day_init, month_init, year_init = d, m, y
    except ValueError:
        pass  # Si la date n'est pas au format JJ-MM-YYYY, on garde les valeurs par défaut

    frame_inputs = ttk.Frame(date_window)
    frame_inputs.pack(pady=10)

    # ---- Saisie du jour (DD) ----
    day_var = tk.StringVar(value=day_init)
    day_entry = ttk.Entry(frame_inputs, textvariable=day_var, width=4)
    day_entry.grid(row=0, column=0, padx=5)
    ttk.Label(frame_inputs, text="/").grid(row=0, column=1)

    # ---- Liste déroulante des mois (01 à 12) ----
    months = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]
    month_var = tk.StringVar(value=month_init)
    month_combo = ttk.Combobox(frame_inputs, textvariable=month_var, values=months, state="readonly", width=3)
    month_combo.grid(row=0, column=2, padx=5)
    ttk.Label(frame_inputs, text="/").grid(row=0, column=3)

    # ---- Liste déroulante des années (2023 à année en cours) ----
    current_year = datetime.now().year
    years_list = [str(y) for y in range(2023, current_year + 1)]
    if year_init not in years_list:
        years_list.insert(0, year_init)

    year_var = tk.StringVar(value=year_init)
    year_combo = ttk.Combobox(frame_inputs, textvariable=year_var, values=years_list, state="readonly", width=5)
    year_combo.grid(row=0, column=4, padx=5)

    def enregistrer_date():
        """Construit la nouvelle date (JJ-MM-YYYY) et l'enregistre."""
        d = day_var.get().strip()
        m = month_var.get().strip()
        y = year_var.get().strip()

        if not (d and m and y):
            return

        new_date = f"{d}-{m}-{y}"

        # Mise à jour dans le Treeview
        tree.item(item_id, values=(
            new_date, values[1], values[2], values[3], values[4], values[5]
        ))

        # Mise à jour dans le DataFrame
        df_factures.loc[df_factures["fichier"] == fichier, "date_facture"] = new_date
        df_factures.to_csv(CSV_FACTURES, index=False)

        text_area.insert("end", f"\n✅ Date mise à jour pour {fichier} : {new_date}\n")
        date_window.destroy()

    frame_buttons = ttk.Frame(date_window)
    frame_buttons.pack(pady=5)

    btn_save = ttk.Button(frame_buttons, text="Enregistrer", command=enregistrer_date)
    btn_save.grid(row=0, column=0, padx=5)

    btn_cancel = ttk.Button(frame_buttons, text="Annuler", command=date_window.destroy)
    btn_cancel.grid(row=0, column=1, padx=5)

    date_window.focus_set()
    date_window.mainloop()


def modifier_valeur(event=None):
    """Ouvre une fenêtre pour modifier les valeurs des colonnes éditables."""
    
    selected_item = tree.selection()
    if not selected_item:
        return

    # Récupération des valeurs de la ligne sélectionnée
    item_id = tree.focus()
    values = tree.item(item_id, "values")
    if not values:
        return

    colonne_index = tree.identify_column(event.x)  # Colonne cliquée
    col_index = int(colonne_index[1:]) - 1  # Convertit "C1" en 0, "C2" en 1, etc.
    col_name = tree["columns"][col_index]
    
    # Si on clique sur la colonne fournisseur ou IBAN, ouvrir la fenêtre spéciale
    if col_name.lower() in ["fournisseur", "iban"]:
        choisir_fournisseur_iban(event, col_name)
        return

    if col_name == "Date":
        modifier_date_facture(event)
        return

    # Gestion spéciale pour "paiement"
    if col_name == "paiement":
        modifier_paiement(event)
        return  # On arrête la fonction ici, pas besoin d'ouvrir une autre fenêtre

    # Création de la fenêtre modale AVANT d'ajouter les boutons
    edit_window = tk.Toplevel(root)
    edit_window.title(f"Modifier {col_name}")
    edit_window.geometry("300x150")
    edit_window.transient(root)
    edit_window.grab_set()

    # On ne peut modifier que Montant ou Référence
    if col_name not in ["Montant", "Référence"]:
        return  # Ignore le clic sur les autres colonnes

    old_value = values[col_index]

    ttk.Label(edit_window, text=f"Modifier {col_name} :").pack(pady=10)

    new_value_var = tk.StringVar(value=old_value)
    new_value_entry = ttk.Entry(edit_window, textvariable=new_value_var, width=30)
    new_value_entry.pack(pady=5)
    new_value_entry.focus()
    new_value_entry.select_range(0, tk.END)

    def enregistrer_modification():
        """Met à jour la valeur et sauvegarde dans le CSV."""
        new_value = new_value_var.get().strip()
        if not new_value:
            return  

        fichier = values[5]  # Récupération du fichier associé

        # Met à jour la valeur éditée dans l'affichage du tableau
        new_values = list(values)
        new_values[col_index] = new_value
        tree.item(item_id, values=new_values)  

        # Met à jour le DataFrame `df_factures`
        col_name_csv = col_name.lower()
        if col_name_csv in df_factures.columns:
            df_factures.loc[df_factures["fichier"] == fichier, col_name_csv] = new_value

        # Enregistre les modifications dans le fichier CSV
        df_factures.to_csv(CSV_FACTURES, index=False)  

        # Affiche un message dans text_area
        text_area.insert("end", f"\n✅ {col_name} mis à jour pour {fichier} : {new_value}\n")

        # Ferme la fenêtre après modification
        edit_window.destroy()

    # Ajout des boutons de validation et d'annulation
    btn_save = ttk.Button(edit_window, text="Enregistrer", command=enregistrer_modification)
    btn_save.pack(pady=5)

    btn_cancel = ttk.Button(edit_window, text="Annuler", command=edit_window.destroy)
    btn_cancel.pack(pady=5)

# Lier le double-clic à la fonction de modification
tree.bind("<Double-1>", modifier_valeur)  # ✅ Maintenant, "paiement" est inclus




# Fonction pour ouvrir un fichier
def ouvrir_fichier():
    selected_item = tree.selection()
    if selected_item:
        fichier = tree.item(selected_item, "values")[4]  # Index 4 = colonne "Fichier"
        chemin_fichier = os.path.join(DOSSIER_FACTURES, fichier)
        if os.path.exists(chemin_fichier):
            subprocess.run(["open", "-a", "PDF Expert", chemin_fichier])  # Ouvre avec PDF Expert
        else:
            text_area.insert("end", f"\n⚠️ Fichier introuvable : {chemin_fichier}\n")

def on_facture_selection(event):
    """Affiche l'aperçu du fichier sélectionné lors de la sélection dans le tableau."""
    selected_items = tree.selection()
    if selected_items:
        # Si plusieurs éléments sont sélectionnés, ne prendre que le premier
        selected_item = selected_items[0]
        fichier = tree.item(selected_item, "values")[5]  # Index 5 = colonne "Fichier"
        chemin_fichier = os.path.join(DOSSIER_FACTURES, fichier)
        afficher_apercu_fichier(chemin_fichier)


def relire_fichier():
    """Relit un fichier sélectionné et met à jour ses informations."""
    selected_items = tree.selection()
    if not selected_items:
        messagebox.showinfo("Information", "Veuillez sélectionner une facture à relire.")
        return
    
    
    # Si plusieurs éléments sont sélectionnés, demander confirmation
    if len(selected_items) > 1:
        confirm = messagebox.askyesno("Confirmation", 
                                      f"Vous avez sélectionné {len(selected_items)} factures. " 
                                      f"Voulez-vous toutes les analyser à nouveau? " 
                                      f"Cela peut prendre du temps.")
        if not confirm:
            return
    
    # Traiter chaque élément sélectionné
    for item_id in selected_items:
        fichier = tree.item(item_id, "values")[5]  # Index de la colonne Fichier
        chemin_fichier = os.path.join(DOSSIER_FACTURES, fichier)
        
        # Effacer le contenu de la zone de texte si c'est le premier élément
        if item_id == selected_items[0]:
            text_area.delete("1.0", "end")
            text_area.insert("end", f"🔄 Analyse de {len(selected_items)} facture(s)...\n\n")
        
        if os.path.exists(chemin_fichier):
            text_area.insert("end", f"📄 Analyse de: {fichier}\n")
            
            # Forcer l'analyse complète en désactivant le cache
            infos = analyser_facture_complete(chemin_fichier, utiliser_cache=False)
            
            # Imprimer les infos pour debug
            print(f"⚙️ Infos extraites pour {fichier}: {infos}")
                        
            # Vérifier si les informations sont valides
            if infos:
                # Afficher les informations extraites dans text_area
                text_area.insert("end", f"   📅 Date: {infos.get('date_facture', 'N/A')}\n")
                text_area.insert("end", f"   💰 Fournisseur: {infos.get('fournisseur', 'N/A')}\n")
                text_area.insert("end", f"   💰 Montant: {infos.get('montant', 'N/A')}\n")
                text_area.insert("end", f"   🏦 IBAN: {infos.get('iban', 'N/A')}\n")
                text_area.insert("end", f"   📌 Référence: {infos.get('reference_facture', 'N/A')}\n\n")
                
                # Mettre à jour les valeurs dans le tableau
                new_values = list(tree.item(item_id, "values"))
                new_values[0] = infos.get('date_facture', new_values[0])     # Date
                new_values[1] = infos.get('fournisseur', new_values[1])      # Fournisseur
                if 'montant' in infos and infos['montant'] != 'N/A':
                    # Convertir en nombre si possible, sinon conserver la valeur existante
                    try:
                        montant_val = float(infos['montant'])
                        new_values[2] = f"{montant_val:.2f}"  # Formater avec 2 décimales
                    except (ValueError, TypeError):
                        # Conserver la valeur existante si conversion impossible
                        pass
                new_values[3] = infos.get('iban', new_values[3])             # IBAN
                new_values[4] = infos.get('reference_facture', new_values[4]) # Référence
                
                # Mettre à jour le tableau
                tree.item(item_id, values=tuple(new_values))
                
                # Mettre à jour le DataFrame
                index = df_factures[df_factures['fichier'] == fichier].index
                if not index.empty:
                    df_factures.loc[index, 'date_facture'] = infos.get('date_facture', df_factures.loc[index, 'date_facture'].values[0])
                    df_factures.loc[index, 'fournisseur'] = infos.get('fournisseur', df_factures.loc[index, 'fournisseur'].values[0])
                    # Correction pour le montant
                    if 'montant' in infos and infos['montant'] != 'N/A':
                        try:
                            df_factures.loc[index, 'montant'] = float(infos['montant'])
                        except (ValueError, TypeError):
                            # Ne pas modifier si la conversion échoue
                            pass
                    df_factures.loc[index, 'iban'] = infos.get('iban', df_factures.loc[index, 'iban'].values[0])
                    df_factures.loc[index, 'reference_facture'] = infos.get('reference_facture', df_factures.loc[index, 'reference_facture'].values[0])
            else:
                text_area.insert("end", f"⚠️ Aucune information n'a pu être extraite pour: {fichier}\n\n")
        else:
            text_area.insert("end", f"⚠️ Fichier introuvable: {chemin_fichier}\n\n")
    
    # Sauvegarder les modifications du DataFrame après avoir traité toutes les factures
    df_factures.to_csv(CSV_FACTURES, index=False)
    text_area.insert("end", "\n✅ Données mises à jour et sauvegardées.\n")
    
    # Afficher l'aperçu du premier fichier sélectionné
    if selected_items:
        first_item = selected_items[0]
        first_fichier = tree.item(first_item, "values")[5]
        first_chemin = os.path.join(DOSSIER_FACTURES, first_fichier)
        if os.path.exists(first_chemin):
            afficher_apercu_fichier(first_chemin)
    
def realiser_virement():
    """Ouvre une fenêtre pour réaliser un virement et met à jour le statut de paiement."""
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Aucune sélection", "Veuillez sélectionner une facture.")
        return

    item_id = selected_item[0]
    values = tree.item(item_id, "values")
    
    if not values:
        return

    # **Récupérer les données de la facture sélectionnée**
    date_facture = values[0] if values[0] else "N/A"
    montant_facture = values[1] if values[1] else "0.00 €"
    iban_facture = values[2] if values[2] else "N/A"
    reference_facture = values[3] if values[3] else "N/A"

    # **Vérifier si l'IBAN est déjà enregistré**
    liste_ibans = charger_liste_ibans()
    ibans_connus = [entry["IBAN"] for entry in liste_ibans]
    iban_inconnu = iban_facture not in ibans_connus and iban_facture != "N/A"

    # **Créer la fenêtre de virement (ajustement taille)**
    virement_window = tk.Toplevel(root)
    virement_window.title("Réaliser un Virement")
    virement_window.geometry("750x280")  # ✅ Ajustement de la hauteur
    virement_window.transient(root)
    virement_window.grab_set()

    frame = ttk.Frame(virement_window, padding=10)
    frame.pack(fill="both", expand=True)

    # **Date du jour pour pré-remplissage**
    date_du_jour = datetime.today().strftime("%d-%m-%Y")

    # **Définition des variables avec pré-remplissage**
    date_virement_var = tk.StringVar(value=date_du_jour)
    reference_var = tk.StringVar(value=reference_facture)
    iban_var = tk.StringVar(value=iban_facture)
    montant_var = tk.StringVar(value=montant_facture)

    # **Champs du formulaire avec pré-remplissage**
    ttk.Label(frame, text="Date de Virement :").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    date_entry = ttk.Entry(frame, textvariable=date_virement_var, width=15)
    date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(frame, text="Référence :").grid(row=1, column=0, padx=5, pady=5, sticky="w")
    reference_entry = ttk.Entry(frame, textvariable=reference_var, width=30)
    reference_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(frame, text="IBAN :").grid(row=2, column=0, padx=5, pady=5, sticky="w")
    iban_entry = ttk.Entry(frame, textvariable=iban_var, width=30)
    iban_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

    # **Bouton pour choisir un IBAN dans la liste**
    btn_choisir_iban = ttk.Button(frame, text="🔍", command=lambda: choisir_iban(iban_var))
    btn_choisir_iban.grid(row=2, column=2, padx=5, pady=5)

    # **Bouton "Ajouter nouvel IBAN" aligné avec la loupe**
    if iban_inconnu:
        lbl_alert = ttk.Label(frame, text="⚠️ Cet IBAN est inconnu de la liste", foreground="red")
        lbl_alert.grid(row=3, column=1, padx=5, pady=2, sticky="w")

        btn_ajouter_iban = ttk.Button(frame, text="Ajouter nouvel IBAN", command=lambda: ajouter_iban(iban_var.get()))
        btn_ajouter_iban.grid(row=2, column=3, padx=5, pady=5, sticky="w")  # ✅ Placement à côté de la loupe

    # ✅ **Correction : placer le montant plus bas pour éviter la superposition**
    ttk.Label(frame, text="Montant (€) :").grid(row=4, column=0, padx=5, pady=5, sticky="w")
    montant_entry = ttk.Entry(frame, textvariable=montant_var, width=15)
    montant_entry.grid(row=4, column=1, padx=5, pady=5, sticky="w")

    def valider_virement():
        """Vérifie et envoie le virement."""
        date_virement = date_virement_var.get().strip()
        reference = reference_var.get().strip()
        iban = iban_var.get().strip()
        montant = montant_var.get().strip()

        if not date_virement or not reference or not iban or not montant:
            messagebox.showwarning("Champs manquants", "Veuillez remplir tous les champs du virement.")
            return

        try:
            # 🔹 Appel du script de génération de virement
            envoyer_virement(date_facture, reference_facture, iban_facture, montant_facture)
            # ✅ Mise à jour de l'interface après un virement réussi
            tree.item(item_id, values=(
                date_facture, montant_facture, iban_facture, reference_facture, values[4], f"✔️ {date_virement}"
            ))

            df_factures.loc[df_factures["fichier"] == values[4], "paiement"] = f"✔️ {date_virement}"
            df_factures.to_csv(CSV_FACTURES, index=False)

            messagebox.showinfo("Succès", f"Virement de {montant} € validé pour {iban}.")
            virement_window.destroy()
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite : {e}")

    btn_valider = ttk.Button(frame, text="Valider Virement", command=valider_virement)
    btn_valider.grid(row=5, column=1, pady=10)

    # ✅ Bouton "Annuler"
    btn_annuler = ttk.Button(frame, text="Annuler", command=virement_window.destroy)
    btn_annuler.grid(row=5, column=2, pady=10)


    def choisir_iban():
        """Ouvre une fenêtre pour choisir un IBAN depuis la liste des fournisseurs."""
        iban_window = tk.Toplevel(virement_window)
        iban_window.title("Choisir un IBAN")
        iban_window.geometry("450x250")  # ✅ Augmentation de la largeur

        ibans = charger_liste_ibans()

        # ✅ Tri des IBANs par ordre alphabétique du fournisseur
        ibans.sort(key=lambda x: x["fournisseur"])

        lb = tk.Listbox(iban_window, height=10, width=50)  # ✅ Augmente la largeur
        lb.pack(padx=10, pady=10, fill="both", expand=True)

        for item in ibans:
            lb.insert("end", f"{item['fournisseur']} - {item['IBAN']}")

        def selectionner_iban():
            selection = lb.curselection()
            if selection:
                iban_var.set(ibans[selection[0]]["IBAN"])
            iban_window.destroy()

        btn_select = ttk.Button(iban_window, text="Sélectionner", command=selectionner_iban)
        btn_select.pack(pady=5)

        # ✅ Ajout d'un bouton Annuler
        btn_cancel = ttk.Button(iban_window, text="Annuler", command=iban_window.destroy)
        btn_cancel.pack(pady=5)

    btn_choisir_iban = ttk.Button(frame, text="🔍", command=lambda: choisir_fournisseur_iban(event, "IBAN"))
    btn_choisir_iban.grid(row=2, column=2, padx=5, pady=5)


    def ajouter_iban(iban_initial=""):
        """Ouvre une fenêtre pour enregistrer un nouvel IBAN."""
        enregistrer_window = tk.Toplevel(root)
        enregistrer_window.title("Enregistrer un nouvel IBAN")
        enregistrer_window.geometry("500x160")
        enregistrer_window.transient(root)
        enregistrer_window.grab_set()

        frame = ttk.Frame(enregistrer_window, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="fournisseur :").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        fournisseur_var = tk.StringVar()
        fournisseur_entry = ttk.Entry(frame, textvariable=fournisseur_var, width=30)
        fournisseur_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(frame, text="IBAN :").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        nouvel_iban_var = tk.StringVar(value=iban_initial)
        nouvel_iban_entry = ttk.Entry(frame, textvariable=nouvel_iban_var, width=30)
        nouvel_iban_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        def sauvegarder_iban():
            fournisseur = fournisseur_var.get().strip()
            nouvel_iban = nouvel_iban_var.get().strip()

            if not fournisseur or not nouvel_iban:
                messagebox.showwarning("Champs incomplets", "Veuillez renseigner tous les champs.")
                return

            enregistrer_iban(fournisseur, nouvel_iban)
            enregistrer_window.destroy()

        ttk.Button(frame, text="Enregistrer", command=sauvegarder_iban).grid(row=2, column=0, padx=5, pady=5)
        ttk.Button(frame, text="Annuler", command=enregistrer_window.destroy).grid(row=2, column=1, padx=5, pady=5)

        enregistrer_window.mainloop()

    virement_window.mainloop()

def verifier_nouvelles_factures():
    """Vérifie la présence de nouvelles factures et gère les fichiers déplacés."""
    global df_factures  # Utiliser la variable globale

    # Charger la base de données existante
    db = charger_base_donnees()
    
    # Set pour suivre les factures trouvées dans cette analyse
    factures_trouvees = set()
    
    # Lister tous les fichiers actuels
    fichiers_actuels = lister_fichiers_factures()
    
    # Nombre de nouvelles factures détectées
    nouvelles_factures = 0
    
    for fichier in fichiers_actuels:
        chemin_fichier = os.path.join(DOSSIER_FACTURES, fichier)
        
        # Ajouter ou mettre à jour dans la base de données
        infos_extraites = {}  # Initialiser avec un dictionnaire vide
        
        # Vérifier si le fichier est déjà dans df_factures
        fichier_existant = df_factures[df_factures["fichier"] == fichier]
        
        if not fichier_existant.empty:
            # Le fichier existe déjà, récupérer ses informations
            row = fichier_existant.iloc[0]
            infos_extraites = {
                "date_facture": row.get("date_facture", "N/A"),
                "fournisseur": row.get("fournisseur", "N/A"),
                "montant": row.get("montant", "N/A"),
                "iban": row.get("iban", "N/A"),
                "reference_facture": row.get("reference_facture", "N/A"),
                "paiement": row.get("paiement", "❌")
            }
        
        # Ajouter ou mettre à jour dans la base de données
        id_facture, est_nouvelle = ajouter_ou_maj_facture(
            db, os.path.basename(fichier), chemin_fichier, infos_extraites
        )
        
        if id_facture:
            factures_trouvees.add(id_facture)
            
            # Si c'est une nouvelle facture, l'analyser
            if est_nouvelle:
                nouvelles_factures += 1
                print(f"🆕 Analyse de la nouvelle facture: {fichier}")
                
                # Analyser la facture avec votre fonction existante
                infos = analyser_facture_complete(chemin_fichier)
                
                # Mettre à jour la base de données avec les nouvelles informations
                db["factures"][id_facture].update(infos)
                
                # Ajouter au DataFrame si ce n'est pas déjà fait
                if fichier_existant.empty:
                    nouvelle_ligne = {
                        "date_facture": infos.get("date_facture", "N/A"),
                        "fournisseur": infos.get("fournisseur", "N/A"),
                        "montant": infos.get("montant", "N/A"),
                        "iban": infos.get("iban", "N/A"),
                        "reference_facture": infos.get("reference_facture", "N/A"),
                        "fichier": fichier,
                        "paiement": "❌"
                    }
                    df_factures = pd.concat([df_factures, pd.DataFrame([nouvelle_ligne])], ignore_index=True)
    
    # Marquer les factures qui n'ont pas été trouvées comme indisponibles
    marquer_factures_manquantes(db, factures_trouvees)
    
    # Sauvegarder la base de données mise à jour
    sauvegarder_base_donnees(db)
    
    if nouvelles_factures > 0:
        print(f"✅ {nouvelles_factures} nouvelles factures ajoutées au fichier CSV.")
        # Sauvegarder le DataFrame mis à jour
        df_factures.to_csv(CSV_FACTURES, index=False)
    else:
        print("✅ Aucune nouvelle facture détectée.")


def mettre_a_jour_listes_filtres():
    """Met à jour les listes déroulantes des filtres avec les valeurs disponibles."""
    # Années
    annees = sorted(list(set(df_factures["date_facture"].str.split("-").str[-1].dropna().tolist())), reverse=True)
    if "---" not in annees:
        annees.insert(0, "---")
    liste_annees["values"] = annees
    
    # Mois
    mois = sorted(list(set(df_factures["date_facture"].str.split("-").str[1].dropna().tolist())))
    if "---" not in mois:
        mois.insert(0, "---")
    liste_mois["values"] = mois
    
    # Fournisseurs
    fournisseurs = sorted(list(set(df_factures["fournisseur"].dropna().tolist())))
    if "Tous" not in fournisseurs:
        fournisseurs.insert(0, "Tous")
    liste_fournisseurs["values"] = fournisseurs
    
    # Par défaut : année et mois courants si disponibles
    current_year = str(datetime.now().year)
    current_month = datetime.now().strftime("%m")
    
    if current_year in annees:
        annee_var.set(current_year)
    else:
        annee_var.set("---")
        
    if current_month in mois:
        mois_var.set(current_month)
    else:
        mois_var.set("---")

def appliquer_filtres():
    """Applique les filtres sélectionnés aux factures et met à jour l'affichage."""
    # Supprimer toutes les entrées du tableau
    for item in tree.get_children():
        tree.delete(item)
    
    # Récupérer les valeurs actuelles des filtres
    annee_courante = annee_var.get()
    mois_courant = mois_var.get()
    fournisseur_courant = fournisseur_var.get()
    statut_courant = statut_var.get()
    
    print(f"🔍 Filtres appliqués - Année: {annee_courante}, Mois: {mois_courant}")
    
    # Appliquer les filtres sélectionnés
    df_filtree = df_factures.copy()
    
    # Filtre année - SEULEMENT si une année spécifique est demandée
    if annee_courante and annee_courante != "---":
        mask_annee = df_filtree["date_facture"].astype(str).str.endswith(annee_courante, na=False)
        df_filtree = df_filtree[mask_annee]
    
    # Filtre mois - SEULEMENT si un mois spécifique est demandé
    if mois_courant and mois_courant != "---":
        # Extraire le mois sous forme "MM" des dates (format "DD-MM-YYYY")
        mask_mois = df_filtree["date_facture"].astype(str).str.split("-").str[1].eq(mois_courant)
        # Gérer les valeurs NaN qui causeront des problèmes avec str.split
        mask_mois = mask_mois.fillna(False)
        df_filtree = df_filtree[mask_mois]
        
    # Filtre fournisseur
    if fournisseur_courant and fournisseur_courant != "Tous":
        df_filtree = df_filtree[df_filtree["fournisseur"] == fournisseur_courant]
    
    # Filtre statut de paiement
    if statut_courant == "Payés":
        df_filtree = df_filtree[df_filtree["paiement"].str.contains("✔️", na=False)]
    elif statut_courant == "Non payés":
        df_filtree = df_filtree[~df_filtree["paiement"].str.contains("✔️", na=False)]
    
    # Filtre montant minimum
    if montant_min_var.get():
        try:
            montant_min = float(montant_min_var.get().replace(",", "."))
            # Convertir les montants en valeurs numériques
            df_filtree["montant_num"] = df_filtree["montant"].str.replace("€", "").str.replace(",", ".").astype(float)
            df_filtree = df_filtree[df_filtree["montant_num"] >= montant_min]
            df_filtree = df_filtree.drop(columns=["montant_num"])
        except ValueError:
            # Ignorer si la valeur n'est pas un nombre valide
            pass
    
    # Tri des factures par date (décroissant)
    df_filtree = df_filtree.sort_values(by=["date_facture"], ascending=False)
    
    # Réinsérer les données filtrées dans le tableau
    for index, row in df_filtree.iterrows():
        paiement = row["paiement"] if pd.notna(row["paiement"]) else "❌"
        item_id = tree.insert("", "end", values=(
            row["date_facture"], 
            row["fournisseur"],
            row["montant"], 
            row["iban"], 
            row["reference_facture"], 
            row["fichier"], 
            paiement
        ))
        # Appliquer le style "disparu" si le fichier n'est plus disponible
        if not row.get("disponible", True):
            tree.item(item_id, tags=("disparu",))
    
    # Afficher un récapitulatif
    text_area.delete("1.0", "end")
    text_area.insert("end", f"📊 Résultats filtrés : {len(df_filtree)} factures trouvées\n")
    text_area.insert("end", f"📅 Filtres: Année {annee_courante}, Mois {mois_courant}\n")
    
    # Calcul du total des montants
    if not df_filtree.empty:
        try:
            montants = df_filtree["montant"].str.replace("€", "").str.replace(",", ".").astype(float)
            total = montants.sum()
            text_area.insert("end", f"💰 Total : {total:.2f} €\n")
        except:
            pass
        
# Modification 2: Gérer le changement d'année pour réinitialiser le mois
def on_annee_change(event):
    """Réinitialise le filtre mois quand l'année change."""
    # Réinitialiser le mois
    mois_var.set("---")
    
    # Appliquer les filtres
    appliquer_filtres()

def reinitialiser_filtres():
    """Réinitialise tous les filtres à leurs valeurs par défaut."""
    annee_var.set("---")
    mois_var.set("---")
    fournisseur_var.set("Tous")
    statut_var.set("Tous")
    montant_min_var.set("")
    appliquer_filtres()

def afficher_analyse_fournisseurs():
    """Affiche une fenêtre avec l'analyse des factures par fournisseur."""
    analyse_window = tk.Toplevel(root)
    analyse_window.title("Analyse par fournisseur")
    analyse_window.geometry("850x500")
    
    # Cadre principal avec padding
    main_frame = ttk.Frame(analyse_window, padding=10)
    main_frame.pack(fill="both", expand=True)
    
    # En-tête
    ttk.Label(main_frame, text="Analyse des factures par fournisseur", 
             font=("Arial", 14, "bold")).pack(pady=10)
    
    # Tableau pour l'analyse
    columns = ("Fournisseur", "Nombre", "Total", "Moyenne", "Dernier", "Statut")
    analyse_tree = ttk.Treeview(main_frame, columns=columns, show="headings", height=15)
    
    # Configuration des colonnes
    for col, width, anchor in [
        ("Fournisseur", 200, "w"),
        ("Nombre", 80, "center"),
        ("Total", 100, "e"),
        ("Moyenne", 100, "e"),
        ("Dernier", 100, "center"),
        ("Statut", 100, "center")
    ]:
        analyse_tree.heading(col, text=col)
        analyse_tree.column(col, width=width, anchor=anchor)
    
    # Scrollbar pour le tableau
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=analyse_tree.yview)
    analyse_tree.configure(yscrollcommand=scrollbar.set)
    
    # Placement du tableau et scrollbar
    analyse_tree.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    # Cadre pour les filtres et contrôles
    control_frame = ttk.Frame(analyse_window, padding=10)
    control_frame.pack(fill="x")
    
    # Filtre par année
    ttk.Label(control_frame, text="Année:").grid(row=0, column=0, padx=5, pady=5)
    annee_analyse_var = tk.StringVar(value="Toutes")
    annees = ["Toutes"] + sorted(list(set(df_factures["date_facture"].str.split("-").str[-1].dropna().tolist())), reverse=True)
    annee_combo = ttk.Combobox(control_frame, textvariable=annee_analyse_var, values=annees, width=10, state="readonly")
    annee_combo.grid(row=0, column=1, padx=5, pady=5)
    
    # Bouton pour l'export CSV
    ttk.Button(control_frame, text="Exporter en CSV", 
              command=lambda: exporter_analyse_csv(analyse_tree)).grid(row=0, column=3, padx=5, pady=5)
    
    # Bouton pour fermer
    ttk.Button(control_frame, text="Fermer", 
              command=analyse_window.destroy).grid(row=0, column=4, padx=5, pady=5)
    
    def remplir_analyse_tree():
        """Remplit le tableau d'analyse avec les données regroupées par fournisseur."""
        # Vider le tableau
        for item in analyse_tree.get_children():
            analyse_tree.delete(item)
        
        # Filtrer par année si nécessaire
        df_analyse = df_factures.copy()
        if annee_analyse_var.get() != "Toutes":
            df_analyse = df_analyse[df_analyse["date_facture"].str.endswith(annee_analyse_var.get())]
        
        # Regrouper par fournisseur
        resultats = []
        
        for fournisseur, groupe in df_analyse.groupby("fournisseur"):
            # Ignorer les fournisseurs vides
            if pd.isna(fournisseur) or fournisseur == "":
                continue
                
            # Nombre de factures
            nombre = len(groupe)
            
            # Montants totaux et moyenne
            try:
                montants = groupe["montant"].str.replace("€", "").str.replace(",", ".").astype(float)
                total = montants.sum()
                moyenne = total / nombre if nombre > 0 else 0
            except:
                total = 0
                moyenne = 0
            
            # Date de la dernière facture
            derniere_date = groupe["date_facture"].max() if not groupe.empty else "N/A"
            
            # Statut (payées / total)
            payees = groupe[groupe["paiement"].str.contains("✔️", na=False)].shape[0]
            statut = f"{payees}/{nombre}"
            
            resultats.append({
                "fournisseur": fournisseur,
                "nombre": nombre,
                "total": total,
                "moyenne": moyenne,
                "derniere": derniere_date,
                "statut": statut
            })
        
        # Trier par montant total décroissant
        resultats.sort(key=lambda x: x["total"], reverse=True)
        
        # Remplir le tableau
        for item in resultats:
            analyse_tree.insert("", "end", values=(
                item["fournisseur"],
                item["nombre"],
                f"{item['total']:.2f} €",
                f"{item['moyenne']:.2f} €",
                item["derniere"],
                item["statut"]
            ))
    
    def exporter_analyse_csv(tree):
        """Exporte l'analyse vers un fichier CSV."""
        try:
            # Demander où enregistrer le fichier
            from tkinter import filedialog
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                title="Enregistrer l'analyse"
            )
            
            if not filename:
                return
            
            # Collecter les données du tableau
            data = []
            for item_id in tree.get_children():
                values = tree.item(item_id, "values")
                data.append({
                    "Fournisseur": values[0],
                    "Nombre": values[1],
                    "Total": values[2],
                    "Moyenne": values[3],
                    "Dernière facture": values[4],
                    "Statut": values[5]
                })
            
            # Créer un DataFrame et sauvegarder en CSV
            df_export = pd.DataFrame(data)
            df_export.to_csv(filename, index=False)
            
            messagebox.showinfo("Export réussi", f"L'analyse a été exportée vers {filename}")
        
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'export : {e}")
    
    # Lier le changement d'année à la mise à jour du tableau
    annee_combo.bind("<<ComboboxSelected>>", lambda e: remplir_analyse_tree())
    
    # Double-clic sur un fournisseur pour voir ses factures
    def voir_factures_fournisseur(event):
        item = analyse_tree.selection()[0]
        fournisseur = analyse_tree.item(item, "values")[0]
        
        # Fermer la fenêtre d'analyse
        analyse_window.destroy()
        
        # Appliquer le filtre pour ce fournisseur
        fournisseur_var.set(fournisseur)
        appliquer_filtres()
    
    analyse_tree.bind("<Double-1>", voir_factures_fournisseur)
    
    # Remplir initialement le tableau
    remplir_analyse_tree()
    
 



# Charger les icônes redimensionnées
icone_ouvrir = charger_icone("open.png")
icone_relire = charger_icone("read.png")
icone_recharger = charger_icone("refresh.png")
icone_virement = charger_icone("transfer.png")
icone_fermer = charger_icone("close.png")

# Boutons d'action placés sous la zone de texte
frame_buttons = ttk.Frame(frame_left)
frame_buttons.pack(side="top", fill="x", pady=10)

btn_ouvrir = ttk.Button(frame_buttons, text="Ouvrir Facture", image=icone_ouvrir, compound="left", command=ouvrir_fichier)
btn_ouvrir.grid(row=0, column=0, padx=10, pady=5)

btn_relire = ttk.Button(frame_buttons, text="Relire Facture", image=icone_relire, compound="left", command=relire_fichier)
btn_relire.grid(row=0, column=1, padx=10, pady=5)

btn_recharger = ttk.Button(frame_buttons, text="Recharger les fichiers", image=icone_recharger, compound="left", command=lambda: recharger_fichiers(True))
btn_recharger.grid(row=1, column=0, padx=10, pady=5)

btn_verifier = ttk.Button(frame_buttons, text="🔄 Vérifier nouvelles factures", command=charger_en_arriere_plan)
btn_verifier.grid(row=1, column=0, padx=10, pady=5)

btn_virement = ttk.Button(frame_buttons, text="Réaliser Virement", image=icone_virement, compound="left", command=realiser_virement)
btn_virement.grid(row=1, column=1, padx=10, pady=5)

btn_analyse = ttk.Button(frame_buttons, text="Analyse par fournisseur", command=afficher_analyse_fournisseurs)
btn_analyse.grid(row=2, column=0, padx=10, pady=5)

btn_gestion_fournisseurs = ttk.Button(frame_buttons, text="Gérer fournisseurs", command=gerer_fournisseurs)
btn_gestion_fournisseurs.grid(row=2, column=1, padx=10, pady=5)

btn_fermer = ttk.Button(frame_buttons, text="Fermer", image=icone_fermer, compound="left", command=fermer)
btn_fermer.grid(row=1, column=2, padx=10, pady=5)

 
# Panneau de filtres amélioré
frame_filtres = ttk.LabelFrame(frame_left, text="Filtres de recherche")
frame_filtres.pack(side="bottom", fill="x", padx=10, pady=10)

# Première ligne : Année et Mois
date_frame = ttk.Frame(frame_filtres)
date_frame.pack(fill="x", padx=5, pady=5)

ttk.Label(date_frame, text="Année:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
annee_var = tk.StringVar()
liste_annees = ttk.Combobox(date_frame, textvariable=annee_var, state="readonly", width=7)
liste_annees.grid(row=0, column=1, padx=5, pady=2)

ttk.Label(date_frame, text="Mois:").grid(row=0, column=2, padx=5, pady=2, sticky="w")
mois_var = tk.StringVar()
liste_mois = ttk.Combobox(date_frame, textvariable=mois_var, state="readonly", width=5)
liste_mois.grid(row=0, column=3, padx=5, pady=2)

# Deuxième ligne : Fournisseur
fournisseur_frame = ttk.Frame(frame_filtres)
fournisseur_frame.pack(fill="x", padx=5, pady=5)

ttk.Label(fournisseur_frame, text="Fournisseur:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
fournisseur_var = tk.StringVar(value="Tous")
liste_fournisseurs = ttk.Combobox(fournisseur_frame, textvariable=fournisseur_var, state="readonly", width=30)
liste_fournisseurs.grid(row=0, column=1, padx=5, pady=2, columnspan=3, sticky="w")

# Troisième ligne : Statut de paiement et montant min
status_frame = ttk.Frame(frame_filtres)
status_frame.pack(fill="x", padx=5, pady=5)

ttk.Label(status_frame, text="Statut:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
statut_var = tk.StringVar(value="Tous")
liste_statuts = ttk.Combobox(status_frame, textvariable=statut_var, state="readonly", width=10, 
                           values=["Tous", "Payés", "Non payés"])
liste_statuts.grid(row=0, column=1, padx=5, pady=2)

ttk.Label(status_frame, text="Montant min:").grid(row=0, column=2, padx=5, pady=2, sticky="w")
montant_min_var = tk.StringVar()
montant_min_entry = ttk.Entry(status_frame, textvariable=montant_min_var, width=10)
montant_min_entry.grid(row=0, column=3, padx=5, pady=2)

# Boutons d'action
boutons_frame = ttk.Frame(frame_filtres)
boutons_frame.pack(fill="x", padx=5, pady=5)

ttk.Button(boutons_frame, text="Appliquer filtres", command=appliquer_filtres).pack(side="left", padx=5)
ttk.Button(boutons_frame, text="Réinitialiser", command=reinitialiser_filtres).pack(side="left", padx=5)
ttk.Button(boutons_frame, text="Analyse par fournisseur", command=afficher_analyse_fournisseurs).pack(side="right", padx=5)

current_year = str(datetime.now().year)
current_month = datetime.now().strftime("%m")


if "annee" in df_factures.columns and current_year in df_factures["annee"].unique():
    annee_var.set(current_year)
else:
    annee_var.set("---")

if "mois" in df_factures.columns and current_month in df_factures["mois"].unique():
    mois_var.set(current_month)
else:
    mois_var.set("---")

initialiser_systeme()


# Mise à jour des filtres en fonction des factures disponibles
def mettre_a_jour_filtres():
    def extraire_mois(date_str):
        """Extrait le mois d'une date si elle est bien formatée (JJ-MM-AAAA), sinon renvoie '---'."""
        if not isinstance(date_str, str):
            return "---"
        
        # Normaliser les séparateurs
        date_str = date_str.replace("/", "-").replace(".", "-")
        
        parties = date_str.split("-")
        if len(parties) == 3:
            # parties[0] = JJ, parties[1] = MM, parties[2] = AAAA
            return parties[1]
        else:
            return "---"
    
    def extraire_annee(date_str):
        """Extrait l'année d'une date si elle est bien formatée (JJ-MM-AAAA), sinon renvoie '---'."""
        if not isinstance(date_str, str):
            return "---"
        
        # Normaliser les séparateurs
        date_str = date_str.replace("/", "-").replace(".", "-")
        
        parties = date_str.split("-")
        if len(parties) == 3:
            # parties[2] = AAAA
            return parties[2]
        else:
            return "---"

    dates_probleme = df_factures[~df_factures["date_facture"].str.match(r"^\d{2}-\d{2}-\d{4}$", na=False)]
    print("⚠️ Problèmes détectés avec ces dates :\n", dates_probleme["date_facture"])

    df_factures["annee"] = df_factures["date_facture"].apply(extraire_annee)
    df_factures["mois"] = df_factures["date_facture"].apply(extraire_mois)

    annees = sorted(df_factures["annee"].unique(), reverse=True)
    mois = sorted(df_factures["mois"].unique())

    # Ajouter "---" en premier pour avoir une option de sélection par défaut
    if "---" not in annees:
        annees.insert(0, "---")
    if "---" not in mois:
        mois.insert(0, "---")

    liste_annees["values"] = annees
    liste_mois["values"] = mois

    # Définir automatiquement les filtres sur l'année et le mois en cours si disponibles
    current_year = str(datetime.now().year)
    current_month = datetime.now().strftime("%m")
    if current_year in annees:
        annee_var.set(current_year)
    else:
        annee_var.set("---")
    if current_month in mois:
        mois_var.set(current_month)
    else:
        mois_var.set("---")
    # Recharge immédiatement les fichiers filtrés
    recharger_fichiers(False)

# Mise à jour des filtres lors de la sélection
liste_annees.bind("<<ComboboxSelected>>", on_annee_change)
liste_mois.bind("<<ComboboxSelected>>", lambda event: appliquer_filtres())
liste_fournisseurs.bind("<<ComboboxSelected>>", lambda event: appliquer_filtres())
liste_statuts.bind("<<ComboboxSelected>>", lambda event: appliquer_filtres())


# Initialisation des filtres
mettre_a_jour_filtres()




# Variable pour suivre l'état du tri
tri_ordre = {"Date": True, "Montant": True}  # True = ordre croissant, False = décroissant

def trier_colonne(colonne):
    """Trie la colonne cliquée en alternant entre croissant et décroissant."""
    global df_factures, tri_ordre
    
    if colonne == "Date":
        df_factures["date_facture"] = pd.to_datetime(df_factures["date_facture"], format="%d-%m-%Y", errors="coerce")
        df_factures.sort_values(by="date_facture", ascending=tri_ordre["Date"], inplace=True)
        df_factures["date_facture"] = df_factures["date_facture"].dt.strftime("%d-%m-%Y")
        tri_ordre["Date"] = not tri_ordre["Date"]
    
    elif colonne == "Montant":
        df_factures["montant_num"] = pd.to_numeric(df_factures["montant"].str.replace("€", "").str.replace(",", "."), errors="coerce")
        df_factures.sort_values(by="montant_num", ascending=tri_ordre["Montant"], inplace=True)
        df_factures.drop(columns=["montant_num"], inplace=True)
        tri_ordre["Montant"] = not tri_ordre["Montant"]
    
    # Supprimer toutes les entrées du tableau
    for item in tree.get_children():
        tree.delete(item)
    print("🔄 Mise à jour du tableau avec les nouvelles valeurs...")
        
    
    # Réinsérer les données triées
    for _, row in df_factures.iterrows():
        tree.insert("", "end", values=(
            row["date_facture"],
            row["montant"],
            row["iban"],
            row["reference_facture"],
            row["fichier"],
            row["paiement"]
        ))

# Associer le tri aux en-têtes de colonnes
tri_ordre = {"Date": True, "Montant": True}
tree.heading("Date", text="Date", command=lambda: trier_colonne("Date"))
tree.heading("Montant", text="Montant", command=lambda: trier_colonne("Montant"))


# Associer le tri aux en-têtes de colonnes
tree.heading("Date", text="Date", command=lambda: trier_colonne("Date"))
tree.heading("Montant", text="Montant", command=lambda: trier_colonne("Montant"))


# Lier les colonnes cliquées à la fonction de tri
tree.heading("Date", text="Date", command=lambda: trier_colonne("Date"))
tree.heading("Montant", text="Montant", command=lambda: trier_colonne("Montant"))

# **Exécuter la initialiser système au lancement**
initialiser_systeme()


# Lancer l'interface
root.mainloop()


if __name__ == "__main__":
    
    print("\n" + "="*50)
    print("🔍 DEBUG - MODULE PRINCIPAL")
    print(f"📂 DOSSIER_FACTURES: {DOSSIER_FACTURES}")
    print(f"🔄 Arguments système: {sys.argv}")
    print("="*50 + "\n")
    
    # Si on a un argument pour un fichier spécifique
    if len(sys.argv) >= 2 and not sys.argv[1].startswith('--'):
        chemin = sys.argv[1]
        print(f"📄 Analyse du fichier spécifique: {chemin}")
        infos = analyser_facture_complete(chemin)
        print("Informations extraites :")
        print(infos)