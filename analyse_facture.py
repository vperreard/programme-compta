import os
import pdfplumber
import re
import pandas as pd
import pytesseract
from PIL import Image, ImageTk
from datetime import datetime
from pdf2image import convert_from_path
import schwifty
import unicodedata
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk, scrolledtext
import subprocess
import openai
import json
import argparse
import sys
import threading
from factures_db_utils import (
    charger_base_donnees, 
    sauvegarder_base_donnees, 
    ajouter_ou_maj_facture, 
    marquer_factures_manquantes,
    calculer_hash_fichier
)

# Classe principale pour l'analyse des factures
class AnalyseFactures:
    def __init__(self, dossier_factures=None):
        # Charger les credentials depuis le fichier JSON
        self.SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
        with open(os.path.join(self.SCRIPT_DIR, 'credentials.json'), 'r') as file:
            self.credentials = json.load(file)
            
        # Configuration de l'API OpenAI
        openai.api_key = self.credentials['api_key']
        
        # Définition des chemins
        self.definir_dossier_factures(dossier_factures)
        
        # Chemins des fichiers CSV
        self.IBAN_LISTE_CSV = os.path.join(self.SCRIPT_DIR, "liste_ibans.csv")
        self.CSV_FACTURES = os.path.join(self.SCRIPT_DIR, "resultats_factures.csv")
        
        # Liste des colonnes nécessaires
        self.REQUIRED_COLUMNS = [
            "iban", "fournisseur", "paiement", "date_facture", "montant", "fichier", "reference_facture"
        ]
        
        # Vérifier si le fichier existe, sinon le créer
        if not os.path.exists(self.IBAN_LISTE_CSV):
            pd.DataFrame(columns=["fournisseur", "IBAN"]).to_csv(self.IBAN_LISTE_CSV, index=False)
            
        # Chargement des fichiers CSV
        self.df_factures = self.charger_dataframe(self.CSV_FACTURES)
        self.df_factures.rename(columns={"Fournisseur": "fournisseur"}, inplace=True)
        self.df_factures = self.validate_dataframe(self.df_factures, self.REQUIRED_COLUMNS)
        
        # Normalisation des colonnes
        self.df_factures.columns = self.df_factures.columns.str.strip().str.lower()
        
        # Chargement des IBANs
        self.df_ibans = pd.read_csv(self.IBAN_LISTE_CSV, sep=",", dtype=str)
        
        # Vérifications et nettoyage
        self.df_ibans.rename(columns={"Fournisseur": "fournisseur", "iban": "IBAN"}, inplace=True)
        if "fournisseur" in self.df_ibans.columns:
            self.df_ibans["fournisseur"] = self.df_ibans["fournisseur"].astype(str).str.strip()
        self.df_ibans["IBAN"] = self.df_ibans["IBAN"].astype(str).str.replace(" ", "")
        
        # Ajout de la colonne "paiement" si absente
        if "paiement" not in self.df_factures.columns or self.df_factures["paiement"].isnull().all():
            self.df_factures["paiement"] = "❌"
        else:
            self.df_factures["paiement"] = self.df_factures["paiement"].fillna("❌").replace("", "❌")
        
        # Sauvegarde après toutes les modifications
        self.df_factures.to_csv(self.CSV_FACTURES, index=False)
        
        # Variables pour la loupe
        self.loupe_window = None
        self.loupe_canvas = None
        self.apercu_img = None
        self.img_tk = None

    def definir_dossier_factures(self, dossier_factures=None):
        """Définit le dossier des factures en utilisant l'argument ou la config."""
        if dossier_factures:
            self.DOSSIER_FACTURES = dossier_factures
            print(f"✅ DOSSIER_FACTURES défini par argument: {self.DOSSIER_FACTURES}")
        else:
            # Essayer depuis config
            try:
                from config import get_file_path
                config_path = get_file_path("dossier_factures", verify_exists=True)
                if config_path:
                    self.DOSSIER_FACTURES = config_path
                    print(f"✅ DOSSIER_FACTURES défini par config: {self.DOSSIER_FACTURES}")
                else:
                    self.DOSSIER_FACTURES = os.path.expanduser("~/Documents/Frais_Factures")
                    print(f"⚠️ Config path vide, utilisation du chemin par défaut: {self.DOSSIER_FACTURES}")
            except Exception as e:
                self.DOSSIER_FACTURES = os.path.expanduser("~/Documents/Frais_Factures")
                print(f"❌ Erreur lors de l'import de config: {e}")
                print(f"⚠️ Utilisation du chemin par défaut: {self.DOSSIER_FACTURES}")
        
        # Vérifier que le dossier existe
        if not os.path.exists(self.DOSSIER_FACTURES):
            print(f"❌ ERREUR: Le dossier {self.DOSSIER_FACTURES} n'existe pas!")
            if not os.path.exists(os.path.dirname(self.DOSSIER_FACTURES)):
                print(f"Le dossier parent {os.path.dirname(self.DOSSIER_FACTURES)} n'existe pas non plus.")
        else:
            print(f"✅ Le dossier {self.DOSSIER_FACTURES} existe.")
            # Lister les fichiers présents
            all_files = os.listdir(self.DOSSIER_FACTURES)
            pdf_files = [f for f in all_files if f.lower().endswith(('.pdf', '.jpg', '.jpeg', '.png'))]
            print(f"📄 {len(pdf_files)} fichiers PDF/images trouvés:")
            for i, file in enumerate(pdf_files[:10]):  # Montrer les 10 premiers fichiers
                print(f"   {i+1}. {file}")
            if len(pdf_files) > 10:
                print(f"   ... et {len(pdf_files)-10} autres fichiers.")

    def charger_dataframe(self, fichier):
        """Charge un fichier CSV s'il existe, sinon retourne un DataFrame vide."""
        if os.path.exists(fichier):
            return pd.read_csv(fichier, dtype=str)
        return pd.DataFrame()

    def validate_dataframe(self, df, required_columns):
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

    def lire_facture(self, chemin_fichier):
        """Lit une facture (PDF ou image)."""
        if chemin_fichier.lower().endswith((".jpg", ".jpeg", ".png")):
            return self.lire_facture_image(chemin_fichier)
        else:
            return self.lire_facture_pdf(chemin_fichier)

    def lire_facture_pdf(self, chemin_fichier):
        """Extrait le texte d'un fichier PDF à l'aide de pdfplumber."""
        try:
            with pdfplumber.open(chemin_fichier) as pdf:
                texte = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
            return texte
        except Exception as e:
            return f"Erreur lors de la lecture du fichier {chemin_fichier} : {str(e)}"

    def lire_facture_image(self, chemin_fichier):
        """Extrait le texte d'une image avec OCR (Tesseract)."""
        try:
            with Image.open(chemin_fichier) as img:
                img.verify()  # Vérifie si l'image est valide
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

    def analyser_facture_api(self, texte):
        """Envoie le texte d'une facture à l'API OpenAI pour extraire les informations clés."""
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
            }

    def analyser_facture_complete(self, chemin_fichier, utiliser_cache=True):
        """Analyse une facture en extrayant son texte depuis un PDF et en l'envoyant à l'API OpenAI."""
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
        texte = self.lire_facture(chemin_fichier)
        result_api = self.analyser_facture_api(texte)
        result_api = {k.lower(): v for k, v in result_api.items()}
        
        # Débogage: voir ce que l'API a renvoyé
        print(f"📋 Résultat API: {result_api}")
        
        # Vérifier si le fournisseur existe déjà dans notre base
        if "fournisseur" in result_api and result_api["fournisseur"] and result_api["fournisseur"] != "N/A":
            detected_fournisseur = result_api["fournisseur"]
            detected_iban = result_api.get("iban", "N/A")
            
            # Chercher une correspondance approximative dans la liste des fournisseurs
            found_match = False
            for index, row in self.df_ibans.iterrows():
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
                if not self.df_ibans[self.df_ibans["IBAN"] == detected_iban].empty:
                    print(f"⚠️ IBAN {detected_iban} existe déjà, pas d'ajout automatique.")
                else:
                    # Ajouter aux ibans connus
                    new_row = pd.DataFrame({"fournisseur": [detected_fournisseur], "IBAN": [detected_iban]})
                    self.df_ibans = pd.concat([self.df_ibans, new_row], ignore_index=True)
                    self.df_ibans.to_csv(self.IBAN_LISTE_CSV, index=False)
                    print(f"✅ Fournisseur automatiquement ajouté à la liste")
        
        # S'assurer que toutes les clés attendues sont présentes
        for key in ["date_facture", "fournisseur", "montant", "iban", "reference_facture"]:
            if key not in result_api or not result_api[key]:
                result_api[key] = "N/A"
        
        return result_api

    def lister_fichiers_factures(self):
        """Liste tous les fichiers PDF et images, y compris ceux dans les sous-dossiers."""
        fichiers = []
        print(f"\n🔍 DEBUG - LISTAGE DES FICHIERS")
        print(f"📂 Recherche dans le dossier: {self.DOSSIER_FACTURES}")
        
        if not os.path.exists(self.DOSSIER_FACTURES):
            print(f"❌ ERREUR: Le dossier {self.DOSSIER_FACTURES} n'existe pas!")
            return fichiers
            
        for root, dirs, files in os.walk(self.DOSSIER_FACTURES):
            print(f"📁 Sous-dossier: {os.path.relpath(root, self.DOSSIER_FACTURES)}")
            for file in files:
                if file.lower().endswith((".pdf", ".jpg", ".jpeg", ".png")):
                    rel_path = os.path.relpath(os.path.join(root, file), self.DOSSIER_FACTURES)
                    fichiers.append(rel_path)
                    print(f"   📄 Trouvé: {rel_path}")
        
        print(f"✅ Total: {len(fichiers)} fichiers trouvés.")
        return fichiers

    def verifier_nouvelles_factures(self):
        """Vérifie la présence de nouvelles factures et gère les fichiers déplacés."""
        # Charger la base de données existante
        db = charger_base_donnees()
        
        # Set pour suivre les factures trouvées dans cette analyse
        factures_trouvees = set()
        
        # Lister tous les fichiers actuels
        fichiers_actuels = self.lister_fichiers_factures()
        
        # Nombre de nouvelles factures détectées
        nouvelles_factures = 0
        
        for fichier in fichiers_actuels:
            chemin_fichier = os.path.join(self.DOSSIER_FACTURES, fichier)
            
            # Ajouter ou mettre à jour dans la base de données
            infos_extraites = {}  # Initialiser avec un dictionnaire vide
            
            # Vérifier si le fichier est déjà dans df_factures
            fichier_existant = self.df_factures[self.df_factures["fichier"] == fichier]
            
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
                    infos = self.analyser_facture_complete(chemin_fichier)
                    
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
                        self.df_factures = pd.concat([self.df_factures, pd.DataFrame([nouvelle_ligne])], ignore_index=True)
        
        # Marquer les factures qui n'ont pas été trouvées comme indisponibles
        marquer_factures_manquantes(db, factures_trouvees)
        
        # Sauvegarder la base de données mise à jour
        sauvegarder_base_donnees(db)
        
        if nouvelles_factures > 0:
            print(f"✅ {nouvelles_factures} nouvelles factures ajoutées au fichier CSV.")
            # Sauvegarder le DataFrame mis à jour
            self.df_factures.to_csv(self.CSV_FACTURES, index=False)
        else:
            print("✅ Aucune nouvelle facture détectée.")
            
        return nouvelles_factures

    def analyser_toutes_les_factures(self):
        """
        Analyse toutes les factures du dossier et retourne un DataFrame.
        Utilise le pipeline complet (API OpenAI avec fallback local).
        """
        fichiers = self.lister_fichiers_factures()
        resultats = []

        for fichier in fichiers:
            chemin_fichier = os.path.join(self.DOSSIER_FACTURES, fichier)
            infos = self.analyser_facture_complete(chemin_fichier)
            infos["fichier"] = fichier
            resultats.append(infos)

        df_resultats = pd.DataFrame(resultats)
        return df_resultats

    def recharger_fichiers(self, forcer_reanalyse=False):
        """Recharge les fichiers et les informations des factures."""
        global df_factures
        print("🔄 Rechargement des données depuis le CSV...")
        print(self.df_factures.head())  # Afficher les premières lignes pour vérification
        if forcer_reanalyse:
            print("🔄 Rechargement complet de toutes les factures...")
            self.df_factures = self.analyser_toutes_les_factures()  # 🔄 Recharger toutes les factures
        else:
            print("🔍 Vérification des nouvelles factures uniquement...")
            self.verifier_nouvelles_factures()  # 🔍 Vérifie seulement les nouvelles factures

        # Charger la base de données pour avoir accès aux informations sur les fichiers manquants
        db = charger_base_donnees()
        
        # Créer un dict des statuts de disponibilité
        disponibilite = {db["factures"][id_facture]["chemin"].split("/")[-1]: db["factures"][id_facture]["disponible"] 
                        for id_facture in db["factures"] if "chemin" in db["factures"][id_facture]}

        self.df_factures["date_facture"] = self.df_factures["date_facture"].astype(str)  # Convertir en string
        
        # Ajouter la colonne "disponible" si elle n'existe pas encore
        if "disponible" not in self.df_factures.columns:
            self.df_factures["disponible"] = True
        
        # Mettre à jour le statut de disponibilité
        for idx, row in self.df_factures.iterrows():
            fichier = row["fichier"]
            if fichier in disponibilite:
                self.df_factures.at[idx, "disponible"] = disponibilite[fichier]    
        
        # Ajouter la colonne "fournisseur" en fusionnant avec la liste des fournisseurs si elle n'existe pas
        if "fournisseur" not in self.df_factures.columns:
            df_ibans_local = pd.read_csv(self.IBAN_LISTE_CSV)
            self.df_factures = pd.merge(self.df_factures, df_ibans_local, left_on="iban", right_on="IBAN", how="left")
            self.df_factures["fournisseur"] = self.df_factures["fournisseur"].fillna("N/A")

        # Extraire l'année et le mois des dates
        if "annee" not in self.df_factures.columns or "mois" not in self.df_factures.columns:
            self.df_factures["annee"] = self.df_factures["date_facture"].apply(
                lambda x: x.split("-")[2] if isinstance(x, str) and len(x.split("-")) == 3 else "N/A"
            )
            self.df_factures["mois"] = self.df_factures["date_facture"].apply(
                lambda x: x.split("-")[1] if isinstance(x, str) and len(x.split("-")) == 3 else "N/A"
            )

        if "paiement" not in self.df_factures.columns:
            self.df_factures["paiement"] = "❌"
        else:
            self.df_factures["paiement"] = self.df_factures["paiement"].fillna("❌").replace("", "❌")
            
        # Tri des factures par date (décroissant)
        self.df_factures = self.df_factures.sort_values(by=["date_facture"], ascending=False)

        # Enregistrer toutes les modifications
        self.df_factures.to_csv(self.CSV_FACTURES, index=False)
        print("✅ Toutes les modifications ont été enregistrées.")
        
        return self.df_factures

    def get_factures_filtrees(self, annee=None, mois=None, fournisseur=None, statut_paiement=None, montant_min=None):
        """Filtre les factures selon les critères spécifiés."""
        df_filtree = self.df_factures.copy()
        
        # Filtre par année
        if annee and annee != "---":
            df_filtree = df_filtree[df_filtree["annee"] == annee]
            
        # Filtre par mois
        if mois and mois != "---":
            df_filtree = df_filtree[df_filtree["mois"] == mois]
            
        # Filtre par fournisseur
        if fournisseur and fournisseur != "Tous":
            df_filtree = df_filtree[df_filtree["fournisseur"] == fournisseur]
            
        # Filtre par statut de paiement
        if statut_paiement:
            if statut_paiement == "Payés":
                df_filtree = df_filtree[df_filtree["paiement"].str.contains("✔️", na=False)]
            elif statut_paiement == "Non payés":
                df_filtree = df_filtree[~df_filtree["paiement"].str.contains("✔️", na=False)]
                
# Filtre par montant minimum
        if montant_min:
            try:
                montant_min = float(str(montant_min).replace(",", "."))
                # Convertir les montants en valeurs numériques
                df_filtree["montant_num"] = df_filtree["montant"].str.replace("€", "").str.replace(",", ".").astype(float)
                df_filtree = df_filtree[df_filtree["montant_num"] >= montant_min]
                df_filtree = df_filtree.drop(columns=["montant_num"])
            except ValueError:
                # Ignorer si la valeur n'est pas un nombre valide
                pass
        
        # Tri des factures par date (décroissant)
        df_filtree = df_filtree.sort_values(by=["date_facture"], ascending=False)
        
        return df_filtree

    def creer_interface(self, parent=None):
        """Crée l'interface graphique d'analyse des factures."""
        if parent is None:
            # Créer une nouvelle fenêtre
            root = tk.Tk()
            root.title("Analyse des Factures")
        else:
            # Utiliser le parent fourni
            root = parent
            
        # Configuration de la zone de gauche (tableau, zone de texte et boutons)
        frame_left = ttk.Frame(root)
        frame_left.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        # Définition des colonnes
        tree = ttk.Treeview(frame_left, columns=("Date","fournisseur", "Montant", "IBAN", "Référence", "Fichier", "paiement"), show="headings")
        tree.heading("Date", text="Date")
        tree.heading("fournisseur", text="Fournisseur")
        tree.column("fournisseur", width=150, anchor="w")
        tree.heading("Montant", text="Montant")
        tree.heading("IBAN", text="IBAN")
        tree.heading("Référence", text="Référence")
        tree.heading("Fichier", text="Fichier")
        tree.heading("paiement", text="Paiement")
        tree.column("paiement", width=100, anchor="center")

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

        # Liaison des événements
        tree.bind("<<TreeviewSelect>>", lambda e: self.on_facture_selection(e, tree, apercu_canvas))
        apercu_canvas.bind("<ButtonPress-1>", lambda e: self.loupe_on(e, apercu_canvas))
        apercu_canvas.bind("<B1-Motion>", lambda e: self.update_loupe(e, apercu_canvas))
        apercu_canvas.bind("<ButtonRelease-1>", lambda e: self.loupe_off(e))
        apercu_canvas.bind("<Configure>", lambda e: self.update_apercu(e, tree, apercu_canvas))

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

        # Liste pour stocker les variables tkinter
        ui_elements = {
            "tree": tree,
            "text_area": text_area,
            "apercu_canvas": apercu_canvas,
            "annee_var": annee_var,
            "mois_var": mois_var,
            "fournisseur_var": fournisseur_var,
            "statut_var": statut_var,
            "montant_min_var": montant_min_var,
            "liste_annees": liste_annees,
            "liste_mois": liste_mois,
            "liste_fournisseurs": liste_fournisseurs
        }

        # Fonction pour mettre à jour les listes des filtres
        def mettre_a_jour_listes_filtres():
            """Met à jour les listes déroulantes des filtres avec les valeurs disponibles."""
            # Années
            annees = sorted(list(set(self.df_factures["annee"].dropna().tolist())), reverse=True)
            if "---" not in annees:
                annees.insert(0, "---")
            liste_annees["values"] = annees
            
            # Mois
            mois = sorted(list(set(self.df_factures["mois"].dropna().tolist())))
            if "---" not in mois:
                mois.insert(0, "---")
            liste_mois["values"] = mois
            
            # Fournisseurs
            fournisseurs = sorted(list(set(self.df_factures["fournisseur"].dropna().tolist())))
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

        # Fonction pour appliquer les filtres
        def appliquer_filtres():
            """Applique les filtres sélectionnés aux factures et met à jour l'affichage."""
            # Supprimer toutes les entrées du tableau
            for item in tree.get_children():
                tree.delete(item)
            
            # Récupérer les valeurs actuelles des filtres
            annee_filtre = annee_var.get()
            mois_filtre = mois_var.get()
            fournisseur_filtre = fournisseur_var.get()
            statut_filtre = statut_var.get()
            montant_min_filtre = montant_min_var.get()
            
            # Obtenir les factures filtrées
            df_filtree = self.get_factures_filtrees(
                annee=annee_filtre if annee_filtre != "---" else None,
                mois=mois_filtre if mois_filtre != "---" else None,
                fournisseur=fournisseur_filtre if fournisseur_filtre != "Tous" else None,
                statut_paiement=statut_filtre if statut_filtre != "Tous" else None,
                montant_min=montant_min_filtre
            )
            
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
            text_area.insert("end", f"📅 Filtres: Année {annee_filtre}, Mois {mois_filtre}\n")
            
            # Calcul du total des montants
            if not df_filtree.empty:
                try:
                    montants = df_filtree["montant"].str.replace("€", "").str.replace(",", ".").astype(float)
                    total = montants.sum()
                    text_area.insert("end", f"💰 Total : {total:.2f} €\n")
                except:
                    pass

        # Fonction pour réinitialiser les filtres
        def reinitialiser_filtres():
            """Réinitialise tous les filtres à leurs valeurs par défaut."""
            annee_var.set("---")
            mois_var.set("---")
            fournisseur_var.set("Tous")
            statut_var.set("Tous")
            montant_min_var.set("")
            appliquer_filtres()

        # Connecter les fonctions aux boutons
        ttk.Button(boutons_frame, text="Appliquer filtres", command=appliquer_filtres).pack(side="left", padx=5)
        ttk.Button(boutons_frame, text="Réinitialiser", command=reinitialiser_filtres).pack(side="left", padx=5)
        
        # Charger les icônes redimensionnées (si disponibles)
        icones = self.charger_icones()
        
        # Boutons d'action placés sous la zone de texte
        frame_buttons = ttk.Frame(frame_left)
        frame_buttons.pack(side="top", fill="x", pady=10)

        # Bouton Ouvrir Facture
        btn_ouvrir = ttk.Button(
            frame_buttons, 
            text="Ouvrir Facture", 
            image=icones.get("icone_ouvrir"), 
            compound="left", 
            command=lambda: self.ouvrir_fichier(tree)
        )
        btn_ouvrir.grid(row=0, column=0, padx=10, pady=5)

        # Bouton Relire Facture
        btn_relire = ttk.Button(
            frame_buttons, 
            text="Relire Facture", 
            image=icones.get("icone_relire"), 
            compound="left", 
            command=lambda: self.relire_fichier(tree, text_area, apercu_canvas)
        )
        btn_relire.grid(row=0, column=1, padx=10, pady=5)

        # Bouton Recharger Fichiers
        btn_recharger = ttk.Button(
            frame_buttons, 
            text="Recharger les fichiers", 
            image=icones.get("icone_recharger"), 
            compound="left", 
            command=lambda: self.on_recharger_fichiers(tree, True, mettre_a_jour_listes_filtres)
        )
        btn_recharger.grid(row=1, column=0, padx=10, pady=5)

        # Bouton Vérifier nouvelles factures
        btn_verifier = ttk.Button(
            frame_buttons, 
            text="🔄 Vérifier nouvelles factures", 
            command=lambda: self.on_verifier_nouvelles_factures(mettre_a_jour_listes_filtres)
        )
        btn_verifier.grid(row=1, column=0, padx=10, pady=5)

        # Bouton Réaliser Virement
        btn_virement = ttk.Button(
            frame_buttons, 
            text="Réaliser Virement", 
            image=icones.get("icone_virement"), 
            compound="left", 
            command=lambda: self.realiser_virement(tree)
        )
        btn_virement.grid(row=1, column=1, padx=10, pady=5)

        # Bouton Fermer
        btn_fermer = ttk.Button(
            frame_buttons, 
            text="Fermer", 
            image=icones.get("icone_fermer"), 
            compound="left", 
            command=lambda: self.fermer(root)
        )
        btn_fermer.grid(row=1, column=2, padx=10, pady=5)

        # Initialisation initiale des filtres
        mettre_a_jour_listes_filtres()
        appliquer_filtres()

        return root, ui_elements

    # Méthodes pour la gestion de l'interface utilisateur
    def charger_icones(self):
        """Charge les icônes depuis le dossier des icônes."""
        icones = {}
        
        ICONS_PATH = "/Users/vincentperreard/script contrats/icons/"
        
        # Liste des noms de fichiers d'icônes à charger
        icon_files = {
            "icone_ouvrir": "open.png",
            "icone_relire": "read.png",
            "icone_recharger": "refresh.png",
            "icone_virement": "transfer.png",
            "icone_fermer": "close.png"
        }
        
        for icon_name, file_name in icon_files.items():
            try:
                image = Image.open(os.path.join(ICONS_PATH, file_name))
                image = image.resize((18, 18), Image.Resampling.LANCZOS)  # Redimensionne proprement
                icones[icon_name] = ImageTk.PhotoImage(image)
            except Exception as e:
                print(f"⚠️ Erreur lors du chargement de l'icône {file_name}: {e}")
                icones[icon_name] = None
                
        return icones

    def afficher_apercu_fichier(self, chemin_fichier, apercu_canvas):
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
            img = img.resize((canvas_width, canvas_height), Image.Resampling.LANCZOS)
            self.apercu_img = img  # Garde l'image originale pour le zoom
            self.img_tk = ImageTk.PhotoImage(img)
            
            apercu_canvas.create_image(canvas_width//2, canvas_height//2, anchor="center", image=self.img_tk)

        except Exception as e:
            print(f"Erreur lors de l'aperçu du fichier : {e}")

    # Méthodes de gestion des événements
    def on_facture_selection(self, event, tree, apercu_canvas):
        """Affiche l'aperçu du fichier sélectionné lors de la sélection dans le tableau."""
        selected_item = tree.selection()
        if selected_item:
            fichier = tree.item(selected_item[0], "values")[5]  # Index 5 = colonne "Fichier"
            chemin_fichier = os.path.join(self.DOSSIER_FACTURES, fichier)
            self.afficher_apercu_fichier(chemin_fichier, apercu_canvas)

    def update_apercu(self, event, tree, apercu_canvas):
        """Met à jour l'aperçu du fichier sélectionné lors du redimensionnement."""
        selected_item = tree.selection()
        if selected_item:
            fichier = tree.item(selected_item[0], "values")[5]  # Colonne "Fichier"
            chemin_fichier = os.path.join(self.DOSSIER_FACTURES, fichier)
            self.afficher_apercu_fichier(chemin_fichier, apercu_canvas)

    def loupe_on(self, event, apercu_canvas):
        """Active la loupe en créant une fenêtre flottante qui s'affiche où la souris clique."""
        if self.loupe_window is None:
            self.loupe_window = tk.Toplevel(apercu_canvas.master)
            self.loupe_window.overrideredirect(True)  # Supprime la barre de titre
            self.loupe_window.resizable(False, False)
            self.loupe_canvas = tk.Canvas(self.loupe_window, width=200, height=200, bg="black")
            self.loupe_canvas.pack()

        self.update_loupe(event, apercu_canvas)
        
    def loupe_off(self, event):
        """Désactive la loupe et ferme la fenêtre flottante."""
        if self.loupe_window:
            self.loupe_window.destroy()
            self.loupe_window = None

    def update_loupe(self, event, apercu_canvas):
        """Met à jour la loupe pour qu'elle s'affiche dans une fenêtre flottante en suivant la souris."""
        if self.apercu_img and self.loupe_window:
            zoom_size = 100  # Taille de la zone capturée
            zoom_factor = 2  # Facteur de zoom
            x, y = event.x, event.y

            img_width, img_height = self.apercu_img.size
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
            cropped_part = self.apercu_img.crop((crop_x1, crop_y1, crop_x2, crop_y2))

            paste_x, paste_y = max(0, -x1), max(0, -y1)
            zoom_area.paste(cropped_part, (paste_x, paste_y))

            # Appliquer le zoom
            zoom_area = zoom_area.resize((zoom_size * zoom_factor, zoom_size * zoom_factor), Image.Resampling.LANCZOS)
            zoom_tk = ImageTk.PhotoImage(zoom_area)

            # Positionner la fenêtre de la loupe près de la souris
            self.loupe_window.geometry(f"200x200+{event.x_root-100}+{event.y_root-100}")
            self.loupe_canvas.delete("all")
            self.loupe_canvas.create_image(100, 100, image=zoom_tk, anchor="center")
            self.loupe_canvas.image = zoom_tk  # Empêche la suppression par le garbage collector

    # Méthodes d'actions utilisateur
    def ouvrir_fichier(self, tree):
        """Ouvre le fichier PDF sélectionné avec le lecteur PDF par défaut."""
        selected_item = tree.selection()
        if selected_item:
            fichier = tree.item(selected_item[0], "values")[5]  # Colonne "Fichier"
            chemin_fichier = os.path.join(self.DOSSIER_FACTURES, fichier)
            
            if os.path.exists(chemin_fichier):
                subprocess.run(["open", "-a", "PDF Expert", chemin_fichier])  # Ouvre avec PDF Expert
            else:
                print(f"⚠️ Fichier introuvable : {chemin_fichier}")

    def relire_fichier(self, tree, text_area, apercu_canvas):
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
            chemin_fichier = os.path.join(self.DOSSIER_FACTURES, fichier)
            
            # Effacer le contenu de la zone de texte si c'est le premier élément
            if item_id == selected_items[0]:
                text_area.delete("1.0", "end")
                text_area.insert("end", f"🔄 Analyse de {len(selected_items)} facture(s)...\n\n")
            
            if os.path.exists(chemin_fichier):
                text_area.insert("end", f"📄 Analyse de: {fichier}\n")
                
                # Forcer l'analyse complète en désactivant le cache
                infos = self.analyser_facture_complete(chemin_fichier, utiliser_cache=False)
                
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
                    index = self.df_factures[self.df_factures['fichier'] == fichier].index
                    if not index.empty:
                        self.df_factures.loc[index, 'date_facture'] = infos.get('date_facture', self.df_factures.loc[index, 'date_facture'].values[0])
                        self.df_factures.loc[index, 'fournisseur'] = infos.get('fournisseur', self.df_factures.loc[index, 'fournisseur'].values[0])
                        # Correction pour le montant
                        if 'montant' in infos and infos['montant'] != 'N/A':
                            try:
                                self.df_factures.loc[index, 'montant'] = float(infos['montant'])
                            except (ValueError, TypeError):
                                # Ne pas modifier si la conversion échoue
                                pass
                        self.df_factures.loc[index, 'iban'] = infos.get('iban', self.df_factures.loc[index, 'iban'].values[0])
                        self.df_factures.loc[index, 'reference_facture'] = infos.get('reference_facture', self.df_factures.loc[index, 'reference_facture'].values[0])
                else:
                    text_area.insert("end", f"⚠️ Aucune information n'a pu être extraite pour: {fichier}\n\n")
            else:
                text_area.insert("end", f"⚠️ Fichier introuvable: {chemin_fichier}\n\n")
        
        # Sauvegarder les modifications du DataFrame après avoir traité toutes les factures
        self.df_factures.to_csv(self.CSV_FACTURES, index=False)
        text_area.insert("end", "\n✅ Données mises à jour et sauvegardées.\n")
        
        # Afficher l'aperçu du premier fichier sélectionné
        if selected_items:
            first_item = selected_items[0]
            first_fichier = tree.item(first_item, "values")[5]
            first_chemin = os.path.join(self.DOSSIER_FACTURES, first_fichier)
            if os.path.exists(first_chemin):
                self.afficher_apercu_fichier(first_chemin, apercu_canvas)
                
                
def on_recharger_fichiers(self, tree, forcer_reanalyse=False, mettre_a_jour_listes_filtres=None):
        """Recharge les fichiers et met à jour l'interface."""
        # Recharger les données
        self.recharger_fichiers(forcer_reanalyse)
        
        # Mettre à jour les listes des filtres
        if mettre_a_jour_listes_filtres:
            mettre_a_jour_listes_filtres()
            
        # Appliquer les filtres pour rafraîchir l'affichage
        appliquer_filtres()
    
def on_verifier_nouvelles_factures(self, mettre_a_jour_listes_filtres=None):
    """Vérifie les nouvelles factures en arrière-plan et met à jour l'interface."""
    def task():
        self.verifier_nouvelles_factures()
        # Mise à jour des filtres et de l'interface
        if mettre_a_jour_listes_filtres:
            mettre_a_jour_listes_filtres()
    
    # Lancer la vérification dans un thread séparé
    thread = threading.Thread(target=task)
    thread.daemon = True
    thread.start()
    
def fermer(self, root):
    """Ferme proprement toutes les fenêtres Tkinter."""
    if self.loupe_window:
        self.loupe_window.destroy()
        self.loupe_window = None

    for window in root.winfo_children():
        window.destroy()  # Détruit toutes les fenêtres enfants

    # Si c'est une fenêtre principale (Tk), la fermer
    if isinstance(root, tk.Tk):
        root.destroy()

def realiser_virement(self, tree):
    """Ouvre une fenêtre pour réaliser un virement et met à jour le statut de paiement."""
    from generer_virement import generer_xml_virements, envoyer_virement_vers_lcl
    
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
    montant_facture = values[2] if values[2] else "0.00 €"
    iban_facture = values[3] if values[3] else "N/A"
    reference_facture = values[4] if values[4] else "N/A"

    # **Vérifier si l'IBAN est déjà enregistré**
    liste_ibans = self.df_ibans.to_dict(orient="records")
    ibans_connus = [entry["IBAN"] for entry in liste_ibans]
    iban_inconnu = iban_facture not in ibans_connus and iban_facture != "N/A"

    # **Créer la fenêtre de virement (ajustement taille)**
    virement_window = tk.Toplevel()
    virement_window.title("Réaliser un Virement")
    virement_window.geometry("750x280")  # ✅ Ajustement de la hauteur
    virement_window.transient()
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
    btn_choisir_iban = ttk.Button(frame, text="🔍", command=lambda: self.choisir_iban(virement_window, iban_var))
    btn_choisir_iban.grid(row=2, column=2, padx=5, pady=5)

    # **Bouton "Ajouter nouvel IBAN" aligné avec la loupe**
    if iban_inconnu:
        lbl_alert = ttk.Label(frame, text="⚠️ Cet IBAN est inconnu de la liste", foreground="red")
        lbl_alert.grid(row=3, column=1, padx=5, pady=2, sticky="w")

        btn_ajouter_iban = ttk.Button(frame, text="Ajouter nouvel IBAN", command=lambda: self.ajouter_iban(virement_window, iban_var.get()))
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
            # 🔹 Préparation du virement
            virement_data = [{
                "beneficiaire": values[1],  # Nom du fournisseur
                "iban": iban,
                "montant": float(montant.replace("€", "").replace(",", ".")),
                "objet": reference
            }]
            
            # 🔹 Génération du fichier XML
            fichier_xml = generer_xml_virements(virement_data)
            
            # 🔹 Envoi du virement
            envoyer_virement_vers_lcl(fichier_xml)
            
            # ✅ Mise à jour de l'interface après un virement réussi
            new_values = list(values)
            new_values[6] = f"✔️ {date_virement}"  # Met à jour le statut de paiement
            tree.item(item_id, values=tuple(new_values))

            # Mise à jour du DataFrame
            fichier = values[5]  # Nom du fichier
            self.df_factures.loc[self.df_factures["fichier"] == fichier, "paiement"] = f"✔️ {date_virement}"
            self.df_factures.to_csv(self.CSV_FACTURES, index=False)

            messagebox.showinfo("Succès", f"Virement de {montant} validé.")
            virement_window.destroy()
            
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite : {e}")

    # Boutons d'action
    btn_valider = ttk.Button(frame, text="Valider Virement", command=valider_virement)
    btn_valider.grid(row=5, column=1, pady=10)

    # Bouton "Annuler"
    btn_annuler = ttk.Button(frame, text="Annuler", command=virement_window.destroy)
    btn_annuler.grid(row=5, column=2, pady=10)

def choisir_iban(self, parent_window, iban_var):
    """Ouvre une fenêtre pour sélectionner un IBAN depuis la liste des fournisseurs."""
    iban_window = tk.Toplevel(parent_window)
    iban_window.title("Choisir un IBAN")
    iban_window.geometry("450x250")
    iban_window.transient(parent_window)
    iban_window.grab_set()

    # Récupérer les IBANs triés par ordre alphabétique du fournisseur
    ibans = self.df_ibans.to_dict(orient="records")
    ibans = sorted(ibans, key=lambda x: x.get("fournisseur", ""))

    lb = tk.Listbox(iban_window, height=10, width=50)
    lb.pack(padx=10, pady=10, fill="both", expand=True)

    for item in ibans:
        lb.insert("end", f"{item['fournisseur']} - {item['IBAN']}")

    def selectionner_iban():
        """Sélectionne l'IBAN et ferme la fenêtre."""
        selection = lb.curselection()
        if selection:
            iban_var.set(ibans[selection[0]]["IBAN"])
        iban_window.destroy()

    # Boutons
    btn_select = ttk.Button(iban_window, text="Sélectionner", command=selectionner_iban)
    btn_select.pack(pady=5)

    btn_cancel = ttk.Button(iban_window, text="Annuler", command=iban_window.destroy)
    btn_cancel.pack(pady=5)

def ajouter_iban(self, parent_window, iban_initial=""):
    """Ouvre une fenêtre pour enregistrer un nouvel IBAN."""
    enregistrer_window = tk.Toplevel(parent_window)
    enregistrer_window.title("Enregistrer un nouvel IBAN")
    enregistrer_window.geometry("500x160")
    enregistrer_window.transient(parent_window)
    enregistrer_window.grab_set()

    frame = ttk.Frame(enregistrer_window, padding=10)
    frame.pack(fill="both", expand=True)

    ttk.Label(frame, text="Fournisseur :").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    fournisseur_var = tk.StringVar()
    fournisseur_entry = ttk.Entry(frame, textvariable=fournisseur_var, width=30)
    fournisseur_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    ttk.Label(frame, text="IBAN :").grid(row=1, column=0, padx=5, pady=5, sticky="w")
    nouvel_iban_var = tk.StringVar(value=iban_initial)
    nouvel_iban_entry = ttk.Entry(frame, textvariable=nouvel_iban_var, width=30)
    nouvel_iban_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    def sauvegarder_iban():
        """Enregistre le nouvel IBAN dans la liste."""
        fournisseur = fournisseur_var.get().strip()
        nouvel_iban = nouvel_iban_var.get().strip()

        if not fournisseur or not nouvel_iban:
            messagebox.showwarning("Champs incomplets", "Veuillez renseigner tous les champs.")
            return

        # Ajouter au DataFrame
        new_row = pd.DataFrame({"fournisseur": [fournisseur], "IBAN": [nouvel_iban]})
        self.df_ibans = pd.concat([self.df_ibans, new_row], ignore_index=True)
        
        # Sauvegarder dans le CSV
        self.df_ibans.to_csv(self.IBAN_LISTE_CSV, index=False)
        
        messagebox.showinfo("Succès", f"IBAN de {fournisseur} ajouté avec succès !")
        enregistrer_window.destroy()

    ttk.Button(frame, text="Enregistrer", command=sauvegarder_iban).grid(row=2, column=0, padx=5, pady=5)
    ttk.Button(frame, text="Annuler", command=enregistrer_window.destroy).grid(row=2, column=1, padx=5, pady=5)

# Point d'entrée principal du module
def main():
    """Fonction principale exécutée lorsque le script est appelé directement."""
    # Analyser les arguments de ligne de commande
    parser = argparse.ArgumentParser(description="Analyse de factures PDF")
    parser.add_argument("--dossier", help="Chemin du dossier contenant les factures")
    args = parser.parse_args()
    
    # Créer l'instance de l'analyseur
    analyseur = AnalyseFactures(args.dossier)
    
    # Créer l'interface graphique
    root, _ = analyseur.creer_interface()
    
    # Lancer l'application
    if isinstance(root, tk.Tk):
        root.mainloop()
    
# Fonction pour utiliser l'analyseur dans une interface existante
def integrer_analyseur_factures(parent_frame, dossier_factures=None):
    """
    Intègre l'analyseur de factures dans un cadre existant.
    
    Args:
        parent_frame: Cadre Tkinter existant dans lequel intégrer l'analyseur
        dossier_factures: Chemin du dossier contenant les factures (optionnel)
        
    Returns:
        analyseur: Instance de la classe AnalyseFactures
        ui_elements: Dictionnaire contenant les éléments d'interface principaux
    """
    # Créer l'instance de l'analyseur
    analyseur = AnalyseFactures(dossier_factures)
    
    # Créer l'interface dans le cadre parent
    _, ui_elements = analyseur.creer_interface(parent_frame)
    
    return analyseur, ui_elements

if __name__ == "__main__":
    main()