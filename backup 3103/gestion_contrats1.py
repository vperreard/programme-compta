import os
import json
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk  # 📌 Ajoute l'import pour la gestion des images
import pandas as pd
import subprocess
import xml.etree.ElementTree as ET
import send2trash  # Module pour envoyer à la corbeille
import webbrowser
from datetime import datetime
import PyPDF2  # Pour lire la date de fin du contrat dans le PDF
import re  # Pour extraire les dates
from config import get_file_path, file_paths
# Charger les credentials depuis le fichier JSON
with open('credentials.json', 'r') as file:
    credentials = json.load(file)

# 📂 Dossiers des contrats
MAR_FOLDER = get_file_path("pdf_mar", verify_exists=True, create_if_missing=True)
IADE_FOLDER = get_file_path("pdf_iade", verify_exists=True, create_if_missing=True)

# 📌 Fichiers de suivi des paiements et signatures
PAYMENT_FILE = os.path.join(os.path.dirname(MAR_FOLDER), "paiements.csv")
CACHE_FILE = os.path.join(os.path.dirname(MAR_FOLDER), "cache_contrats.json")
SIGNATURES_FILE = os.path.join(os.path.dirname(MAR_FOLDER), "signatures.json")


# Définir les chemins par défaut si les clés nécessaires ne sont pas présentes dans le fichier
file_paths.setdefault("excel_mar", "/Users/vincentperreard/Dropbox/SEL:SPFPL Mathilde/contrat remplacement MAR/MARS SELARL.xlsx")
file_paths.setdefault("excel_iade", "/Users/vincentperreard/Dropbox/SEL:SPFPL Mathilde/Compta SEL/contrats/CDD à faire signer/IADE remplaçants.xlsx")

# Vérification de la présence des fichiers nécessaires
if "excel_mar" not in file_paths or "excel_iade" not in file_paths:
    print("❌ Fichiers Excel manquants dans la configuration.")
else:
    print(f"Fichiers chargés : {file_paths['excel_mar']} et {file_paths['excel_iade']}")

# 📌 Namespace XML pour SEPA
NAMESPACE = "urn:iso:std:iso:20022:tech:xsd:pain.001.001.02"
ET.register_namespace("", NAMESPACE)

# 📌 Mapping des BIC pour les banques
bic_mapping = {
    "10107": "PSSTFRPPXXX",   # La Banque Postale
    "10278": "CMCIFR2A",      # CIC
    "11425": "CEPAFRPP142",   # Caisse d'Épargne Normandie
    "14518": "FTNOFRP1XXX",   # Arkéa Direct Bank (Fortuneo)
    "18306": "BREDFRPPXXX",   # BRED Banque Populaire
    "18706": "CEPAFRPP142",   # Caisse d'Épargne
    "30002": "CRLYFRPPXXX",   # LCL (Crédit Lyonnais)
    "30003": "SOGEFRPPXXX",   # Société Générale
    "30004": "BNPAFRPPXXX",   # BNP Paribas
    "30027": "CEPAFRPPXXX",   # Caisse d'Épargne
    "30056": "SOGEFRPPXXX",   # Société Générale
    "30066": "CCBPFRPPMTG",   # Banque Populaire Occitane
    "30087": "AGRIFRPPXXX",   # Crédit Agricole
    "30129": "CCBPFRPPMTG",   # Banque Populaire
    "30438": "CCBPFRPPXXX",   # Banque Populaire Val de France
    "30678": "CEPAFRPPXXX",   # Caisse d'Épargne Rhône Alpes
    "30688": "CEPAFRPPXXX",   # Caisse d'Épargne Côte d'Azur
    "30738": "CEPAFRPPXXX",   # Caisse d'Épargne Bretagne Pays de Loire
    "30778": "CEPAFRPPXXX",   # Caisse d'Épargne Languedoc-Roussillon
    "30938": "CCBPFRPPXXX",   # Banque Populaire Rives de Paris
    "31039": "CCBPFRPPXXX",   # Banque Populaire Nord
    "31539": "CCBPFRPPXXX",   # Banque Populaire Sud
    "39999": "INGBFRPPXXX",   # ING France
    "40618": "BANKFRPPXXX",
    "42559": "BPRIFRPPXXX",   # Banque de la Réunion
    "42599": "BPRIFRPPXXX",   # Banque de Nouvelle-Calédonie
    "50000": "CCBPFRPPXXX",   # Banque Populaire Alsace Lorraine Champagne
    "50088": "AXABFRPPXXX",   # AXA Banque
    "50100": "CCBPFRPPXXX",   # Banque Populaire du Nord
    "50783": "CMCIUS3MXXX",   # Crédit Mutuel
    "52933": "SOGEFRPPXXX",   # Société Générale (Particuliers)
    "53778": "BDFEFRPPXXX",   # Banque de France
    "55178": "CCBPFRPPXXX",   # Banque Populaire Méditerranée
    "55208": "CCBPFRPPXXX",   # Banque Populaire Bourgogne Franche Comté
    "55218": "CCBPFRPPXXX",   # Banque Populaire Auvergne Rhône Alpes
    "55278": "CCBPFRPPXXX",   # Banque Populaire Loire et Lyonnais
    "55348": "CCBPFRPPXXX",   # Banque Populaire Grand Ouest
    "55388": "CCBPFRPPXXX",   # Banque Populaire Val de France
    "65729": "PSSTFRPPXXX",   # Banque Postale
    "84998": "MONAFRPPXXX",   # Monabanq
    "88258": "CCBPFRPPXXX",   # Banque Populaire Occitane
    "88897": "REVOFRP2XXX",   # Revolut France
    "88899": "QNBPFRPPXXX",   # Qonto
    "89002": "TREEFRP1XXX",   # Treezor
    "89004": "MAXEFRPPXXX",   # Ma French Bank (groupe La Banque Postale)
    "89009": "NIKIFRPPXXX",   # Nickel (BNP Paribas)
    "89013": "BUNQNL2AXXX",   # Bunq
    "89015": "LYDEFRPPXXX",   # Lydia
    "89025": "HDBKFRPPXXX",   # Hello Bank! (BNP Paribas)
    "89139": "FRLYFRPPXXX",   # Fortuneo Banque
    "89169": "PAYNFRP1XXX",   # Anytime (Orange Bank)
    "89201": "SPELFRPPXXX",   # Shine (Société Générale)
    "89889": "REVOFRP2XXX",   # Revolut France
    "89999": "N26PFRPPXXX"    # N26 France
}



### 📖 Lecture et stockage des contrats ###
def charger_cache():
    """Charge le cache des contrats si disponible."""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return {}  
    return {}

def sauvegarder_cache(contrats):
    """Sauvegarde les contrats en cache."""
    with open(CACHE_FILE, "w") as f:
        json.dump(contrats, f)

def vider_cache():
    """Vide le cache et affiche un message."""
    if os.path.exists(CACHE_FILE):
        os.remove(CACHE_FILE)
    messagebox.showinfo("Cache vidé", "Le cache des contrats a été supprimé.")


def extraire_date_fin(chemin_fichier):
    try:
        if not os.path.exists(chemin_fichier):
            print(f"🚨 Fichier introuvable : {chemin_fichier}")
            return "Date inconnue"

        with open(chemin_fichier, "rb") as f:
            lecteur = PyPDF2.PdfReader(f)
            texte_complet = "\n".join([page.extract_text() or "" for page in lecteur.pages])

        # 🔍 Dictionnaire des mois en français
        mois_francais = {
            "janvier": "01", "février": "02", "mars": "03", "avril": "04", "mai": "05", "juin": "06",
            "juillet": "07", "août": "08", "septembre": "09", "octobre": "10", "novembre": "11", "décembre": "12"
        }

        ### 🏥 1️⃣ Contrats MAR : Recherche des dates de début et de fin ###
        # 📌 Format "du yyyy-mm-dd à hh:mm au yyyy-mm-dd à hh:mm compris."
        match_mar_iso = re.search(r"du\s+(\d{4}-\d{2}-\d{2})\s+à.*?au\s+(\d{4}-\d{2}-\d{2})", texte_complet, re.IGNORECASE)
        if match_mar_iso:
            date_debut, date_fin = match_mar_iso.groups()
            return datetime.strptime(date_fin, "%Y-%m-%d").strftime("%d/%m/%Y")

        # 📌 Format "du 28 février 2025 à 8h au 28 février 2025 à 18h30 compris."
        match_mar_fr = re.search(r"du\s+(\d{1,2})\s+(\w+)\s+(\d{4})\s+à.*?au\s+(\d{1,2})\s+(\w+)\s+(\d{4})", texte_complet, re.IGNORECASE)
        if match_mar_fr:
            jour_debut, mois_debut, annee_debut, jour_fin, mois_fin, annee_fin = match_mar_fr.groups()
            mois_fin_num = mois_francais.get(mois_fin.lower(), "??")
            return f"{jour_fin.zfill(2)}/{mois_fin_num}/{annee_fin}"

        ### 🩺 2️⃣ Contrats IADE : Recherche des dates de début et de fin ###
        # 📌 Format "Le présent contrat débutera le 2025-02-21 et s’achèvera automatiquement le 2025-02-21 inclus."
        match_iade_iso = re.search(r"débutera le\s+(\d{4}-\d{2}-\d{2})\s+et s’achèvera automatiquement le\s+(\d{4}-\d{2}-\d{2})", texte_complet, re.IGNORECASE)
        if match_iade_iso:
            date_debut, date_fin = match_iade_iso.groups()
            return datetime.strptime(date_fin, "%Y-%m-%d").strftime("%d/%m/%Y")

        # 📌 Format "Le présent contrat débutera le vendredi 28 février 2025 et s’achèvera automatiquement le vendredi 28 février 2025 inclus."
        match_iade_fr = re.search(r"débutera le\s+\w+\s+(\d{1,2})\s+(\w+)\s+(\d{4})\s+et s’achèvera automatiquement le\s+\w+\s+(\d{1,2})\s+(\w+)\s+(\d{4})", texte_complet, re.IGNORECASE)
        if match_iade_fr:
            jour_debut, mois_debut, annee_debut, jour_fin, mois_fin, annee_fin = match_iade_fr.groups()
            mois_fin_num = mois_francais.get(mois_fin.lower(), "??")
            return f"{jour_fin.zfill(2)}/{mois_fin_num}/{annee_fin}"

        return "Date inconnue"

    except Exception as e:
        print(f"⚠️ Erreur lors de la lecture du PDF {chemin_fichier} : {e}")
        return "Date inconnue"


def charger_signatures():
    """Charge l'état des signatures des contrats depuis un fichier JSON."""
    if os.path.exists(SIGNATURES_FILE):
        try:
            with open(SIGNATURES_FILE, "r") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return {}  
    return {}

def sauvegarder_signature(fichier):
    """Marque un contrat comme signé et met à jour le fichier de suivi."""
    signatures = charger_signatures()
    signatures[fichier] = True  # ✅ Marque le fichier comme signé

    with open(SIGNATURES_FILE, "w") as f:
        json.dump(signatures, f)


def recuperer_contrat_docusign(tree, item):
    """ Demande confirmation et exécute le script pour récupérer le contrat signé via DocuSign. """
    values = list(tree.item(item, "values"))
    fichier = values[-1]  # Nom du fichier PDF
    
    confirmation = messagebox.askyesno("Téléchargement DocuSign",
        f"Voulez-vous chercher le contrat signé pour '{fichier}' dans DocuSign ?")

    if confirmation:
        try:
            # 📌 Exécuter le script recup_docusign.py avec le nom du fichier en argument
            subprocess.run(["python3", "/Users/vincentperreard/script contrats/recup_docusign.py", fichier], check=True)
            
            # 📌 Après récupération réussie, mettre à jour le cache
            mettre_a_jour_cache_apres_signature(fichier)

            # 📌 Mise à jour visuelle : remplacer la ❌ par ✅ dans l’interface
            values[-2] = "✅"
            values[-1] = f"{fichier}_docusign.pdf"  # 🔹 Met à jour le fichier associé
            tree.item(item, values=values)

            messagebox.showinfo("Succès", f"Le contrat signé '{fichier}' a été récupéré avec succès !")

        except subprocess.CalledProcessError:
            messagebox.showerror("Échec", f"Le contrat '{fichier}' n'a pas pu être récupéré depuis DocuSign.")


def menu_gestion_contrats():
    """Fenêtre de sélection MAR / IADE avec un design corrigé."""
    fenetre = tk.Toplevel()
    fenetre.title("Gestion des contrats")
    fenetre.geometry("400x250")  # 🔹 Taille ajustée
    fenetre.configure(bg="#f2f7ff")  # 🔹 Fond bleu clair

    # 📌 Titre
    label = tk.Label(fenetre, text="📋 Sélectionnez un type de contrat", font=("Arial", 14, "bold"),
                     bg="#4a90e2", fg="white", pady=10)
    label.pack(fill="x")

    # 📌 Bouton MAR
    bouton_mar = tk.Button(fenetre, text="📄 Contrats MAR", command=lambda: afficher_contrats("MAR"), width=25,
                           font=("Arial", 11, "bold"), bg="#007ACC", fg="black", activebackground="#005f99",
                           relief="raised", activeforeground="white")
    bouton_mar.pack(pady=10)

    # 📌 Bouton IADE
    bouton_iade = tk.Button(fenetre, text="📄 Contrats IADE", command=lambda: afficher_contrats("IADE"), width=25,
                            font=("Arial", 11, "bold"), bg="#007ACC", fg="black", activebackground="#005f99",
                            relief="raised", activeforeground="white")
    bouton_iade.pack(pady=10)

    # 📌 Bouton Retour
    bouton_retour = tk.Button(fenetre, text="🔙 Retour", command=fenetre.destroy, width=25,
                              font=("Arial", 11, "bold"), bg="#DDDDDD", fg="black", activebackground="#BBBBBB",
                              relief="raised", activeforeground="black")
    bouton_retour.pack(pady=10)

    # Centrage de la fenêtre
    fenetre.update_idletasks()
    x = (fenetre.winfo_screenwidth() - fenetre.winfo_width()) // 2
    y = (fenetre.winfo_screenheight() - fenetre.winfo_height()) // 2
    fenetre.geometry(f"+{x}+{y}")


    


def ajouter_bouton_gestion(menu_principal):
    bouton_gestion = tk.Button(menu_principal, text="📋 Gestion Contrats", command=menu_gestion_contrats)
    bouton_gestion.pack(pady=10)

def ajuster_colonnes(tree):
    """Réduit la largeur des colonnes 'Début' et 'Fin' pour voir tout le tableau."""
    tree.column("Début", width=100)  # 🔹 Réduction de la largeur
    tree.column("Fin", width=100)    # 🔹 Réduction de la largeur
    tree.column("Montant", width=120)  # 🔹 Augmente légèrement pour afficher "XXXX €"
    tree.column("Payé", width=60)  # 🔹 Ajuste pour la case à cocher
    tree.column("Fichier", width=80)  # 🔹 Ajuste la colonne Fichier


def recharger_selection(tree, dossier, type_contrat):
    """Recharge uniquement les contrats sélectionnés, met à jour le cache et force l'affichage des nouvelles valeurs."""
    selected_items = tree.selection()

    cache = charger_cache()  # 🛠 Charger le cache pour éviter l'erreur "NameError"

    if not selected_items:
        # Si aucune sélection, recharger tout
        if messagebox.askyesno("Recharger tout", "Aucune sélection. Recharger tous les contrats ?"):
            for item in tree.get_children():  # Effacer l'affichage
                tree.delete(item)
            cache.clear()  # 🔥 Efface le cache pour forcer la relecture des fichiers
            contrats = lire_contrats(dossier, type_contrat)
            for contrat in contrats:
                tree.insert("", "end", values=contrat)
            sauvegarder_cache(cache)  # ✅ Sauvegarde après mise à jour
        return

    for item in selected_items:
        values = list(tree.item(item, "values"))
        fichier = values[-1]  # 📂 Nom du fichier PDF
        
        chemin_fichier = os.path.join(dossier, fichier)
        
        # 🔥 Forcer la suppression du fichier dans le cache avant relecture
        if fichier in cache:
            del cache[fichier]

        # 🔍 Extraire les nouvelles données
        date_fin = extraire_date_fin(chemin_fichier)
        if date_fin != "Date inconnue":
            date_fin = " / ".join(date_fin.split("/"))  # Ajoute des espaces entre JJ / MM / AAAA
        montant_remplacement = extraire_montant_remplacement(chemin_fichier)

        # 📌 Mise à jour des valeurs
        new_values = list(values)
        new_values[1] = date_fin  # Met à jour la colonne "Fin"
        new_values[4] = f"{montant_remplacement} €" if montant_remplacement else "---"

        # 🔥 Mettre à jour l'affichage
        tree.item(item, values=new_values)

        # ✅ Sauvegarde dans le cache mis à jour
        cache[fichier] = new_values

    sauvegarder_cache(cache)  # ✅ Sauvegarde du cache après mise à jour



        
### 💳 Gestion des paiements ###
# ✅ Chargement des données
def charger_paiements():
    if os.path.exists(PAYMENT_FILE):
        try:
            with open(PAYMENT_FILE, "r") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return {}  
    return {}

def sauvegarder_paiements(data):
    with open(PAYMENT_FILE, "w") as f:
        json.dump(data, f)

def deduire_bic_depuis_code_banque(code_banque):
    """ Récupère le BIC en fonction du code banque """
    return bic_mapping.get(str(code_banque).zfill(5), "UNKNOWN BIC")

def generer_xml_virement(beneficiaire, iban, montant, objet="Virement SEPA"):
    """
    Génère un fichier XML de virement conforme à la norme SEPA pain.001.001.02.
    
    Args:
        beneficiaire (str): Nom du bénéficiaire du virement
        iban (str): IBAN du bénéficiaire
        montant (float): Montant du virement en euros
        objet (str, optional): Objet du virement (par défaut : "Virement SEPA")

    Returns:
        str: Chemin du fichier XML généré
    """
    date_execution = datetime.now().strftime("%Y%m%d")
    nom_fichier_xml = f"/Users/vincentperreard/Documents/Contrats/virements/virement_{beneficiaire}.xml"

    # 🔍 Extraction du code banque
    code_banque = iban.replace(" ", "")[4:9]  # Récupère les chiffres 4 à 9
    bic = deduire_bic_depuis_code_banque(code_banque)
    if bic == "UNKNOWN BIC":
        print(f"⚠️ BIC inconnu pour l'IBAN {iban}, assurez-vous qu'il est correct.")

    # 📌 Création de la structure XML
    document = ET.Element(f"{{{NAMESPACE}}}Document")
    pain = ET.SubElement(document, f"{{{NAMESPACE}}}pain.001.001.02")

    # 🏦 En-tête du fichier
    grpHdr = ET.SubElement(pain, f"{{{NAMESPACE}}}GrpHdr")
    ET.SubElement(grpHdr, f"{{{NAMESPACE}}}MsgId").text = f"SELARLANESTHES-{date_execution}"
    ET.SubElement(grpHdr, f"{{{NAMESPACE}}}CreDtTm").text = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    ET.SubElement(grpHdr, f"{{{NAMESPACE}}}BtchBookg").text = "true"
    ET.SubElement(grpHdr, f"{{{NAMESPACE}}}NbOfTxs").text = "1"
    ET.SubElement(grpHdr, f"{{{NAMESPACE}}}CtrlSum").text = f"{montant:.2f}"
    ET.SubElement(grpHdr, f"{{{NAMESPACE}}}Grpg").text = "MIXD"

    initgPty = ET.SubElement(grpHdr, f"{{{NAMESPACE}}}InitgPty")
    ET.SubElement(initgPty, f"{{{NAMESPACE}}}Nm").text = "SELARL ANESTHESISTES DE LA CLINIQUE MATHILDE"

    # 📜 Détails du paiement
    pmtInf = ET.SubElement(pain, f"{{{NAMESPACE}}}PmtInf")
    ET.SubElement(pmtInf, f"{{{NAMESPACE}}}PmtInfId").text = f"SELARLANESTHES-{date_execution}"
    ET.SubElement(pmtInf, f"{{{NAMESPACE}}}PmtMtd").text = "TRF"

    pmtTpInf = ET.SubElement(pmtInf, f"{{{NAMESPACE}}}PmtTpInf")
    svcLvl = ET.SubElement(pmtTpInf, f"{{{NAMESPACE}}}SvcLvl")
    ET.SubElement(svcLvl, f"{{{NAMESPACE}}}Cd").text = "SEPA"

    ET.SubElement(pmtInf, f"{{{NAMESPACE}}}ReqdExctnDt").text = datetime.now().strftime("%Y-%m-%d")

    # Débiteur
    dbtr = ET.SubElement(pmtInf, f"{{{NAMESPACE}}}Dbtr")
    ET.SubElement(dbtr, f"{{{NAMESPACE}}}Nm").text = "SELARL DES ANESTHESISTES DE LA CLINIQUE MATHILDE"

    dbtrAcct = ET.SubElement(pmtInf, f"{{{NAMESPACE}}}DbtrAcct")
    id_dbtrAcct = ET.SubElement(dbtrAcct, f"{{{NAMESPACE}}}Id")
    ET.SubElement(id_dbtrAcct, f"{{{NAMESPACE}}}IBAN").text = credentials['company_iban']


    dbtrAgt = ET.SubElement(pmtInf, f"{{{NAMESPACE}}}DbtrAgt")
    finInstnId = ET.SubElement(dbtrAgt, f"{{{NAMESPACE}}}FinInstnId")
    ET.SubElement(finInstnId, f"{{{NAMESPACE}}}BIC").text = "CRLYFRPP"

    # 💳 Bénéficiaire
    cdtTrfTxInf = ET.SubElement(pmtInf, f"{{{NAMESPACE}}}CdtTrfTxInf")

    # PmtId (Correction : ajout des sous-balises obligatoires)
    pmtId = ET.SubElement(cdtTrfTxInf, f"{{{NAMESPACE}}}PmtId")
    ET.SubElement(pmtId, f"{{{NAMESPACE}}}EndToEndId").text = f"SELARLANE-000001-{date_execution}"

    # PmtTpInf (Correction : cette balise doit être ajoutée)
    pmtTpInf = ET.SubElement(cdtTrfTxInf, f"{{{NAMESPACE}}}PmtTpInf")
    svcLvl = ET.SubElement(pmtTpInf, f"{{{NAMESPACE}}}SvcLvl")
    ET.SubElement(svcLvl, f"{{{NAMESPACE}}}Cd").text = "SEPA"

    # Montant (Correction : format correct avec devise EUR)
    amt = ET.SubElement(cdtTrfTxInf, f"{{{NAMESPACE}}}Amt")
    ET.SubElement(amt, f"{{{NAMESPACE}}}InstdAmt", Ccy="EUR").text = f"{montant:.2f}"

    # BIC du bénéficiaire
    cdtrAgt = ET.SubElement(cdtTrfTxInf, f"{{{NAMESPACE}}}CdtrAgt")
    finInstnId = ET.SubElement(cdtrAgt, f"{{{NAMESPACE}}}FinInstnId")
    ET.SubElement(finInstnId, f"{{{NAMESPACE}}}BIC").text = bic

    # Nom du bénéficiaire
    cdtr = ET.SubElement(cdtTrfTxInf, f"{{{NAMESPACE}}}Cdtr")
    ET.SubElement(cdtr, f"{{{NAMESPACE}}}Nm").text = beneficiaire

    # Compte du bénéficiaire
    cdtrAcct = ET.SubElement(cdtTrfTxInf, f"{{{NAMESPACE}}}CdtrAcct")
    id_cdtrAcct = ET.SubElement(cdtrAcct, f"{{{NAMESPACE}}}Id")
    ET.SubElement(id_cdtrAcct, f"{{{NAMESPACE}}}IBAN").text = iban.replace(" ", "")

    # Objet du virement
    rmtInf = ET.SubElement(cdtTrfTxInf, f"{{{NAMESPACE}}}RmtInf")
    ET.SubElement(rmtInf, f"{{{NAMESPACE}}}Ustrd").text = objet

    # Sauvegarde du fichier XML
    tree = ET.ElementTree(document)
    tree.write(nom_fichier_xml, encoding="utf-8", xml_declaration=True)

    print(f"✅ Fichier XML généré : {nom_fichier_xml}")
    return nom_fichier_xml

def recuperer_iban(file_path, beneficiaire, contrat_type="MAR"):
    """
    Récupère l'IBAN du bénéficiaire dans le fichier Excel correspondant.
    
    Args:
        file_path (str): Chemin du fichier Excel contenant les données.
        beneficiaire (str): Nom du bénéficiaire (remplaçant MAR ou IADE).
        contrat_type (str, optional): Type de contrat ("MAR" ou "IADE").
        
    Returns:
        str: IBAN du bénéficiaire ou None si introuvable.
    """
    # Sélection de la feuille et des colonnes en fonction du type de contrat
    if contrat_type == "MAR":
        sheet_name = "MARS Remplaçants"
        colonne_nom = "NOMR"
        colonne_prenom = "PRENOMR"
        colonne_iban = "IBAN"
    elif contrat_type == "IADE":
        sheet_name = "Coordonnées IADEs"
        colonne_nom = "NOMR"
        colonne_prenom = "PRENOMR"
        colonne_iban = "IBAN"
    else:
        print(f"❌ Erreur : Type de contrat inconnu {contrat_type}")
        return None

    try:
        # Charger le fichier Excel
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Vérifier si les colonnes existent
        if colonne_nom not in df.columns or colonne_prenom not in df.columns or colonne_iban not in df.columns:
            print(f"❌ Erreur : Une ou plusieurs colonnes sont introuvables dans la feuille '{sheet_name}'")
            return None

        # 🔍 Vérification que le bénéficiaire n'est pas vide
        if not beneficiaire.strip():
            print("❌ Erreur : Le bénéficiaire est vide.")
            return None

        # 🔍 Extraction du nom et prénom en gérant les noms composés
        beneficiaire_split = beneficiaire.strip().split(" ")
        if len(beneficiaire_split) < 2:
            print(f"❌ Erreur : Nom mal formaté pour '{beneficiaire}'.")
            return None

        nom_beneficiaire = beneficiaire_split[-1]  # Dernier mot = NOM
        prenom_beneficiaire = " ".join(beneficiaire_split[:-1])  # Tous les autres = PRÉNOM

        print(f"🔍 Recherche IBAN pour : Nom='{nom_beneficiaire}', Prénom='{prenom_beneficiaire}'")

        # Filtrer les résultats pour trouver l’IBAN correspondant
        data_match = df[
            (df[colonne_nom].str.strip().str.upper() == nom_beneficiaire.strip().upper()) &
            (df[colonne_prenom].str.strip().str.upper() == prenom_beneficiaire.strip().upper())
        ]

        if data_match.empty:
            print(f"❌ IBAN introuvable pour '{beneficiaire}'. Vérifiez l'orthographe.")
            return None

        # Récupérer l'IBAN
        iban = data_match.iloc[0][colonne_iban]

        if pd.isna(iban) or not str(iban).strip():
            print(f"❌ IBAN non renseigné pour '{beneficiaire}'.")
            return None

        print(f"✅ IBAN trouvé pour {beneficiaire}: {iban}")
        return iban.strip()

    except Exception as e:
        print(f"❌ Erreur lors de la récupération de l'IBAN : {e}")
        return None


def effectuer_virement(tree, dossier, contrat_type):
    """
    Génère un fichier de virement XML et l'enregistre pour l'export bancaire.
    
    Args:
        tree: Tableau Tkinter contenant les contrats sélectionnés.
        dossier (str): Dossier des contrats MAR ou IADE.
        contrat_type (str): Type de contrat ("MAR" ou "IADE").
    """
    selected_item = tree.selection()
    if not selected_item:
        print("⚠️ Veuillez sélectionner un contrat.")
        return

    # Récupérer les valeurs du contrat sélectionné
    values = tree.item(selected_item[0])["values"]
    fichier = values[-1]  # Nom du fichier PDF du contrat
    beneficiaire = values[3]  # Nom du remplaçant
    montant = values[4]  # Montant du virement

    # Vérification du montant
    if not montant or montant == "---":
        print(f"❌ Aucun montant défini pour {beneficiaire}.")
        return

    # Nettoyage du montant
    montant = montant.replace("€", "").strip().replace(",", ".")
    montant = float(montant)

    # Sélection du fichier Excel approprié
    excel_path = file_paths["excel_mar"] if contrat_type == "MAR" else file_paths["excel_iade"]

    # Récupération de l'IBAN du bénéficiaire
    iban = recuperer_iban(excel_path, beneficiaire, contrat_type)
    if not iban:
        print(f"❌ IBAN introuvable pour {beneficiaire}.")
        return

    # Générer le fichier XML de virement
    xml_file = generer_xml_virement(beneficiaire, iban, montant)

    if not xml_file:
        print("❌ Erreur lors de la génération du fichier XML.")
        return

    print(f"✅ Virement généré avec succès : {xml_file}")

# 📖 Lire et analyser les contrats
def lire_contrats(dossier, type_contrat):
    """Lit les contrats depuis le cache ou les analyse si non en cache."""
    contrats = []
    paiements = charger_paiements()
    signatures = charger_signatures()
    cache = charger_cache()

    for fichier in os.listdir(dossier):
        if not fichier.endswith(".pdf"):
            continue
        
        if fichier in cache:
            # ✅ Utiliser les données en cache si existantes
            contrat = cache[fichier]
        else:
            # 📄 Extraire les infos du contrat depuis le PDF
            parts = fichier.replace(".pdf", "").split("_")
            if len(parts) < 3:
                continue
            
            date_contrat = parts[0]
            try:
                date_contrat = datetime.strptime(date_contrat, "%Y%m%d").strftime("%d / %m / %Y")
            except ValueError:
                continue

            remplacé = parts[-1] if type_contrat == "MAR" else "---"
            remplaçant = parts[2] if type_contrat == "MAR" else parts[-1]

            date_fin = extraire_date_fin(os.path.join(dossier, fichier))

            montant = paiements.get(fichier, {}).get("Montant", "---")
            paye = "✅" if paiements.get(fichier, {}).get("Payé", False) else "❌"
            signature = "✅" if signatures.get(fichier, False) else "❌"

            contrat = [date_contrat, date_fin, remplacé, remplaçant, montant, paye, signature, fichier]
            cache[fichier] = contrat  # 🔥 Mettre en cache
        
        contrats.append(contrat)

    sauvegarder_cache(cache)  # 📌 Sauvegarde immédiate du cache mis à jour

    contrats.sort(reverse=True, key=lambda x: datetime.strptime(x[0], "%d / %m / %Y"))
    return contrats


def extraire_montant_remplacement(chemin_fichier):
    """Extrait le montant du remplacement dans un contrat MAR en recherchant la phrase contenant 'versera un montant forfaitaire'."""
    try:
        if not os.path.exists(chemin_fichier):
            return None

        with open(chemin_fichier, "rb") as f:
            lecteur = PyPDF2.PdfReader(f)
            texte_complet = ""
            for page in lecteur.pages:
                texte_complet += page.extract_text() or ""

        # 🔍 Recherche du montant dans la phrase type
        match_montant = re.search(r"versera .*? montant forfaitaire de\s+([\d\s]+)[€]?", texte_complet, re.IGNORECASE)
        if match_montant:
            montant = match_montant.group(1).replace(" ", "").strip()  # 🔥 Retire les espaces
            return float(montant)  # ✅ Convertit en float pour normalisation

        return None  # Aucun montant trouvé

    except Exception as e:
        print(f"⚠️ Erreur lors de la lecture du PDF {chemin_fichier} : {e}")
        return None



def on_treeview_double_click(event, tree):
    """ Gère les interactions au double-clic sur les colonnes Montant, Payé et Signature """
    selected_item = tree.selection()
    if not selected_item:
        return

    item = selected_item[0]
    column_id = tree.identify_column(event.x)  # Détecte la colonne cliquée
    colonne_index = int(column_id[1:]) - 1  # Convertit '#5' en index 4

    print(f"DEBUG: Clic dans colonne index {colonne_index}")

    if colonne_index == 4:  # ✅ Montant (5e colonne)
        modifier_montant(tree, item)
    elif colonne_index == 5:  # ✅ Payé (6e colonne)
        toggle_paiement(tree, item)
    elif colonne_index == 6:  # ✅ Signature (7e colonne)
        toggle_signature(tree, item)


def toggle_signature(tree, item):
    """ Gère l'ouverture de la récupération du contrat signé avec confirmation ou la dévalidation """
    values = list(tree.item(item, "values"))
    fichier = values[-1]  # 📌 Récupération du fichier associé
    signature = values[-2]  # 📌 État de la signature actuelle ("✅" ou "❌")

    popup = tk.Toplevel()
    popup.title("Gestion de la signature")
    popup.geometry("400x180")  # 📌 Ajustement de la taille

    # ✅ Ajout de la gestion de l'image du contrat
    try:
        contrat_icone = Image.open("/Users/vincentperreard/script contrats/icone_contrat.png")
        contrat_icone = contrat_icone.resize((40, 40), Image.LANCZOS)
        contrat_icone = ImageTk.PhotoImage(contrat_icone)
    except Exception as e:
        print(f"⚠️ Erreur lors du chargement de l'image : {e}")
        contrat_icone = None  # Si l'image ne charge pas, éviter un plantage

    if contrat_icone:
        tk.Label(popup, image=contrat_icone).pack(pady=10)
        popup.image = contrat_icone  # 📌 Éviter que l'image ne disparaisse

    boutons_frame = tk.Frame(popup)
    boutons_frame.pack(pady=10)

    if signature == "✅":
        # 📌 Le contrat est déjà signé → Proposer de dévalider
        tk.Label(popup, text="Le contrat est déjà marqué comme signé.", font=("Arial", 12, "bold")).pack(pady=10)

        def devalider_signature():
            """ Dévalide la signature et remet la croix rouge ❌ """
            values[-2] = "❌"
            tree.item(item, values=tuple(values))
            popup.destroy()

        bouton_devalider = tk.Button(boutons_frame, text="Dévalider", command=devalider_signature,
                                     font=("Arial", 10, "bold"), width=12, bg="#f44336", fg="white")
        bouton_devalider.pack(pady=5)

    else:
        # 📌 Le contrat n'est pas signé → Proposer d'aller chercher DocuSign
        message = f"Voulez-vous chercher le contrat signé pour :\n {fichier} dans DocuSign ?"
        tk.Label(popup, text=message, font=("Arial", 12, "bold"), wraplength=350).pack(pady=10)

        def recuperer_et_fermer():
            """ Lance la récupération et ferme la fenêtre """
            popup.destroy()  # ✅ Ferme la popup
            script_path = "/Users/vincentperreard/script contrats/recup_docusign.py"  # 📂 Chemin du script
            subprocess.Popen(["python3", script_path, fichier])  # ✅ Lance le script

        # ✅ Boutons Oui (gauche) & Non (droite)
        bouton_oui = tk.Button(boutons_frame, text="Oui", command=recuperer_et_fermer,
                               font=("Arial", 10, "bold"), width=10, bg="#DDDDDD", fg="black")
        bouton_oui.pack(side="left", padx=5)

        bouton_non = tk.Button(boutons_frame, text="Non", command=popup.destroy,
                               font=("Arial", 10, "bold"), width=10, bg="#DDDDDD", fg="black")
        bouton_non.pack(side="right", padx=5)

        bouton_oui.focus_set()
        popup.bind("<Return>", lambda event: recuperer_et_fermer())  # ✅ Entrée valide "Oui"




def demander_recuperation_signature(tree, item):
    """Affiche une popup pour confirmer la récupération de la signature via DocuSign."""
    values = list(tree.item(item, "values"))
    fichier = values[-1]  # Nom du fichier
    dossier = MAR_FOLDER if "MAR" in fichier else IADE_FOLDER  # Détermine le dossier du contrat

    popup = tk.Toplevel()
    popup.title("Signature DocuSign")
    popup.geometry("400x200")
    popup.configure(bg="#f2f7ff")

    tk.Label(popup, text="Voulez-vous chercher le contrat signé dans DocuSign ?", font=("Arial", 12, "bold"), bg="#f2f7ff").pack(pady=20)

    def lancer_recuperation():
        popup.destroy()
        recuperer_signature_docusign(tree, item, fichier, dossier)

    bouton_oui = tk.Button(popup, text="Oui", command=lancer_recuperation, width=10, bg="#DDDDDD", fg="black", font=("Arial", 10, "bold"))
    bouton_non = tk.Button(popup, text="Non", command=popup.destroy, width=10, bg="#DDDDDD", fg="black", font=("Arial", 10, "bold"))

    bouton_oui.pack(side="left", padx=20, pady=10, expand=True)
    bouton_non.pack(side="right", padx=20, pady=10, expand=True)

    popup.bind("<Return>", lambda event: lancer_recuperation())  # Valide par "Entrée"
    popup.transient(tree.winfo_toplevel())
    popup.grab_set()




def afficher_contrats(type_contrat, recharger=False):
    """ Affiche la liste des contrats avec un design amélioré et correction de l'importation """
    dossier = MAR_FOLDER if type_contrat == "MAR" else IADE_FOLDER

    fenetre = tk.Toplevel()
    fenetre.title(f"Liste des contrats {type_contrat}")
    fenetre.geometry("1100x550")
    fenetre.configure(bg="#f2f7ff")

    # ✅ Label titre
    label_titre = tk.Label(fenetre, text=f"📋 Liste des contrats {type_contrat}", font=("Arial", 14, "bold"),
                           bg="#4a90e2", fg="white")
    label_titre.pack(fill="x", pady=10)

    # ✅ Chargement des contrats
    contrats = lire_contrats(dossier, type_contrat)

    # ✅ Vérification si des contrats sont chargés
    if not contrats:
        messagebox.showinfo("Information", f"Aucun contrat trouvé dans {dossier}.")

    # ✅ Définition des colonnes
    colonnes = ["Début", "Fin", "Remplacé", "Remplaçant", "Montant", "Payé", "Signature", "Fichier"]
    tree = ttk.Treeview(fenetre, columns=colonnes, show="headings")

    tree.bind("<Double-1>", lambda event: on_treeview_double_click(event, tree))

    for col in colonnes:
        tree.heading(col, text=col)
        tree.column(col, anchor="center", width=120)


    # ✅ Insertion des contrats
    for contrat in contrats:
        tree.insert("", "end", values=contrat)

    tree.pack(expand=True, fill="both")

    # ✅ Cadre des boutons
    frame_boutons = tk.Frame(fenetre, bg="#f2f7ff")
    frame_boutons.pack(pady=10)

    # ✅ Correction couleur des boutons
    tk.Button(frame_boutons, text="📂 Ouvrir contrat", command=lambda: ouvrir_selection(tree, dossier), width=18, bg="#DDDDDD", fg="black", font=("Arial", 10, "bold")).pack(side="left", padx=5)
    tk.Button(frame_boutons, text="🔄 Recharger sélection", command=lambda: recharger_selection(tree, dossier, type_contrat), width=18, bg="#DDDDDD", fg="black", font=("Arial", 10, "bold")).pack(side="left", padx=5)
    tk.Button(frame_boutons, text="🗑 Supprimer", command=lambda: supprimer_contrat(tree, dossier), width=18,
              font=("Arial", 11, "bold"), bg="#f44336", fg="black", activebackground="#d32f2f", activeforeground="white",
              relief="raised").pack(side="left", padx=5)
    tk.Button(
        frame_boutons, text="💰 Effectuer un virement",
        command=lambda: effectuer_virement(tree, dossier, type_contrat),
        width=18, bg="#DDDDDD", fg="black", font=("Arial", 10, "bold")
    ).pack(side="left", padx=5)
    tk.Button(frame_boutons, text="🔙 Retour", command=fenetre.destroy, width=18,
              font=("Arial", 11, "bold"), bg="#DDDDDD", fg="black", activebackground="#BBBBBB", activeforeground="black",
              relief="raised").pack(side="left", padx=5)


    ajuster_colonnes(tree)



def toggle_paiement(tree, item):
    """ Ouvre une fenêtre popup pour valider ou invalider un paiement """
    values = list(tree.item(item, "values"))
    fichier = values[-1]  # Récupération du fichier associé

    paiements = charger_paiements()
    paiement_actuel = paiements.get(fichier, {}).get("Payé", False)

    def enregistrer_paiement(valide):
        paiements[fichier] = {"Montant": paiements.get(fichier, {}).get("Montant", "---"), "Payé": valide}
        sauvegarder_paiements(paiements)

        values[5] = "✅" if valide else "❌"  # ✅ Met bien à jour la colonne 5 (Payé)
        tree.item(item, values=tuple(values))
        popup.destroy()

    popup = tk.Toplevel()
    popup.title("Validation du paiement")
    popup.geometry("300x150")

    popup.update_idletasks()
    x = (popup.winfo_screenwidth() - popup.winfo_width()) // 2
    y = (popup.winfo_screenheight() - popup.winfo_height()) // 2
    popup.geometry(f"+{x}+{y}")

    message = "Valider le paiement ?" if not paiement_actuel else "Invalider le paiement ?"
    tk.Label(popup, text=message, font=("Arial", 12, "bold")).pack(pady=15)

    bouton_oui = tk.Button(popup, text="Oui", command=lambda: enregistrer_paiement(True), font=("Arial", 10))
    bouton_non = tk.Button(popup, text="Non", command=lambda: enregistrer_paiement(False), font=("Arial", 10))

    bouton_oui.pack(side="left", padx=20, pady=10, expand=True)
    bouton_non.pack(side="right", padx=20, pady=10, expand=True)

    popup.bind("<Return>", lambda event: enregistrer_paiement(True))
    popup.transient(tree.winfo_toplevel())
    popup.grab_set()




def ouvrir_selection(tree, dossier):
    """Ouvre le fichier PDF sélectionné."""
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Attention", "Veuillez sélectionner un contrat.")
        return

    fichier = tree.item(selected_item[0])["values"][-1]  # Dernière colonne contient le fichier
    ouvrir_contrat(fichier, dossier)


def ouvrir_contrat(fichier, dossier):
    """Ouvre un contrat PDF avec l'application par défaut (PDF Expert si dispo, sinon Aperçu)."""
    chemin_fichier = os.path.join(dossier, fichier)
    
    if os.path.exists(chemin_fichier):
        try:
            subprocess.run(["open", "-a", "PDF Expert", chemin_fichier])
        except FileNotFoundError:
            subprocess.run(["open", chemin_fichier])  # Ouvre avec Aperçu si PDF Expert n'est pas installé
    else:
        messagebox.showerror("Erreur", f"Le fichier {fichier} n'existe pas.")



def supprimer_contrat(tree, dossier):
    """ Supprime un contrat après confirmation et le retire du cache """
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("Attention", "Veuillez sélectionner un contrat à supprimer.")
        return

    item = selected_item[0]
    values = tree.item(item, "values")
    fichier = values[-1]  # Récupère le nom du fichier
    chemin_fichier = os.path.join(dossier, fichier)

    # 📌 Création de la fenêtre popup
    popup = tk.Toplevel()
    popup.title("Confirmation de suppression")
    popup.geometry("450x200")  # 🔹 Fenêtre plus grande pour meilleure visibilité
    popup.configure(bg="#f2f7ff")  # 🔹 Fond bleu clair

    # Centrage de la fenêtre sur l’écran
    popup.update_idletasks()
    x = (popup.winfo_screenwidth() - popup.winfo_width()) // 2
    y = (popup.winfo_screenheight() - popup.winfo_height()) // 2
    popup.geometry(f"+{x}+{y}")

    # 📌 Message de confirmation
    label_message = tk.Label(
        popup, text=f"Voulez-vous vraiment supprimer\nle contrat '{fichier}' ?", 
        font=("Arial", 12, "bold"), bg="#f2f7ff", fg="#333333", justify="center"
    )
    label_message.pack(pady=15)

    def confirmer_suppression():
        """ Supprime le fichier et l'entrée du cache """
        if os.path.exists(chemin_fichier):
            send2trash.send2trash(chemin_fichier)  # 📂 Envoie à la corbeille

        # 🔹 Suppression du cache
        cache = charger_cache()
        if fichier in cache:
            del cache[fichier]
            sauvegarder_cache(cache)

        # 🔹 Suppression de la ligne du tableau
        tree.delete(item)
        popup.destroy()

    # 📌 Cadre pour organiser les boutons
    frame_boutons = tk.Frame(popup, bg="#f2f7ff")
    frame_boutons.pack(pady=10)

    # 📌 Bouton "Oui" (Vert)
    bouton_oui = tk.Button(
        frame_boutons, text="✅ Oui", command=confirmer_suppression, font=("Arial", 12, "bold"),
        width=12, bg="#DDDDDD", fg="black", activebackground="#45a049"
    )
    bouton_oui.pack(side="left", padx=20, pady=10)

    # 📌 Bouton "Non" (Rouge, mis en évidence)
    bouton_non = tk.Button(
        frame_boutons, text="❌ Non", command=popup.destroy, font=("Arial", 12, "bold"),
        width=12, bg="#DDDDDD", fg="black", activebackground="#d32f2f"
    )
    bouton_non.pack(side="right", padx=20, pady=10)

    # 📌 Mettre le focus sur "Non" par défaut
    bouton_non.focus_set()

    # 📌 Associer la touche "Entrée" à "Non" par défaut
    popup.bind("<Return>", lambda event: popup.destroy())

    popup.transient(tree.winfo_toplevel())  # Associer la popup à la fenêtre principale
    popup.grab_set()  # Gérer le focus de la fenêtre
    

# ✅ Modifier le paiement (cocher/décocher)
def valider_paiement(event, tree):
    selected_item = tree.selection()
    if not selected_item:
        return

    item = selected_item[0]
    values = tree.item(item, "values")
    fichier = values[-1]
    paiements = charger_paiements()

    new_state = not paiements.get(fichier, {}).get("Payé", False)
    paiements[fichier] = {"Montant": paiements.get(fichier, {}).get("Montant", "---"), "Payé": new_state}
    sauvegarder_paiements(paiements)
    
    tree.item(item, values=values[:-2] + [values[-2], "✅" if new_state else "❌", values[-1]])







# 💳 Lancer le paiement en ligne
def effectuer_paiement(fichier):
    url = f"https://www.ta-banque.fr/paiement?ref={fichier}"
    webbrowser.open(url)
    
def modifier_montant(tree, item):
    """ Ouvre une fenêtre pour modifier le montant et enregistre la modification dans le cache """
    values = list(tree.item(item, "values"))
    fichier = values[-1]  # 📂 Nom du fichier PDF du contrat

    paiements = charger_paiements()
    cache = charger_cache()  # 🔥 Charger le cache pour mise à jour

    montant_actuel = paiements.get(fichier, {}).get("Montant", "").replace(" €", "")  # 🔍 Retirer le € pour édition

    popup = tk.Toplevel()
    popup.title("Modifier le montant")
    popup.geometry("350x150")

    label = tk.Label(popup, text="Entrer le montant :", font=("Arial", 12, "bold"))
    label.pack(pady=10)

    frame = tk.Frame(popup)
    frame.pack(pady=5)

    entry = tk.Entry(frame, font=("Arial", 12), width=10, justify="right")
    entry.pack(side="left", padx=5)
    entry.insert(0, montant_actuel)  
    entry.select_range(0, "end")  
    entry.focus_set()

    tk.Label(frame, text="€", font=("Arial", 12, "bold")).pack(side="left")

    def enregistrer():
        montant = entry.get().strip()
        if montant.isdigit():  
            montant += " €"

        # 🔥 Mettre à jour dans le cache et le fichier de paiements
        paiements[fichier] = {"Montant": montant, "Payé": paiements.get(fichier, {}).get("Payé", False)}
        sauvegarder_paiements(paiements)

        # 🔥 Mise à jour dans le cache global
        if fichier in cache:
            cache[fichier][4] = montant  # ✅ Met à jour la colonne "Montant" dans le cache

        sauvegarder_cache(cache)  # ✅ Sauvegarde après mise à jour

        values[4] = montant  
        tree.item(item, values=tuple(values))
        popup.destroy()

    bouton_enregistrer = tk.Button(popup, text="Enregistrer", command=enregistrer, font=("Arial", 10))
    bouton_enregistrer.pack(pady=10)

    popup.bind("<Return>", lambda event: enregistrer())  

    popup.transient(tree.winfo_toplevel())  
    popup.grab_set()






# ✅ Case à cocher pour "Payé"
def basculer_paiement(event, tree):
    """Permet de cocher ou décocher un paiement en double-cliquant sur 'Payé'."""
    selected_item = tree.selection()
    if not selected_item:
        return
    
    item = selected_item[0]
    values = list(tree.item(item, "values"))  # 🔹 Convertit en liste modifiable
    fichier = values[-1]
    paiements = charger_paiements()

    est_paye = paiements.get(fichier, {}).get("Payé", False)
    nouveau_statut = not est_paye  # 🔹 Inverse le statut

    paiements[fichier] = {"Montant": paiements.get(fichier, {}).get("Montant", "---"), "Payé": nouveau_statut}
    sauvegarder_paiements(paiements)

    icone_paye = "✅" if nouveau_statut else "❌"
    values[-2] = icone_paye  # 🔹 Met à jour la colonne "Payé"
    tree.item(item, values=values)


  

# 🚀 Lancer l'application
def main_menu():
    root = tk.Tk()
    root.title("Gestion des Contrats")
    tk.Button(root, text="📋 Contrats MAR", command=lambda: afficher_contrats("MAR")).pack()
    tk.Button(root, text="📋 Contrats IADE", command=lambda: afficher_contrats("IADE")).pack()
    root.mainloop()
    

if __name__ == "__main__":
    main_menu()