import os
import json
import time
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(SCRIPT_DIR, 'credentials.json'), 'r') as file:
    credentials = json.load(file)


# 📌 Chargement des paramètres depuis config.json
CONFIG_FILE = "config.json"

def charger_config():
    """Charge les paramètres bancaires depuis config.json."""
    if not os.path.exists(CONFIG_FILE):
        return {"bank_url": "https://espacepro.secure.lcl.fr/", "bank_id": ""}
    
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)

config = charger_config()
BANK_URL = config.get("bank_url", "https://espacepro.secure.lcl.fr/")
BANK_ID = config.get("bank_id", "VPERRE0003")


# 📌 Namespace XML pour SEPA
NAMESPACE = "urn:iso:std:iso:20022:tech:xsd:pain.001.001.02"
ET.register_namespace("", NAMESPACE)

# 📌 Mapping des BIC pour les banques
bic_mapping = {
    "10107": "PSSTFRPPXXX",   # La Banque Postale
    "10278": "CMCIFR2A",      # CIC
    "11315": "CMBRFR2BXXX",  # Crédit Mutuel Bretagne
    "11425": "CEPAFRPP142",   # Caisse d'Épargne Normandie
    "12968": "HSBCFRPP",     # HSBC France
    "13507": "CCBPFRPPMTG",  # Banque Populaire Grand Ouest
    "14518": "FTNOFRP1XXX",   # Arkéa Direct Bank (Fortuneo)
    "15589": "CRLYFRPP",     # LCL (Crédit Lyonnais)
    "16607": "SOGEFRPPXXX",  # Société Générale
    "18300": "BREDFRPPXXX",
    "18306": "BREDFRPPXXX",   # BRED Banque Populaire
    "18706": "CEPAFRPP142",   # Caisse d'Épargne
    "19506": "AGRIFRPP895",   # Ajoutez ici le BIC correspondant au code banque 19506
    "30002": "CRLYFRPPXXX",   # LCL (Crédit Lyonnais)
    "30003": "SOGEFRPPXXX",   # Société Générale
    "30004": "BNPAFRPPXXX",   # BNP Paribas
    "30006": "CCBPFRPPPPG",  # Banque Populaire du Grand Paris
    "30027": "CEPAFRPPXXX",   # Caisse d'Épargne
    "30056": "SOGEFRPPXXX",   # Société Générale
    "30066": "CCBPFRPPMTG",   # Banque Populaire Occitane
    "30087": "AGRIFRPPXXX",   # Crédit Agricole
    "30108": "CCBPFRPPDPE",  # Banque Populaire Val de France
    "30129": "CCBPFRPPMTG",   # Banque Populaire
    "30278": "CMCIFR2A",     # CIC
    "30438": "CCBPFRPPXXX",   # Banque Populaire Val de France
    "30475": "CCBPFRPPTOU",  # Banque Populaire Occitane
    "30588": "CCBPFRPPLYO",  # Banque Populaire Auvergne Rhône Alpes
    "30628": "CCBPFRPPMAR",  # Banque Populaire Méditerranée
    "30678": "CEPAFRPPXXX",   # Caisse d'Épargne Rhône Alpes
    "30688": "CEPAFRPPXXX",   # Caisse d'Épargne Côte d'Azur
    "30738": "CEPAFRPPXXX",   # Caisse d'Épargne Bretagne Pays de Loire
    "30778": "CEPAFRPPXXX",   # Caisse d'Épargne Languedoc-Roussillon
    "30788": "CCBPFRPPGRE",  # Banque Populaire Auvergne Rhône Alpes
    "30938": "CCBPFRPPXXX",   # Banque Populaire Rives de Paris
    "31039": "CCBPFRPPXXX",   # Banque Populaire Nord
    "31048": "CEPAFRPP316",  # Caisse d'Épargne Aquitaine Poitou-Charentes
    "31006": "CEPAFRPP310",  # Caisse d'Épargne Midi-Pyrénées
    "31539": "CCBPFRPPXXX",   # Banque Populaire Sud
    "32039": "CMBRFR2BXXX",  # Crédit Mutuel Bretagne
    "32349": "CEPAFRPP330",  # Caisse d'Épargne Bretagne Pays de Loire
    "39999": "INGBFRPPXXX",   # ING France
    "40618": "BANKFRPPXXX",
    "40949": "CCBPFRPPCHO",  # Banque Populaire Centre Ouest
    "42559": "BPRIFRPPXXX",   # Banque de la Réunion
    "42599": "BPRIFRPPXXX",   # Banque de Nouvelle-Calédonie
    "50000": "CCBPFRPPXXX",   # Banque Populaire Alsace Lorraine Champagne
    "50088": "AXABFRPPXXX",   # AXA Banque
    "50100": "CCBPFRPPXXX",   # Banque Populaire du Nord
    "50107": "CCBPFRPPXXX",  # Banque Populaire BRED
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
    "88999": "N26PFRPPXXX",  # N26 France
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
    "89999": "N26PFRPPXXX",    # N26 France
    "98727": "TRZOFR21XXX",  # Treezor (Fintech)
    "99999": "REVOLT21",     # Revolut France
}

def deduire_bic_depuis_code_banque(code_banque):
    """ Récupère le BIC en fonction du code banque """
    return bic_mapping.get(str(code_banque).zfill(5), "UNKNOWN BIC")

def generer_xml_virements(virements, fichier_sortie=None, date_execution=None):
    """
    Génère un fichier XML de virements SEPA avec plusieurs transactions.

    Args:
        virements (list): Liste de dictionnaires contenant les informations des virements.
                          Chaque dictionnaire doit contenir :
                          - 'beneficiaire' (str) : Nom du bénéficiaire
                          - 'iban' (str) : IBAN du bénéficiaire
                          - 'montant' (float) : Montant du virement en euros
                          - 'objet' (str) : Objet du virement (peut contenir la date)
        fichier_sortie (str, optional): Chemin du fichier XML à générer (par défaut, un fichier basé sur la date).
        date_execution (str, optional): Date d'exécution au format YYYY-MM-DD. Si non spécifiée, utilise la date du jour.

    Returns:
        str: Chemin du fichier XML généré.
    """

    if not virements:
        print("❌ Aucune transaction fournie pour générer un virement.")
        return None

    # Si date_execution n'est pas fournie, utiliser la date du jour
    if not date_execution:
        date_execution = datetime.now().strftime("%Y-%m-%d")
    
    # Extraire la date du premier virement si disponible dans l'objet (format: "YYYY-MM-DD - Référence")
    if not date_execution and virements and "objet" in virements[0]:
        objet_parts = virements[0]["objet"].split(" - ")
        if len(objet_parts) >= 2:
            try:
                # Essayer de parser la date depuis l'objet
                date_from_objet = datetime.strptime(objet_parts[0], "%Y-%m-%d")
                date_execution = date_from_objet.strftime("%Y-%m-%d")
            except:
                # Si échec, utiliser la date du jour
                date_execution = datetime.now().strftime("%Y-%m-%d")
    
    # Formater la date pour les identifiants (sans tirets)
    date_id = date_execution.replace("-", "")
    
    if not fichier_sortie:
        fichier_sortie = f"virements_{date_id}.xml"

    # 📌 Création de la structure XML SEPA
    document = ET.Element(f"{{{NAMESPACE}}}Document")
    pain = ET.SubElement(document, f"{{{NAMESPACE}}}pain.001.001.02")

    # 🏦 En-tête du fichier
    grpHdr = ET.SubElement(pain, f"{{{NAMESPACE}}}GrpHdr")
    ET.SubElement(grpHdr, f"{{{NAMESPACE}}}MsgId").text = f"SELARLANESTHES-{date_id}"
    ET.SubElement(grpHdr, f"{{{NAMESPACE}}}CreDtTm").text = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    ET.SubElement(grpHdr, f"{{{NAMESPACE}}}BtchBookg").text = "true"
    ET.SubElement(grpHdr, f"{{{NAMESPACE}}}NbOfTxs").text = str(len(virements))
    ET.SubElement(grpHdr, f"{{{NAMESPACE}}}CtrlSum").text = f"{sum(v['montant'] for v in virements):.2f}"
    ET.SubElement(grpHdr, f"{{{NAMESPACE}}}Grpg").text = "MIXD"

    initgPty = ET.SubElement(grpHdr, f"{{{NAMESPACE}}}InitgPty")
    ET.SubElement(initgPty, f"{{{NAMESPACE}}}Nm").text = "SELARL ANESTHESISTES DE LA CLINIQUE MATHILDE"

    # 🏦 Détails des paiements
    pmtInf = ET.SubElement(pain, f"{{{NAMESPACE}}}PmtInf")
    ET.SubElement(pmtInf, f"{{{NAMESPACE}}}PmtInfId").text = f"SELARLANESTHES-{date_id}"
    ET.SubElement(pmtInf, f"{{{NAMESPACE}}}PmtMtd").text = "TRF"

    pmtTpInf = ET.SubElement(pmtInf, f"{{{NAMESPACE}}}PmtTpInf")
    svcLvl = ET.SubElement(pmtTpInf, f"{{{NAMESPACE}}}SvcLvl")
    ET.SubElement(svcLvl, f"{{{NAMESPACE}}}Cd").text = "SEPA"

    # Utiliser la date d'exécution spécifiée au lieu de la date du jour
    ET.SubElement(pmtInf, f"{{{NAMESPACE}}}ReqdExctnDt").text = date_execution

    # Débiteur
    dbtr = ET.SubElement(pmtInf, f"{{{NAMESPACE}}}Dbtr")
    ET.SubElement(dbtr, f"{{{NAMESPACE}}}Nm").text = "SELARL DES ANESTHESISTES DE LA CLINIQUE MATHILDE"

    dbtrAcct = ET.SubElement(pmtInf, f"{{{NAMESPACE}}}DbtrAcct")
    id_dbtrAcct = ET.SubElement(dbtrAcct, f"{{{NAMESPACE}}}Id")
    ET.SubElement(id_dbtrAcct, f"{{{NAMESPACE}}}IBAN").text = credentials['company_iban']

    dbtrAgt = ET.SubElement(pmtInf, f"{{{NAMESPACE}}}DbtrAgt")
    finInstnId = ET.SubElement(dbtrAgt, f"{{{NAMESPACE}}}FinInstnId")
    ET.SubElement(finInstnId, f"{{{NAMESPACE}}}BIC").text = "CRLYFRPP"

    # 📌 Ajouter chaque virement
    for virement in virements:
        beneficiaire = virement["beneficiaire"]
        iban = virement["iban"].replace(" ", "")
        montant = virement["montant"]
        objet = virement.get("objet", "Virement SEPA")

        # Récupérer le BIC en fonction du code banque extrait de l'IBAN
        code_banque = iban[4:9]
        bic = deduire_bic_depuis_code_banque(code_banque)

        cdtTrfTxInf = ET.SubElement(pmtInf, f"{{{NAMESPACE}}}CdtTrfTxInf")
        pmtId = ET.SubElement(cdtTrfTxInf, f"{{{NAMESPACE}}}PmtId")

        # Générer l'EndToEndId au format "VIR-SELANE-{ref_benef}-{month_abbr}{year_last2}"
        # Si le bénéficiaire est composé (ex: "Christiane Passerieu"), on prend la première lettre du prénom et le nom complet en majuscules.
        parts = beneficiaire.split()
        if len(parts) >= 2:
            ref_benef = parts[0][0].upper() + parts[-1].upper()
        else:
            ref_benef = beneficiaire.upper()
        # Calcul de la longueur maximale autorisée pour ref_benef
        # Format total = len("VIR-SELANE-") (11) + len(ref_benef) + 1 (tiret) + 3 (mois abrégé) + 2 (année) ≤ 35
        # On obtient : len(ref_benef) ≤ 35 - (11+1+3+2) = 18
        if len(ref_benef) > 18:
            ref_benef = ref_benef[:18]

        # Extraire le mois abrégé en français et l'année à partir de la date d'exécution
        try:
            dt = datetime.strptime(date_execution, "%Y-%m-%d")
            french_abbr = {1:"JAN",2:"FEV",3:"MAR",4:"AVR",5:"MAI",6:"JUN",7:"JUL",8:"AOU",9:"SEP",10:"OCT",11:"NOV",12:"DEC"}
            month_abbr = french_abbr.get(dt.month, "")
            year_last2 = dt.strftime("%y")  # Les deux derniers chiffres de l'année
        except Exception as e:
            print(f"⚠️ Erreur lors du formatage de la date: {e}")
            month_abbr = ""
            year_last2 = datetime.now().strftime("%y")

        end_to_end_id = f"VIR-SELANE-{ref_benef}-{month_abbr}{year_last2}"
        ET.SubElement(pmtId, f"{{{NAMESPACE}}}EndToEndId").text = end_to_end_id

        amt = ET.SubElement(cdtTrfTxInf, f"{{{NAMESPACE}}}Amt")
        ET.SubElement(amt, f"{{{NAMESPACE}}}InstdAmt", Ccy="EUR").text = f"{montant:.2f}"

        cdtrAgt = ET.SubElement(cdtTrfTxInf, f"{{{NAMESPACE}}}CdtrAgt")
        finInstnId = ET.SubElement(cdtrAgt, f"{{{NAMESPACE}}}FinInstnId")
        ET.SubElement(finInstnId, f"{{{NAMESPACE}}}BIC").text = bic

        cdtr = ET.SubElement(cdtTrfTxInf, f"{{{NAMESPACE}}}Cdtr")
        ET.SubElement(cdtr, f"{{{NAMESPACE}}}Nm").text = beneficiaire

        cdtrAcct = ET.SubElement(cdtTrfTxInf, f"{{{NAMESPACE}}}CdtrAcct")
        id_cdtrAcct = ET.SubElement(cdtrAcct, f"{{{NAMESPACE}}}Id")
        ET.SubElement(id_cdtrAcct, f"{{{NAMESPACE}}}IBAN").text = iban

        rmtInf = ET.SubElement(cdtTrfTxInf, f"{{{NAMESPACE}}}RmtInf")
        ET.SubElement(rmtInf, f"{{{NAMESPACE}}}Ustrd").text = objet

    # 📂 Sauvegarde du fichier XML
    tree = ET.ElementTree(document)
    tree.write(fichier_sortie, encoding="utf-8", xml_declaration=True)

    print(f"✅ Virements générés avec succès : {fichier_sortie}")
    return fichier_sortie

def envoyer_virement_vers_lcl(fichier_xml):
    """Automatise la connexion et l’envoi du fichier XML via Selenium."""
    # ✅ Création de la fenêtre Tkinter
    virement_window = tk.Toplevel()
    virement_window.title("Virement en cours")
    virement_window.geometry("400x200")

    label_instruction = tk.Label(virement_window, text="Procédure en cours...\nVeuillez ne pas fermer cette fenêtre.", font=("Arial", 12))
    label_instruction.pack(pady=20)

    # ✅ Création du navigateur Selenium
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")  # Ouvre Chrome en plein écran
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)


    try:
        # 🏦 Accéder au site de la banque
        driver.get(BANK_URL)

        # 🔑 Entrer l'identifiant bancaire
        login_input = wait.until(EC.presence_of_element_located((By.ID, "login")))
        login_input.send_keys(BANK_ID)

        # 🔘 Sélectionner "LCL m-Identité"
        mode_identite_label = wait.until(EC.element_to_be_clickable((By.XPATH, "//label[contains(., 'LCL m-Identité')]")))
        mode_identite_label.click()

        # 📲 Pause pour validation sur mobile
        print("✅ Validez votre connexion sur votre téléphone...")
        time.sleep(15)  # Temps d'attente pour validation manuelle
        
        # 🔄 Vérifier si la pop-up de cookies apparaît après connexion
        try:
            accept_cookies = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "popin_tc_privacy_button_3"))
            )
            accept_cookies.click()
            print("✅ Cookies acceptés après connexion.")
        except Exception:
            print("ℹ️ Aucun bandeau cookies détecté après connexion.")

        # 📌 Accéder à "Règlements"
        reglements = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[@class='avecLayer' and text()='Règlements']")))
        reglements.click()

        # 📌 Accéder à "Virements Europe (SEPA)"
        virements = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Virements Europe (SEPA)']")))
        virements.click()

        # 📌 Accéder à "Faire une remise"
        faire_remise = wait.until(EC.element_to_be_clickable((By.ID, "onglet1")))
        faire_remise.click()

        # 📂 Sélectionner le fichier XML
        fichier_input = wait.until(EC.presence_of_element_located((By.ID, "fichier")))
        fichier_input.send_keys(os.path.abspath(fichier_xml))

        # ✅ Valider l’envoi du fichier
        valider_button = wait.until(EC.element_to_be_clickable((By.ID, "veu_validation_fichier")))
        valider_button.click()

        # ✅ Mise à jour du message de suivi
        label_instruction.config(text="Virement effectué avec succès !\nCliquez sur 'Terminer' pour fermer.")

    except Exception as e:
        label_instruction.config(text=f"❌ Erreur : {e}")

    # ✅ Ajout du bouton "Terminer"
    def terminer_virement():
        """Ferme Selenium et clôture la fenêtre."""
        try:
            driver.quit()  # ✅ Ferme uniquement Selenium
            print("✅ Fenêtre Selenium fermée avec succès.")
        except Exception as e:
            print(f"⚠️ Erreur lors de la fermeture de Selenium : {e}")

        virement_window.destroy()  # ✅ Ferme la fenêtre Tkinter

    btn_terminer = tk.Button(virement_window, text="Terminer", command=terminer_virement, font=("Arial", 12), bg="green", fg="black")
    btn_terminer.pack(pady=20)

    virement_window.mainloop()

def lire_donnees_virement(chemin_fichier):
    """Lit la feuille 'Virement' et extrait les données sous forme de liste de dictionnaires."""
    try:
        df = pd.read_excel(chemin_fichier, sheet_name="Virement")
        
        # Vérification des colonnes attendues
        colonnes_attendues = ["Nom_Beneficiaire", "IBAN", "Montant", "Objet"]
        for col in colonnes_attendues:
            if col not in df.columns:
                raise ValueError(f"Colonne manquante : {col} dans la feuille 'Virement'")
        
        # Nettoyage des données et conversion du montant
        virements = []
        for _, row in df.iterrows():
            try:
                montant = float(str(row["Montant"]).replace(" ", "").replace(",", "").replace("€", ""))
                virement = {
                    "beneficiaire": str(row["Nom_Beneficiaire"]).strip(),
                    "iban": str(row["IBAN"]).replace(" ", ""),
                    "montant": montant,
                    "objet": str(row["Objet"]).strip()
                }
                virements.append(virement)
            except Exception as e:
                print(f"⚠️ Erreur de conversion pour {row['Nom_Beneficiaire']} : {e}")
        
        return virements
    except Exception as e:
        print(f"❌ Erreur lors de la lecture du fichier de virement : {e}")
        return []


def valider_virement():
    """Lit le fichier Excel et envoie les données pour générer un virement."""
    fichier_virement = file_paths.get("chemin_fichier_virement", "")
    if not os.path.exists(fichier_virement):
        print(f"❌ Fichier de virement introuvable : {fichier_virement}")
        return
    
    virements = lire_donnees_virement(fichier_virement)
    if virements:
        fichier_xml = generer_xml_virements(virements)

if __name__ == "__main__":
    # Exemple de virements
    virements_test = [
        {"beneficiaire": "Jean Dupont", "iban": "FR7630003030001234567890185", "montant": 1500.50, "objet": "Facture 2024"},
        {"beneficiaire": "SARL Alpha", "iban": "FR7630003030009876543210123", "montant": 2300.75, "objet": "Fournisseur"}
    ]

    # 1️⃣ Générer le fichier XML
    fichier_xml = generer_xml_virements(virements_test)

    # 2️⃣ Envoyer le fichier XML vers LCL
    if fichier_xml:
        envoyer_virement_vers_lcl(fichier_xml)