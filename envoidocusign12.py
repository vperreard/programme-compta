from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import time
import os
import locale
import json
import sys
import random
from datetime import datetime

chromedriver_path = "/opt/homebrew/bin/chromedriver"
service = Service(chromedriver_path)

options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)

def safe_click(driver, element_xpath):
    """Attend que l'élément soit cliquable et utilise JavaScript si nécessaire."""
    try:
        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, element_xpath)))
        driver.execute_script("arguments[0].scrollIntoView();", element)  # Scroll vers l'élément
        time.sleep(1)  # Pause pour éviter les problèmes d'affichage
        element.click()
        print(f"✅ Élément cliqué avec succès : {element_xpath}")
    except Exception as e:
        print(f"⚠️ Erreur lors du clic : {e}. Tentative avec JavaScript...")
        try:
            element = driver.find_element(By.XPATH, element_xpath)
            driver.execute_script("arguments[0].click();", element)  # Clic forcé via JS
            print(f"✅ Clic forcé réussi sur l'élément : {element_xpath}")
        except Exception as e:
            print(f"❌ Impossible de cliquer sur l'élément {element_xpath} : {e}")
            
def split_name(full_name):
    """Sépare prénom et nom correctement, même si un seul élément est donné."""
    parts = full_name.strip().split(" ", 1)  # Sépare au premier espace
    if len(parts) == 2:
        return parts[0], parts[1]  # ✅ Prénom, Nom
    else:
        return "", parts[0]  # ✅ Pas de prénom trouvé

def extract_first_name(full_name):
    """Extrait le prénom correctement, même si la chaîne contient des espaces en trop."""
    if not full_name or full_name.strip() == "":
        return "Inconnu"
    return full_name.strip().split(" ")[0]  # Prend le premier mot

def find_element_by_partial_xpath(driver, xpath_start):
    """Tente de trouver un élément dont l'ID change à chaque session, en cherchant avec un XPath partiel."""
    try:
        possible_elements = driver.find_elements(By.XPATH, f"//*[contains(@id, '{xpath_start}')]")
        if possible_elements:
            return possible_elements[0]  # Prend le premier élément trouvé
        return None
    except Exception:
        return None

def ajouter_destinataires_et_message():
    """Ajoute les destinataires et remplit les champs email et message après l'application du modèle."""
    print("\n⏳ Ajout des destinataires...")
    destinataires_button_xpath = "//*[@id='root']/div/div/div/div[1]/div/div[3]/div/div/div/div[3]/div/div[2]/div[2]/div/div[1]/h3/span"
    safe_click(driver, destinataires_button_xpath)

    time.sleep(2)
    input_fields = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input[type='text']")))

    try:
        if contract_type == "MAR":
            # Destinataire 1 : Médecin remplacé
            input_fields[0].clear()
            input_fields[0].send_keys(replaced_name)
            input_fields[1].clear()
            input_fields[1].send_keys(replaced_email)

            # Destinataire 2 : Médecin remplaçant
            input_fields[2].clear()
            input_fields[2].send_keys(replacing_name)
            input_fields[3].clear()
            input_fields[3].send_keys(replacing_email)

        print("✅ Destinataires ajoutés avec succès.")
    except Exception as e:
        print(f"❌ Erreur lors de l'ajout des destinataires : {e}")

    # 7️⃣ Ajout du message
    print("\n⏳ Ajout du message...")
    message_encart_xpath = "//*[@id='root']/div/div/div/div[1]/div/div[3]/div/div/div/div[3]/div/div[2]/div[3]/div/div[1]/h3/span"
    safe_click(driver, message_encart_xpath)
    print("✅ Encart 'Ajouter un message' ouvert.")
    time.sleep(2)

    # Détection des champs
    input_fields = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@type='text'] | //textarea")))
    objet_email_field, message_field = None, None

    for field in input_fields:
        placeholder = field.get_attribute("placeholder")
        if field.tag_name == "input" and "Complétez avec Docusign" in placeholder:
            objet_email_field = field
        elif field.tag_name == "textarea" and "Saisir un message" in placeholder:
            message_field = field

    # ✅ Formatage correct de la date
    try:
        date_formattee = datetime.strptime(date_contrat, "%d%m%Y").strftime("%d/%m/%Y")
    except ValueError:
        date_formattee = date_contrat  # Si erreur, garde la date brute

    # Effacer et remplir l'objet de l'email
    if objet_email_field:
        effacer_champ(objet_email_field)
        objet_email_field.send_keys(subject)
        print("✅ Objet de l'email rempli.")
    else:
        print("❌ Champ 'Objet' non trouvé.")

    # Effacer et remplir le message de l'email
    if message_field:
        effacer_champ(message_field)
        message_field.send_keys(message)
        print("✅ Message rempli.")
    else:
        print("❌ Champ 'Message' non trouvé.")
    
    print("\n🛑 Script terminé. Vérifiez les informations et SIGNEZ le document.")


options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")  # Ouvre Chrome en grand écran
options.add_argument("--disable-infobars") # Désactive les popups gênantes
options.add_argument("--remote-debugging-port=9222")  # Debugging

CHROMEDRIVER_PATH = "/usr/local/bin/chromedriver"
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=options)


# 📌 Définir la langue en français pour afficher les dates en français
locale.setlocale(locale.LC_TIME, "fr_FR.UTF-8")

# Vérification des arguments reçus
print(f"🛠️ DEBUG Arguments reçus dans envoidocusign11.py : {sys.argv}")

if len(sys.argv) < 7:  # Pour IADE, 7 arguments minimum
    print("❌ Erreur : Nombre incorrect d'arguments fournis !")
    sys.exit(1)

pdf_path = sys.argv[1]
contract_type = sys.argv[2]
start_date = sys.argv[3]
end_date = sys.argv[4]
replacing_name = sys.argv[5]
replacing_email = sys.argv[6]


if contract_type == "MAR":
    if len(sys.argv) < 9:  # Pour MAR, 9 arguments minimum
        print("❌ Erreur : Nombre incorrect d'arguments fournis pour MAR !")
        sys.exit(1)
    replaced_name = sys.argv[7]
    replaced_email = sys.argv[8]
else:  # IADE ou autre
    replaced_name = "Non précisé"
    replaced_email = "email_inconnu@exemple.com"

print(f"📆 DEBUG : pdf_path={pdf_path}, contract_type={contract_type}, start_date={start_date}, end_date={end_date}, replacing_name={replacing_name}, replacing_email={replacing_email}, replaced_name={replaced_name}, replaced_email={replaced_email}")

print(f"📆 DEBUG : pdf_path={pdf_path}, contract_type={contract_type}, start_date={start_date}, end_date={end_date}, "
      f"replacing_name={replacing_name}, replacing_email={replacing_email}, replaced_name={replaced_name}, replaced_email={replaced_email}")


# Vérification des valeurs
if not replacing_name or not replacing_email:
    print("❌ Erreur : Informations du remplaçant incomplètes.")
    sys.exit(1)

# Si nécessaire, tu peux corriger ici l'affectation des noms
med_remplace_nom_complet = replaced_name
med_remplace_email = replaced_email
med_remplacant_nom_complet = replacing_name
med_remplacant_email = replacing_email
replaced_last_name = split_name(replaced_name)[1] if replaced_name != "Non précisé" else "Inconnu"
replacing_last_name = split_name(replacing_name)[1] if replacing_name else "Inconnu"
date_contrat = datetime.today().strftime("%d%m%Y")
formatted_date = datetime.strptime(start_date, "%Y-%m-%d").strftime("%A %d %B %Y")


if contract_type == "MAR":
    subject = f"Contrat {replaced_last_name} {replacing_last_name} {date_contrat}"
    message = (f"Bonjour "
            f"{extract_first_name(replacing_name)},"
            f"{' , ' if replaced_name != 'Non précisé' else ''}"
            f"{extract_first_name(replaced_name) if replaced_name != 'Non précisé' else ''},\n\n"
            f"Ci-joint le contrat de remplacement pour la journée du {formatted_date}.\n\n"
            "Merci de le signer.\n\n"
            "V. PERREARD")


elif contract_type == "IADE":
    replaced_name = "Non précisé"
    replaced_email = "email_inconnu@exemple.com"
    formatted_start_date = datetime.strptime(start_date, "%Y-%m-%d").strftime("%d/%m/%Y")
    formatted_end_date = datetime.strptime(end_date, "%Y-%m-%d").strftime("%d/%m/%Y")

    if start_date == end_date:
        date_text = f"pour le {formatted_start_date}"
    else:
        date_text = f"pour la période du {formatted_start_date} au {formatted_end_date}"

    subject = f"CDD {replacing_last_name} {formatted_start_date}"
    prenom_remplacant = extract_first_name(replacing_name)
    print(f"🔍 DEBUG : replacing_name = {replacing_name}")
    print(f"🔍 DEBUG : extract_first_name(replacing_name) = {extract_first_name(replacing_name)}")


    message = (f"Bonjour {prenom_remplacant},\n\n"
            f"Ci-joint le contrat pour le {formatted_date}.\n\n"
            "Merci de le signer, à bientôt.\n\n"
            "V. PERREARD")



try:
    med_remplace_prenom, med_remplace_nom = split_name(med_remplace_nom_complet)
    med_remplacant_prenom, med_remplacant_nom = split_name(med_remplacant_nom_complet)
except Exception as e:
    print(f"❌ Erreur lors de la séparation des noms : {e}")
    sys.exit(1)

# 📆 Date du jour pour le contrat
date_contrat = datetime.today().strftime("%d%m%Y")  # Format ddmmyyyy

# ✅ Affichage pour vérification
print(f"📄 Fichier PDF : {pdf_path}")
print(f"👨‍⚕️ Médecin remplacé : {med_remplace_prenom} {med_remplace_nom} ({med_remplace_email})")
print(f"👨‍⚕️ Médecin remplaçant : {med_remplacant_prenom} {med_remplacant_nom} ({med_remplacant_email})")
print(f"📆 Date du contrat : {date_contrat}")

# 🔽 Ici, intègre l'envoi du fichier à DocuSign avec ces informations

def load_docusign_credentials():
    """Charge les identifiants DocuSign depuis config.json"""
    config_file = "config.json"
    try:
        with open(config_file, "r") as f:
            config = json.load(f)
    except FileNotFoundError:
        config = {
            "docusign_login_page": "https://account.docusign.com",
            "docusign_email": "",
            "docusign_password": ""
        }

    return config

try:
    # Initialisation de l'action chain
    actions = ActionChains(driver)

    # Fonction pour effacer tout le champ en sélectionnant tout le texte
    def effacer_champ(element):
        element.click()  # Activer le champ
        time.sleep(0.5)

        # 🔹 Étape 1 : Aller au début et tout sélectionner vers le haut
        element.send_keys(Keys.HOME)
        for _ in range(20):  
            element.send_keys(Keys.SHIFT, Keys.ARROW_UP)
            time.sleep(0.05)
        element.send_keys(Keys.DELETE)  # Supprimer ce qui est sélectionné

        # 🔹 Étape 2 : Aller à la fin et tout sélectionner vers le bas
        for _ in range(20):  
            element.send_keys(Keys.SHIFT, Keys.ARROW_DOWN)
            time.sleep(0.05)
        element.send_keys(Keys.DELETE)  # Supprimer ce qui est sélectionné

    # Charger les identifiants DocuSign
    docusign_config = load_docusign_credentials()

    # 1️⃣ Connexion à DocuSign
    print("🔹 Lancement de Chrome via Selenium...")
    driver.get(docusign_config["docusign_login_page"])
    wait = WebDriverWait(driver, 20)

    # Saisie de l'email
    email_field = wait.until(EC.presence_of_element_located((By.NAME, "email")))
    email_field.clear()
    email_field.send_keys(docusign_config["docusign_email"])
    print("✅ Email saisi.")

    # "Suivant"
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-qa='submit-username']"))).click()
    print("✅ Bouton 'Suivant' cliqué.")
    time.sleep(2)

    # Saisie du mot de passe
    password_field = wait.until(EC.presence_of_element_located((By.NAME, "password")))
    password_field.clear()

    if docusign_config["docusign_password"]:  # Si un mot de passe est enregistré
        password_field.send_keys(docusign_config["docusign_password"])
        print("✅ Mot de passe saisi.")
    else:
        print("⚠️ Mot de passe non enregistré, saisissez-le manuellement.")

    # Connexion
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-qa='submit-password']"))).click()
    print("✅ Bouton 'Se connecter' cliqué.")

    # Envoi du code SMS
    send_code_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='send-code']")))
    send_code_button.click()
    print("✅ Bouton 'Send Code' cliqué.")

    # 🔄 Optimisation : Attente de la validation SMS + chargement de la page d'accueil
    print("\n⏳ **Attente de la validation du code SMS et chargement de la page d'accueil...**")

    try:
        WebDriverWait(driver, 60).until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//h1[contains(text(), 'Bienvenue')]")),
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'dashboard')]"))  # Vérification alternative
            )
        )
        print("✅ Connexion réussie, page d'accueil détectée.")
    except Exception as e:
        print(f"❌ Erreur : La page d'accueil ne s'est pas chargée. Détails : {e}")

    # Maximiser la fenêtre et activer l'onglet Selenium
    driver.switch_to.window(driver.window_handles[0])
    driver.maximize_window()
    print("✅ Fenêtre Selenium activée et maximisée.")

    # 🔄 Gestion améliorée des pop-ups
    try:
        alert = driver.switch_to.alert
        print(f"⚠️ Alerte détectée : {alert.text}")
        alert.accept()
        print("✅ Alerte fermée.")
    except:
        print("✅ Aucune alerte détectée après connexion.")

    # 📌 Vérification et clic sur le bouton "Commencer"
    try:
        print("⏳ Recherche du bouton 'Commencer'...")

        commencer_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@data-qa='manage-sidebar-actions-ndse-trigger']"))
        )
        driver.execute_script("arguments[0].scrollIntoView();", commencer_button)  # S'assure qu'il est bien visible
        time.sleep(1)
        commencer_button.click()
        print("✅ Bouton 'Commencer' cliqué avec succès.")

    except Exception as e:
        print(f"❌ Erreur : Impossible de cliquer sur 'Commencer'. Détails : {e}")
        sys.exit(1)

    # 📌 Recherche de la zone "Déposer les fichiers ici"
    try:
        print("⏳ Recherche de la zone de dépôt des fichiers...")

        drop_zone = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//p[@data-qa='upload-file-prompt-text']"))
        )
        driver.execute_script("arguments[0].scrollIntoView();", drop_zone)  # Scroll vers la zone
        time.sleep(1)

        print("✅ Zone de dépôt détectée.")

        # 📌 Vérification et envoi unique du fichier sans double clic
        print("⏳ Vérification du champ d'importation...")

        file_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='file' and @data-qa='upload-file-input-file-drop']"))
        )

        print("⚠️ Vérification avant envoi du fichier dans <input type='file'>")
        file_input.send_keys(pdf_path)  # ✅ Envoi direct du fichier SANS cliquer avant
        print(f"✅ Fichier '{pdf_path}' chargé avec succès.")

    except Exception as e:
        print(f"❌ Erreur lors de l'importation du fichier : {e}")
        sys.exit(1)


    # 4️⃣ Détection du modèle proposé
    print("⏳ Vérification si DocuSign propose un modèle...")
    try:
        modele_popup_xpath = "//*[@id='ModalContainer']/div[2]/div[2]/div/div/div[2]/h2/span"
        WebDriverWait(driver, 2.5).until(EC.presence_of_element_located((By.XPATH, modele_popup_xpath)))
        print("🛑 Modèle proposé détecté !")

        appliquer_button_xpath = "//*[@id='ModalContainer']/div[2]/div[2]/div/div/div[4]/div/div/div[1]/button"
        appliquer_button = wait.until(EC.element_to_be_clickable((By.XPATH, appliquer_button_xpath)))
        appliquer_button.click()
        print("✅ Modèle appliqué automatiquement.")
        time.sleep(2)

        print("⏳ Poursuite du processus d'ajout des destinataires et du message...")
            # Poursuite du script même si un modèle a été appliqué automatiquement
        ajouter_destinataires_et_message()

    except Exception:
        print("✅ Aucun modèle détecté, on continue.")
        
except Exception as e:
    print(f"❌ Erreur détectée : {e}")
    print("\n🛑 Une erreur est survenue. Vérifiez le navigateur et SIGNEZ manuellement si nécessaire.")
    input("\n➡️ **Appuyez sur Entrée pour fermer Selenium et le navigateur...**")

# ➡️ Appliquer modèle "CDD 5 pages" pour IADE (avant ajout des destinataires)
if contract_type == "IADE":
    print("⏳ Application du modèle 'CDD 5 pages'...")

    # XPath plus générique pour les trois petits points
    trois_points_xpath = "//button[@data-qa='file-menu' and contains(@aria-label, 'Plus d’options')]"

    try:
        safe_click(driver, trois_points_xpath)
        print("✅ Ouverture des options du document réussie.")
    except Exception as e:
        print(f"❌ Erreur lors du clic sur les 3 petits points : {e}")



  # 2️⃣ Clic sur "Appliquer les modèles"
    try:
        time.sleep(2)
        appliquer_modeles_xpath = "//*[@id='windows-drag-handler-wrapper']/div/div[1]/div[3]/div/div[2]/div/div/div[2]/div[1]/button/span/span"
        safe_click(driver, appliquer_modeles_xpath)
        print("✅ Ouverture du menu d'application des modèles.")
    except Exception as e:
        print(f"❌ Erreur lors du clic sur 'Appliquer les modèles' : {e}")

    # 3️⃣ Sélection du modèle "CDD 5 pages"
    cdd_xpath = "//label[contains(@data-qa, 'label-select-template') and contains(., 'CDD 5 pages')]"

    try:
        safe_click(driver, cdd_xpath)
        print("✅ Modèle 'CDD 5 pages' sélectionné.")
    except Exception as e:
        print(f"❌ Erreur lors de la sélection du modèle 'CDD 5 pages' : {e}")


    # 4️⃣ Validation en cliquant sur "Appliquer la sélection"
    try:
        time.sleep(2)
        appliquer_selection_xpath = "//*[@id='OverlayContainer']/div/div[2]/div[2]/div[3]/div/button"
        safe_click(driver, appliquer_selection_xpath)
        print("✅ Modèle appliqué.")
    except Exception as e:
        print(f"❌ Erreur lors de l'application du modèle : {e}")


    # 5️⃣ Vérification de la fenêtre SMS
    print("\n⏳ Vérification de la fenêtre SMS...")
    try:
        sms_popup_xpath = "//*[@id='add-recipients-content']/div/div/div[5]/div[1]/div/div/div/div[1]/div[2]/ul/li[2]/div/div[1]/fieldset/div[2]/div[2]/div[2]/div/div/div[2]"
        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, sms_popup_xpath)))

        close_sms_xpath = "//*[@id='add-recipients-content']/div/div/div[5]/div[1]/div/div/div/div[1]/div[2]/ul/li[2]/div/div[1]/fieldset/div[2]/div[2]/div[2]/div/div/div[1]/div/button"
        wait.until(EC.element_to_be_clickable((By.XPATH, close_sms_xpath))).click()
        print("✅ Fenêtre SMS détectée et fermée.")

    except Exception:
        print("✅ Aucune fenêtre SMS détectée.")

    print("\n⏳ Ajout des destinataires...")
    destinataires_button_xpath = "//*[@id='root']/div/div/div/div[1]/div/div[3]/div/div/div/div[3]/div/div[2]/div[2]/div/div[1]/h3/span"
    safe_click(driver, destinataires_button_xpath)

    time.sleep(2)

    # ✅ Suppression du popup "first-run-callout-SMS-container" qui bloque les clics
    try:
        sms_block_xpath = "//div[@data-qa='first-run-callout-SMS-container']"
        sms_block = driver.find_element(By.XPATH, sms_block_xpath)
        driver.execute_script("arguments[0].style.display = 'none';", sms_block)
        print("✅ Fenêtre bloquante supprimée.")
    except Exception:
        print("✅ Aucune fenêtre bloquante détectée.")

    # ✅ Trouver tous les champs "Nom" et "Email" des destinataires
    recipient_name_fields = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input[data-qa='recipient-name']")))
    recipient_email_fields = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input[data-qa='recipient-email']")))

    try:
        if contract_type == "IADE":
            try:
                # Rechercher la croix de fermeture si elle est présente
                close_button = driver.find_element("xpath", "//button[@data-qa='close-first-run-callout-SMS']")
                close_button.click()
                print("✅ Fenêtre popup fermée avec succès.")
            except NoSuchElementException:
                print("ℹ️ Aucune fenêtre popup à fermer.")

            # 📌 1️⃣ Destinataire 1 : Vincent PERREARD (Nom & Email de l'entreprise)
            entreprise_nom_field = recipient_name_fields[0]
            entreprise_email_field = recipient_email_fields[0]

            driver.execute_script("arguments[0].scrollIntoView();", entreprise_nom_field)
            time.sleep(0.5)
            entreprise_nom_field.click()  # 🔹 On s'assure que le champ est actif
            for _ in range(50):  # 🔹 On efface caractère par caractère
                entreprise_nom_field.send_keys(Keys.BACKSPACE)
            entreprise_nom_field.send_keys("PERREARD Vincent")

            driver.execute_script("arguments[0].scrollIntoView();", entreprise_email_field)
            time.sleep(0.5)
            entreprise_email_field.click()
            for _ in range(50):
                entreprise_email_field.send_keys(Keys.BACKSPACE)
            entreprise_email_field.send_keys("vincent.perreard@outlook.fr")

            print("✅ Entreprise renseignée correctement.")

            # 📌 2️⃣ Destinataire 2 : IADE Remplaçant
            iade_nom_field = recipient_name_fields[1]
            iade_email_field = recipient_email_fields[1]

            driver.execute_script("arguments[0].scrollIntoView();", iade_nom_field)
            time.sleep(0.5)
            iade_nom_field.click()
            for _ in range(50):
                iade_nom_field.send_keys(Keys.BACKSPACE)
            iade_nom_field.send_keys(replacing_name)

            driver.execute_script("arguments[0].scrollIntoView();", iade_email_field)
            time.sleep(0.5)
            iade_email_field.click()
            for _ in range(50):
                iade_email_field.send_keys(Keys.BACKSPACE)
            iade_email_field.send_keys(replacing_email)

            print(f"✅ IADE ajouté : {replacing_name} - {replacing_email}")

        elif contract_type == "MAR":
            # Destinataire 1 : Médecin remplacé
            input_fields[0].clear()
            input_fields[0].send_keys(replaced_name)
            input_fields[1].clear()
            input_fields[1].send_keys(replaced_email)

            # Destinataire 2 : Médecin remplaçant
            input_fields[2].clear()
            input_fields[2].send_keys(replacing_name)
            input_fields[3].clear()
            input_fields[3].send_keys(replacing_email)

        print("✅ Destinataires ajoutés avec succès.")

    except Exception as e:
        print(f"❌ Erreur lors de l'ajout des destinataires : {e}")



    # 7️⃣ Ajout du message
    print("\n⏳ Ajout du message...")
    message_encart_xpath = "//*[@id='root']/div/div/div/div[1]/div/div[3]/div/div/div/div[3]/div/div[2]/div[3]/div/div[1]/h3/span"
    safe_click(driver, message_encart_xpath)
    print("✅ Encart 'Ajouter un message' ouvert.")
    time.sleep(2)

    # Détection des champs
    input_fields = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//input[@type='text'] | //textarea")))
    objet_email_field, message_field = None, None

    for field in input_fields:
        placeholder = field.get_attribute("placeholder")
        if field.tag_name == "input" and "Complétez avec Docusign" in placeholder:
            objet_email_field = field
        elif field.tag_name == "textarea" and "Saisir un message" in placeholder:
            message_field = field

    # ✅ Formatage correct de la date
    try:
        date_formattee = datetime.strptime(date_contrat, "%d%m%Y").strftime("%d/%m/%Y")
    except ValueError:
        date_formattee = date_contrat  # Si erreur, garde la date brute

    # Effacer et remplir l'objet de l'email
    if objet_email_field:
        effacer_champ(objet_email_field)
        objet_email_field.send_keys(subject)
        print("✅ Objet de l'email rempli.")
    else:
        print("❌ Champ 'Objet' non trouvé.")

    # Effacer et remplir le message de l'email
    if message_field:
        effacer_champ(message_field)
        message_field.send_keys(message)
        print("✅ Message rempli.")
    else:
        print("❌ Champ 'Message' non trouvé.")
        

# ✅ Pause pour permettre la vérification avant fermeture (même en cas de succès)
print("\n🛑 Script terminé. Vérifiez les informations et SIGNEZ le document.")
input("\n➡️ **Appuyez sur Entrée pour fermer manuellement Selenium et le navigateur...**")