import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_path
import os

# 📂 Chemins des dossiers
INPUT_FOLDER = r"/Users/vincentperreard/script contrats/dossier factures test/factures à tester"
OUTPUT_FOLDER = r"/Users/vincentperreard/script contrats/dossier factures test/factures redressees"

# Créer le dossier de sortie s'il n'existe pas
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def deskew_region(image, x, y, w, h):
    """Corrige l'inclinaison globale de l'image uniquement si nécessaire"""
    roi = image[y:y+h, x:x+w]
    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 50, 150)

    # Détecter les lignes principales
    lines = cv2.HoughLinesP(edges, 1, np.pi / 180, 100, minLineLength=50, maxLineGap=5)

    # Calculer l'angle moyen
    angles = []
    if lines is not None:
        for line in lines:
            x1, y1, x2, y2 = line[0]
            angle = np.arctan2(y2 - y1, x2 - x1) * 180 / np.pi
            angles.append(angle)

    median_angle = np.median(angles) if angles else 0.0

    # Appliquer une correction uniquement si l'angle dépasse 5°
    if abs(median_angle) > 5:
        center = (w // 2, h // 2)
        M = cv2.getRotationMatrix2D(center, median_angle, 1.0)
        rotated = cv2.warpAffine(roi, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
        image[y:y+h, x:x+w] = rotated  # Remettre la région corrigée
        print(f"📌 Redressement appliqué : {median_angle:.2f}°")
    else:
        print("✅ Pas besoin de redressement, l'angle détecté est trop faible.")

    return image

def process_invoice(file_path):
    """Lit une facture, redresse les zones inclinées et applique l'OCR"""
    filename = os.path.basename(file_path)
    output_image_path = os.path.join(OUTPUT_FOLDER, os.path.splitext(filename)[0] + ".jpg")  # Enregistrement en JPG

    if file_path.lower().endswith(".pdf"):
        try:
            images = convert_from_path(file_path, poppler_path="/opt/homebrew/bin/")  # Vérifie ton chemin poppler
            if not images:
                print(f"❌ Erreur : Aucun visuel extrait du PDF {file_path}")
                return
            image = np.array(images[0])  # Prend la première page
            image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)
        except Exception as e:
            print(f"❌ Erreur lors de la conversion du PDF : {str(e)}")
            return
    else:
        image = cv2.imread(file_path)

    # Vérifier si l'image a été chargée correctement
    if image is None:
        print(f"❌ Erreur : Impossible de charger {file_path}")
        return

    # Détection des blocs de texte
    data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)
    n_boxes = len(data['text'])

    for i in range(n_boxes):
        if int(data['conf'][i]) > 50:  # On corrige uniquement si l'OCR détecte du texte
            x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
            image = deskew_region(image, x, y, w, h)

    # Sauvegarde de l'image corrigée en JPG
    success = cv2.imwrite(output_image_path, image)
    if not success:
        print(f"❌ Erreur lors de la sauvegarde de {output_image_path}")
        return

    print(f"✅ Image corrigée sauvegardée : {output_image_path}")

    # Extraction OCR après correction
    text = pytesseract.image_to_string(image, lang="eng")

    # Sauvegarde du texte OCR dans un fichier
    text_output_path = os.path.join(OUTPUT_FOLDER, os.path.splitext(filename)[0] + "_OCR.txt")
    with open(text_output_path, "w", encoding="utf-8") as f:
        f.write(text)
    
    print(f"📄 Texte extrait sauvegardé dans : {text_output_path}\n")
    return text

# 🔍 Traiter tous les fichiers du dossier
files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith((".jpg", ".jpeg", ".png", ".pdf"))]

if not files:
    print("⚠️ Aucun fichier trouvé dans le dossier : ", INPUT_FOLDER)
else:
    for file in files:
        file_path = os.path.join(INPUT_FOLDER, file)
        print(f"\n📌 Traitement de {file}...")
        process_invoice(file_path)