import os
import json
import hashlib
from datetime import datetime

# Chemin du fichier pour stocker la base de données des factures
FACTURES_DB = os.path.join(os.path.expanduser("~"), "script contrats", "factures_db.json")

def calculer_hash_fichier(chemin_fichier):
    """
    Calcule l'empreinte MD5 d'un fichier pour l'identifier de façon unique.
    """
    hash_md5 = hashlib.md5()
    try:
        with open(chemin_fichier, "rb") as f:
            # Lire le fichier par morceaux pour gérer les gros fichiers
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except Exception as e:
        print(f"Erreur lors du calcul du hash pour {chemin_fichier}: {e}")
        return None

def charger_base_donnees():
    """
    Charge la base de données des factures depuis le fichier JSON.
    """
    if os.path.exists(FACTURES_DB):
        try:
            with open(FACTURES_DB, "r") as f:
                return json.load(f)
        except json.JSONDecodeError:
            print("Erreur: Base de données corrompue, création d'une nouvelle base.")
            return {"factures": {}, "derniere_mise_a_jour": datetime.now().isoformat()}
    else:
        return {"factures": {}, "derniere_mise_a_jour": datetime.now().isoformat()}

def sauvegarder_base_donnees(db):
    """
    Sauvegarde la base de données des factures dans le fichier JSON.
    """
    db["derniere_mise_a_jour"] = datetime.now().isoformat()
    try:
        with open(FACTURES_DB, "w") as f:
            json.dump(db, f, indent=2)
    except Exception as e:
        print(f"Erreur lors de la sauvegarde de la base de données: {e}")

def ajouter_ou_maj_facture(db, fichier, chemin_complet, infos_facture):
    """
    Ajoute une nouvelle facture à la base de données ou met à jour une facture existante.
    
    Args:
        db: Base de données des factures
        fichier: Nom du fichier
        chemin_complet: Chemin complet du fichier
        infos_facture: Dictionnaire avec les informations de la facture
    
    Returns:
        id_facture: Identifiant de la facture dans la base
        est_nouvelle: True si c'est une nouvelle facture, False si mise à jour
    """
    # Calculer le hash du fichier
    hash_fichier = calculer_hash_fichier(chemin_complet)
    if not hash_fichier:
        return None, False
    
    # Créer un dictionnaire inversé pour recherche par hash
    hash_to_id = {info["hash"]: id_facture for id_facture, info in db["factures"].items()}
    
    # Vérifier si ce fichier existe déjà dans la base (même hash)
    if hash_fichier in hash_to_id:
        id_facture = hash_to_id[hash_fichier]
        # Mettre à jour l'emplacement si nécessaire
        if db["factures"][id_facture]["chemin"] != chemin_complet:
            print(f"Fichier déplacé détecté: {db['factures'][id_facture]['chemin']} -> {chemin_complet}")
            db["factures"][id_facture]["chemin"] = chemin_complet
        
        # Marquer comme disponible
        db["factures"][id_facture]["disponible"] = True
        
        return id_facture, False
    else:
        # Nouvelle facture
        id_facture = str(len(db["factures"]) + 1).zfill(5)  # ID simple séquentiel
        
        # Stocker les informations avec le hash
        db["factures"][id_facture] = {
            "nom": fichier,
            "chemin": chemin_complet,
            "hash": hash_fichier,
            "date_analyse": datetime.now().isoformat(),
            "disponible": True,
            **infos_facture  # Ajouter toutes les autres infos extraites
        }
        
        return id_facture, True

def marquer_factures_manquantes(db, factures_trouvees):
    """
    Marque comme indisponibles les factures qui n'ont pas été trouvées.
    
    Args:
        db: Base de données des factures
        factures_trouvees: Set des IDs des factures trouvées
    """
    for id_facture in db["factures"]:
        if id_facture not in factures_trouvees:
            db["factures"][id_facture]["disponible"] = False