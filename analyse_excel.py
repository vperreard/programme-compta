import pandas as pd
from datetime import datetime
import re
import os

def est_mars(nom):
    """Renvoie True si le nom est en majuscules (MARS) et de longueur > 1."""
    return nom.isupper() and len(nom) > 1

def est_iade(nom):
    """Renvoie True si le nom n'est pas considéré comme MARS."""
    return not est_mars(nom)

def propager_dates(row):
    """
    Parcourt la ligne contenant les dates (sous forme de strings "dd/mm/yyyy")
    et propage la dernière date trouvée sur toutes les colonnes (pour gérer les cellules fusionnées).
    """
    dates = []
    current = None
    for i, cell in enumerate(row):
        cell_str = str(cell) if not pd.isna(cell) else ""
        # Cherche un motif type 03/03/2025
        match = re.search(r'\d{2}/\d{2}/\d{4}', cell_str)
        if match:
            current = match.group(0)  # ex: "03/03/2025"
        dates.append(current)
    return dates

def charger_fichier_momentum(fichier):
    """
    Tente de charger le fichier :
    - d'abord en Excel (pd.read_excel),
    - sinon en HTML (pd.read_html).
    Si HTML, parcourt les tables à la recherche d'une qui contient :
        - "Absences" et "Présences IADE" en colonne 0
        - une ligne contenant un motif de date (ex: 03/03/2025 ou lun, 03/03/2025)
    Renvoie : df, is_html
    """
    import pandas as pd
    import re

    def table_contient_date(df):
        """Vérifie si le DataFrame contient au moins une ligne avec une date type dd/mm/yyyy."""
        for i in range(min(len(df), 30)):  # On limite à 30 lignes pour la perf
            row = df.iloc[i]
            row_str = " ".join(str(x) for x in row if pd.notna(x))
            if re.search(r'\d{2}[/-]\d{2}[/-]\d{2,4}', row_str):
                return True
        return False

    # --- 1) Tentative de lecture Excel ---
    try:
        df = pd.read_excel(fichier, header=None)
        print(f"Debug: Lecture avec read_excel OK, shape={df.shape}")
        return df, False
    except Exception as e:
        print(f"Debug: Echec de read_excel : {e}")

    # --- 2) Lecture HTML avec inspection de toutes les tables ---
    try:
        tables = pd.read_html(fichier, header=None)
        print(f"Debug: Lecture avec read_html OK, nb_tables={len(tables)}")

        for i, t in enumerate(tables[:100]):  # On limite à 100 pour ne pas tout scanner
            if t.shape[1] == 0:
                continue

            col0_str = t.iloc[:, 0].astype(str)
            has_absences = col0_str.str.contains("Absences", case=False, na=False).any()
            has_presences = col0_str.str.contains("Présences IADE", case=False, na=False).any()
            has_date = table_contient_date(t)

            if has_absences and has_presences and has_date:
                print(f"✅ Table #{i} sélectionnée (Absences + Présences IADE + date). shape={t.shape}")
                return t, True


        # Si aucune table ne convient, on prend quand même la 1re par défaut
        if len(tables) > 0:
            print("⚠️ Aucune table satisfaisante, on prend la table #0 par défaut.")
            return tables[0], True
        else:
            print("❌ Aucune table détectée.")
            return None, True

    except Exception as e:
        print(f"❌ Echec de read_html : {e}")
        return None, False


def trouver_ligne_dates(df):
    """
    Détermine si les dates sont dans les noms de colonnes ou dans une ligne du DataFrame.
    Renvoie :
    - ('columns_header', None) si les colonnes contiennent les dates (ex: "lun, 03/03/2025")
    - ('rows', i) si une ligne i contient des dates (fusionnées)
    - (None, None) si rien trouvé
    """
    import re

    # Vérifie si les colonnes contiennent les dates
    for col in df.columns:
        if isinstance(col, str) and re.search(r'\d{2}/\d{2}/\d{4}', col):
            return 'columns_header', None

    # Sinon, cherche dans les lignes
    for i in range(len(df)):
        row = df.iloc[i]
        row_str = " ".join(str(x) for x in row if pd.notna(x))
        if re.search(r'\d{2}/\d{2}/\d{4}', row_str):
            return 'rows', i

    return None, None

def analyser_fichier(fichier, convertir_excel=False, convertir_csv=False):
    print(f"📂 Analyse du fichier : {fichier}")

    df, is_html = charger_fichier_momentum(fichier)
    # Juste après avoir chargé df
    print("=== Debug : Aperçu de la table sélectionnée ===")
    for i in range(min(500, len(df))):  # on affiche les 500 premières lignes
        row_list = [str(x) for x in df.iloc[i].values if pd.notna(x)]
        row_str = " | ".join(row_list)
        print(f"Ligne {i}: {row_str}")
    print("===============================================")
    if df is None:
        print("Debug: Impossible de charger le fichier sous forme DataFrame.")
        return None

    print("Debug: DataFrame chargé avec shape :", df.shape)

    # Optionnel : si c'est du HTML, on peut le convertir en vrai XLSX/CSV
    if is_html and convertir_excel:
        out_xlsx = fichier + ".converted.xlsx"
        df.to_excel(out_xlsx, index=False)
        print(f"Debug: Fichier converti en {out_xlsx}")

    if is_html and convertir_csv:
        out_csv = fichier + ".converted.csv"
        df.to_csv(out_csv, index=False, sep=";")
        print(f"Debug: Fichier converti en {out_csv}")

    # --- 1) Identifier la ligne qui contient les dates ---
    orientation, index_dates = trouver_ligne_dates(df)
    if orientation is None:
        print("❌ Impossible de trouver une ligne ou des colonnes contenant des dates.")
        return None

    if orientation == 'columns_header':
        print("✅ Les dates ont été trouvées dans les noms de colonnes.")
        dates_par_colonne = list(df.columns)
        df.columns = range(len(df.columns))  # Reset des noms de colonnes pour faciliter la suite
        df = df.reset_index(drop=True)
    else:
        print(f"✅ Les dates ont été trouvées dans la ligne {index_dates}")
        header_row = df.iloc[index_dates]
        dates_par_colonne = propager_dates(header_row)


    def nettoyer_date(date):
        if not date or not isinstance(date, str):
            return None
        # Cherche une date dans le texte (comme "lun, 03/03/2025")
        match = re.search(r'(\d{2}/\d{2}/\d{4})', date)
        if match:
            return match.group(1)
        return None

    dates_par_colonne = [nettoyer_date(str(d)) for d in dates_par_colonne]
    print("🧪 Clés uniformisées :", dates_par_colonne[:5])

    # --- 2) Repérer la 2e occurrence de la ligne "Absences" dans la colonne 0 ---
    col0_str = df.iloc[:, 0].astype(str)
    lignes_absences = col0_str[col0_str.str.contains("Absences", case=False, na=False)].index.tolist()
    print("Debug: Lignes contenant 'Absences' :", lignes_absences)

    if len(lignes_absences) < 2:
        print("❌ Impossible de trouver la 2e ligne 'Absences'")
        return None
    ligne_absences = lignes_absences[1]
    print(f"Debug: On utilisera la ligne d'absences à l'index {ligne_absences}")

    # --- 3) Repérer la ligne "Présences IADE" ---
    lignes_presences = col0_str[col0_str.str.contains("Présences IADE", case=False, na=False)].index.tolist()
    print("Debug: Lignes contenant 'Présences IADE' :", lignes_presences)

    if not lignes_presences:
        print("❌ Ligne 'Présences IADE' introuvable")
        return None

    idx_presences = lignes_presences[0]
    
    # Recherche dynamique plus souple des lignes contenant "GF", "G" et "P"
    ligne_GF = ligne_G = ligne_P = None
    for i in range(idx_presences, min(idx_presences + 600, len(df))):
        ligne_vals = [str(cell).strip().upper() for cell in df.iloc[i] if pd.notna(cell)]
        if "GF" in ligne_vals and ligne_GF is None:
            ligne_GF = i
        if "G" in ligne_vals and ligne_G is None:
            ligne_G = i
        if "P" in ligne_vals and ligne_P is None:
            ligne_P = i
        if all([ligne_GF, ligne_G, ligne_P]):
            break



    # Vérification obligatoire
    if None in (ligne_GF, ligne_G, ligne_P):
        print("\n❌ Erreur : Impossible d'identifier toutes les lignes GF, G et P automatiquement.")
        print("   ➤ Vérifiez dans la colonne 0 du fichier HTML que les intitulés 'GF', 'G' et 'P' sont bien présents entre les lignes", idx_presences, "et", idx_presences + 50)
        print("   💡 Sinon, vous pouvez fixer les lignes à la main comme ceci :")
        print("       ligne_GF = 155\n       ligne_G = 165\n       ligne_P = 166\n")
        return None



    # --- 4) Extraire absences et présences ---
    absents_par_date = {}
    presences_par_date = {}

    # On itère sur les colonnes pour extraire
    for col, date in enumerate(dates_par_colonne):
        if not date:
            # Pas de date pour cette colonne
            continue

        print(f"\n📅 {date} — colonne {col}")
        absents_par_date.setdefault(date, {"mars": [], "iade": []})
        presences_par_date.setdefault(date, {"GF": [], "G": [], "P": []})

        # a) Absences
        cell_absence = df.iloc[ligne_absences, col]
        print(f"  🔸 Absences : {cell_absence}")
        if pd.notna(cell_absence):
            val_str = str(cell_absence).strip()
            if val_str and val_str.lower() != "nan":
                # Sépare correctement les noms même s’ils sont collés avec plusieurs espaces
                noms = re.split(r'[,\n\r|]+', val_str)
                for nom in noms:
                    sous_noms = re.split(r'\s{2,}', nom.strip())  # coupe si 2+ espaces
                    for sous_nom in sous_noms:
                        nom_clean = sous_nom.strip()
                        if nom_clean:
                            if est_mars(nom_clean):
                                absents_par_date[date]["mars"].append(nom_clean)
                            else:
                                absents_par_date[date]["iade"].append(nom_clean)


        # b) Présences GF, G, P
        for label, ligne in zip(["GF", "G", "P"], [ligne_GF, ligne_G, ligne_P]):
            if ligne < len(df):
                val = df.iloc[ligne, col]
                # Debug : contenu brut + ligne complète
                ligne_complete = df.iloc[ligne].tolist()
                print(f"  🔹 {label} [{ligne}] → cellule: {val}")
                print(f"     ➤ Ligne {label} complète : {[str(x) for x in ligne_complete]}")

                if pd.notna(val):
                    val_str = str(val).strip()
                    if val_str and val_str.lower() != "nan":
                        noms = re.split(r'[,\n\r|]+', val_str)
                        for nom in noms:
                            sous_noms = re.split(r'\s{2,}', nom.strip())
                            for sous_nom in sous_noms:
                                nom_clean = sous_nom.strip()
                                if nom_clean:
                                    presences_par_date[date][label].append(nom_clean)
                                
    return absents_par_date, presences_par_date

def analyse_depuis_script(fichier):
    """
    Fonction réutilisable depuis un autre script.
    Prend en entrée un fichier HTML/XLS et renvoie :
        - absents_par_date
        - presences_par_date
    """
    return analyser_fichier(fichier, convertir_excel=False, convertir_csv=False)

# ─────────────────────────────────────────────────────────────
#                      SCRIPT PRINCIPAL
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    fichier = input("👉 Glissez ici le fichier téléchargé depuis Momentum (.xls ou .html), puis appuyez sur Entrée :\n").strip()
    if not os.path.exists(fichier):
        print("❌ Fichier introuvable.")
    else:
        result = analyser_fichier(fichier, convertir_excel=False, convertir_csv=False)
        if result:
            absents, presences = result

            print("\n📊 Résumé des données par date :")
            for date in sorted(absents):
                mars = absents[date].get("mars", [])
                iades_abs = absents[date].get("iade", [])

                gf = presences[date].get("GF", [])
                g  = presences[date].get("G", [])
                p  = presences[date].get("P", [])

                grande_journee = gf + g
                petite_journee = p

                print(f"📅 {date}")
                print(f"   🔴 Absents : {len(mars)} MAR ({', '.join(mars) if mars else '-'}) | {len(iades_abs)} IADEs ({', '.join(iades_abs) if iades_abs else '-'})")
                print(f"   ✅ Présents : {len(grande_journee)} IADEs grande journée (GF+G) | {len(petite_journee)} petite journée (P)\n")
        else:
            print("Debug: Aucune donnée n'a été extraite.")