import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import json
import re


def ajouter_info_bulle(widget, texte):
    def on_enter(event):
        x, y, _, _ = widget.bbox("insert")
        x += widget.winfo_rootx() + 25
        y += widget.winfo_rooty() + 20
        global info_bulle
        info_bulle = tk.Toplevel(widget)
        info_bulle.wm_overrideredirect(True)
        info_bulle.wm_geometry(f"+{x}+{y}")
        label = tk.Label(info_bulle, text=texte, background="lightyellow", relief='solid', borderwidth=1, font=("Helvetica", 9))
        label.pack(ipadx=1)
    def on_leave(event):
        global info_bulle
        if info_bulle:
            info_bulle.destroy()
    widget.bind("<Enter>", on_enter)
    widget.bind("<Leave>", on_leave)
    
def extraire_mars_depuis_fichier(chemin_fichier):
    try:
        tables = pd.read_html(chemin_fichier)  # Lecture des tableaux HTML
        df = tables[0]  # Le premier tableau contient généralement les données utiles

        # 🔧 Afficher toutes les colonnes et les 100 premières lignes
        pd.set_option("display.max_rows", 500)
        pd.set_option("display.max_columns", None)
        pd.set_option("display.width", None)
        pd.set_option("display.max_colwidth", None)

        print(df.head(20))  # Affichage debug
        return df
    except Exception as e:
        print(f"Erreur lors de la lecture HTML : {e}")
        return pd.DataFrame()
    
def extraire_noms_mars(df):
    mots_cles = ["Gardes", "CS1", "CS2", "Absences"]
    noms_mars = set()

    for index, row in df.iterrows():
        cellule_de_gauche = str(row.iloc[0]).strip()
        if any(cellule_de_gauche.startswith(cle) for cle in mots_cles):
            for cell in row:
                if isinstance(cell, str):
                    # On sépare sur les espaces
                    noms_possibles = cell.strip().split()
                    for nom in noms_possibles:
                        if nom.isupper() and re.fullmatch(r"[A-ZÉÈÀÂÊÎÔÛÄËÏÖÜÇ\-]{2,}", nom):
                            noms_mars.add(nom)

    return sorted(noms_mars)

def gestion_mars_gui():
    chemin_json = "liste_mars.json"

    if os.path.exists(chemin_json):
        with open(chemin_json, "r") as f:
            data = json.load(f)
    else:
        data = {}

    def importer_mars():
        chemin = filedialog.askopenfilename(title="Sélectionner le fichier Momentum", filetypes=[("Fichiers XLS/HTML", "*.xls *.html")])
        if not chemin:
            return
        df = extraire_mars_depuis_fichier(chemin)
        mars_extraits = extraire_noms_mars(df)
        for nom in mars_extraits:
            if nom not in data:
                data[nom] = {
                    "mi_temps": False,
                    "paires": [],
                    "impaires": []
                }
        data[nom]["paires"] = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi"]
        data[nom]["impaires"] = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi"]
        charger_interface()

    def sauvegarder():
        for nom in data:
            data[nom]["paires"] = [j for j, var in vars_widgets[nom]["paires"].items() if var.get()]
            data[nom]["impaires"] = [j for j, var in vars_widgets[nom]["impaires"].items() if var.get()]
            data[nom]["mi_temps_pairs"] = vars_widgets[nom]["mi_temps_pairs"].get()
            data[nom]["mi_temps_impairs"] = vars_widgets[nom]["mi_temps_impairs"].get()
        with open(chemin_json, "w") as f:
            json.dump(data, f, indent=2)
        messagebox.showinfo("✅ Sauvegarde", "Données enregistrées avec succès.")

    def charger_interface():
        for widget in cadre_scrollable.winfo_children():
            widget.destroy()
        vars_widgets.clear()

        jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi"]

        for idx, nom in enumerate(sorted(data.keys())):
            paire_vars = {}
            impaire_vars = {}

            vars_widgets[nom] = {
                "paires": paire_vars,
                "impaires": impaire_vars,
                "mi_temps_pairs": tk.BooleanVar(value=data[nom].get("mi_temps_pairs", False)),
                "mi_temps_impairs": tk.BooleanVar(value=data[nom].get("mi_temps_impairs", False))
            }

            # Ligne du nom avec bouton de suppression
            cadre_nom = ttk.Frame(cadre_scrollable)
            cadre_nom.grid(row=idx*4, column=0, columnspan=6, sticky="w", pady=(10, 0))

            label_nom = ttk.Label(cadre_nom, text=nom, font=('Helvetica', 10, 'bold'))
            label_nom.pack(side="left")

            def supprimer_mars(nom_mars=nom):
                if messagebox.askyesno("🗑 Confirmation", f"Supprimer {nom_mars} de la liste ?"):
                    del data[nom_mars]
                    charger_interface()

            bouton_supprimer = ttk.Button(cadre_nom, text="🗑 Supprimer", command=supprimer_mars)
            bouton_supprimer.pack(side="left", padx=10)

            # ✅ Semaine paire
            cadre_paire = ttk.Frame(cadre_scrollable)
            cadre_paire.grid(row=idx*4+1, column=0, columnspan=6, sticky="w", padx=20)
            ttk.Label(cadre_paire, text="Semaine paire").pack(side="left", padx=(0, 10))
            for jour in jours:
                var = tk.BooleanVar(value=True)
                cb = ttk.Checkbutton(cadre_paire, text=jour, variable=var)
                cb.pack(side="left", padx=6)
                paire_vars[jour] = var

            # ✅ Semaine impaire
            cadre_impaire = ttk.Frame(cadre_scrollable)
            cadre_impaire.grid(row=idx*4+2, column=0, columnspan=6, sticky="w", padx=20)
            ttk.Label(cadre_impaire, text="Semaine impaire").pack(side="left", padx=(0, 10))
            for jour in jours:
                var = tk.BooleanVar(value=True)
                cb = ttk.Checkbutton(cadre_impaire, text=jour, variable=var)
                cb.pack(side="left", padx=6)
                impaire_vars[jour] = var

            # Options mi-temps mois pairs / impairs
            mt_pairs = vars_widgets[nom]["mi_temps_pairs"]
            mt_impairs = vars_widgets[nom]["mi_temps_impairs"]

            cb1 = ttk.Checkbutton(cadre_scrollable, text="🗓 Mi-temps mois pairs", variable=mt_pairs)
            cb1.grid(row=idx*4+3, column=1, sticky="w", padx=5)
            ajouter_info_bulle(cb1, "Le médecin ne travaille qu’un mois sur deux : mois pairs (février, avril, etc.)")

            cb2 = ttk.Checkbutton(cadre_scrollable, text="📆 Mi-temps mois impairs", variable=mt_impairs)
            cb2.grid(row=idx*4+3, column=2, sticky="w", padx=5)
            ajouter_info_bulle(cb2, "Le médecin ne travaille qu’un mois sur deux : mois impairs (janvier, mars, etc.)")




    # Création de la fenêtre principale
    fenetre = tk.Toplevel()
    fenetre.title("🩺 Gestion des MARS")
    fenetre.geometry("1000x800")
    fenetre.configure(bg="#f0f4f7")

    label = ttk.Label(fenetre, text="🩺 Déclaration des MARS à mi-temps et leurs jours de travail :", font=('Helvetica', 14, 'bold'))
    label.pack(pady=10)

    bouton_import = ttk.Button(fenetre, text="📥 Importer depuis Momentum", command=importer_mars)
    bouton_import.pack(pady=5)

    # ➕ Ajout manuel d’un MARS
    def ajouter_mars_manuellement():
        nom = entree_nom.get().strip().upper()
        if not nom:
            messagebox.showwarning("⚠️ Nom vide", "Veuillez entrer un nom.")
            return
        if not re.fullmatch(r"[A-ZÉÈÀÂÊÎÔÛÄËÏÖÜÇ\-]{2,}", nom):
            messagebox.showwarning("⚠️ Format incorrect", "Le nom doit être en majuscules et sans chiffres.")
            return
        if nom in data:
            messagebox.showinfo("ℹ️ Déjà présent", f"{nom} est déjà dans la liste.")
            return
        data[nom] = {
            "mi_temps": False,
            "paires": [],
            "impaires": [],
            "mi_temps_pairs": False,
            "mi_temps_impairs": False
        }
        entree_nom.delete(0, tk.END)
        data[nom]["paires"] = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi"]
        data[nom]["impaires"] = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi"]
        charger_interface()

    # ➕ Ajout manuel d’un MARS (zone texte + bouton)
    cadre_ajout = ttk.Frame(fenetre)
    cadre_ajout.pack(pady=(5, 0))

    label_ajout = ttk.Label(cadre_ajout, text="Ajouter un MAR :")
    label_ajout.grid(row=0, column=0, padx=5)

    entree_nom = ttk.Entry(cadre_ajout, width=30)
    entree_nom.grid(row=0, column=1, padx=5)

    bouton_ajout = ttk.Button(cadre_ajout, text="➕ Ajouter MARS", command=ajouter_mars_manuellement)
    bouton_ajout.grid(row=0, column=2, padx=5)

    # ➕ Zone scrollable dans un cadre
    cadre_scroll = ttk.Frame(fenetre)
    cadre_scroll.pack(padx=10, pady=(5, 0), fill="both", expand=True)

    # Zone scrollable avec scroll uniquement sur la largeur du contenu (colonne médecins)
    canvas = tk.Canvas(cadre_scroll, height=500, borderwidth=0, background="#f0f4f7")
    scrollbar = ttk.Scrollbar(cadre_scroll, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    # Placement du canvas et scrollbar côte à côte
    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0, column=1, sticky="ns")

    # Faire en sorte que la scrollzone prenne seulement l'espace nécessaire
    cadre_scroll.columnconfigure(0, weight=1)
    cadre_scroll.rowconfigure(0, weight=1)

    # Frame scrollable (la zone avec médecins)
    cadre_scrollable = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=cadre_scrollable, anchor="nw")

    # Mise à jour de la zone scrollable
    cadre_scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))


    vars_widgets = {}
    charger_interface()

    # 📌 Note explicative
    label_explication = ttk.Label(
        fenetre,
        text="ℹ️ Si vous cochez 'mi-temps mois pairs' ou 'mois impairs', les absences ne seront comptées que dans les mois concernés,\nà condition que les jours de semaine sélectionnés soient aussi actifs.",
        font=("Helvetica", 8),
        foreground="gray"
    )
    label_explication.pack(pady=(10, 2))

    bouton_sauvegarder = ttk.Button(fenetre, text="💾 Sauvegarder", command=sauvegarder)
    bouton_sauvegarder.pack(pady=10)

    bouton_fermer = ttk.Button(fenetre, text="❌ Fermer", command=fenetre.destroy)
    bouton_fermer.pack(pady=5)