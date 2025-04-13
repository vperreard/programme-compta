"""
Module de widgets communs pour l'application de gestion des contrats.
Ce module fournit des widgets personnalisés et réutilisables pour assurer
une cohérence visuelle et fonctionnelle à travers les différents modules de l'application.
"""

import tkinter as tk
from tkinter import ttk, filedialog
import os
from datetime import datetime
from tkcalendar import Calendar
import pandas as pd

# Import des modules personnalisés
from ui_styles import (
    COLORS, FONTS, create_button, create_label, create_frame, 
    create_entry, create_modal_dialog, configure_ttk_style
)
from utils import (
    logger, show_error, show_warning, show_info, show_success, 
    ask_question, format_date, center_window
)

class HeaderFrame(tk.Frame):
    """
    Cadre d'en-tête standard avec titre et sous-titre optionnel.
    """
    def __init__(self, parent, title, subtitle=None, **kwargs):
        """
        Initialise un cadre d'en-tête.
        
        Args:
            parent: Widget parent
            title (str): Titre principal
            subtitle (str, optional): Sous-titre
            **kwargs: Arguments supplémentaires pour le Frame
        """
        bg_color = kwargs.pop("bg", COLORS["primary"])
        super().__init__(parent, bg=bg_color, **kwargs)
        
        # Titre principal
        self.title_label = tk.Label(
            self, 
            text=title,
            font=FONTS["title"],
            bg=bg_color,
            fg=COLORS["text_light"],
            padx=10,
            pady=5
        )
        self.title_label.pack(fill="x", pady=(5, 0))
        
        # Sous-titre optionnel
        if subtitle:
            self.subtitle_label = tk.Label(
                self, 
                text=subtitle,
                font=FONTS["body"],
                bg=bg_color,
                fg=COLORS["text_light"],
                padx=10,
                pady=2
            )
            self.subtitle_label.pack(fill="x", pady=(0, 5))

class FooterFrame(tk.Frame):
    """
    Cadre de pied de page standard avec boutons.
    """
    def __init__(self, parent, **kwargs):
        """
        Initialise un cadre de pied de page.
        
        Args:
            parent: Widget parent
            **kwargs: Arguments supplémentaires pour le Frame
        """
        bg_color = kwargs.pop("bg", COLORS["secondary"])
        super().__init__(parent, bg=bg_color, **kwargs)
        
        # Cadre pour les boutons
        self.button_frame = tk.Frame(self, bg=bg_color)
        self.button_frame.pack(fill="x", pady=10)
    
    def add_button(self, text, command, style="primary", side="left", **kwargs):
        """
        Ajoute un bouton au pied de page.
        
        Args:
            text (str): Texte du bouton
            command: Fonction à exécuter lors du clic
            style (str): Style du bouton
            side (str): Côté où placer le bouton (left, right)
            **kwargs: Arguments supplémentaires pour le bouton
            
        Returns:
            tk.Button: Le bouton créé
        """
        button = create_button(self.button_frame, text, command, style=style, **kwargs)
        button.pack(side=side, padx=5, pady=5)
        return button

class SectionFrame(tk.Frame):
    """
    Cadre de section avec titre et séparateur.
    """
    def __init__(self, parent, title, **kwargs):
        """
        Initialise un cadre de section.
        
        Args:
            parent: Widget parent
            title (str): Titre de la section
            **kwargs: Arguments supplémentaires pour le Frame
        """
        bg_color = kwargs.pop("bg", COLORS["secondary"])
        super().__init__(parent, bg=bg_color, **kwargs)
        
        # Titre de la section
        self.title_label = tk.Label(
            self, 
            text=title,
            font=FONTS["subtitle"],
            bg=bg_color,
            fg=COLORS["primary"],
            padx=5,
            pady=5
        )
        self.title_label.pack(anchor="w", pady=(10, 5))
        
        # Séparateur
        self.separator = ttk.Separator(self, orient="horizontal")
        self.separator.pack(fill="x", pady=(0, 10))
        
        # Cadre pour le contenu
        self.content_frame = tk.Frame(self, bg=bg_color)
        self.content_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
    def get_content_frame(self):
        """
        Retourne le cadre de contenu.
        
        Returns:
            tk.Frame: Le cadre de contenu
        """
        return self.content_frame

class FormField(tk.Frame):
    """
    Champ de formulaire avec label et entrée.
    """
    def __init__(self, parent, label_text, variable=None, entry_type="entry", **kwargs):
        """
        Initialise un champ de formulaire.
        
        Args:
            parent: Widget parent
            label_text (str): Texte du label
            variable (tk.Variable, optional): Variable Tkinter à lier
            entry_type (str): Type d'entrée (entry, combobox, readonly, etc.)
            **kwargs: Arguments supplémentaires pour l'entrée
        """
        bg_color = kwargs.pop("bg", COLORS["secondary"])
        super().__init__(parent, bg=bg_color)
        
        # Label
        self.label = tk.Label(
            self, 
            text=label_text,
            font=FONTS["body"],
            bg=bg_color,
            fg=COLORS["text_dark"],
            anchor="w"
        )
        self.label.pack(fill="x", pady=(5, 2))
        
        # Variable Tkinter
        if variable is None:
            self.variable = tk.StringVar()
        else:
            self.variable = variable
        
        # Entrée selon le type
        if entry_type == "entry":
            self.entry = tk.Entry(
                self, 
                textvariable=self.variable,
                font=FONTS["body"],
                **kwargs
            )
            self.entry.pack(fill="x", pady=(0, 5))
            
        elif entry_type == "combobox":
            self.entry = ttk.Combobox(
                self, 
                textvariable=self.variable,
                font=FONTS["body"],
                state="readonly",
                **kwargs
            )
            self.entry.pack(fill="x", pady=(0, 5))
            
        elif entry_type == "readonly":
            self.entry = tk.Entry(
                self, 
                textvariable=self.variable,
                font=FONTS["body"],
                state="readonly",
                **kwargs
            )
            self.entry.pack(fill="x", pady=(0, 5))
            
        elif entry_type == "text":
            self.entry = tk.Text(
                self, 
                font=FONTS["body"],
                height=kwargs.pop("height", 5),
                width=kwargs.pop("width", 30),
                **kwargs
            )
            self.entry.pack(fill="both", expand=True, pady=(0, 5))
            
            # Méthode pour définir et récupérer le texte
            def set_text(text):
                self.entry.delete("1.0", "end")
                self.entry.insert("1.0", text)
            
            def get_text():
                return self.entry.get("1.0", "end-1c")
            
            self.variable.set = set_text
            self.variable.get = get_text
            
        elif entry_type == "date":
            self.entry = tk.Entry(
                self, 
                textvariable=self.variable,
                font=FONTS["body"],
                **kwargs
            )
            self.entry.pack(side="left", fill="x", expand=True, pady=(0, 5))
            
            # Bouton pour ouvrir le calendrier
            self.calendar_button = create_button(
                self, 
                text="📅", 
                command=self.open_calendar,
                style="neutral",
                width=3,
                height=1
            )
            self.calendar_button.pack(side="left", padx=(5, 0), pady=(0, 5))
    
    def get(self):
        """
        Récupère la valeur du champ.
        
        Returns:
            str: Valeur du champ
        """
        return self.variable.get()
    
    def set(self, value):
        """
        Définit la valeur du champ.
        
        Args:
            value: Valeur à définir
        """
        self.variable.set(value)
    
    def open_calendar(self):
        """
        Ouvre un calendrier pour sélectionner une date.
        """
        def on_date_select():
            selected_date = calendar.get_date()  # Format YYYY-MM-DD
            formatted_date = format_date(selected_date, "%Y-%m-%d", "%d-%m-%Y")
            self.variable.set(formatted_date)
            calendar_window.destroy()
        
        # Fenêtre du calendrier
        calendar_window = tk.Toplevel(self)
        calendar_window.title("Sélectionner une date")
        calendar_window.geometry("300x300")
        calendar_window.transient(self)
        calendar_window.grab_set()
        
        # Calendrier
        calendar = Calendar(
            calendar_window,
            selectmode="day",
            date_pattern="yyyy-mm-dd",
            background=COLORS["secondary"],
            foreground=COLORS["text_dark"],
            selectbackground=COLORS["primary"],
            selectforeground=COLORS["text_light"],
            normalbackground="white",
            weekendbackground="#f5f5f5",
            weekendforeground=COLORS["primary"],
            othermonthforeground="#aaaaaa",
            font=FONTS["body"]
        )
        calendar.pack(padx=10, pady=10, fill="both", expand=True)
        
        # Bouton de validation
        create_button(
            calendar_window, 
            text="Valider", 
            command=on_date_select,
            style="primary"
        ).pack(pady=10)
        
        # Centrer la fenêtre
        center_window(calendar_window)

class FileSelector(tk.Frame):
    """
    Sélecteur de fichier avec label, entrée et bouton.
    """
    def __init__(self, parent, label_text, variable=None, file_type="file", **kwargs):
        """
        Initialise un sélecteur de fichier.
        
        Args:
            parent: Widget parent
            label_text (str): Texte du label
            variable (tk.Variable, optional): Variable Tkinter à lier
            file_type (str): Type de fichier (file, directory)
            **kwargs: Arguments supplémentaires pour l'entrée
        """
        bg_color = kwargs.pop("bg", COLORS["secondary"])
        super().__init__(parent, bg=bg_color)
        
        # Label
        self.label = tk.Label(
            self, 
            text=label_text,
            font=FONTS["body"],
            bg=bg_color,
            fg=COLORS["text_dark"],
            anchor="w"
        )
        self.label.pack(fill="x", pady=(5, 2))
        
        # Cadre pour l'entrée et le bouton
        self.entry_frame = tk.Frame(self, bg=bg_color)
        self.entry_frame.pack(fill="x", pady=(0, 5))
        
        # Variable Tkinter
        if variable is None:
            self.variable = tk.StringVar()
        else:
            self.variable = variable
        
        # Entrée
        self.entry = tk.Entry(
            self.entry_frame, 
            textvariable=self.variable,
            font=FONTS["body"],
            **kwargs
        )
        self.entry.pack(side="left", fill="x", expand=True)
        
        # Type de fichier
        self.file_type = file_type
        
        # Bouton pour sélectionner le fichier
        self.select_button = create_button(
            self.entry_frame, 
            text="...", 
            command=self.select_file,
            style="neutral",
            width=3,
            height=1
        )
        self.select_button.pack(side="left", padx=(5, 0))
    
    def get(self):
        """
        Récupère le chemin du fichier.
        
        Returns:
            str: Chemin du fichier
        """
        return self.variable.get()
    
    def set(self, value):
        """
        Définit le chemin du fichier.
        
        Args:
            value (str): Chemin du fichier
        """
        self.variable.set(value)
    
    def select_file(self):
        """
        Ouvre une boîte de dialogue pour sélectionner un fichier ou un dossier.
        """
        if self.file_type == "file":
            path = filedialog.askopenfilename(title="Sélectionner un fichier")
        else:  # directory
            path = filedialog.askdirectory(title="Sélectionner un dossier")
        
        if path:
            self.variable.set(path)

class SearchableListbox(tk.Frame):
    """
    Listbox avec champ de recherche.
    """
    def __init__(self, parent, items=None, **kwargs):
        """
        Initialise une listbox avec champ de recherche.
        
        Args:
            parent: Widget parent
            items (list, optional): Liste d'éléments à afficher
            **kwargs: Arguments supplémentaires pour la Listbox
        """
        bg_color = kwargs.pop("bg", COLORS["secondary"])
        super().__init__(parent, bg=bg_color)
        
        # Cadre pour le champ de recherche
        self.search_frame = tk.Frame(self, bg=bg_color)
        self.search_frame.pack(fill="x", pady=(0, 5))
        
        # Champ de recherche
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_items)
        
        self.search_entry = tk.Entry(
            self.search_frame, 
            textvariable=self.search_var,
            font=FONTS["body"],
            width=kwargs.pop("width", 30)
        )
        self.search_entry.pack(side="left", fill="x", expand=True)
        
        # Bouton pour effacer la recherche
        self.clear_button = create_button(
            self.search_frame, 
            text="✕", 
            command=lambda: self.search_var.set(""),
            style="warning",
            width=2,
            height=1
        )
        self.clear_button.pack(side="left", padx=(5, 0))
        
        # Cadre pour la listbox et la scrollbar
        self.listbox_frame = tk.Frame(self, bg=bg_color)
        self.listbox_frame.pack(fill="both", expand=True)
        
        # Listbox
        self.listbox = tk.Listbox(
            self.listbox_frame,
            font=FONTS["body"],
            selectbackground=COLORS["primary"],
            selectforeground=COLORS["text_light"],
            **kwargs
        )
        self.listbox.pack(side="left", fill="both", expand=True)
        
        # Scrollbar
        self.scrollbar = ttk.Scrollbar(
            self.listbox_frame, 
            orient="vertical", 
            command=self.listbox.yview
        )
        self.scrollbar.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=self.scrollbar.set)
        
        # Liste des éléments
        self.all_items = items or []
        self.update_items(self.all_items)
    
    def filter_items(self, *args):
        """
        Filtre les éléments selon le texte de recherche.
        """
        search_text = self.search_var.get().lower()
        
        if not search_text:
            filtered_items = self.all_items
        else:
            filtered_items = [item for item in self.all_items if search_text in str(item).lower()]
        
        self.update_listbox(filtered_items)
    
    def update_items(self, items):
        """
        Met à jour la liste des éléments.
        
        Args:
            items (list): Nouvelle liste d'éléments
        """
        self.all_items = items
        self.filter_items()
    
    def update_listbox(self, items):
        """
        Met à jour la listbox avec les éléments filtrés.
        
        Args:
            items (list): Liste d'éléments à afficher
        """
        self.listbox.delete(0, "end")
        for item in items:
            self.listbox.insert("end", item)
    
    def get_selection(self):
        """
        Récupère l'élément sélectionné.
        
        Returns:
            str or None: Élément sélectionné ou None si aucun
        """
        selection = self.listbox.curselection()
        if selection:
            return self.listbox.get(selection[0])
        return None
    
    def get_selections(self):
        """
        Récupère les éléments sélectionnés.
        
        Returns:
            list: Liste des éléments sélectionnés
        """
        selections = self.listbox.curselection()
        return [self.listbox.get(i) for i in selections]
    
    def select(self, index):
        """
        Sélectionne un élément par son index.
        
        Args:
            index (int): Index de l'élément à sélectionner
        """
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(index)
        self.listbox.see(index)
    
    def bind(self, sequence, func):
        """
        Lie un événement à la listbox.
        
        Args:
            sequence (str): Séquence d'événement
            func: Fonction à exécuter
        """
        self.listbox.bind(sequence, func)

class DataTable(tk.Frame):
    """
    Table de données avec fonctionnalités de tri et filtrage.
    """
    def __init__(self, parent, columns, data=None, **kwargs):
        """
        Initialise une table de données.
        
        Args:
            parent: Widget parent
            columns (list): Liste des colonnes (tuples (id, label, width))
            data (list, optional): Liste des données (listes de valeurs)
            **kwargs: Arguments supplémentaires pour le Frame
        """
        bg_color = kwargs.pop("bg", COLORS["secondary"])
        super().__init__(parent, bg=bg_color)
        
        # Configuration des colonnes
        self.columns = columns
        self.column_ids = [col[0] for col in columns]
        
        # Cadre pour la table et la scrollbar
        self.table_frame = tk.Frame(self, bg=bg_color)
        self.table_frame.pack(fill="both", expand=True)
        
        # Treeview
        self.tree = ttk.Treeview(
            self.table_frame, 
            columns=self.column_ids,
            show="headings",
            selectmode="extended",
            **kwargs
        )
        self.tree.pack(side="left", fill="both", expand=True)
        
        # Configuration des colonnes
        for col_id, col_label, col_width in columns:
            self.tree.heading(col_id, text=col_label, command=lambda c=col_id: self.sort_by(c))
            self.tree.column(col_id, width=col_width, anchor="center")
        
        # Scrollbar verticale
        self.vsb = ttk.Scrollbar(
            self.table_frame, 
            orient="vertical", 
            command=self.tree.yview
        )
        self.vsb.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=self.vsb.set)
        
        # Scrollbar horizontale
        self.hsb = ttk.Scrollbar(
            self, 
            orient="horizontal", 
            command=self.tree.xview
        )
        self.hsb.pack(fill="x")
        self.tree.configure(xscrollcommand=self.hsb.set)
        
        # Données
        self.all_data = data or []
        self.filtered_data = self.all_data.copy()
        
        # Remplir la table
        self.update_table()
        
        # Variables pour le tri
        self.sort_column = None
        self.sort_reverse = False
    
    def update_data(self, data):
        """
        Met à jour les données de la table.
        
        Args:
            data (list): Nouvelle liste de données
        """
        self.all_data = data
        self.filtered_data = data.copy()
        self.update_table()
    
    def update_table(self):
        """
        Met à jour l'affichage de la table.
        """
        # Effacer la table
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Remplir avec les données filtrées
        for row in self.filtered_data:
            self.tree.insert("", "end", values=row)
    
    def sort_by(self, column):
        """
        Trie les données par colonne.
        
        Args:
            column (str): ID de la colonne
        """
        # Inverser l'ordre si on clique sur la même colonne
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False
        
        # Index de la colonne
        col_idx = self.column_ids.index(column)
        
        # Trier les données
        self.filtered_data.sort(key=lambda x: x[col_idx], reverse=self.sort_reverse)
        
        # Mettre à jour la table
        self.update_table()
    
    def filter_data(self, filter_func):
        """
        Filtre les données selon une fonction.
        
        Args:
            filter_func (callable): Fonction de filtrage
        """
        self.filtered_data = [row for row in self.all_data if filter_func(row)]
        self.update_table()
    
    def reset_filters(self):
        """
        Réinitialise les filtres.
        """
        self.filtered_data = self.all_data.copy()
        self.update_table()
    
    def get_selection(self):
        """
        Récupère la ligne sélectionnée.
        
        Returns:
            list or None: Valeurs de la ligne sélectionnée ou None si aucune
        """
        selection = self.tree.selection()
        if selection:
            return self.tree.item(selection[0], "values")
        return None
    
    def get_selections(self):
        """
        Récupère les lignes sélectionnées.
        
        Returns:
            list: Liste des valeurs des lignes sélectionnées
        """
        selections = self.tree.selection()
        return [self.tree.item(item, "values") for item in selections]
    
    def bind(self, sequence, func):
        """
        Lie un événement à la table.
        
        Args:
            sequence (str): Séquence d'événement
            func: Fonction à exécuter
        """
        self.tree.bind(sequence, func)

class StatusBar(tk.Frame):
    """
    Barre de statut avec message et indicateur de progression.
    """
    def __init__(self, parent, **kwargs):
        """
        Initialise une barre de statut.
        
        Args:
            parent: Widget parent
            **kwargs: Arguments supplémentaires pour le Frame
        """
        bg_color = kwargs.pop("bg", COLORS["secondary"])
        super().__init__(parent, bg=bg_color, **kwargs)
        
        # Message
        self.message_var = tk.StringVar()
        self.message_label = tk.Label(
            self, 
            textvariable=self.message_var,
            font=FONTS["small"],
            bg=bg_color,
            fg=COLORS["text_dark"],
            anchor="w"
        )
        self.message_label.pack(side="left", fill="x", expand=True, padx=5)
        
        # Indicateur de progression
        self.progress = ttk.Progressbar(
            self, 
            orient="horizontal",
            length=100,
            mode="determinate"
        )
        self.progress.pack(side="right", padx=5)
        
        # Masquer l'indicateur par défaut
        self.progress.pack_forget()
    
    def set_message(self, message):
        """
        Définit le message de la barre de statut.
        
        Args:
            message (str): Message à afficher
        """
        self.message_var.set(message)
    
    def start_progress(self, maximum=100):
        """
        Démarre l'indicateur de progression.
        
        Args:
            maximum (int): Valeur maximale de la progression
        """
        self.progress["maximum"] = maximum
        self.progress["value"] = 0
        self.progress.pack(side="right", padx=5)
    
    def update_progress(self, value):
        """
        Met à jour l'indicateur de progression.
        
        Args:
            value (int): Nouvelle valeur
        """
        self.progress["value"] = value
    
    def stop_progress(self):
        """
        Arrête l'indicateur de progression.
        """
        self.progress.pack_forget()

class TabView(ttk.Notebook):
    """
    Vue à onglets.
    """
    def __init__(self, parent, **kwargs):
        """
        Initialise une vue à onglets.
        
        Args:
            parent: Widget parent
            **kwargs: Arguments supplémentaires pour le Notebook
        """
        super().__init__(parent, **kwargs)
        
        # Configurer le style ttk
        configure_ttk_style()
        
        # Dictionnaire des onglets
        self.tabs = {}
    
    def add_tab(self, title, **kwargs):
        """
        Ajoute un onglet.
        
        Args:
            title (str): Titre de l'onglet
            **kwargs: Arguments supplémentaires pour le Frame
            
        Returns:
            tk.Frame: Le cadre de l'onglet
        """
        # Créer le cadre de l'onglet
        tab_frame = tk.Frame(self, **kwargs)
        
        # Ajouter l'onglet
        self.add(tab_frame, text=title)
        
        # Stocker l'onglet
        self.tabs[title] = tab_frame
        
        return tab_frame
    
    def select_tab(self, title):
        """
        Sélectionne un onglet.
        
        Args:
            title (str): Titre de l'onglet
        """
        if title in self.tabs:
            tab_id = self.index(self.tabs[title])
            self.select(tab_id)
    
    def get_tab(self, title):
        """
        Récupère le cadre d'un onglet.
        
        Args:
            title (str): Titre de l'onglet
            
        Returns:
            tk.Frame: Le cadre de l'onglet ou None si non trouvé
        """
        return self.tabs.get(title)

class DateRangeSelector(tk.Frame):
    """
    Sélecteur de plage de dates.
    """
    def __init__(self, parent, **kwargs):
        """
        Initialise un sélecteur de plage de dates.
        
        Args:
            parent: Widget parent
            **kwargs: Arguments supplémentaires pour le Frame
        """
        bg_color = kwargs.pop("bg", COLORS["secondary"])
        super().__init__(parent, bg=bg_color, **kwargs)
        
        # Variables pour les dates
        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()
        
        # Cadre pour les champs de date
        self.dates_frame = tk.Frame(self, bg=bg_color)
        self.dates_frame.pack(fill="x")
        
        # Champ de date de début
        self.start_date_field = FormField(
            self.dates_frame, 
            "Date de début :", 
            variable=self.start_date_var,
            entry_type="date",
            bg=bg_color
        )
        self.start_date_field.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # Champ de date de fin
        self.end_date_field = FormField(
            self.dates_frame, 
            "Date de fin :", 
            variable=self.end_date_var,
            entry_type="date",
            bg=bg_color
        )
        self.end_date_field.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # Bouton pour sélectionner les deux dates
        self.select_button = create_button(
            self, 
            text="📅 Sélectionner les dates", 
            command=self.select_dates,
            style="neutral"
        )
        self.select_button.pack(pady=5)
    
    def get_dates(self):
        """
        Récupère les dates sélectionnées.
        
        Returns:
            tuple: (date_début, date_fin)
        """
        return (self.start_date_var.get(), self.end_date_var.get())
    
    def set_dates(self, start_date, end_date):
        """
        Définit les dates.
        
        Args:
            start_date (str): Date de début
            end_date (str): Date de fin
        """
        self.start_date_var.set(start_date)
        self.end_date_var.set(end_date)
    
    def select_dates(self):
        """
        Ouvre un calendrier pour sélectionner une plage de dates.
        """
        selected_dates = []

        def on_date_select():
            """Capture la date sélectionnée."""
            selected_date = calendar.get_date()  # Format YYYY-MM-DD
            selected_dates.append(selected_date)
            
            if len(selected_dates) == 1:
                # Première date sélectionnée (début)
                self.start_date_var.set(format_date(selected_date, "%Y-%m-%d", "%d-%m-%Y"))
                message_var.set("Sélectionnez la date de fin.")
            elif len(selected_dates) == 2:
                # Deuxième date sélectionnée (fin)
                # Trier les dates si nécessaire
                start, end = sorted(selected_dates)
                self.start_date_var.set(format_date(start, "%Y-%m-%d", "%d-%m-%Y"))
                self.end_date_var.set(format_date(end, "%Y-%m-%d", "%d-%m-%Y"))
                calendar_window.destroy()

        def close_calendar():
            """Ferme le calendrier en ne retenant qu'une seule date si nécessaire."""
            if len(selected_dates) == 1:
                # Une seule date = même début et fin
                date = selected_dates[0]
                formatted_date = format_date(date, "%Y-%m-%d", "%d-%m-%Y")
                self.start_date_var.set(formatted_date)
                self.end_date_var.set(formatted_date)
            calendar_window.destroy()

        # Fenêtre du calendrier
        calendar_window = tk.Toplevel(self)
        calendar_window.title("Sélectionner les dates")
        calendar_window.geometry("400x400")
        calendar_window.transient(self)
        calendar_window.grab_set()
        
        # Message d'instruction
        message_var = tk.StringVar(value="Sélectionnez la date de début.")
        tk.Label(
            calendar_window, 
            textvariable=message_var,
            font=FONTS["body"],
            bg=COLORS["secondary"],
            fg=COLORS["text_dark"]
        ).pack(pady=10)
        
        # Calendrier
        calendar = Calendar(
            calendar_window,
            selectmode="day",
            date_pattern="yyyy-mm-dd",
            background=COLORS["secondary"],
            foreground=COLORS["text_dark"],
            selectbackground=COLORS["primary"],
            selectforeground=COLORS["text_light"],
            normalbackground="white",
            weekendbackground="#f5f5f5",
            weekendforeground=COLORS["primary"],
            othermonthforeground="#aaaaaa",
            font=FONTS["body"]
        )
        calendar.pack(padx=10, pady=10, fill="both", expand=True)
        
        # Boutons
        buttons_frame = tk.Frame(calendar_window, bg=COLORS["secondary"])
        buttons_frame.pack(pady=10, fill="x")
        
        create_button(
            buttons_frame, 
            text="Valider", 
            command=on_date_select,
            style="primary"
        ).pack(side="left", padx=10)
        
        create_button(
            buttons_frame, 
            text="Fermer", 
            command=close_calendar,
            style="warning"
        ).pack(side="right", padx=10)
        
        # Centrer la fenêtre
        center_window(calendar_window)
