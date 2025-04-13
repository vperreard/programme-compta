"""
Module de styles UI communs pour l'application de gestion des contrats.
Ce module centralise tous les éléments de style pour assurer une cohérence visuelle
à travers les différents modules de l'application.
"""

import tkinter as tk
from tkinter import ttk
import os

# Couleurs principales
COLORS = {
    "primary": "#2c5aa0",       # Bleu principal plus foncé
    "primary_dark": "#1e3c6a",  # Bleu foncé (pour hover)
    "secondary": "#f5f5f5",     # Gris clair (fond)
    "accent": "#228B22",        # Vert (actions positives) plus foncé
    "accent_dark": "#006400",   # Vert foncé (hover)
    "warning": "#c62828",       # Rouge (actions négatives) plus foncé
    "warning_dark": "#8e0000",  # Rouge foncé (hover)
    "info": "#1a5276",          # Bleu ciel (informations) plus foncé
    "info_dark": "#0c2c40",     # Bleu ciel foncé (hover)
    "neutral": "#7B8FA1",       # Bleu lavande (actions neutres) plus foncé
    "neutral_dark": "#5d6b78",  # Bleu lavande foncé (hover)
    "text_dark": "#333333",     # Texte foncé
    "text_light": "#ffffff",    # Texte clair
    "border": "#dddddd",        # Bordures
    "highlight": "#D35400",     # Orange (mise en évidence) plus foncé
    "highlight_dark": "#A04000" # Orange foncé (hover)
}

# Polices
FONTS = {
    "title": ("Arial", 14, "bold"),
    "subtitle": ("Arial", 12, "bold"),
    "body": ("Arial", 10),
    "body_bold": ("Arial", 10, "bold"),
    "small": ("Arial", 9),
    "button": ("Arial", 10, "bold"),
    "large_button": ("Arial", 12, "bold")
}

# Styles de boutons
BUTTON_STYLES = {
    "primary": {
        "bg": COLORS["primary"],
        "fg": COLORS["text_dark"],
        "font": FONTS["button"],
        "activebackground": COLORS["primary_dark"],
        "activeforeground": COLORS["text_dark"],
        "relief": "flat",
        "borderwidth": 0,
        "padx": 10,
        "pady": 5
    },
    "secondary": {
        "bg": COLORS["secondary"],
        "fg": COLORS["text_dark"],
        "font": FONTS["button"],
        "activebackground": COLORS["border"],
        "activeforeground": COLORS["text_dark"],
        "relief": "flat",
        "borderwidth": 0,
        "padx": 10,
        "pady": 5
    },
    "accent": {
        "bg": COLORS["accent"],
        "fg": COLORS["text_dark"],
        "font": FONTS["button"],
        "activebackground": COLORS["accent_dark"],
        "activeforeground": COLORS["text_dark"],
        "relief": "flat",
        "borderwidth": 0,
        "padx": 10,
        "pady": 5
    },
    "warning": {
        "bg": COLORS["warning"],
        "fg": COLORS["text_dark"],
        "font": FONTS["button"],
        "activebackground": COLORS["warning_dark"],
        "activeforeground": COLORS["text_dark"],
        "relief": "flat",
        "borderwidth": 0,
        "padx": 10,
        "pady": 5
    },
    "info": {
        "bg": COLORS["info"],
        "fg": COLORS["text_dark"],
        "font": FONTS["button"],
        "activebackground": COLORS["info_dark"],
        "activeforeground": COLORS["text_dark"],
        "relief": "flat",
        "borderwidth": 0,
        "padx": 10,
        "pady": 5
    },
    "neutral": {
        "bg": COLORS["neutral"],
        "fg": COLORS["text_dark"],
        "font": FONTS["button"],
        "activebackground": COLORS["neutral_dark"],
        "activeforeground": COLORS["text_dark"],
        "relief": "flat",
        "borderwidth": 0,
        "padx": 10,
        "pady": 5
    },
    "highlight": {
        "bg": COLORS["highlight"],
        "fg": COLORS["text_dark"],
        "font": FONTS["button"],
        "activebackground": COLORS["highlight_dark"],
        "activeforeground": COLORS["text_dark"],
        "relief": "flat",
        "borderwidth": 0,
        "padx": 10,
        "pady": 5
    }
}

# Styles pour les frames
FRAME_STYLES = {
    "main": {
        "bg": COLORS["secondary"],
        "padx": 20,
        "pady": 20,
        "relief": "flat"
    },
    "header": {
        "bg": COLORS["primary"],
        "padx": 10,
        "pady": 10,
        "relief": "flat"
    },
    "content": {
        "bg": COLORS["secondary"],
        "padx": 10,
        "pady": 10,
        "relief": "flat"
    },
    "footer": {
        "bg": COLORS["secondary"],
        "padx": 10,
        "pady": 10,
        "relief": "flat"
    }
}

# Styles pour les labels
LABEL_STYLES = {
    "title": {
        "bg": COLORS["primary"],
        "fg": COLORS["text_dark"],
        "font": FONTS["title"],
        "padx": 10,
        "pady": 5
    },
    "subtitle": {
        "bg": COLORS["secondary"],
        "fg": COLORS["primary"],
        "font": FONTS["subtitle"],
        "padx": 5,
        "pady": 5
    },
    "body": {
        "bg": COLORS["secondary"],
        "fg": COLORS["text_dark"],
        "font": FONTS["body"],
        "padx": 5,
        "pady": 2
    },
    "info": {
        "bg": COLORS["secondary"],
        "fg": COLORS["info"],
        "font": FONTS["body_bold"],
        "padx": 5,
        "pady": 2
    },
    "warning": {
        "bg": COLORS["secondary"],
        "fg": COLORS["warning"],
        "font": FONTS["body_bold"],
        "padx": 5,
        "pady": 2
    }
}

# Styles pour les entrées
ENTRY_STYLES = {
    "normal": {
        "font": FONTS["body"],
        "borderwidth": 1,
        "relief": "solid"
    },
    "readonly": {
        "font": FONTS["body"],
        "borderwidth": 1,
        "relief": "solid",
        "state": "readonly"
    }
}

# Icônes communes
ICONS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")

# Fonction pour appliquer un style de bouton
def apply_button_style(button, style_name="primary", width=None, height=None):
    """
    Applique un style prédéfini à un bouton.
    
    Args:
        button: Le widget bouton à styliser
        style_name: Le nom du style à appliquer (primary, secondary, etc.)
        width: Largeur optionnelle du bouton
        height: Hauteur optionnelle du bouton
    """
    style = BUTTON_STYLES.get(style_name, BUTTON_STYLES["primary"])
    for key, value in style.items():
        try:
            button[key] = value
        except:
            pass
    
    if width:
        button["width"] = width
    if height:
        button["height"] = height

# Fonction pour appliquer un style de frame
def apply_frame_style(frame, style_name="main"):
    """
    Applique un style prédéfini à un frame.
    
    Args:
        frame: Le widget frame à styliser
        style_name: Le nom du style à appliquer (main, header, etc.)
    """
    style = FRAME_STYLES.get(style_name, FRAME_STYLES["main"])
    for key, value in style.items():
        try:
            frame[key] = value
        except:
            pass

# Fonction pour appliquer un style de label
def apply_label_style(label, style_name="body"):
    """
    Applique un style prédéfini à un label.
    
    Args:
        label: Le widget label à styliser
        style_name: Le nom du style à appliquer (title, subtitle, etc.)
    """
    style = LABEL_STYLES.get(style_name, LABEL_STYLES["body"])
    for key, value in style.items():
        try:
            label[key] = value
        except:
            pass

# Fonction pour appliquer un style d'entrée
def apply_entry_style(entry, style_name="normal"):
    """
    Applique un style prédéfini à une entrée.
    
    Args:
        entry: Le widget entry à styliser
        style_name: Le nom du style à appliquer (normal, readonly)
    """
    style = ENTRY_STYLES.get(style_name, ENTRY_STYLES["normal"])
    for key, value in style.items():
        try:
            entry[key] = value
        except:
            pass

# Fonction pour créer un bouton stylisé
def create_button(parent, text, command, style="primary", width=None, height=None, **kwargs):
    """
    Crée un bouton avec un style prédéfini.
    
    Args:
        parent: Le widget parent
        text: Le texte du bouton
        command: La fonction à exécuter lors du clic
        style: Le nom du style à appliquer
        width: Largeur optionnelle du bouton
        height: Hauteur optionnelle du bouton
        **kwargs: Arguments supplémentaires pour le bouton
        
    Returns:
        Le bouton créé
    """
    button_style = BUTTON_STYLES.get(style, BUTTON_STYLES["primary"]).copy()
    button_style.update(kwargs)
    
    if width:
        button_style["width"] = width
    if height:
        button_style["height"] = height
        
    button = tk.Button(parent, text=text, command=command, **button_style)
    return button

# Fonction pour créer un label stylisé
def create_label(parent, text, style="body", **kwargs):
    """
    Crée un label avec un style prédéfini.
    
    Args:
        parent: Le widget parent
        text: Le texte du label
        style: Le nom du style à appliquer
        **kwargs: Arguments supplémentaires pour le label
        
    Returns:
        Le label créé
    """
    label_style = LABEL_STYLES.get(style, LABEL_STYLES["body"]).copy()
    label_style.update(kwargs)
    
    label = tk.Label(parent, text=text, **label_style)
    return label

# Fonction pour créer un frame stylisé
def create_frame(parent, style="main", **kwargs):
    """
    Crée un frame avec un style prédéfini.
    
    Args:
        parent: Le widget parent
        style: Le nom du style à appliquer
        **kwargs: Arguments supplémentaires pour le frame
        
    Returns:
        Le frame créé
    """
    frame_style = FRAME_STYLES.get(style, FRAME_STYLES["main"]).copy()
    frame_style.update(kwargs)
    
    frame = tk.Frame(parent, **frame_style)
    return frame

# Fonction pour créer une entrée stylisée
def create_entry(parent, textvariable=None, style="normal", **kwargs):
    """
    Crée une entrée avec un style prédéfini.
    
    Args:
        parent: Le widget parent
        textvariable: La variable tkinter à lier
        style: Le nom du style à appliquer
        **kwargs: Arguments supplémentaires pour l'entrée
        
    Returns:
        L'entrée créée
    """
    entry_style = ENTRY_STYLES.get(style, ENTRY_STYLES["normal"]).copy()
    entry_style.update(kwargs)
    
    entry = tk.Entry(parent, textvariable=textvariable, **entry_style)
    return entry

# Fonction pour configurer le style ttk
def configure_ttk_style():
    """Configure le style ttk pour les widgets ttk."""
    style = ttk.Style()
    
    # Style pour Treeview
    style.configure("Treeview", 
                    background=COLORS["secondary"],
                    foreground=COLORS["text_dark"],
                    rowheight=25,
                    fieldbackground=COLORS["secondary"])
    
    style.configure("Treeview.Heading", 
                    background=COLORS["primary"],
                    foreground=COLORS["text_dark"],
                    font=FONTS["body_bold"])
    
    # Style pour les onglets
    style.configure("TNotebook", 
                    background=COLORS["secondary"])
    
    style.configure("TNotebook.Tab", 
                    background=COLORS["neutral"],
                    foreground=COLORS["text_dark"],
                    padding=[10, 5],
                    font=FONTS["body"])
    
    style.map("TNotebook.Tab",
              background=[("selected", COLORS["primary"])],
              foreground=[("selected", COLORS["text_dark"])])
    
    # Style pour les boutons
    style.configure("TButton", 
                    background=COLORS["primary"],
                    foreground=COLORS["text_dark"],
                    font=FONTS["button"],
                    padding=[10, 5])
    
    style.map("TButton",
              background=[("active", COLORS["primary_dark"])],
              foreground=[("active", COLORS["text_dark"])])
    
    # Style pour les combobox
    style.configure("TCombobox", 
                    background=COLORS["secondary"],
                    foreground=COLORS["text_dark"],
                    font=FONTS["body"])
    
    # Style pour les entrées
    style.configure("TEntry", 
                    background=COLORS["secondary"],
                    foreground=COLORS["text_dark"],
                    font=FONTS["body"])
    
    # Style pour les checkbuttons
    style.configure("TCheckbutton", 
                    background=COLORS["secondary"],
                    foreground=COLORS["text_dark"],
                    font=FONTS["body"])
    
    # Style pour les radiobuttons
    style.configure("TRadiobutton", 
                    background=COLORS["secondary"],
                    foreground=COLORS["text_dark"],
                    font=FONTS["body"])
    
    # Style pour les séparateurs
    style.configure("TSeparator", 
                    background=COLORS["border"])
    
    # Style pour les progressbars
    style.configure("TProgressbar", 
                    background=COLORS["primary"],
                    troughcolor=COLORS["secondary"],
                    borderwidth=0)
    
    return style

# Fonction pour créer un titre de section
def create_section_title(parent, text):
    """
    Crée un titre de section avec une ligne de séparation.
    
    Args:
        parent: Le widget parent
        text: Le texte du titre
        
    Returns:
        Le frame contenant le titre et la ligne
    """
    frame = create_frame(parent)
    label = create_label(frame, text, style="subtitle")
    label.pack(anchor="w", pady=(10, 5))
    
    separator = ttk.Separator(frame, orient="horizontal")
    separator.pack(fill="x", pady=(0, 10))
    
    return frame

# Fonction pour créer un message d'information
def create_info_message(parent, text, style="info"):
    """
    Crée un message d'information stylisé.
    
    Args:
        parent: Le widget parent
        text: Le texte du message
        style: Le style du message (info, warning)
        
    Returns:
        Le label créé
    """
    return create_label(parent, text, style=style)

# Fonction pour créer une boîte de dialogue modale
def create_modal_dialog(parent, title, width=400, height=300):
    """
    Crée une boîte de dialogue modale.
    
    Args:
        parent: La fenêtre parente
        title: Le titre de la boîte de dialogue
        width: La largeur de la boîte de dialogue
        height: La hauteur de la boîte de dialogue
        
    Returns:
        La fenêtre de dialogue
    """
    dialog = tk.Toplevel(parent)
    dialog.title(title)
    dialog.geometry(f"{width}x{height}")
    dialog.transient(parent)
    dialog.grab_set()
    dialog.focus_set()
    
    # Centrer la fenêtre
    dialog.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
    dialog.geometry(f"+{x}+{y}")
    
    # Appliquer le style
    dialog.configure(bg=COLORS["secondary"])
    
    return dialog
