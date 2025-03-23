# === AVEVA Tool Mise à jour avec CTkImage, ToolTip custom et thème blanc modernisé ===
import os
import sys
import clr
import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime
from threading import Thread
from System import String, Boolean, Array
from PIL import Image
from customtkinter import CTkImage
import tkinter as tk

# === ToolTip personnalisé ===
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)

    def show(self, event=None):
        if self.tipwindow or not self.text:
            return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, background="#ffffe0", relief="solid", borderwidth=1, font=("Arial", 9))
        label.pack()

    def hide(self, event=None):
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None

# === CHEMIN DES DLLs ===
if hasattr(sys, '_MEIPASS'):
    base_path = sys._MEIPASS
else:
    base_path = os.getcwd()

dll_path = os.path.join(base_path, "dll")
os.environ["PATH"] += os.pathsep + dll_path
if dll_path not in sys.path:
    sys.path.append(dll_path)

# === CONFIGURATION DE L'ICÔNE ===
try:
    icon_path = os.path.join(base_path, "icon.ico")
    if os.path.exists(icon_path):
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("aveva.export.tool")
except:
    pass

# === INITIALISATION UI ===
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("AVEVA Export & Import XML Tool")
app.geometry("800x700")  # Moins haut
app.configure(fg_color="#f2f4f8")  # Fond moderne clair
if os.path.exists(icon_path):
    app.iconbitmap(icon_path)

# === CENTRER FENÊTRE ===
app.update_idletasks()
width = 800
height = 700
x = (app.winfo_screenwidth() // 2) - (width // 2)
y = (app.winfo_screenheight() // 2) - (height // 2)
app.geometry(f"{width}x{height}+{x}+{y}")

# === LOGO AVEVA ===
try:
    aveva_logo_path = os.path.join(base_path, "aveva.png")
    aveva_logo = CTkImage(Image.open(aveva_logo_path).convert("RGBA"), size=(200, 50))
    logo_label = ctk.CTkLabel(app, image=aveva_logo, text="")
    logo_label.pack(pady=(5, 0))  # Moins d’espace
except Exception as e:
    print(f"Erreur chargement logo AVEVA : {e}")




def show_conditions():
    msg = """
--- CONDITIONS GÉNÉRALES D'UTILISATION ---

Ce logiciel a été conçu et réalisé par Project Collaborative Development EIRL.
Il est distribué librement via la plateforme GitHub.

Licence : MIT (Massachusetts Institute of Technology License)
- Vous êtes libre d'utiliser, copier, modifier et redistribuer le code.
- Ce logiciel est fourni "EN L'ÉTAT", sans aucune garantie expresse ou implicite.
- En aucun cas les auteurs ou les détenteurs de droits ne pourront être tenus responsables de tout dommage ou réclamation.

Ce programme est destiné à faciliter l'export et l'import d'objets ou de graphiques dans un environnement AVEVA.

Compatibilité actuelle : AVEVA System Platform 2023 R2

En utilisant ce logiciel, vous acceptez ces conditions.
"""
    top = tk.Toplevel()
    top.title("Conditions générales d'utilisation")
    top.geometry("600x400")
    top.resizable(False, False)
    if os.path.exists(icon_path):
        top.iconbitmap(icon_path)
    text = tk.Text(top, wrap="word", font=("Segoe UI", 10), padx=10, pady=10)
    text.insert("1.0", msg)
    text.config(state="disabled")
    text.pack(expand=True, fill="both", padx=10, pady=10)
    ctk.CTkButton(top, text="Fermer", command=top.destroy).pack(pady=(0, 10))

conditions_button = ctk.CTkButton(app, text="Conditions générales", command=show_conditions, width=160)
conditions_button.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-5)



# === VARIABLES ===
selected_galaxy = ctk.StringVar(value="Chargement...")
username = ctk.StringVar(value="admin")
password = ctk.StringVar()
object_name = ctk.StringVar()
export_folder = ctk.StringVar()
export_type = ctk.StringVar(value="Graphic")
import_graphic_name = ctk.StringVar()
import_xml_path = ctk.StringVar()

galaxy_list = []

# === STYLED FRAME ACCORDION ===
accordion = ctk.CTkTabview(app, segmented_button_fg_color="#dee2e6", segmented_button_selected_color="#0d6efd", segmented_button_selected_hover_color="#0b5ed7", segmented_button_unselected_color="#ffffff")
accordion.pack(fill="both", expand=True, padx=20, pady=10)
accordion.add("Export")
accordion.add("Import")

# === ICONES ===
def load_icon(name):
    path = os.path.join(base_path, name)
    if os.path.exists(path):
        return CTkImage(Image.open(path).convert("RGBA"), size=(28, 28))
    return None

export_icon = load_icon("export.png")
import_icon = load_icon("import.png")

# === EXPORT UI ===
export_tab = accordion.tab("Export")
ctk.CTkLabel(export_tab, text="Type d'export :").grid(row=0, column=0, sticky="w")
export_menu = ctk.CTkOptionMenu(export_tab, variable=export_type, values=["Graphic", "Instance", "Template"])
export_menu.grid(row=0, column=1)
ctk.CTkLabel(export_tab, text="Galaxy :").grid(row=1, column=0, sticky="w")
galaxy_combo = ctk.CTkComboBox(export_tab, variable=selected_galaxy, values=["Chargement..."])
galaxy_combo.grid(row=1, column=1)
ctk.CTkLabel(export_tab, text="Utilisateur :").grid(row=2, column=0, sticky="w")
ctk.CTkEntry(export_tab, textvariable=username).grid(row=2, column=1)
ctk.CTkLabel(export_tab, text="Mot de passe :").grid(row=3, column=0, sticky="w")
ctk.CTkEntry(export_tab, textvariable=password, show="*").grid(row=3, column=1)
ctk.CTkLabel(export_tab, text="Nom de l'objet ou fichier .txt :").grid(row=4, column=0, sticky="w")
ctk.CTkEntry(export_tab, textvariable=object_name).grid(row=4, column=1)
select_file_btn = ctk.CTkButton(export_tab, text="...", width=30, command=lambda: select_txt_file(), hover=True)
select_file_btn.grid(row=4, column=2)
ToolTip(select_file_btn, "Parcourir fichier .txt")
ctk.CTkLabel(export_tab, text="Dossier d'export :").grid(row=5, column=0, sticky="w")
ctk.CTkEntry(export_tab, textvariable=export_folder, state="readonly").grid(row=5, column=1)
select_folder_btn = ctk.CTkButton(export_tab, text="...", width=30, command=lambda: select_folder(), hover=True)
select_folder_btn.grid(row=5, column=2)
ToolTip(select_folder_btn, "Parcourir dossier")
export_button = ctk.CTkButton(export_tab, image=export_icon, text="Exporter", compound="left", font=("Segoe UI", 13, "bold"), width=140, height=48, command=lambda: threaded(run_export))
export_button.grid(row=6, column=1, pady=10)
ToolTip(export_button, "Exporter les objets ou graphiques")

# === IMPORT UI ===
import_tab = accordion.tab("Import")
ctk.CTkLabel(import_tab, text="Nom du graphique :").grid(row=0, column=0, sticky="w")
ctk.CTkEntry(import_tab, textvariable=import_graphic_name).grid(row=0, column=1)
ctk.CTkLabel(import_tab, text="Fichier XML :").grid(row=1, column=0, sticky="w")
ctk.CTkEntry(import_tab, textvariable=import_xml_path, state="readonly").grid(row=1, column=1)
select_xml_btn = ctk.CTkButton(import_tab, text="...", width=30, command=lambda: select_xml_file(), hover=True)
select_xml_btn.grid(row=1, column=2)
ToolTip(select_xml_btn, "Parcourir fichier XML")
import_button = ctk.CTkButton(import_tab, image=import_icon, text="Importer", compound="left", font=("Segoe UI", 13, "bold"), width=140, height=48, command=lambda: threaded(run_import_xml))
import_button.grid(row=2, column=1, pady=10)
ToolTip(import_button, "Importer un graphique XML")

# === STATUS & FOOTER ===
def update_status(msg):
    status_label.configure(text=f"  État : {msg}")
    app.update_idletasks()

status_label = ctk.CTkLabel(app, text="  Lancement...", anchor="w", padx=10)
status_label.pack(fill="x", pady=(0, 5))
progress = ctk.CTkProgressBar(app, mode="indeterminate")
progress.pack(fill="x", pady=(0, 5))
footer = ctk.CTkLabel(app, text="Project Collaborative Development EIRL", font=("Segoe UI", 10, "italic"), anchor="center", text_color="#6c757d")
footer.pack(pady=(0, 2))



# === OUTILS UI ===
def update_status(msg):
    status_label.configure(text=f"État : {msg}")
    app.update_idletasks()

def start_loading():
    progress.start()
    update_status("En cours...")

def stop_loading():
    progress.stop()
    update_status("Prêt")

def threaded(func):
    Thread(target=lambda: [start_loading(), func(), stop_loading()]).start()

# === FONCTIONS DE FICHIERS ===
def select_txt_file():
    result = os.popen(
        'powershell -Command "[System.Reflection.Assembly]::LoadWithPartialName(\'System.Windows.Forms\') | Out-Null; '
        '$ofd = New-Object System.Windows.Forms.OpenFileDialog; '
        '$ofd.Filter = \'Text Files (*.txt)|*.txt\'; '
        '$ofd.Title = \'Choisir un fichier .txt\'; '
        'if ($ofd.ShowDialog() -eq \'OK\') { $ofd.FileName }"'
    ).read().strip()
    if result and result.lower().endswith(".txt"):
        object_name.set(result)

def select_folder():
    result = os.popen('powershell "(new-object -COM \'Shell.Application\').BrowseForFolder(0,\'Choisir un dossier\',0,0).self.path"').read().strip()
    if result:
        export_folder.set(result)

def select_xml_file():
    result = os.popen(
        'powershell -Command "[System.Reflection.Assembly]::LoadWithPartialName(\'System.Windows.Forms\') | Out-Null; '
        '$ofd = New-Object System.Windows.Forms.OpenFileDialog; '
        '$ofd.Filter = \'XML Files (*.xml)|*.xml\'; '
        '$ofd.Title = \'Choisir un fichier graphique .xml\'; '
        'if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $ofd.FileName }"'
    ).read().strip()
    if result and result.lower().endswith(".xml"):
        import_xml_path.set(result)

def write_log(entry):
    with open("aveva_export_log.txt", "a", encoding="utf-8") as log_file:
        timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
        log_file.write(f"{timestamp} {entry}\n")

# === INITIALISATION BACKEND EN THREAD ===
def initialize_backend():
    try:
        update_status("Chargement des DLLs...")
        clr.AddReference(os.path.join(dll_path, "ArchestrA.Visualization.GraphicAccess.dll"))
        global GRAccessApp, EgObjectIsTemplateOrInstance, EExportType, GraphicAccess
        from ArchestrA.GRAccess import GRAccessApp, EgObjectIsTemplateOrInstance, EExportType
        from ArchestrA.Visualization.GraphicAccess import GraphicAccess

        update_status("Connexion au serveur GRAccess...")
        global gr_access, galaxies, galaxy_list
        gr_access = GRAccessApp()
        galaxies = gr_access.QueryGalaxiesEx(String("localhost"))
        galaxy_list = [galaxies[i + 1].Name for i in range(galaxies.count)] if galaxies.count > 0 else []

        selected_galaxy.set(galaxy_list[0] if galaxy_list else "")
        galaxy_combo.configure(values=galaxy_list)
        update_status("Prêt")
    except Exception as e:
        update_status("Erreur de chargement.")
        messagebox.showerror("Erreur DLL", f"Erreur lors du chargement : {e}")
        sys.exit(1)

# === EXPORT ===
def export_single(galaxy, obj_name, export_type_str, folder):
    try:
        export_path = os.path.join(folder, f"{obj_name}.xml" if export_type_str == "Graphic" else f"{obj_name}.aaPKG")
        update_status(f"Export de {obj_name}...")
        if export_type_str == "Graphic":
            graphic_access = GraphicAccess()
            result = graphic_access.ExportGraphicToXml(galaxy, String(obj_name), String(export_path), Boolean(True))
            success = result.Successful
        else:
            obj_type = EgObjectIsTemplateOrInstance.gObjectIsInstance if export_type_str == "Instance" else EgObjectIsTemplateOrInstance.gObjectIsTemplate
            obj_array = Array[String]([String(obj_name)])
            objects, result = galaxy.QueryObjectsByName(obj_type, obj_array)
            if objects.count == 0:
                write_log(f"❌ Aucun objet trouvé : {obj_name}")
                return False
            objects.ExportObjects(EExportType.exportAsPDF, String(export_path))
            success = os.path.exists(export_path)
        if success:
            write_log(f"✅ Export réussi : {export_path}")
            return True
        else:
            write_log(f"❌ Échec export : {obj_name}")
            return False
    except Exception as e:
        write_log(f"❌ Exception pour {obj_name} : {e}")
        return False

def run_export():
    update_status("Connexion à la Galaxy...")
    galaxy_name = selected_galaxy.get()
    user = username.get()
    pwd = password.get()
    input_value = object_name.get().strip()
    folder = export_folder.get()
    export_type_str = export_type.get()
    batch_mode = use_txt_file.get()

    if not input_value or not folder:
        messagebox.showerror("Erreur", "Veuillez saisir un objet ou fichier .txt et un dossier d'export.")
        return
    try:
        galaxy = galaxies[String(galaxy_name)]
        galaxy.Login(String(user), String(pwd))
        if not galaxy.CommandResult.Successful:
            messagebox.showerror("Connexion échouée", galaxy.CommandResult.Text)
            return
        if batch_mode:
            if not os.path.exists(input_value):
                messagebox.showerror("Erreur", "Fichier .txt introuvable.")
                return
            with open(input_value, "r", encoding="utf-8") as f:
                lines = [line.strip() for line in f if line.strip()]
            total = len(lines)
            success_count = 0
            for obj_name in lines:
                if export_single(galaxy, obj_name, export_type_str, folder):
                    success_count += 1
            messagebox.showinfo("Terminé", f"Export batch terminé : {success_count}/{total} objets exportés.\nVoir le log.")
        else:
            if export_single(galaxy, input_value, export_type_str, folder):
                messagebox.showinfo("Succès", f"{export_type_str} exporté avec succès.")
            else:
                messagebox.showerror("Erreur", f"Échec de l'export de : {input_value}")
    except Exception as e:
        write_log(f"❌ Exception globale : {e}")
        messagebox.showerror("Erreur", str(e))

# === IMPORT XML ===
def run_import_xml():
    update_status("Connexion à la Galaxy pour import...")
    galaxy_name = selected_galaxy.get()
    user = username.get()
    pwd = password.get()
    graphic_name = import_graphic_name.get().strip()
    xml_path = import_xml_path.get().strip()

    if not graphic_name or not xml_path:
        messagebox.showerror("Erreur", "Veuillez saisir un nom de graphique et sélectionner un fichier XML.")
        return
    try:
        galaxy = galaxies[String(galaxy_name)]
        galaxy.Login(String(user), String(pwd))
        if not galaxy.CommandResult.Successful:
            messagebox.showerror("Connexion échouée", galaxy.CommandResult.Text)
            return
        update_status("Import du graphique en cours...")
        graphic_access = GraphicAccess()
        result = graphic_access.ImportGraphicFromXml(galaxy, String(graphic_name), String(xml_path), Boolean(True))
        if result.Successful:
            write_log(f"✅ Import XML réussi : {graphic_name} depuis {xml_path}")
            messagebox.showinfo("Succès", f"Import XML réussi : {graphic_name}")
        else:
            write_log(f"❌ Échec import XML : {graphic_name} → {result.Text}")
            messagebox.showerror("Erreur", f"Échec de l'import : {result.Text}")
    except Exception as e:
        write_log(f"❌ Exception import XML : {e}")
        messagebox.showerror("Exception", str(e))

# === DÉMARRAGE ===
threaded(initialize_backend)
app.mainloop()
