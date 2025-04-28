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
import tkinter.messagebox as messagebox
import tkinter.filedialog as filedialog
import platform
import subprocess
import shlex
import concurrent.futures

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

# === DÉTERMINER BASE PATH ===
if hasattr(sys, '_MEIPASS'):
    base_path = sys._MEIPASS
else:
    base_path = os.getcwd()

if base_path not in sys.path:
    sys.path.append(base_path)

# === GESTION DU DOSSIER PERSISTANT ===
def get_persistent_storage_path():      
    if platform.system() == "Windows":
        return os.path.join(os.environ["LOCALAPPDATA"], "AvevaTool")
    else:
        return os.path.expanduser("~/.avevatool")

persistent_path = get_persistent_storage_path()
os.makedirs(persistent_path, exist_ok=True)

memory_file = os.path.join(persistent_path, "dll_paths.txt")

def load_custom_dll_paths():
    paths = []

    # Chemins définis par l'utilisateur
    if os.path.exists(memory_file):
        with open(memory_file, "r", encoding="utf-8") as f:
            paths += [line.strip() for line in f if os.path.isdir(line.strip())]

    # Recherche automatique dans les répertoires standards AVEVA
    default_dirs = [
        r"C:\Program Files (x86)\Common Files\ArchestrA",
        r"C:\Program Files (x86)\ArchestrA\Framework\Bin",
    ]

    for root_dir in default_dirs:
        for root, dirs, files in os.walk(root_dir):
            if "ArchestrA.Visualization.GraphicAccess.dll" in files:
                # Évite doublon
                if root not in paths:
                    paths.append(root)
                break  # On garde le premier trouvé

    return paths


def save_custom_dll_paths(paths):
    with open(memory_file, "w", encoding="utf-8") as f:
        for path in paths:
            f.write(path + "\n")

def reset_custom_dll_paths():
    if os.path.exists(memory_file):
        os.remove(memory_file)
    messagebox.showinfo("DLL Reset", "Les chemins DLL personnalisés ont été réinitialisés. L'application va redémarrer.")
    app.destroy()
    subprocess.Popen([sys.executable] + sys.argv, shell=True)
    sys.exit()

def update_export_mode():
    if export_type.get() == "Graphic":
        # Si Graphic => forcer 'Un fichier par objet' et désactiver
        export_mode.set("Un fichier par objet")
        export_mode_menu.configure(state="disabled")
    else:
        # Sinon, permettre de choisir
        export_mode_menu.configure(state="normal")


# === CHARGEMENT DES DLLs ===

custom_dll_paths = load_custom_dll_paths()

if not custom_dll_paths:
    print("Aucun dossier DLL personnalisé enregistré. Aucun chargement effectué.")
else:
    dll_loaded = False
    for path in custom_dll_paths:
        try:
            dll_file = os.path.join(path, "ArchestrA.Visualization.GraphicAccess.dll")
            if os.path.isfile(dll_file):
                clr.AddReference(dll_file)
                # Ajouter le chemin pour les importations suivantes
                if path not in sys.path:
                    sys.path.append(path)
                dll_loaded = True
                print(f"DLL chargée depuis : {dll_file}")
                break
        except Exception as e:
            print(f"Échec chargement DLL depuis {path} : {e}")

    if not dll_loaded:
        messagebox.showerror("Erreur DLL", "Impossible de charger 'ArchestrA.Visualization.GraphicAccess.dll' depuis les dossiers enregistrés.")
        sys.exit(1)


# === CONFIGURATION DE L'ICÔNE ===
try:
    icon_path = os.path.join(base_path, "icon.ico")
    if os.path.exists(icon_path):
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("aveva.export.tool")
        icon_available = True
    else:
        icon_available = False
except:
    icon_available = False

# === INITIALISATION UI ===
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
if icon_available:
    try:
        app.iconbitmap(icon_path)
    except:
        pass
top_button_frame = ctk.CTkFrame(app, fg_color="transparent")
top_button_frame.pack(fill="x", padx=10, pady=(10, 0))

app.title("AVEVA Export & Import XML Tool")
app.geometry("700x450")  # Taille étendue pour une meilleure disposition  # Fenêtre plus compacte
app.configure(fg_color="#f2f4f8")  # Fond moderne clair

# === CENTRER LA FENÊTRE ===
app.update_idletasks()
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
x = (screen_width - 700) // 2
y = (screen_height - 450) // 2
app.geometry(f"700x450+{x}+{y}")



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


# === VARIABLES ===
selected_galaxy = ctk.StringVar(value="Chargement...")
username = ctk.StringVar(value="")
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
accordion.add("Galaxy")
accordion.add("Export")
accordion.add("Import")
accordion.add("Suppression")

# === GALAXY UI ===
export_tab = accordion.tab("Galaxy")
ctk.CTkLabel(export_tab, text="Galaxy :").grid(row=1, column=0, sticky="w")
galaxy_combo = ctk.CTkComboBox(export_tab, variable=selected_galaxy, values=["Chargement..."])
galaxy_combo.grid(row=1, column=1)
ctk.CTkLabel(export_tab, text="Utilisateur :").grid(row=2, column=0, sticky="w")
ctk.CTkEntry(export_tab, textvariable=username).grid(row=2, column=1)
ctk.CTkLabel(export_tab, text="Mot de passe :").grid(row=3, column=0, sticky="w")
ctk.CTkEntry(export_tab, textvariable=password, show="*").grid(row=3, column=1)

# === EXPORT UI ===
export_tab = accordion.tab("Export")

# Type d'export
ctk.CTkLabel(export_tab, text="Type d'export :").grid(row=0, column=0, sticky="w")
export_type = ctk.StringVar(value="Graphic")  # Déclare la variable pour le type d'export
export_menu = ctk.CTkOptionMenu(export_tab, variable=export_type, values=["Graphic", "Instance", "Template"])
export_menu.grid(row=0, column=1)

# Mode d'export
ctk.CTkLabel(export_tab, text="Mode d'export :").grid(row=1, column=0, sticky="w")
export_mode = ctk.StringVar(value="Un fichier par objet")  # Déclare la variable pour le mode d'export
export_mode_menu = ctk.CTkOptionMenu(export_tab, variable=export_mode, values=["Un fichier par objet", "Un seul fichier"])
export_mode_menu.grid(row=1, column=1)

# Nom de l'objet ou fichier .txt
ctk.CTkLabel(export_tab, text="Nom de l'objet ou fichier .txt :").grid(row=4, column=0, sticky="w")
ctk.CTkEntry(export_tab, textvariable=object_name).grid(row=4, column=1)

# Sélecteur de fichier .txt
select_file_btn = ctk.CTkButton(export_tab, text="...", width=30, command=lambda: select_txt_file(), hover=True)
select_file_btn.grid(row=4, column=2)
ToolTip(select_file_btn, "Parcourir fichier .txt")

# Dossier d'export
ctk.CTkLabel(export_tab, text="Dossier d'export :").grid(row=5, column=0, sticky="w")
ctk.CTkEntry(export_tab, textvariable=export_folder, state="readonly").grid(row=5, column=1)

# Sélecteur de dossier
select_folder_btn = ctk.CTkButton(export_tab, text="...", width=30, command=lambda: select_folder(), hover=True)
select_folder_btn.grid(row=5, column=2)
ToolTip(select_folder_btn, "Parcourir dossier")

# Bouton Export
export_button = ctk.CTkButton(export_tab, text="Exporter", compound="left", font=("Segoe UI", 13, "bold"), width=140, height=48, command=lambda: threaded(run_export))
export_button.grid(row=6, column=1, pady=10)
ToolTip(export_button, "Exporter les objets ou graphiques")


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

def manage_dll_paths():
    new_path = filedialog.askdirectory(title="Ajouter un dossier de DLL")
    if new_path:
        if new_path not in custom_dll_paths:
            custom_dll_paths.append(new_path)
            save_custom_dll_paths(custom_dll_paths)
            messagebox.showinfo("Ajouté", "Chemin ajouté. L'application va se fermer, veuillez la relancer manuellement.")
            app.quit()
            sys.exit()
        else:
            messagebox.showinfo("Info", "Ce chemin est déjà présent.")
      


add_dll_button = ctk.CTkButton(top_button_frame, text="Ajouter DLL", width=100, command=manage_dll_paths)
add_dll_button.pack(side="left", padx=(0, 10))

reset_dll_button = ctk.CTkButton(top_button_frame, text="Reset DLL", width=100, command=reset_custom_dll_paths)
reset_dll_button.pack(side="left", padx=(0, 10))

conditions_button = ctk.CTkButton(top_button_frame, text="Conditions générales", width=160, command=show_conditions)
conditions_button.pack(side="right")

def initialize_backend():
    try:
        update_status("Chargement des DLLs...")

        # Charger les chemins personnalisés
        custom_dll_paths = load_custom_dll_paths()
        dll_loaded = False
        last_exception = None

        for path in custom_dll_paths:
            try:
                dll_file = os.path.join(path, "ArchestrA.Visualization.GraphicAccess.dll")
                if os.path.isfile(dll_file):
                    clr.AddReference(dll_file)
                    if path not in sys.path:
                        sys.path.append(path)
                    dll_loaded = True
                    break
            except Exception as e:
                last_exception = e

        if not dll_loaded:
            raise Exception(f"Aucune DLL valide trouvée. Dernière erreur : {last_exception}")

        # Imports après chargement réussi
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
        retry = messagebox.askyesno(
            "Erreur DLL",
            f"Erreur lors du chargement :\n{e}\n\nVoulez-vous ajouter un dossier contenant les DLL ?"
        )
        if retry:
            manage_dll_paths()  # Cette fonction doit ouvrir un file dialog + enregistrer le dossier
        else:
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
    success_count = 0
    total = 0
    update_status("Connexion à la Galaxy...")
    galaxy_name = selected_galaxy.get()
    user = username.get()
    pwd = password.get()
    input_value = object_name.get().strip()
    folder = export_folder.get()
    export_type_str = export_type.get()
    export_mode_str = export_mode.get()  # Récupérer le mode d'export
    batch_mode = os.path.isfile(input_value) and input_value.lower().endswith(".txt")

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
            
            if export_mode_str == "Un seul fichier" and export_type_str in ("Instance", "Template"):
                # === Exporter tous les objets ensemble dans UN fichier ===
                update_status(f"Export groupé de {len(lines)} objets...")
                
                export_path = os.path.join(folder, f"{os.path.splitext(os.path.basename(input_value))[0]}_grouped.aaPKG")
                
                obj_type = EgObjectIsTemplateOrInstance.gObjectIsInstance if export_type_str == "Instance" else EgObjectIsTemplateOrInstance.gObjectIsTemplate
                obj_array = Array[String]([String(name) for name in lines])
                
                objects, result = galaxy.QueryObjectsByName(obj_type, obj_array)
                if objects.count == 0:
                    write_log("❌ Aucun objet trouvé pour export groupé.")
                    messagebox.showerror("Erreur", "Aucun objet trouvé pour l'export groupé.")
                    return
                
                # L'export se fait avec succès, alors on met à jour le compteur
                objects.ExportObjects(EExportType.exportAsPDF, String(export_path))  # Peut-être exportAsPKG à confirmer
                if os.path.exists(export_path):
                    write_log(f"✅ Export groupé réussi : {export_path}")
                    success_count = len(lines)  # Le nombre d'objets exportés est égal à la longueur de `lines`
                    total = len(lines)  # Le total d'objets à exporter est également le même
                else:
                    write_log(f"❌ Échec export groupé.")
                    messagebox.showerror("Erreur", "Échec de l'export groupé.")
            else:
                # === Export individuel, comme actuellement ===
                update_status(f"Export de {len(lines)} objets en parallèle (max 2)...")
                with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
                    futures = [executor.submit(export_single, galaxy, obj_name, export_type_str, folder) for obj_name in lines]
                    total = len(futures)
                    success_count = 0
                    for future in concurrent.futures.as_completed(futures):
                        if future.result():
                            success_count += 1

            # Affichage de la popup une seule fois après l'export batch
            messagebox.showinfo("Terminé", f"Export batch terminé : {success_count}/{total} objets exportés.\nVoir le log.")

        else:
            # Export pour un seul objet
            if export_single(galaxy, input_value, export_type_str, folder):
                messagebox.showinfo("Succès", f"{export_type_str} exporté avec succès.")
            else:
                messagebox.showerror("Erreur", f"Échec de l'export de : {input_value}")
    except Exception as e:
        write_log(f"❌ Exception globale : {e}")
        messagebox.showerror("Erreur", str(e))


def run_import_xml():
    update_status("Connexion à la Galaxy pour import...")
    galaxy_name = selected_galaxy.get()
    user = username.get()
    pwd = password.get()
    xml_path = import_xml_path.get().strip()

    if not xml_path:
        messagebox.showerror("Erreur", "Veuillez sélectionner un fichier XML.")
        return

    graphic_name = os.path.splitext(os.path.basename(xml_path))[0]

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

def select_txt_file_delete():
    result = os.popen(
        'powershell -Command "[System.Reflection.Assembly]::LoadWithPartialName(\'System.Windows.Forms\') | Out-Null; '
        '$ofd = New-Object System.Windows.Forms.OpenFileDialog; '
        '$ofd.Filter = \'Text Files (*.txt)|*.txt\'; '
        '$ofd.Title = \'Choisir un fichier .txt\'; '
        'if ($ofd.ShowDialog() -eq \'OK\') { $ofd.FileName }"'
    ).read().strip()

    if result and result.lower().endswith(".txt"):
        delete_object_name.set(result)


def run_delete():
    update_status("Connexion à la Galaxy pour suppression...")
    galaxy_name = selected_galaxy.get()
    user = username.get()
    pwd = password.get()
    input_value = delete_object_name.get().strip()
    delete_type_str = delete_type.get()
    batch_mode = os.path.isfile(input_value) and input_value.lower().endswith(".txt")

    if not input_value:
        messagebox.showerror("Erreur", "Veuillez saisir un objet ou fichier .txt.")
        return
    try:
        galaxy = galaxies[String(galaxy_name)]
        galaxy.Login(String(user), String(pwd))
        if not galaxy.CommandResult.Successful:
            messagebox.showerror("Connexion échouée", galaxy.CommandResult.Text)
            return

        obj_type = EgObjectIsTemplateOrInstance.gObjectIsInstance if delete_type_str == "Instance" else EgObjectIsTemplateOrInstance.gObjectIsTemplate

        if batch_mode:
            with open(input_value, "r", encoding="utf-8") as f:
                lines = [line.strip() for line in f if line.strip()]
            total = len(lines)
            success_count = 0
            for obj_name in lines:
                objects, result = galaxy.QueryObjectsByName(obj_type, Array[String]([String(obj_name)]))
                if objects.count > 0:
                    objects.DeleteAllObjects()
                    success_count += 1
                    write_log(f"✅ Objet supprimé : {obj_name}")
                else:
                    write_log(f"❌ Objet introuvable : {obj_name}")
            messagebox.showinfo("Terminé", f"Suppression batch terminée : {success_count}/{total} objets supprimés.\nVoir le log.")
        else:
            objects, result = galaxy.QueryObjectsByName(obj_type, Array[String]([String(input_value)]))
            if objects.count > 0:
                objects.DeleteAllObjects()
                write_log(f"✅ Objet supprimé : {input_value}")
                messagebox.showinfo("Succès", f"Objet supprimé : {input_value}")
            else:
                write_log(f"❌ Objet introuvable : {input_value}")
                messagebox.showerror("Erreur", f"Objet introuvable : {input_value}")
    except Exception as e:
        write_log(f"❌ Exception suppression : {e}")
        messagebox.showerror("Exception", str(e))

# === DÉMARRAGE ===
threaded(initialize_backend)
app.mainloop()
