import tkinter as tk
from tkinter import ttk, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import xml.etree.ElementTree as ET
import subprocess
import os
import sys
from PIL import Image, ImageTk  # Für Bildverarbeitung
import locale

# Zusätzliche Importe für die Erstellung einer Desktop-Verknüpfung
try:
    import winshell
    import win32com.client
except ImportError:
    # Wenn winshell und pywin32 nicht installiert sind, den Benutzer informieren
    messagebox.showerror("Module Error", "Bitte installieren Sie die Module 'winshell' und 'pywin32':\npip install winshell pywin32")
    sys.exit(1)

class Application(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        # Internationalisierung einrichten
        self.translations = {}
        self.current_language = 'en_US'  # Standard-Sprache

        self.load_config()  # Laden der gespeicherten Sprache aus config.xml
        self.load_translation(self.current_language)

        self.title(self.translations.get("title", "Conversion Program"))
        self.geometry("550x600")  # Angepasste Fenstergröße, um das Sprach-Dropdown aufzunehmen
        self.iconbitmap(r'resources\icon.ico')
        self.resizable(False, False)  # Fenstergröße kann nicht mehr geändert werden

        # Variablen
        self.path1_var = tk.StringVar()
        self.path2_var = tk.StringVar()
        self.file_list = []  # Liste der Dateien
        self.selected_image = None  # Aktuell ausgewähltes Bild
        self.compiler_var = tk.StringVar(value="mesh")  # Variable für die Compiler-Auswahl
        self.option_var = tk.StringVar(value="none")  # Variable für die zusätzlichen Optionen

        # Widgets erstellen
        self.create_widgets()

        # Pfade aus XML laden, falls vorhanden
        self.load_from_xml()

    def load_translation(self, language_code):
        translation_file = os.path.join('languages', f'{language_code}.xml')
        if not os.path.exists(translation_file):
            if language_code != 'en_US':
                messagebox.showerror("Error", self.translations.get("translation_error", "Automatic translation is not available. The GUI will be in English."))
            language_code = 'en_US'  # Fallback zu Englisch
            translation_file = os.path.join('languages', f'{language_code}.xml')
            if not os.path.exists(translation_file):
                messagebox.showerror("Error", "Default language file 'en_US.xml' is missing.")
                sys.exit(1)

        try:
            tree = ET.parse(translation_file)
            root = tree.getroot()
            self.translations = {}
            for text in root.findall('text'):
                key = text.get('key')
                value = text.text
                if key:
                    self.translations[key] = value
            self.current_language = language_code
        except ET.ParseError:
            messagebox.showerror("Error", f"Failed to parse translation file '{translation_file}'. Falling back to English.")
            if language_code != 'en_US':
                self.load_translation('en_US')
            else:
                sys.exit(1)

    def get_available_languages(self):
        languages_dir = 'languages'
        languages = []
        if not os.path.exists(languages_dir):
            messagebox.showerror("Error", "Languages directory 'languages/' not found.")
            return languages
        for file in os.listdir(languages_dir):
            if file.endswith('.xml'):
                language_code = file[:-4]  # Entfernt '.xml'
                translation = self.load_single_translation(language_code)
                if translation and 'language_name' in translation:
                    language_name = translation['language_name']
                else:
                    # Fallback zu Sprachcode
                    language_name = language_code
                languages.append((language_code, language_name))
        return languages

    def load_single_translation(self, language_code):
        translation_file = os.path.join('languages', f'{language_code}.xml')
        if not os.path.exists(translation_file):
            return None
        try:
            tree = ET.parse(translation_file)
            root = tree.getroot()
            translation = {}
            for text in root.findall('text'):
                key = text.get('key')
                value = text.text
                if key:
                    translation[key] = value
            return translation
        except ET.ParseError:
            return None

    def create_widgets(self):
        # Language Selection Dropdown
        languages = self.get_available_languages()
        if not languages:
            languages = [('en_US', 'English')]

        self.language_frame = ttk.LabelFrame(self, text="Language")
        self.language_frame.pack(pady=10, padx=10, fill='x')

        self.language_var = tk.StringVar()
        language_codes = [lang[0] for lang in languages]
        language_names = [lang[1] for lang in languages]

        self.language_dropdown = ttk.Combobox(self.language_frame, textvariable=self.language_var, values=language_names, state='readonly')
        self.language_dropdown.pack(pady=5, padx=5, fill='x')

        # Set the current language in the dropdown
        if self.translations.get("language_name"):
            current_language_name = self.translations.get("language_name")
            if current_language_name in language_names:
                self.language_dropdown.current(language_names.index(current_language_name))
            else:
                self.language_dropdown.set(language_names[0])
        else:
            self.language_dropdown.set(language_names[0])

        self.language_dropdown.bind("<<ComboboxSelected>>", self.change_language)

        # Entry for Path1
        self.path1_label = ttk.Label(self, text=self.translations.get("path1_label", "Path1:"))
        self.path1_label.pack(pady=2, anchor='w', padx=10)
        self.path1_entry = ttk.Entry(self, textvariable=self.path1_var)
        self.path1_entry.pack(pady=2, padx=10, fill='x')

        # Entry for Path2
        self.path2_label = ttk.Label(self, text=self.translations.get("path2_label", "Path2:"))
        self.path2_label.pack(pady=2, anchor='w', padx=10)
        self.path2_entry = ttk.Entry(self, textvariable=self.path2_var)
        self.path2_entry.pack(pady=2, padx=10, fill='x')

        # Compiler Selection
        self.compiler_frame = ttk.LabelFrame(self, text=self.translations.get("compiler_selection", "Compiler Selection"))
        self.compiler_frame.pack(pady=10, padx=10, fill='x')

        self.mesh_radio = ttk.Radiobutton(
            self.compiler_frame, text=self.translations.get("mesh_compiler", "Mesh Compiler"), variable=self.compiler_var, value="mesh", command=self.update_compiler_selection
        )
        self.mesh_radio.grid(row=0, column=0, padx=5, pady=2, sticky='w')

        self.texture_radio = ttk.Radiobutton(
            self.compiler_frame, text=self.translations.get("texture_compiler", "Texture Compiler"), variable=self.compiler_var, value="texture", command=self.update_compiler_selection
        )
        self.texture_radio.grid(row=0, column=1, padx=5, pady=2, sticky='w')

        # Options
        self.options_frame = ttk.LabelFrame(self, text=self.translations.get("options", "Options"))
        self.options_frame.pack(pady=10, padx=10, fill='x')

        self.none_radio = ttk.Radiobutton(
            self.options_frame, text=self.translations.get("none", "None"), variable=self.option_var, value="none"
        )
        self.none_radio.grid(row=0, column=0, padx=5, pady=2, sticky='w')

        self.physics_mesh_radio = ttk.Radiobutton(
            self.options_frame, text=self.translations.get("enable_physics_mesh", "Enable Physics Mesh"), variable=self.option_var, value="physics_mesh"
        )
        self.physics_mesh_radio.grid(row=0, column=1, padx=5, pady=2, sticky='w')

        self.physics_object_radio = ttk.Radiobutton(
            self.options_frame, text=self.translations.get("enable_physics_object", "Enable Physics Object"), variable=self.option_var, value="physics_object"
        )
        self.physics_object_radio.grid(row=1, column=1, padx=5, pady=2, sticky='w')

        self.compression_radio = ttk.Radiobutton(
            self.options_frame, text=self.translations.get("enable_compression", "Enable Compression"), variable=self.option_var, value="compression"
        )
        self.compression_radio.grid(row=0, column=3, padx=5, pady=2, sticky='w')

        # Standardauswahl "None"
        self.none_radio.invoke()

        # Buttons Frame
        self.buttons_frame = ttk.Frame(self)
        self.buttons_frame.pack(pady=10, padx=10)

        # Convert Button
        self.convert_button = ttk.Button(self.buttons_frame, text=self.translations.get("convert", "Convert"), command=self.convert)
        self.convert_button.grid(row=0, column=0, padx=5)

        # Help Button
        self.help_button = ttk.Button(self.buttons_frame, text=self.translations.get("help", "Help"), command=self.show_help)
        self.help_button.grid(row=0, column=1, padx=5)

        # Create Shortcut Button
        self.shortcut_button = ttk.Button(self.buttons_frame, text=self.translations.get("create_shortcut", "Create Shortcut"), command=self.create_shortcut)
        self.shortcut_button.grid(row=0, column=2, padx=5)

        # Drag-and-Drop Area
        self.drag_drop_label = ttk.Label(self, text=self.translations.get("drag_drop", "Drag and drop files here:"))
        self.drag_drop_label.pack(pady=2, anchor='w', padx=10)
        self.drop_area = tk.Frame(self, relief='groove', borderwidth=2)
        self.drop_area.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        # Canvas für die Anzeige der Bilder
        self.canvas = tk.Canvas(self.drop_area)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbar
        self.scrollbar = ttk.Scrollbar(
            self.drop_area, orient=tk.VERTICAL, command=self.canvas.yview
        )
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Frame innerhalb des Canvas für die Bilder
        self.image_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.image_frame, anchor='nw')

        # Scroll-Konfiguration
        self.image_frame.bind(
            "<Configure>",
            lambda event: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )

        # Drag-and-Drop aktivieren
        self.canvas.drop_target_register(DND_FILES)
        self.canvas.dnd_bind('<<Drop>>', self.handle_drop)

        # Event Bindings
        self.bind('<Delete>', self.delete_selected)

    def change_language(self, event):
        selected_language_name = self.language_var.get()
        languages = self.get_available_languages()
        selected_language_code = None
        for code, name in languages:
            if name == selected_language_name:
                selected_language_code = code
                break
        if selected_language_code:
            self.load_translation(selected_language_code)
            self.update_gui_texts()
            self.save_config()
        else:
            messagebox.showerror("Error", "Selected language not found.")

    def update_gui_texts(self):
        # Update window title
        self.title(self.translations.get("title", "Conversion Program"))

        # Update all labels and buttons
        self.language_frame.config(text="Language")  # Optional: übersetzen

        self.path1_label.config(text=self.translations.get("path1_label", "Path1:"))
        self.path2_label.config(text=self.translations.get("path2_label", "Path2:"))

        self.compiler_frame.config(text=self.translations.get("compiler_selection", "Compiler Selection"))
        self.mesh_radio.config(text=self.translations.get("mesh_compiler", "Mesh Compiler"))
        self.texture_radio.config(text=self.translations.get("texture_compiler", "Texture Compiler"))

        self.options_frame.config(text=self.translations.get("options", "Options"))
        self.none_radio.config(text=self.translations.get("none", "None"))
        self.physics_mesh_radio.config(text=self.translations.get("enable_physics_mesh", "Enable Physics Mesh"))
        self.physics_object_radio.config(text=self.translations.get("enable_physics_object", "Enable Physics Object"))
        self.compression_radio.config(text=self.translations.get("enable_compression", "Enable Compression"))

        self.convert_button.config(text=self.translations.get("convert", "Convert"))
        self.help_button.config(text=self.translations.get("help", "Help"))
        self.shortcut_button.config(text=self.translations.get("create_shortcut", "Create Shortcut"))

        self.drag_drop_label.config(text=self.translations.get("drag_drop", "Drag and drop files here:"))

        # Update existing image labels if needed (optional)

        # Aktualisieren geöffneter Toplevel-Fenster (z.B. Help Output)
        for window in self.winfo_children():
            if isinstance(window, tk.Toplevel):
                if window.title() in [
                    "Help Output",
                    "Help Output (Français)",
                    "Help Output (Español)",
                    "Help Output (Italiano)",
                    "Help Output (Русский)",
                    "Help Output (简体中文)",
                    "Help Output (日本語)",
                    "Help Output (हिन्दी)"
                ]:
                    window.title(self.translations.get("help_output_title", "Help Output"))

    def load_config(self):
        config_file = 'config.xml'
        if not os.path.exists(config_file):
            self.current_language = 'en_US'  # Standardmäßig Englisch
            return
        try:
            tree = ET.parse(config_file)
            root = tree.getroot()
            language_element = root.find('language')
            if language_element is not None:
                self.current_language = language_element.text
            else:
                self.current_language = 'en_US'
        except ET.ParseError:
            messagebox.showerror("Error", "Failed to parse 'config.xml'. Using default language.")
            self.current_language = 'en_US'

    def save_config(self):
        config_file = 'config.xml'
        root = ET.Element("config")
        language_element = ET.SubElement(root, "language")
        language_element.text = self.current_language
        tree = ET.ElementTree(root)
        try:
            tree.write(config_file)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration:\n{e}")

    def update_compiler_selection(self):
        # Optionen zurücksetzen
        self.option_var.set("none")
        if self.compiler_var.get() == "mesh":
            # Enable Physics Mesh und Enable Physics Object verfügbar, Enable Compression deaktiviert
            self.physics_mesh_radio.config(state='normal')
            self.physics_object_radio.config(state='normal')
            self.compression_radio.config(state='disabled')
        elif self.compiler_var.get() == "texture":
            # Enable Compression verfügbar, Enable Physics Mesh und Enable Physics Object deaktiviert
            self.physics_mesh_radio.config(state='disabled')
            self.physics_object_radio.config(state='disabled')
            self.compression_radio.config(state='normal')
        else:
            # Alle deaktivieren
            self.physics_mesh_radio.config(state='disabled')
            self.physics_object_radio.config(state='disabled')
            self.compression_radio.config(state='disabled')
        # "None" immer aktivieren
        self.none_radio.config(state='normal')

    def calculate_new_size(self, image, width):
        w_percent = width / float(image.size[0])
        h_size = int((float(image.size[1]) * float(w_percent)))
        return (width, h_size)

    def handle_drop(self, event):
        files = self.tk.splitlist(event.data)
        for file_path in files:
            file_path = file_path.strip('{}')  # Entfernt geschweifte Klammern
            self.add_file(file_path)

    def add_file(self, file_path):
        if file_path not in self.file_list:
            self.file_list.append(file_path)
            # Aktualisiere das Bildraster
            self.update_image_grid()

    def update_image_grid(self):
        # Bestehende Widgets löschen
        for widget in self.image_frame.winfo_children():
            widget.destroy()

        # Anzahl der Bilder pro Zeile anpassen basierend auf der Fenstergröße
        images_per_row = 6  # Bei Bedarf anpassen

        # Bilder hinzufügen, max images_per_row pro Zeile
        for index, file_path in enumerate(self.file_list):
            row = index // images_per_row
            column = index % images_per_row

            # Container für Bild und Label
            container = tk.Frame(self.image_frame)

            # Bild laden und skalieren
            image_path = os.path.join('resources', 'file_image.png')
            if os.path.exists(image_path):
                try:
                    image = Image.open(image_path)
                    resized_image = image.resize(
                        self.calculate_new_size(image, 80), Image.LANCZOS
                    )
                    photo = ImageTk.PhotoImage(resized_image)
                    image_label = tk.Label(container, image=photo)
                    image_label.image = photo  # Referenz behalten
                except Exception:
                    image_label = tk.Label(
                        container, text=self.translations.get("image_not_found", "Image not found"), width=10, height=5
                    )
            else:
                # Platzhalter, falls das Bild nicht gefunden wird
                image_label = tk.Label(
                    container, text=self.translations.get("image_not_found", "Image not found"), width=10, height=5
                )

            image_label.file_path = file_path  # Datei-Pfad speichern
            image_label.pack()

            # Dateiname extrahieren und kürzen
            filename = os.path.basename(file_path)
            # Text kürzen, falls zu lang
            max_chars = 15  # Maximale Anzahl Zeichen, angepasst für bessere Darstellung
            if len(filename) > max_chars:
                filename = filename[:max_chars - 3] + "..."

            # Label für den Dateinamen
            name_label = tk.Label(container, text=filename, width=15)
            name_label.pack()

            # Event Bindings
            image_label.bind('<Button-1>', self.select_image)
            image_label.bind('<Button-3>', self.show_context_menu)
            name_label.bind('<Button-1>', self.select_image)
            name_label.bind('<Button-3>', self.show_context_menu)

            # Container im Grid platzieren
            container.grid(row=row, column=column, padx=2, pady=2)

        # Scrollbereich aktualisieren
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def select_image(self, event):
        # Markiert das ausgewählte Bild
        if self.selected_image:
            self.selected_image.config(bg='SystemButtonFace')
        event.widget.config(bg='lightblue')
        self.selected_image = event.widget

    def show_context_menu(self, event):
        # Kontextmenü zum Entfernen des Bildes
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(
            label=self.translations.get("remove", "Remove"), command=lambda: self.remove_image(event.widget)
        )
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def remove_image(self, widget):
        file_path = widget.file_path
        if file_path in self.file_list:
            self.file_list.remove(file_path)
        self.update_image_grid()
        if self.selected_image == widget:
            self.selected_image = None

    def delete_selected(self, event):
        # Entfernt das ausgewählte Bild
        if self.selected_image:
            self.remove_image(self.selected_image)
            self.selected_image = None

    def save_to_xml(self):
        root = ET.Element("paths")
        # Speichern von path1, path2, compiler_var und option_var
        path1_element = ET.SubElement(root, "path1")
        path1_element.text = self.path1_var.get()
        path2_element = ET.SubElement(root, "path2")
        path2_element.text = self.path2_var.get()
        compiler_element = ET.SubElement(root, "compiler")
        compiler_element.text = self.compiler_var.get()
        option_element = ET.SubElement(root, "option")
        option_element.text = self.option_var.get()

        tree = ET.ElementTree(root)
        tree.write("paths.xml")

    def load_from_xml(self):
        if os.path.exists("paths.xml"):
            try:
                tree = ET.parse("paths.xml")
                root = tree.getroot()
                path1_element = root.find('path1')
                path2_element = root.find('path2')
                compiler_element = root.find('compiler')
                option_element = root.find('option')
                if path1_element is not None:
                    self.path1_var.set(path1_element.text)
                if path2_element is not None:
                    self.path2_var.set(path2_element.text)
                if compiler_element is not None:
                    self.compiler_var.set(compiler_element.text)
                    self.update_compiler_selection()  # Auswahl aktualisieren
                if option_element is not None:
                    self.option_var.set(option_element.text)
                else:
                    self.option_var.set("none")
                # Optionen basierend auf der Compiler-Auswahl aktualisieren
                self.update_compiler_selection()
            except ET.ParseError:
                messagebox.showerror("Error", "Failed to parse 'paths.xml'.")

    def convert(self):
        # Speichern der Pfade in XML
        self.save_to_xml()

        texts = self.translations

        # Basispfad für path1 gemäß Benutzereingabe
        path1_base = self.path1_var.get()
        if not path1_base:
            messagebox.showerror(texts.get("error", "Error"), texts.get("enter_path1", "Please enter Path1."))
            return

        # Compiler-Pfad an path1 anhängen
        if self.compiler_var.get() == "mesh":
            path1 = os.path.join(path1_base, "mesh_compiler.com")
        elif self.compiler_var.get() == "texture":
            path1 = os.path.join(path1_base, "texture_compiler.com")
        else:
            messagebox.showerror(texts.get("error", "Error"), texts.get("select_compiler", "Please select a compiler."))
            return

        if not os.path.exists(path1):
            messagebox.showerror(texts.get("error", "Error"), f"{path1} not found.")
            return

        path2 = self.path2_var.get()
        if not path2:
            messagebox.showerror(texts.get("error", "Error"), texts.get("enter_path2", "Please enter Path2."))
            return

        for file_path in self.file_list:
            command = f'{path1} "{file_path}" -o "{path2}"'
            # Option hinzufügen
            if self.option_var.get() == "physics_mesh" and self.compiler_var.get() == "mesh":
                command += ' -m physics_mesh'
            elif self.option_var.get() == "physics_object" and self.compiler_var.get() == "mesh":
                command += ' -m physics'
            elif self.option_var.get() == "compression" and self.compiler_var.get() == "texture":
                command += ' -c'
            # Keine zusätzliche Option für "none"

            if sys.platform == "win32":
                CREATE_NO_WINDOW = 0x08000000
            else:
                CREATE_NO_WINDOW = 0  # Für Nicht-Windows-Systeme

            try:
                # Befehl in PowerShell unter C:\WINDOWS\system32 ausführen
                subprocess.run(
                    ["powershell", "-Command", command],
                    cwd=r"C:\WINDOWS\system32",
                    check=True,
                    creationflags=CREATE_NO_WINDOW
                )
            except subprocess.CalledProcessError as e:
                # Fehlermeldung anzeigen
                messagebox.showerror(texts.get("error", "Error"), f"{texts.get('an_error_occurred', 'An error occurred:')}\n{e}")
            except Exception as e:
                messagebox.showerror(texts.get("error", "Error"), f"{texts.get('unexpected_error_occurred', 'An unexpected error occurred:')}\n{e}")

    def show_help(self):
        texts = self.translations

        # Basispfad für path1 gemäß Benutzereingabe
        path1_base = self.path1_var.get()
        if not path1_base:
            messagebox.showerror(texts.get("error", "Error"), texts.get("enter_path1", "Please enter Path1."))
            return

        # Compiler-Pfad an path1 anhängen
        if self.compiler_var.get() == "mesh":
            path1 = os.path.join(path1_base, "mesh_compiler.com")
        elif self.compiler_var.get() == "texture":
            path1 = os.path.join(path1_base, "texture_compiler.com")
        else:
            messagebox.showerror(texts.get("error", "Error"), texts.get("select_compiler", "Please select a compiler."))
            return

        if not os.path.exists(path1):
            messagebox.showerror(texts.get("error", "Error"), f"{path1} not found.")
            return

        command = f'"{path1}" -h'
        try:
            # Befehl ausführen und Ausgabe erfassen
            result = subprocess.run(
                command, shell=True, capture_output=True, text=True
            )
            output = result.stdout + result.stderr
            # Ausgabe in einem neuen Fenster anzeigen
            self.display_help_output(output)
        except Exception as e:
            messagebox.showerror(texts.get("error", "Error"), f"{texts.get('failed_to_get_help', 'Failed to get help:')}\n{e}")

    def display_help_output(self, output):
        texts = self.translations
        # Neues Fenster zum Anzeigen der Hilfeausgabe erstellen
        help_window = tk.Toplevel(self)
        help_window.title(texts.get("help_output_title", "Help Output"))
        help_window.geometry("600x400+600+500")  # Position angepasst
        help_window.focus_force()

        # Text-Widget hinzufügen, um die Ausgabe anzuzeigen
        text_widget = tk.Text(help_window, wrap=tk.WORD)
        text_widget.insert(tk.END, output)
        text_widget.configure(state='disabled')  # Nur-Lese-Modus
        text_widget.pack(expand=True, fill=tk.BOTH)

    def create_shortcut(self):
        texts = self.translations
        try:
            # Pfad zum Desktop
            desktop = winshell.desktop()
            # Pfad zu diesem Skript oder ausführbaren Datei
            if getattr(sys, 'frozen', False):
                # Wenn das Skript zu einer ausführbaren Datei kompiliert wurde
                target = sys.executable
                args = ''
            else:
                # Wenn das Skript ausgeführt wird
                target = sys.executable
                script = os.path.abspath(__file__)
                args = f'"{script}"'

            # Pfad zur Verknüpfung
            shortcut_path = os.path.join(desktop, "ConversionProgram.lnk")

            # Verwenden der Windows Shell, um die Verknüpfung zu erstellen oder zu aktualisieren
            shell = win32com.client.Dispatch('WScript.Shell')

            # Überprüfen, ob die Verknüpfung bereits existiert
            if os.path.exists(shortcut_path):
                # Vorhandene Verknüpfung laden
                shortcut = shell.CreateShortCut(shortcut_path)
                current_target = shortcut.Targetpath
                current_args = shortcut.Arguments
                expected_target = target
                expected_args = args

                if current_target == expected_target and current_args == expected_args:
                    # Verknüpfung zeigt bereits auf dieses Programm
                    messagebox.showinfo(texts.get("shortcut_already_exists", "Shortcut already exists."), texts.get("shortcut_already_exists", "Shortcut already exists."))
                    return
                else:
                    # Verknüpfung aktualisieren
                    shortcut.Targetpath = expected_target
                    shortcut.Arguments = expected_args
                    shortcut.WorkingDirectory = os.path.dirname(expected_target)
                    shortcut.IconLocation = expected_target
                    shortcut.save()
                    messagebox.showinfo(texts.get("shortcut_updated", "Shortcut updated."), texts.get("shortcut_updated", "Shortcut updated."))
            else:
                # Neue Verknüpfung erstellen
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = target
                shortcut.Arguments = args
                shortcut.WorkingDirectory = os.path.dirname(target)
                shortcut.IconLocation = target
                shortcut.save()
                messagebox.showinfo(texts.get("shortcut_created", "Desktop shortcut created successfully."), texts.get("shortcut_created", "Desktop shortcut created successfully."))
        except Exception as e:
            messagebox.showerror(texts.get("error", "Error"), f"{texts.get('failed_to_create_shortcut', 'Failed to create shortcut:')}\n{e}")

if __name__ == "__main__":
    app = Application()
    app.mainloop()
