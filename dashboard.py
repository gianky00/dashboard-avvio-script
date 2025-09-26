import customtkinter as ctk
import json
import subprocess
import os
from tkinter import filedialog

class ScriptDialog(ctk.CTkToplevel):
    """
    Finestra di dialogo per aggiungere o modificare una configurazione di script.
    """
    def __init__(self, parent, title="Aggiungi Script", script_data=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x400")  # Aumentata l'altezza per il nuovo campo
        self.transient(parent)  # Mantiene la finestra in primo piano
        self.result = None

        if script_data:
            self.name_var = ctk.StringVar(value=script_data.get("name", ""))
            self.desc_var = ctk.StringVar(value=script_data.get("description", ""))
            self.path_var = ctk.StringVar(value=script_data.get("path", ""))
            self.excel_path_var = ctk.StringVar(value=script_data.get("excel_path", ""))
        else:
            self.name_var = ctk.StringVar()
            self.desc_var = ctk.StringVar()
            self.path_var = ctk.StringVar()
            self.excel_path_var = ctk.StringVar()

        # Creazione dei widget
        ctk.CTkLabel(self, text="Nome:").pack(pady=(10, 0))
        self.name_entry = ctk.CTkEntry(self, textvariable=self.name_var, width=300)
        self.name_entry.pack(pady=5)

        ctk.CTkLabel(self, text="Descrizione:").pack(pady=(10, 0))
        self.desc_entry = ctk.CTkEntry(self, textvariable=self.desc_var, width=300)
        self.desc_entry.pack(pady=5)

        ctk.CTkLabel(self, text="Percorso Script (.bat):").pack(pady=(10, 0))
        path_frame = ctk.CTkFrame(self)
        path_frame.pack(pady=5)
        self.path_entry = ctk.CTkEntry(path_frame, textvariable=self.path_var, width=250)
        self.path_entry.pack(side="left", padx=(0, 5))
        ctk.CTkButton(path_frame, text="Sfoglia", width=50, command=self.browse_bat_file).pack(side="left")

        ctk.CTkLabel(self, text="Percorso File Excel (Opzionale):").pack(pady=(10, 0))
        excel_path_frame = ctk.CTkFrame(self)
        excel_path_frame.pack(pady=5)
        self.excel_path_entry = ctk.CTkEntry(excel_path_frame, textvariable=self.excel_path_var, width=250)
        self.excel_path_entry.pack(side="left", padx=(0, 5))
        ctk.CTkButton(excel_path_frame, text="Sfoglia", width=50, command=self.browse_excel_file).pack(side="left")

        # Pulsanti di conferma e annullamento
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=20)
        ctk.CTkButton(button_frame, text="Salva", command=self.on_save).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Annulla", command=self.destroy).pack(side="left", padx=10)

    def browse_bat_file(self):
        filepath = filedialog.askopenfilename(
            title="Seleziona un file .bat",
            filetypes=(("Batch files", "*.bat"), ("All files", "*.*"))
        )
        if filepath:
            self.path_var.set(filepath)

    def browse_excel_file(self):
        filepath = filedialog.askopenfilename(
            title="Seleziona un file Excel",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filepath:
            self.excel_path_var.set(filepath)

    def on_save(self):
        self.result = {
            "name": self.name_var.get(),
            "description": self.desc_var.get(),
            "path": self.path_var.get(),
            "excel_path": self.excel_path_var.get()
        }
        self.destroy()

class App(ctk.CTk):
    """
    Applicazione Dashboard principale.
    """
    def __init__(self):
        super().__init__()

        self.title("Dashboard di Avvio Script")
        self.geometry("700x500")
        self.attributes('-fullscreen', True)  # Avvia a schermo intero
        ctk.set_appearance_mode("System")  # o "Dark", "Light"
        ctk.set_default_color_theme("blue")

        self.data_file = "data.json"
        self.scripts = self.load_data()

        # Layout principale
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Interfaccia a Schede ---
        self.tab_view = ctk.CTkTabview(self, width=250)
        self.tab_view.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        self.tab_view.add("Schede")
        self.tab_view.add("Contabilità")
        self.tab_view.add("Programmazione")
        self.tab_view.add("Report Giornaliere")

        # --- Contenuto della scheda "Schede" ---
        schede_tab = self.tab_view.tab("Schede")
        schede_tab.grid_columnconfigure(0, weight=1)
        schede_tab.grid_rowconfigure(0, weight=1)

        self.scrollable_frame = ctk.CTkScrollableFrame(schede_tab, label_text="Script Disponibili")
        self.scrollable_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.add_button = ctk.CTkButton(schede_tab, text="Aggiungi Script", command=self.add_script)
        self.add_button.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

        # --- Contenuto delle altre schede (segnaposto) ---
        contabilita_tab = self.tab_view.tab("Contabilità")
        ctk.CTkLabel(contabilita_tab, text="Contenuto della sezione Contabilità.").pack(padx=20, pady=20)

        programmazione_tab = self.tab_view.tab("Programmazione")
        ctk.CTkLabel(programmazione_tab, text="Contenuto della sezione Programmazione.").pack(padx=20, pady=20)

        report_tab = self.tab_view.tab("Report Giornaliere")
        ctk.CTkLabel(report_tab, text="Contenuto della sezione Report Giornaliere.").pack(padx=20, pady=20)

        self.refresh_script_list()

    def load_data(self):
        if not os.path.exists(self.data_file):
            return []
        try:
            with open(self.data_file, "r") as f:
                return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError):
            return []

    def save_data(self):
        with open(self.data_file, "w") as f:
            json.dump(self.scripts, f, indent=4)

    def refresh_script_list(self):
        # Pulisce il frame prima di aggiornare
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        # Aggiunge una riga per ogni script
        for i, script in enumerate(self.scripts):
            self.create_script_entry(i, script)

    def create_script_entry(self, index, script):
        entry_frame = ctk.CTkFrame(self.scrollable_frame)
        entry_frame.pack(fill="x", expand=True, padx=10, pady=5)

        info_frame = ctk.CTkFrame(entry_frame)
        info_frame.pack(side="left", fill="x", expand=True, padx=5, pady=5)

        name_label = ctk.CTkLabel(info_frame, text=script["name"], font=ctk.CTkFont(size=14, weight="bold"))
        name_label.pack(anchor="w")

        desc_label = ctk.CTkLabel(info_frame, text=script["description"], anchor="w")
        desc_label.pack(anchor="w")

        # Frame per i pulsanti di azione
        action_frame = ctk.CTkFrame(entry_frame)
        action_frame.pack(side="right", padx=5, pady=5)

        # Aggiungi il pulsante "Apri Excel" solo se il percorso è specificato
        excel_path = script.get("excel_path")
        if excel_path:
            ctk.CTkButton(action_frame, text="Apri Excel", width=90, command=lambda p=excel_path: self.open_excel(p)).pack(side="left", padx=5)

        ctk.CTkButton(action_frame, text="Avvia", width=80, command=lambda p=script["path"]: self.launch_script(p)).pack(side="left", padx=5)
        ctk.CTkButton(action_frame, text="Modifica", width=80, command=lambda i=index: self.edit_script(i)).pack(side="left", padx=5)
        ctk.CTkButton(action_frame, text="Elimina", width=80, fg_color="red", hover_color="darkred", command=lambda i=index: self.delete_script(i)).pack(side="left", padx=5)

    def launch_script(self, path):
        if os.path.exists(path):
            try:
                # Esegue lo script in una nuova finestra di terminale
                subprocess.Popen(f'cmd /c start "{os.path.basename(path)}" "{path}"', shell=True)
            except Exception as e:
                print(f"Errore durante l'avvio dello script: {e}")
        else:
            print(f"Percorso non trovato: {path}")

    def open_excel(self, path):
        if os.path.exists(path):
            try:
                # Apre il file con l'applicazione predefinita (es. Excel)
                os.startfile(path)
            except Exception as e:
                print(f"Errore durante l'apertura del file Excel: {e}")
        else:
            print(f"Percorso non trovato: {path}")

    def add_script(self):
        dialog = ScriptDialog(self, title="Aggiungi Nuovo Script")
        self.wait_window(dialog)

        if dialog.result:
            # Validazione base
            if dialog.result["name"] and dialog.result["path"]:
                self.scripts.append(dialog.result)
                self.save_data()
                self.refresh_script_list()

    def edit_script(self, index):
        script_to_edit = self.scripts[index]
        dialog = ScriptDialog(self, title="Modifica Script", script_data=script_to_edit)
        self.wait_window(dialog)

        if dialog.result:
            if dialog.result["name"] and dialog.result["path"]:
                self.scripts[index] = dialog.result
                self.save_data()
                self.refresh_script_list()

    def delete_script(self, index):
        self.scripts.pop(index)
        self.save_data()
        self.refresh_script_list()

if __name__ == "__main__":
    app = App()
    app.mainloop()