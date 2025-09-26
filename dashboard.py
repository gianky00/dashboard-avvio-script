import customtkinter as ctk
import json
import subprocess
import os
import threading
from tkinter import filedialog

class ScriptDialog(ctk.CTkToplevel):
    """
    Finestra di dialogo per aggiungere o modificare una configurazione di script.
    """
    def __init__(self, parent, tab_names, title="Aggiungi Script", script_data=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x450")  # Aumentata l'altezza per il nuovo campo
        self.transient(parent)
        self.result = None

        # Inizializzazione delle variabili
        self.name_var = ctk.StringVar()
        self.desc_var = ctk.StringVar()
        self.path_var = ctk.StringVar()
        self.excel_path_var = ctk.StringVar()
        self.tab_var = ctk.StringVar(value=tab_names[0])  # Default alla prima scheda

        if script_data:
            self.name_var.set(script_data.get("name", ""))
            self.desc_var.set(script_data.get("description", ""))
            self.path_var.set(script_data.get("path", ""))
            self.excel_path_var.set(script_data.get("excel_path", ""))
            self.tab_var.set(script_data.get("tab", tab_names[0]))

        # Creazione dei widget
        ctk.CTkLabel(self, text="Nome:").pack(pady=(10, 0))
        self.name_entry = ctk.CTkEntry(self, textvariable=self.name_var, width=300)
        self.name_entry.pack(pady=5)

        ctk.CTkLabel(self, text="Descrizione:").pack(pady=(10, 0))
        self.desc_entry = ctk.CTkEntry(self, textvariable=self.desc_var, width=300)
        self.desc_entry.pack(pady=5)

        ctk.CTkLabel(self, text="Assegna a Scheda:").pack(pady=(10, 0))
        self.tab_menu = ctk.CTkOptionMenu(self, variable=self.tab_var, values=tab_names, width=300)
        self.tab_menu.pack(pady=5)

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
            "excel_path": self.excel_path_var.get(),
            "tab": self.tab_var.get()
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
        self.state('zoomed')  # Avvia massimizzato
        ctk.set_appearance_mode("System")  # o "Dark", "Light"
        ctk.set_default_color_theme("blue")

        self.data_file = "data.json"
        self.scripts = self.load_data()

        # Layout principale
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)  # Riga per la tab view (espandibile)
        self.grid_rowconfigure(1, weight=0)  # Riga per il pulsante Add (fissa)
        self.grid_rowconfigure(2, weight=0)  # Riga per i log (fissa)

        # --- Interfaccia a Schede ---
        self.tab_view = ctk.CTkTabview(self, width=250)
        self.tab_view.grid(row=0, column=0, padx=20, pady=(20, 5), sticky="nsew")

        self.tab_names = ["Generale", "Schede", "Contabilità", "Programmazione", "Report Giornaliere"]
        self.scrollable_frames = {}

        for tab_name in self.tab_names:
            tab = self.tab_view.add(tab_name)
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)

            label = "Tutti gli Script" if tab_name == "Generale" else f"Script in {tab_name}"
            scrollable_frame = ctk.CTkScrollableFrame(tab, label_text=label)
            scrollable_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
            self.scrollable_frames[tab_name] = scrollable_frame

        # Pulsante per aggiungere nuovi script
        self.add_button = ctk.CTkButton(self, text="Aggiungi Nuovo Script", command=self.add_script)
        self.add_button.grid(row=1, column=0, padx=20, pady=(5, 10), sticky="ew")

        # --- Area Log ---
        log_frame = ctk.CTkFrame(self)
        log_frame.grid(row=2, column=0, padx=20, pady=(5, 10), sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)

        self.log_textbox = ctk.CTkTextbox(log_frame, height=150, activate_scrollbars=True)
        self.log_textbox.grid(row=0, column=0, sticky="nsew")
        self.log_textbox.configure(state="disabled") # Rendila non modificabile dall'utente

        self.clear_log_button = ctk.CTkButton(log_frame, text="Pulisci Log", command=self.clear_log)
        self.clear_log_button.grid(row=1, column=0, pady=(5, 0), sticky="e")

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
        # Pulisce tutti i frame scorrevoli prima di aggiornare
        for frame in self.scrollable_frames.values():
            for widget in frame.winfo_children():
                widget.destroy()

        # Aggiunge una riga per ogni script nella scheda corretta e in "Generale"
        for i, script in enumerate(self.scripts):
            # Gestisce la retrocompatibilità per gli script senza una scheda assegnata
            target_tab_name = script.get("tab", "Schede")

            # Aggiunge lo script alla sua scheda specifica
            if target_tab_name in self.scrollable_frames:
                parent_frame = self.scrollable_frames[target_tab_name]
                self.create_script_entry(parent_frame, i, script)

            # Aggiunge lo script anche alla scheda "Generale", ma solo se non è già la sua scheda di destinazione
            if target_tab_name != "Generale":
                generale_frame = self.scrollable_frames["Generale"]
                self.create_script_entry(generale_frame, i, script)

    def create_script_entry(self, parent_frame, index, script):
        entry_frame = ctk.CTkFrame(parent_frame)
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

    def _append_log_message(self, message):
        """METODO PRIVATO: Aggiunge un messaggio alla textbox dei log. DEVE essere chiamato dal thread principale."""
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message)
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def _read_process_output(self, process, script_name):
        """Legge l'output di un processo e programma l'aggiornamento della GUI."""
        self.after(0, self._append_log_message, f"--- Avvio del processo: {script_name} ---\n")
        for line in iter(process.stdout.readline, ''):
            self.after(0, self._append_log_message, line)
        process.stdout.close()
        return_code = process.wait()
        self.after(0, self._append_log_message, f"\n--- Processo '{script_name}' terminato con codice d'uscita: {return_code} ---\n")

    def launch_script(self, path):
        """Lancia uno script e cattura il suo output in un thread separato."""
        if not os.path.exists(path):
            self._append_log_message(f"Errore: Percorso non trovato: {path}\n")
            return

        try:
            process = subprocess.Popen(
                ['cmd', '/c', path],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                errors='replace',
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            script_name = os.path.basename(path)
            thread = threading.Thread(target=self._read_process_output, args=(process, script_name))
            thread.daemon = True
            thread.start()

        except Exception as e:
            self._append_log_message(f"Errore durante l'avvio dello script: {e}\n")

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
        # Passa i nomi delle schede (escludendo "Generale") alla finestra di dialogo
        assignable_tabs = [name for name in self.tab_names if name != "Generale"]
        dialog = ScriptDialog(self, tab_names=assignable_tabs, title="Aggiungi Nuovo Script")
        self.wait_window(dialog)

        if dialog.result:
            # Validazione base
            if dialog.result["name"] and dialog.result["path"]:
                self.scripts.append(dialog.result)
                self.save_data()
                self.refresh_script_list()

    def edit_script(self, index):
        script_to_edit = self.scripts[index]
        assignable_tabs = [name for name in self.tab_names if name != "Generale"]
        dialog = ScriptDialog(self, tab_names=assignable_tabs, title="Modifica Script", script_data=script_to_edit)
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

    def clear_log(self):
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

if __name__ == "__main__":
    app = App()
    app.mainloop()