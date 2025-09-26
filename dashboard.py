import customtkinter as ctk
import json
import subprocess
import os
import threading
from tkinter import filedialog
from tkinter import messagebox

class ScriptDialog(ctk.CTkToplevel):
    """
    Finestra di dialogo per aggiungere o modificare una configurazione di script.
    """
    def __init__(self, parent, tab_names, title="Aggiungi Script", script_data=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x450")
        self.transient(parent)
        self.result = None

        self.name_var = ctk.StringVar()
        self.desc_var = ctk.StringVar()
        self.path_var = ctk.StringVar()
        self.excel_path_var = ctk.StringVar()
        self.tab_var = ctk.StringVar(value=tab_names[0] if tab_names else "")

        if script_data:
            self.name_var.set(script_data.get("name", ""))
            self.desc_var.set(script_data.get("description", ""))
            self.path_var.set(script_data.get("path", ""))
            self.excel_path_var.set(script_data.get("excel_path", ""))
            self.tab_var.set(script_data.get("tab", tab_names[0] if tab_names else ""))

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
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=20)
        ctk.CTkButton(button_frame, text="Salva", command=self.on_save).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Annulla", command=self.destroy).pack(side="left", padx=10)

    def browse_bat_file(self):
        filepath = filedialog.askopenfilename(title="Seleziona un file .bat", filetypes=(("Batch files", "*.bat"), ("All files", "*.*")))
        if filepath: self.path_var.set(filepath)

    def browse_excel_file(self):
        filepath = filedialog.askopenfilename(title="Seleziona un file Excel", filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))
        if filepath: self.excel_path_var.set(filepath)

    def on_save(self):
        self.result = {"name": self.name_var.get(), "description": self.desc_var.get(), "path": self.path_var.get(), "excel_path": self.excel_path_var.get(), "tab": self.tab_var.get()}
        self.destroy()

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Dashboard di Avvio Script")
        self.geometry("700x500")
        self.state('zoomed')
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.data_file = "data.json"
        self.config_file = "config.json"
        self.scripts = self.load_data()
        self.config = self.load_config()
        self.user_tab_names = self.config.get("tabs", ["Schede"])

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=0)

        self.tab_view = ctk.CTkTabview(self, width=250)
        self.tab_view.grid(row=0, column=0, padx=20, pady=(20, 5), sticky="nsew")

        self.fixed_tab_names = ["Generale", "Configurazione"]
        self.tab_names = self.fixed_tab_names + self.user_tab_names
        self.scrollable_frames = {}
        for tab_name in self.tab_names:
            tab = self.tab_view.add(tab_name)
            if tab_name != "Configurazione":
                tab.grid_columnconfigure(0, weight=1)
                tab.grid_rowconfigure(0, weight=1)
                label_text = "Tutti gli Script" if tab_name == "Generale" else f"Script in {tab_name}"
                scrollable_frame = ctk.CTkScrollableFrame(tab, label_text=label_text)
                scrollable_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
                self.scrollable_frames[tab_name] = scrollable_frame

        self.add_button = ctk.CTkButton(self, text="Aggiungi Nuovo Script", command=self.add_script)
        self.add_button.grid(row=1, column=0, padx=20, pady=(5, 10), sticky="ew")

        log_frame = ctk.CTkFrame(self)
        log_frame.grid(row=2, column=0, padx=20, pady=(5, 10), sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        self.log_textbox = ctk.CTkTextbox(log_frame, height=150, activate_scrollbars=True)
        self.log_textbox.grid(row=0, column=0, sticky="nsew")
        self.log_textbox.configure(state="disabled")
        self.clear_log_button = ctk.CTkButton(log_frame, text="Pulisci Log", command=self.clear_log)
        self.clear_log_button.grid(row=1, column=0, pady=(5, 0), sticky="e")

        self._setup_config_tab()
        self.refresh_script_list()

    def _setup_config_tab(self):
        config_tab = self.tab_view.tab("Configurazione")
        config_tab.grid_columnconfigure(0, weight=1)
        config_tab.grid_rowconfigure(1, weight=1)

        add_frame = ctk.CTkFrame(config_tab)
        add_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        add_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(add_frame, text="Aggiungi Nuova Scheda:").grid(row=0, column=0, padx=10, pady=(10,0), sticky="w")
        self.new_tab_entry = ctk.CTkEntry(add_frame, placeholder_text="Nome nuova scheda")
        self.new_tab_entry.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        ctk.CTkButton(add_frame, text="Aggiungi", command=self._add_tab).grid(row=1, column=1, padx=10, pady=5)

        self.manage_tabs_frame = ctk.CTkScrollableFrame(config_tab, label_text="Gestisci Schede Esistenti")
        self.manage_tabs_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.manage_tabs_frame.grid_columnconfigure(0, weight=1)
        self._refresh_config_tab_list()

    def _refresh_config_tab_list(self):
        for widget in self.manage_tabs_frame.winfo_children():
            widget.destroy()
        for i, tab_name in enumerate(self.user_tab_names):
            tab_entry_frame = ctk.CTkFrame(self.manage_tabs_frame)
            tab_entry_frame.grid(row=i, column=0, padx=5, pady=5, sticky="ew")
            tab_entry_frame.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(tab_entry_frame, text=tab_name).grid(row=0, column=0, padx=10, pady=5, sticky="w")
            ctk.CTkButton(tab_entry_frame, text="Rinomina", width=80, command=lambda name=tab_name: self._rename_tab(name)).grid(row=0, column=1, padx=5, pady=5)
            if len(self.user_tab_names) > 1:
                ctk.CTkButton(tab_entry_frame, text="Elimina", width=80, fg_color="red", hover_color="darkred", command=lambda name=tab_name: self._delete_tab(name)).grid(row=0, column=2, padx=5, pady=5)

    def load_data(self):
        if not os.path.exists(self.data_file): return []
        try:
            with open(self.data_file, "r", encoding='utf-8') as f: return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError): return []

    def save_data(self):
        with open(self.data_file, "w", encoding='utf-8') as f: json.dump(self.scripts, f, indent=4, ensure_ascii=False)

    def load_config(self):
        default_config = {"tabs": ["Schede", "Contabilità", "Programmazione", "Report Giornaliere", "Strumenti Campione"]}
        if not os.path.exists(self.config_file):
            self.save_config(default_config)
            return default_config
        try:
            with open(self.config_file, "r", encoding='utf-8') as f: return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError): return default_config

    def save_config(self, config_data=None):
        if config_data is None: config_data = self.config
        with open(self.config_file, "w", encoding='utf-8') as f: json.dump(config_data, f, indent=4, ensure_ascii=False)

    def refresh_script_list(self):
        for frame in self.scrollable_frames.values():
            for widget in frame.winfo_children(): widget.destroy()
        for i, script in enumerate(self.scripts):
            target_tab_name = script.get("tab", self.user_tab_names[0] if self.user_tab_names else "Schede")
            if target_tab_name in self.scrollable_frames:
                self.create_script_entry(self.scrollable_frames[target_tab_name], i, script)
            if target_tab_name != "Generale":
                self.create_script_entry(self.scrollable_frames["Generale"], i, script)

    def create_script_entry(self, parent_frame, index, script):
        entry_frame = ctk.CTkFrame(parent_frame)
        entry_frame.pack(fill="x", expand=True, padx=10, pady=5)
        info_frame = ctk.CTkFrame(entry_frame)
        info_frame.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        ctk.CTkLabel(info_frame, text=script["name"], font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(info_frame, text=script["description"], anchor="w").pack(anchor="w")
        action_frame = ctk.CTkFrame(entry_frame)
        action_frame.pack(side="right", padx=5, pady=5)
        excel_path = script.get("excel_path")
        if excel_path:
            ctk.CTkButton(action_frame, text="Apri Excel", width=90, command=lambda p=excel_path: self.open_excel(p)).pack(side="left", padx=5)
        ctk.CTkButton(action_frame, text="Avvia", width=80, command=lambda p=script["path"]: self.launch_script(p)).pack(side="left", padx=5)
        ctk.CTkButton(action_frame, text="Modifica", width=80, command=lambda i=index: self.edit_script(i)).pack(side="left", padx=5)
        ctk.CTkButton(action_frame, text="Elimina", width=80, fg_color="red", hover_color="darkred", command=lambda i=index: self.delete_script(i)).pack(side="left", padx=5)

    def _append_log_message(self, message):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message)
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def _read_process_output(self, process, script_name):
        self.after(0, self._append_log_message, f"--- Avvio del processo: {script_name} ---\n")
        for line in iter(process.stdout.readline, ''): self.after(0, self._append_log_message, line)
        process.stdout.close()
        return_code = process.wait()
        self.after(0, self._append_log_message, f"\n--- Processo '{script_name}' terminato con codice d'uscita: {return_code} ---\n")

    def launch_script(self, path):
        if not os.path.exists(path):
            self._append_log_message(f"Errore: Percorso non trovato: {path}\n")
            return
        try:
            process = subprocess.Popen(['cmd', '/c', path], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='replace', creationflags=subprocess.CREATE_NO_WINDOW)
            thread = threading.Thread(target=self._read_process_output, args=(process, os.path.basename(path)))
            thread.daemon = True
            thread.start()
        except Exception as e:
            self._append_log_message(f"Errore durante l'avvio dello script: {e}\n")

    def open_excel(self, path):
        if os.path.exists(path):
            try: os.startfile(path)
            except Exception as e: print(f"Errore durante l'apertura del file Excel: {e}")
        else: print(f"Percorso non trovato: {path}")

    def add_script(self):
        dialog = ScriptDialog(self, tab_names=self.user_tab_names, title="Aggiungi Nuovo Script")
        self.wait_window(dialog)
        if dialog.result and dialog.result["name"] and dialog.result["path"]:
            self.scripts.append(dialog.result)
            self.save_data()
            self.refresh_script_list()

    def edit_script(self, index):
        script_to_edit = self.scripts[index]
        dialog = ScriptDialog(self, tab_names=self.user_tab_names, title="Modifica Script", script_data=script_to_edit)
        self.wait_window(dialog)
        if dialog.result and dialog.result["name"] and dialog.result["path"]:
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

    def _show_restart_dialog(self):
        messagebox.showinfo("Riavvio Richiesto", "La modifica è stata salvata. Per favore, riavvia l'applicazione per vedere le modifiche.")

    def _add_tab(self):
        new_tab_name = self.new_tab_entry.get().strip()
        if not new_tab_name:
            messagebox.showerror("Errore", "Il nome della scheda non può essere vuoto.")
            return
        if new_tab_name in self.tab_names:
            messagebox.showerror("Errore", f"La scheda '{new_tab_name}' esiste già.")
            return
        self.user_tab_names.append(new_tab_name)
        self.config["tabs"] = self.user_tab_names
        self.save_config()
        self.new_tab_entry.delete(0, "end")
        self._refresh_config_tab_list()
        self._show_restart_dialog()

    def _rename_tab(self, old_name):
        dialog = ctk.CTkInputDialog(text=f"Inserisci il nuovo nome per la scheda '{old_name}':", title="Rinomina Scheda")
        new_name = dialog.get_input()
        if not new_name or not new_name.strip(): return
        new_name = new_name.strip()
        if new_name == old_name: return
        if new_name in self.tab_names:
            messagebox.showerror("Errore", f"La scheda '{new_name}' esiste già.")
            return
        try:
            index = self.user_tab_names.index(old_name)
            self.user_tab_names[index] = new_name
            self.config["tabs"] = self.user_tab_names
            self.save_config()
        except ValueError:
            messagebox.showerror("Errore", "Impossibile trovare la scheda da rinominare.")
            return
        for script in self.scripts:
            if script.get("tab") == old_name:
                script["tab"] = new_name
        self.save_data()
        self._refresh_config_tab_list()
        self._show_restart_dialog()

    def _delete_tab(self, tab_name_to_delete):
        if len(self.user_tab_names) <= 1:
            messagebox.showerror("Errore", "Non puoi eliminare l'ultima scheda utente.")
            return
        if not messagebox.askyesno("Conferma Eliminazione", f"Sei sicuro di voler eliminare la scheda '{tab_name_to_delete}'?\nGli script associati verranno spostati nella prima scheda disponibile."):
            return
        fallback_tab = next(tab for tab in self.user_tab_names if tab != tab_name_to_delete)
        self.user_tab_names.remove(tab_name_to_delete)
        self.config["tabs"] = self.user_tab_names
        self.save_config()
        for script in self.scripts:
            if script.get("tab") == tab_name_to_delete:
                script["tab"] = fallback_tab
        self.save_data()
        self._refresh_config_tab_list()
        self._show_restart_dialog()

if __name__ == "__main__":
    app = App()
    app.mainloop()