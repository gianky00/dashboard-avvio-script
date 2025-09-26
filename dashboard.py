import customtkinter as ctk
import json
import subprocess
import os
import threading
from tkinter import filedialog, Menu, messagebox
from datetime import datetime
import itertools
import collections

class ScriptDialog(ctk.CTkToplevel):
    def __init__(self, parent, tab_names, title="Aggiungi Script", script_data=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x650")
        self.transient(parent)
        self.result = None

        self.name_var = ctk.StringVar()
        self.desc_var = ctk.StringVar()
        self.path_var = ctk.StringVar()
        self.excel_path_var = ctk.StringVar()
        self.tab_var = ctk.StringVar(value=tab_names[0] if tab_names else "")
        self.group_var = ctk.StringVar()

        if script_data:
            self.name_var.set(script_data.get("name", ""))
            self.desc_var.set(script_data.get("description", ""))
            self.path_var.set(script_data.get("path", ""))
            self.excel_path_var.set(script_data.get("excel_path", ""))
            self.tab_var.set(script_data.get("tab", tab_names[0] if tab_names else ""))
            self.group_var.set(script_data.get("group", ""))

        ctk.CTkLabel(self, text="Nome:").pack(padx=20, pady=(10, 0), anchor="w")
        self.name_entry = ctk.CTkEntry(self, textvariable=self.name_var)
        self.name_entry.pack(padx=20, pady=5, fill="x")
        ctk.CTkLabel(self, text="Descrizione:").pack(padx=20, pady=(10, 0), anchor="w")
        self.desc_entry = ctk.CTkEntry(self, textvariable=self.desc_var)
        self.desc_entry.pack(padx=20, pady=5, fill="x")
        ctk.CTkLabel(self, text="Assegna a Scheda:").pack(padx=20, pady=(10, 0), anchor="w")
        self.tab_menu = ctk.CTkOptionMenu(self, variable=self.tab_var, values=tab_names)
        self.tab_menu.pack(padx=20, pady=5, fill="x")
        ctk.CTkLabel(self, text="Nome Gruppo (Opzionale):").pack(padx=20, pady=(10, 0), anchor="w")
        self.group_entry = ctk.CTkEntry(self, textvariable=self.group_var)
        self.group_entry.pack(padx=20, pady=5, fill="x")
        path_frame = ctk.CTkFrame(self)
        path_frame.pack(padx=20, pady=5, fill="x")
        ctk.CTkLabel(path_frame, text="Percorso Script (.bat, .vbs):").pack(anchor="w")
        self.path_entry = ctk.CTkEntry(path_frame, textvariable=self.path_var)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        ctk.CTkButton(path_frame, text="Sfoglia", width=80, command=self.browse_bat_file).pack(side="left")
        excel_path_frame = ctk.CTkFrame(self)
        excel_path_frame.pack(padx=20, pady=5, fill="x")
        ctk.CTkLabel(excel_path_frame, text="Percorso File Excel (Opzionale):").pack(anchor="w")
        self.excel_path_entry = ctk.CTkEntry(excel_path_frame, textvariable=self.excel_path_var)
        self.excel_path_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        ctk.CTkButton(excel_path_frame, text="Sfoglia", width=80, command=self.browse_excel_file).pack(side="left")
        ctk.CTkLabel(self, text="Note Dettagliate:").pack(padx=20, pady=(10, 0), anchor="w")
        self.notes_textbox = ctk.CTkTextbox(self, height=100)
        self.notes_textbox.pack(padx=20, pady=5, fill="both", expand=True)
        if script_data and script_data.get("notes"):
            self.notes_textbox.insert("1.0", script_data.get("notes"))
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=20)
        ctk.CTkButton(button_frame, text="Salva", command=self.on_save).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Annulla", command=self.destroy).pack(side="left", padx=10)

    def browse_bat_file(self):
        filepath = filedialog.askopenfilename(title="Seleziona un file .bat o .vbs", filetypes=(("Script files", "*.bat *.vbs"), ("All files", "*.*")))
        if filepath: self.path_var.set(filepath)

    def browse_excel_file(self):
        filepath = filedialog.askopenfilename(title="Seleziona un file Excel", filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*")))
        if filepath: self.excel_path_var.set(filepath)

    def on_save(self):
        self.result = {"name": self.name_var.get(), "description": self.desc_var.get(), "path": self.path_var.get(), "excel_path": self.excel_path_var.get(), "tab": self.tab_var.get(), "notes": self.notes_textbox.get("1.0", "end-1c"), "group": self.group_var.get().strip()}
        self.destroy()

class SelectTabDialog(ctk.CTkToplevel):
    """Finestra di dialogo per selezionare una scheda da un elenco."""
    def __init__(self, parent, tab_names, title, prompt):
        super().__init__(parent)
        self.title(title)
        self.geometry("350x150")
        self.transient(parent)
        self.result = None

        self.label = ctk.CTkLabel(self, text=prompt)
        self.label.pack(padx=20, pady=10)

        self.tab_var = ctk.StringVar(value=tab_names[0])
        self.tab_menu = ctk.CTkOptionMenu(self, variable=self.tab_var, values=tab_names, width=300)
        self.tab_menu.pack(padx=20, pady=5)

        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=15)
        ctk.CTkButton(button_frame, text="OK", command=self.on_ok).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Annulla", command=self.destroy).pack(side="left", padx=10)

        self.grab_set()

    def on_ok(self):
        self.result = self.tab_var.get()
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

        self._create_menubar()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)
        self.grid_rowconfigure(3, weight=0)
        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", lambda name, index, mode: self._on_search())
        search_entry = ctk.CTkEntry(self, textvariable=self.search_var, placeholder_text="Cerca per nome o descrizione...")
        search_entry.grid(row=0, column=0, padx=20, pady=(10, 5), sticky="ew")
        self.tab_view = ctk.CTkTabview(self, width=250)
        self.tab_view.grid(row=1, column=0, padx=20, pady=(5, 5), sticky="nsew")
        self.tab_names = ["Generale"] + self.user_tab_names
        self.scrollable_frames = {}
        for tab_name in self.tab_names:
            tab = self.tab_view.add(tab_name)
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)
            label_text = "Tutti gli Script" if tab_name == "Generale" else f"Script in {tab_name}"
            scrollable_frame = ctk.CTkScrollableFrame(tab, label_text=label_text)
            scrollable_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
            self.scrollable_frames[tab_name] = scrollable_frame
        self.add_button = ctk.CTkButton(self, text="Aggiungi Nuovo Script", command=self.add_script)
        self.add_button.grid(row=2, column=0, padx=20, pady=(5, 10), sticky="ew")
        log_frame = ctk.CTkFrame(self)
        log_frame.grid(row=3, column=0, padx=20, pady=(5, 10), sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        self.log_textbox = ctk.CTkTextbox(log_frame, height=150, activate_scrollbars=True)
        self.log_textbox.grid(row=0, column=0, sticky="nsew")
        self.log_textbox.configure(state="disabled")
        self.clear_log_button = ctk.CTkButton(log_frame, text="Pulisci Log", command=self.clear_log)
        self.clear_log_button.grid(row=1, column=0, pady=(5, 0), sticky="e")
        self.refresh_script_list()

    def _create_menubar(self):
        menubar = Menu(self)
        file_menu = Menu(menubar, tearoff=0)
        file_menu.add_command(label="Esci", command=self.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        config_menu = Menu(menubar, tearoff=0)
        config_menu.add_command(label="Aggiungi Scheda...", command=self._add_tab)
        config_menu.add_command(label="Rinomina Scheda...", command=self._rename_tab)
        config_menu.add_command(label="Elimina Scheda...", command=self._delete_tab)
        menubar.add_cascade(label="Configurazione", menu=config_menu)

        self.configure(menu=menubar)

    def load_data(self):
        if not os.path.exists(self.data_file): return []
        try:
            with open(self.data_file, "r", encoding='utf-8') as f: scripts = json.load(f)
            needs_saving = False
            for i, script in enumerate(scripts):
                if "order" not in script:
                    script["order"] = i
                    needs_saving = True
            if needs_saving: self.save_data(scripts)
            return sorted(scripts, key=lambda s: s.get("order", 0))
        except (json.JSONDecodeError, FileNotFoundError): return []

    def save_data(self, scripts_data=None):
        if scripts_data is None: scripts_data = self.scripts
        with open(self.data_file, "w", encoding='utf-8') as f: json.dump(scripts_data, f, indent=4, ensure_ascii=False)

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

    def _populate_frame_with_grouped_scripts(self, parent_frame, scripts_list):
        key_func = lambda s: (s.get("group") or "").lower()
        sorted_scripts = sorted(scripts_list, key=key_func)
        for group_name, group_items_iterator in itertools.groupby(sorted_scripts, key=key_func):
            group_items = list(group_items_iterator)
            if group_name:
                group_header = ctk.CTkLabel(parent_frame, text=group_name.upper(), font=ctk.CTkFont(size=12, weight="bold"))
                group_header.pack(fill="x", padx=5, pady=(15, 5))
                separator = ctk.CTkFrame(parent_frame, height=1, fg_color="gray50")
                separator.pack(fill="x", padx=5, pady=(0, 10))
            for script in group_items:
                original_index = self.scripts.index(script)
                self.create_script_entry(parent_frame, original_index, script)

    def refresh_script_list(self):
        search_term = self.search_var.get().lower().strip()
        if search_term:
            scripts_to_display = [s for s in self.scripts if search_term in s.get("name", "").lower() or search_term in s.get("description", "").lower()]
        else:
            scripts_to_display = self.scripts
        scripts_by_tab = collections.defaultdict(list)
        for script in scripts_to_display:
            tab_name = script.get("tab", self.user_tab_names[0] if self.user_tab_names else "Schede")
            scripts_by_tab[tab_name].append(script)
        for frame in self.scrollable_frames.values():
            for widget in frame.winfo_children(): widget.destroy()
        for tab_name in self.user_tab_names:
            if tab_name in self.scrollable_frames:
                self._populate_frame_with_grouped_scripts(self.scrollable_frames[tab_name], scripts_by_tab[tab_name])
        self._populate_frame_with_grouped_scripts(self.scrollable_frames["Generale"], scripts_to_display)

    def create_script_entry(self, parent_frame, index, script):
        entry_frame = ctk.CTkFrame(parent_frame)
        entry_frame.pack(fill="x", expand=True, padx=10, pady=5)
        info_frame = ctk.CTkFrame(entry_frame)
        info_frame.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        ctk.CTkLabel(info_frame, text=script["name"], font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(info_frame, text=script["description"], anchor="w").pack(anchor="w")
        last_executed_ts = script.get("last_executed", "N/D")
        ctk.CTkLabel(info_frame, text=f"Ultima Esecuzione: {last_executed_ts}", text_color="gray", font=ctk.CTkFont(size=10)).pack(anchor="w", pady=(5,0))
        action_frame = ctk.CTkFrame(entry_frame)
        action_frame.pack(side="right", padx=5, pady=5)
        order_frame = ctk.CTkFrame(action_frame)
        order_frame.pack(side="left", padx=5)
        ctk.CTkButton(order_frame, text="▲", width=20, command=lambda i=index: self._move_script(i, -1)).pack()
        ctk.CTkButton(order_frame, text="▼", width=20, command=lambda i=index: self._move_script(i, 1)).pack()
        notes = script.get("notes")
        if notes:
            ctk.CTkButton(action_frame, text="Note", width=60, command=lambda n=notes, name=script.get("name"): self._show_notes(name, n)).pack(side="left", padx=5)
        excel_path = script.get("excel_path")
        if excel_path:
            ctk.CTkButton(action_frame, text="Apri Excel", width=90, command=lambda p=excel_path: self.open_excel(p)).pack(side="left", padx=5)
        ctk.CTkButton(action_frame, text="Avvia", width=80, command=lambda i=index: self.launch_script(i)).pack(side="left", padx=5)
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

    def launch_script(self, script_index):
        script_to_run = self.scripts[script_index]
        path = script_to_run.get("path")
        if not path or not os.path.exists(path):
            self._append_log_message(f"Errore: Percorso non valido o non trovato per lo script '{script_to_run.get('name')}'\n")
            return
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.scripts[script_index]["last_executed"] = now_str
        self.save_data()
        self.refresh_script_list()
        try:
            process = subprocess.Popen(['cmd', '/c', path], stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='replace', creationflags=subprocess.CREATE_NO_WINDOW)
            thread = threading.Thread(target=self._read_process_output, args=(process, os.path.basename(path)))
            thread.daemon = True
            thread.start()
        except Exception as e:
            self._append_log_message(f"Errore durante l'avvio dello script '{script_to_run.get('name')}': {e}\n")

    def open_excel(self, path):
        if os.path.exists(path):
            try: os.startfile(path)
            except Exception as e: print(f"Errore durante l'apertura del file Excel: {e}")
        else: print(f"Percorso non trovato: {path}")

    def add_script(self):
        dialog = ScriptDialog(self, tab_names=self.user_tab_names, title="Aggiungi Nuovo Script")
        self.wait_window(dialog)
        if dialog.result and dialog.result["name"] and dialog.result["path"]:
            new_script = dialog.result
            new_script["order"] = len(self.scripts)
            self.scripts.append(new_script)
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
        for i, script in enumerate(self.scripts):
            script["order"] = i
        self.save_data()
        self.refresh_script_list()

    def clear_log(self):
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

    def _move_script(self, index, direction):
        if direction == -1 and index > 0:
            self.scripts[index], self.scripts[index - 1] = self.scripts[index - 1], self.scripts[index]
        elif direction == 1 and index < len(self.scripts) - 1:
            self.scripts[index], self.scripts[index + 1] = self.scripts[index + 1], self.scripts[index]
        else:
            return
        for i, script in enumerate(self.scripts):
            script["order"] = i
        self.save_data()
        self.refresh_script_list()

    def _show_notes(self, script_name, notes):
        dialog = ctk.CTkToplevel(self)
        dialog.title(f"Note per: {script_name}")
        dialog.geometry("450x350")
        dialog.transient(self)
        dialog.attributes("-topmost", True)
        textbox = ctk.CTkTextbox(dialog, wrap="word")
        textbox.pack(expand=True, fill="both", padx=10, pady=10)
        textbox.insert("1.0", notes)
        textbox.configure(state="disabled")
        ok_button = ctk.CTkButton(dialog, text="OK", command=dialog.destroy, width=100)
        ok_button.pack(pady=10)
        dialog.grab_set()

    def _show_restart_dialog(self):
        messagebox.showinfo("Riavvio Richiesto", "La modifica è stata salvata. Per favore, riavvia l'applicazione per vedere le modifiche.")

    def _on_search(self):
        self.refresh_script_list()

    def _add_tab(self):
        dialog = ctk.CTkInputDialog(text="Inserisci il nome della nuova scheda:", title="Aggiungi Scheda")
        new_tab_name = dialog.get_input()
        if not new_tab_name or not new_tab_name.strip(): return
        new_tab_name = new_tab_name.strip()
        if new_tab_name in self.tab_names:
            messagebox.showerror("Errore", f"La scheda '{new_tab_name}' esiste già.")
            return
        self.user_tab_names.append(new_tab_name)
        self.config["tabs"] = self.user_tab_names
        self.save_config()
        self._show_restart_dialog()

    def _rename_tab(self):
        select_dialog = SelectTabDialog(self, self.user_tab_names, "Rinomina Scheda", "Seleziona la scheda da rinominare:")
        self.wait_window(select_dialog)
        old_name = select_dialog.result

        if not old_name: return

        new_name_dialog = ctk.CTkInputDialog(text=f"Inserisci il nuovo nome per la scheda '{old_name}':", title="Rinomina Scheda")
        new_name = new_name_dialog.get_input()

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
        self._show_restart_dialog()

    def _delete_tab(self):
        if len(self.user_tab_names) <= 1:
            messagebox.showerror("Errore", "Non puoi eliminare l'ultima scheda utente.")
            return

        select_dialog = SelectTabDialog(self, self.user_tab_names, "Elimina Scheda", "Seleziona la scheda da eliminare:")
        self.wait_window(select_dialog)
        tab_to_delete = select_dialog.result

        if not tab_to_delete: return

        if not messagebox.askyesno("Conferma Eliminazione", f"Sei sicuro di voler eliminare la scheda '{tab_to_delete}'?\nGli script associati verranno spostati nella prima scheda disponibile."):
            return

        fallback_tab = next(tab for tab in self.user_tab_names if tab != tab_to_delete)
        self.user_tab_names.remove(tab_to_delete)
        self.config["tabs"] = self.user_tab_names
        self.save_config()

        for script in self.scripts:
            if script.get("tab") == tab_to_delete:
                script["tab"] = fallback_tab
        self.save_data()
        self._show_restart_dialog()

if __name__ == "__main__":
    app = App()
    app.mainloop()