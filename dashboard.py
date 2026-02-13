import customtkinter as ctk
import json
import subprocess
import os
import threading
import locale
import shutil
import tempfile
from tkinter import filedialog, Menu, messagebox
from datetime import datetime
import itertools
import collections

# --- COSTANTI E CONFIGURAZIONE ---
APP_NAME = "Dashboard Avvio Script"
DATA_FILE = "data.json"
CONFIG_FILE = "config.json"
DEFAULT_TABS = ["Schede", "Contabilità", "Programmazione", "Report Giornaliere", "Strumenti Campione"]
VERSION = "1.1.0"

class ScriptDialog(ctk.CTkToplevel):
    """Dialogo per aggiungere o modificare uno script."""
    def __init__(self, parent, tab_names, title="Aggiungi Script", script_data=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x700")
        self.transient(parent)
        self.resizable(False, False)
        self.result = None

        # Variabili
        self.name_var = ctk.StringVar(value=script_data.get("name", "") if script_data else "")
        self.desc_var = ctk.StringVar(value=script_data.get("description", "") if script_data else "")
        self.path_var = ctk.StringVar(value=script_data.get("path", "") if script_data else "")
        self.excel_path_var = ctk.StringVar(value=script_data.get("excel_path", "") if script_data else "")
        
        default_tab = tab_names[0] if tab_names else ""
        current_tab = script_data.get("tab", default_tab) if script_data else default_tab
        if current_tab not in tab_names:
            current_tab = default_tab
            
        self.tab_var = ctk.StringVar(value=current_tab)
        self.group_var = ctk.StringVar(value=script_data.get("group", "") if script_data else "")

        # Layout UI
        self._build_ui(tab_names, script_data)
        
        # Focus iniziale
        self.after(100, lambda: self.name_entry.focus())

    def _build_ui(self, tab_names, script_data):
        # Nome
        ctk.CTkLabel(self, text="Nome Script:", font=("Arial", 12, "bold")).pack(padx=20, pady=(15, 0), anchor="w")
        self.name_entry = ctk.CTkEntry(self, textvariable=self.name_var)
        self.name_entry.pack(padx=20, pady=5, fill="x")

        # Descrizione
        ctk.CTkLabel(self, text="Descrizione:").pack(padx=20, pady=(10, 0), anchor="w")
        self.desc_entry = ctk.CTkEntry(self, textvariable=self.desc_var)
        self.desc_entry.pack(padx=20, pady=5, fill="x")

        # Scheda e Gruppo
        row_frame = ctk.CTkFrame(self, fg_color="transparent")
        row_frame.pack(padx=20, pady=5, fill="x")
        
        # Scheda
        tab_frame = ctk.CTkFrame(row_frame, fg_color="transparent")
        tab_frame.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ctk.CTkLabel(tab_frame, text="Scheda:").pack(anchor="w")
        self.tab_menu = ctk.CTkOptionMenu(tab_frame, variable=self.tab_var, values=tab_names)
        self.tab_menu.pack(fill="x")

        # Gruppo
        group_frame = ctk.CTkFrame(row_frame, fg_color="transparent")
        group_frame.pack(side="right", fill="x", expand=True, padx=(5, 0))
        ctk.CTkLabel(group_frame, text="Gruppo (Opz.):").pack(anchor="w")
        self.group_entry = ctk.CTkEntry(group_frame, textvariable=self.group_var)
        self.group_entry.pack(fill="x")

        # Percorso Script
        ctk.CTkLabel(self, text="Percorso Script (.bat, .vbs, .cmd):", font=("Arial", 12, "bold")).pack(padx=20, pady=(15, 0), anchor="w")
        path_frame = ctk.CTkFrame(self)
        path_frame.pack(padx=20, pady=5, fill="x")
        self.path_entry = ctk.CTkEntry(path_frame, textvariable=self.path_var)
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(5, 5), pady=5)
        ctk.CTkButton(path_frame, text="...", width=40, command=self.browse_bat_file).pack(side="right", padx=(0, 5), pady=5)

        # Percorso Excel
        ctk.CTkLabel(self, text="Percorso File Excel (Opzionale):").pack(padx=20, pady=(10, 0), anchor="w")
        excel_frame = ctk.CTkFrame(self)
        excel_frame.pack(padx=20, pady=5, fill="x")
        self.excel_path_entry = ctk.CTkEntry(excel_frame, textvariable=self.excel_path_var)
        self.excel_path_entry.pack(side="left", fill="x", expand=True, padx=(5, 5), pady=5)
        ctk.CTkButton(excel_frame, text="...", width=40, command=self.browse_excel_file).pack(side="right", padx=(0, 5), pady=5)

        # Note
        ctk.CTkLabel(self, text="Note / Istruzioni:").pack(padx=20, pady=(10, 0), anchor="w")
        self.notes_textbox = ctk.CTkTextbox(self, height=100)
        self.notes_textbox.pack(padx=20, pady=5, fill="both", expand=True)
        if script_data and script_data.get("notes"):
            self.notes_textbox.insert("1.0", script_data.get("notes"))

        # Bottoni
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(pady=20, fill="x")
        ctk.CTkButton(button_frame, text="Salva", command=self.on_save, fg_color="green", hover_color="darkgreen").pack(side="right", padx=20)
        ctk.CTkButton(button_frame, text="Annulla", command=self.destroy, fg_color="gray", hover_color="gray30").pack(side="right", padx=(0, 10))

    def browse_bat_file(self):
        filepath = filedialog.askopenfilename(
            title="Seleziona Script",
            filetypes=(("Script", "*.bat *.vbs *.cmd *.ps1 *.py"), ("Tutti i file", "*.*"))
        )
        if filepath: self.path_var.set(os.path.normpath(filepath))

    def browse_excel_file(self):
        filepath = filedialog.askopenfilename(
            title="Seleziona Excel",
            filetypes=(("Excel", "*.xlsx *.xls *.xlsm *.csv"), ("Tutti i file", "*.*"))
        )
        if filepath: self.excel_path_var.set(os.path.normpath(filepath))

    def on_save(self):
        name = self.name_var.get().strip()
        path = self.path_var.get().strip()
        
        if not name:
            messagebox.showwarning("Dati Mancanti", "Il nome dello script è obbligatorio.")
            self.name_entry.focus()
            return
        if not path:
            messagebox.showwarning("Dati Mancanti", "Il percorso dello script è obbligatorio.")
            self.path_entry.focus()
            return
            
        self.result = {
            "name": name,
            "description": self.desc_var.get().strip(),
            "path": path,
            "excel_path": self.excel_path_var.get().strip(),
            "tab": self.tab_var.get(),
            "notes": self.notes_textbox.get("1.0", "end-1c").strip(),
            "group": self.group_var.get().strip()
        }
        self.destroy()

class SelectTabDialog(ctk.CTkToplevel):
    """Dialogo generico per selezione tab."""
    def __init__(self, parent, tab_names, title, prompt):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x180")
        self.transient(parent)
        self.resizable(False, False)
        self.result = None

        ctk.CTkLabel(self, text=prompt, wraplength=350).pack(padx=20, pady=20)
        self.tab_var = ctk.StringVar(value=tab_names[0] if tab_names else "")
        self.tab_menu = ctk.CTkOptionMenu(self, variable=self.tab_var, values=tab_names, width=250)
        self.tab_menu.pack(padx=20, pady=5)

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=20)
        ctk.CTkButton(btn_frame, text="Conferma", command=self.on_ok).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Annulla", command=self.destroy, fg_color="gray").pack(side="left", padx=10)
        self.grab_set()

    def on_ok(self):
        self.result = self.tab_var.get()
        self.destroy()

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{VERSION}")
        self.geometry("900x650")
        self.minsize(800, 600)
        
        # Tema
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Dati
        self.scripts = []
        self.config = {}
        self.user_tab_names = []
        self.running_processes = {}  # {index: subprocess.Popen}
        
        self.load_configuration_and_data()
        self._setup_ui()
        self._create_menubar()
        
        # Caricamento iniziale
        self.refresh_script_list()

    def load_configuration_and_data(self):
        """Carica config e dati con gestione errori robusta."""
        # Config
        if not os.path.exists(CONFIG_FILE):
            self.config = {"tabs": DEFAULT_TABS}
            self._save_json(CONFIG_FILE, self.config)
        else:
            try:
                with open(CONFIG_FILE, "r", encoding='utf-8') as f:
                    self.config = json.load(f)
            except Exception as e:
                messagebox.showerror("Errore Config", f"Configurazione corrotta: {e}\nRipristino default.")
                self.config = {"tabs": DEFAULT_TABS}
        
        self.user_tab_names = self.config.get("tabs", DEFAULT_TABS)

        # Scripts
        if not os.path.exists(DATA_FILE):
            self.scripts = []
        else:
            try:
                with open(DATA_FILE, "r", encoding='utf-8') as f:
                    self.scripts = json.load(f)
                
                # Normalizzazione dati mancanti
                dirty = False
                for i, s in enumerate(self.scripts):
                    if "order" not in s: 
                        s["order"] = i
                        dirty = True
                    if "group" not in s:
                        s["group"] = ""
                        dirty = True
                
                if dirty: self._save_json(DATA_FILE, self.scripts)
                self.scripts.sort(key=lambda x: x.get("order", 0))

            except Exception as e:
                messagebox.showerror("Errore Dati", f"File dati corrotto: {e}.\nIl file verrà rinominato in .bak")
                try:
                    shutil.copy(DATA_FILE, DATA_FILE + ".bak")
                except: pass
                self.scripts = []

    def _save_json(self, filepath, data):
        """Helper sicuro per salvare JSON con scrittura atomica."""
        try:
            # Scrivi su file temporaneo prima
            dir_name = os.path.dirname(filepath) or "."
            with tempfile.NamedTemporaryFile("w", delete=False, dir=dir_name, encoding='utf-8', suffix=".tmp") as tmp_file:
                json.dump(data, tmp_file, indent=4, ensure_ascii=False)
                tmp_name = tmp_file.name
            
            # Sostituzione atomica
            shutil.move(tmp_name, filepath)
        except Exception as e:
            messagebox.showerror("Errore Salvataggio", f"Impossibile salvare su {filepath}:\n{e}")
            if 'tmp_name' in locals() and os.path.exists(tmp_name):
                os.unlink(tmp_name)

    def _setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # Main content area

        # 1. Header & Search
        top_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_frame.grid(row=0, column=0, padx=20, pady=(10, 5), sticky="ew")
        top_frame.grid_columnconfigure(0, weight=1)

        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.refresh_script_list())
        self.search_entry = ctk.CTkEntry(top_frame, textvariable=self.search_var, placeholder_text="🔍 Cerca script...", height=35)
        self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))

        self.add_btn = ctk.CTkButton(top_frame, text="+ Nuovo Script", command=self.add_script, width=120, height=35)
        self.add_btn.grid(row=0, column=1)

        # 2. Tab View
        self.tab_view = ctk.CTkTabview(self)
        self.tab_view.grid(row=1, column=0, padx=20, pady=(5, 5), sticky="nsew")
        
        self.scrollable_frames = {}
        # Le tab vengono create dinamicamente in _rebuild_tabs

        # 3. Log Area
        log_frame = ctk.CTkFrame(self, height=150)
        log_frame.grid(row=2, column=0, padx=20, pady=(5, 20), sticky="ew")
        log_frame.grid_columnconfigure(0, weight=1)
        
        log_label = ctk.CTkLabel(log_frame, text="Log Esecuzione", font=("Arial", 10, "bold"))
        log_label.grid(row=0, column=0, sticky="w", padx=5)

        self.log_textbox = ctk.CTkTextbox(log_frame, height=120, activate_scrollbars=True, font=("Consolas", 12))
        self.log_textbox.grid(row=1, column=0, sticky="nsew", padx=5, pady=(0, 5))
        self.log_textbox.configure(state="disabled")

        clear_btn = ctk.CTkButton(log_frame, text="Pulisci", width=60, height=20, command=self.clear_log, font=("Arial", 10))
        clear_btn.grid(row=0, column=0, sticky="e", padx=5, pady=2)

    def _rebuild_tabs(self):
        """Ricostruisce le tab in base a self.user_tab_names. Fix per il refresh dinamico."""
        # 1. Rimuovi tab esistenti
        try:
            # CTkTabview non ha un metodo 'delete_all', dobbiamo iterare
            existing_tabs = list(self.scrollable_frames.keys())
            for tab in existing_tabs:
                try:
                    self.tab_view.delete(tab)
                except: pass # Potrebbe non esistere più
        except Exception as e:
            print(f"Warn: errore pulizia tab: {e}")

        self.scrollable_frames = {}
        
        # 2. Crea Tab "Generale" (sempre presente)
        all_tabs = ["Generale"] + self.user_tab_names
        
        for tab_name in all_tabs:
            try:
                self.tab_view.add(tab_name)
                tab_frame = self.tab_view.tab(tab_name)
                tab_frame.grid_columnconfigure(0, weight=1)
                tab_frame.grid_rowconfigure(0, weight=1)
                
                scroll_frame = ctk.CTkScrollableFrame(tab_frame, fg_color="transparent")
                scroll_frame.grid(row=0, column=0, sticky="nsew")
                scroll_frame.grid_columnconfigure(0, weight=1)
                
                self.scrollable_frames[tab_name] = scroll_frame
            except ValueError:
                # Tab già esistente (raro se puliamo bene, ma sicurezza)
                pass

    def _create_menubar(self):
        menubar = Menu(self)
        
        # File
        file_menu = Menu(menubar, tearoff=0)
        file_menu.add_command(label="Salva Dati", command=lambda: self._save_json(DATA_FILE, self.scripts))
        file_menu.add_separator()
        file_menu.add_command(label="Esci", command=self.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        # Configurazione (Tab Management)
        tab_menu = Menu(menubar, tearoff=0)
        tab_menu.add_command(label="Nuova Scheda...", command=self._add_tab)
        tab_menu.add_command(label="Rinomina Scheda...", command=self._rename_tab)
        tab_menu.add_command(label="Elimina Scheda...", command=self._delete_tab)
        menubar.add_cascade(label="Schede", menu=tab_menu)

        # Help
        help_menu = Menu(menubar, tearoff=0)
        help_menu.add_command(label="Info", command=lambda: messagebox.showinfo("Info", f"{APP_NAME}\nVersione: {VERSION}"))
        menubar.add_cascade(label="?", menu=help_menu)

        self.configure(menu=menubar)

    # --- LOGICA DATI E VISUALIZZAZIONE ---

    def refresh_script_list(self, rebuild_tabs_structure=True):
        if rebuild_tabs_structure:
            self._rebuild_tabs()

        search_term = self.search_var.get().lower().strip()
        
        # Filtro
        filtered_scripts = []
        for s in self.scripts:
            if search_term:
                if (search_term in s.get("name", "").lower() or 
                    search_term in s.get("description", "").lower() or
                    search_term in s.get("group", "").lower()):
                    filtered_scripts.append(s)
            else:
                filtered_scripts.append(s)

        # Raggruppamento per Tab
        scripts_by_tab = collections.defaultdict(list)
        for script in filtered_scripts:
            tab = script.get("tab", "Schede")
            if tab not in self.user_tab_names: 
                tab = "Schede" # Fallback
                if tab not in self.user_tab_names and self.user_tab_names:
                    tab = self.user_tab_names[0] # Super fallback
            
            scripts_by_tab[tab].append(script)
            scripts_by_tab["Generale"].append(script) # Tutti in generale

        # Render
        for tab_name, frame in self.scrollable_frames.items():
            # Pulisci frame
            for widget in frame.winfo_children(): widget.destroy()
            
            tab_scripts = scripts_by_tab.get(tab_name, [])
            self._render_scripts_in_frame(frame, tab_scripts)

    def _render_scripts_in_frame(self, frame, scripts_list):
        # Ordina per gruppo poi per ordine personalizzato
        scripts_list.sort(key=lambda s: (s.get("group", "").lower(), s.get("order", 0)))
        
        # Raggruppa visivamente
        grouped = itertools.groupby(scripts_list, key=lambda s: s.get("group", ""))
        
        for group_name, group_iter in grouped:
            group_items = list(group_iter)
            
            # Header gruppo (solo se c'è un nome gruppo o se non siamo nella tab Generale con mix)
            if group_name:
                header = ctk.CTkLabel(frame, text=group_name.upper(), font=("Arial", 11, "bold"), text_color="gray70", anchor="w")
                header.pack(fill="x", padx=10, pady=(15, 2))
                ctk.CTkFrame(frame, height=2, fg_color="gray40").pack(fill="x", padx=10, pady=(0, 5))

            for script in group_items:
                self._create_script_card(frame, script)

    def _create_script_card(self, parent, script):
        real_index = self.scripts.index(script)
        
        # Validazione percorso
        path = script.get("path", "")
        exists = os.path.exists(path) if path else False

        card = ctk.CTkFrame(parent)
        if not exists:
            card.configure(border_color="#FF5555", border_width=2) # Bordo rosso se manca

        card.pack(fill="x", padx=5, pady=4)
        
        # Colonna Info
        info_col = ctk.CTkFrame(card, fg_color="transparent")
        info_col.pack(side="left", fill="both", expand=True, padx=10, pady=5)
        
        name_text = script["name"]
        if not exists:
            name_text += " ⚠️ (File mancante)"

        name_lbl = ctk.CTkLabel(info_col, text=name_text, font=("Arial", 13, "bold"))
        if not exists:
            name_lbl.configure(text_color=("#FF5555", "#FF5555")) # Rosso in entrambi i temi
        
        name_lbl.pack(anchor="w")
        
        if script.get("description"):
            ctk.CTkLabel(info_col, text=script["description"], font=("Arial", 11), text_color="gray").pack(anchor="w")
        
        last_run = script.get("last_executed", "Mai")
        ctk.CTkLabel(info_col, text=f"Ultimo avvio: {last_run}", font=("Arial", 9), text_color="gray60").pack(anchor="w", pady=(2,0))

        # Colonna Azioni
        action_col = ctk.CTkFrame(card, fg_color="transparent")
        action_col.pack(side="right", padx=10, pady=5)

        # Icone/Bottoni compatti
        if script.get("notes"):
            ctk.CTkButton(action_col, text="📝", width=30, command=lambda s=script: self._show_notes(s)).pack(side="left", padx=2)
        
        if script.get("excel_path"):
            ctk.CTkButton(action_col, text="📊 Excel", width=60, fg_color="green", command=lambda p=script["excel_path"]: self.open_excel(p)).pack(side="left", padx=2)

        # Start/Stop Button
        if real_index in self.running_processes:
            ctk.CTkButton(action_col, text="⏹ Stop", width=60, fg_color="#AA0000", hover_color="#880000", command=lambda i=real_index: self.stop_script(i)).pack(side="left", padx=5)
        else:
            ctk.CTkButton(action_col, text="▶ Avvia", width=60, command=lambda i=real_index: self.launch_script(i)).pack(side="left", padx=5)
        
        # Menu Altro
        settings_btn = ctk.CTkButton(action_col, text="⚙", width=30, fg_color="gray30", command=lambda i=real_index: self.edit_script(i))
        settings_btn.pack(side="left", padx=2)
        
        del_btn = ctk.CTkButton(action_col, text="🗑", width=30, fg_color="#AA0000", hover_color="#880000", command=lambda i=real_index: self.delete_script(i))
        del_btn.pack(side="left", padx=2)

    # --- LOGICA ESECUZIONE ---

    def _log(self, msg):
        def _write():
            self.log_textbox.configure(state="normal")
            self.log_textbox.insert("end", msg)
            self.log_textbox.see("end")
            self.log_textbox.configure(state="disabled")
        self.after(0, _write)

    def launch_script(self, index):
        if index in self.running_processes:
            messagebox.showinfo("Info", "Lo script è già in esecuzione.")
            return

        script = self.scripts[index]
        path = script.get("path", "")
        name = script.get("name", "Unknown")

        if not path or not os.path.exists(path):
            self._log(f"❌ ERRORE: File non trovato: {path}\n")
            return

        # Aggiorna timestamp
        script["last_executed"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._save_json(DATA_FILE, self.scripts)
        
        # Aggiorna UI immediato per mostrare stato "Running" (se gestito in refresh) o log
        self._log(f"🚀 Avvio: {name} ({path})...\n")

        # Thread separato per non bloccare UI
        threading.Thread(target=self._run_process_thread, args=(path, name, index), daemon=True).start()
        
        # Refresh UI per mostrare pulsante Stop
        self.after(100, lambda: self.refresh_script_list(rebuild_tabs_structure=False))

    def stop_script(self, index):
        if index in self.running_processes:
            proc = self.running_processes[index]
            try:
                self._log(f"🛑 Richiesto arresto forzato (incluso background) per: {self.scripts[index]['name']}...\n")
                
                # Usiamo taskkill su Windows per uccidere l'intero albero dei processi (/T) in modo forzato (/F)
                # Questo risolve il problema dei processi orfani lanciati da script batch o cmd.
                subprocess.run(
                    ["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                
                # Fallback di sicurezza: se taskkill fallisce o non siamo su Windows (improbabile dato il contesto)
                if proc.poll() is None:
                    proc.terminate()
                    
            except Exception as e:
                self._log(f"❌ Errore durante l'arresto: {e}\n")
        else:
            messagebox.showinfo("Info", "Lo script non risulta in esecuzione.")
            self.refresh_script_list(rebuild_tabs_structure=False)

    def _run_process_thread(self, path, name, index):
        try:
            # FIX: Gestione encoding Windows e quoting automatico
            # Usiamo shell=True con path tra virgolette se necessario, o meglio una lista diretta se eseguibile.
            # Per .bat e .vbs su windows, 'cmd /c' è sicuro.
            
            # Determina encoding console sistema
            sys_encoding = locale.getpreferredencoding()
            
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW # Nasconde finestra cmd pop-up

            # Costruzione comando sicura
            # Se è un .bat, .cmd o .vbs, meglio chiamarli tramite cmd /c "path"
            cmd = ['cmd', '/c', path]
            
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL, # Previene hang se lo script chiede input
                text=True,
                encoding=sys_encoding, # Encoding corretto per evitare crash di decodifica
                errors='replace',      # Sostituisci caratteri strani invece di crashare
                startupinfo=startupinfo,
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            # Registra processo
            self.running_processes[index] = process

            for line in iter(process.stdout.readline, ''):
                self._log(f"[{name}] {line}")
            
            process.stdout.close()
            rc = process.wait()
            
            status_icon = "✅" if rc == 0 else "⚠️"
            # Se rc è negativo (es. -15 o altro), potrebbe essere stato killato
            if rc != 0 and index in self.running_processes:
                 # Se è ancora in lista ma rc != 0, è crashato o finito male.
                 # Se lo abbiamo killato noi, potrebbe essere già uscito.
                 pass

            self._log(f"{status_icon} Terminato: {name} (Codice: {rc})\n--------------------------------------------------\n")

        except Exception as e:
            self._log(f"❌ ECCEZIONE AVVIO {name}: {str(e)}\n")
        
        finally:
            # Cleanup
            if index in self.running_processes:
                del self.running_processes[index]
            
            # Refresh UI (torna verde)
            self.after(0, lambda: self.refresh_script_list(rebuild_tabs_structure=False))

    def open_excel(self, path):
        if not os.path.exists(path):
            messagebox.showerror("Errore", f"File Excel non trovato:\n{path}")
            return
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aprire Excel:\n{e}")

    # --- CRUD SCRIPT ---

    def add_script(self):
        dialog = ScriptDialog(self, self.user_tab_names, title="Nuovo Script")
        self.wait_window(dialog)
        if dialog.result:
            new_script = dialog.result
            new_script["order"] = len(self.scripts)
            new_script["last_executed"] = "Mai"
            self.scripts.append(new_script)
            self._save_json(DATA_FILE, self.scripts)
            self.refresh_script_list(rebuild_tabs_structure=False)

    def edit_script(self, index):
        script = self.scripts[index]
        dialog = ScriptDialog(self, self.user_tab_names, title="Modifica Script", script_data=script)
        self.wait_window(dialog)
        if dialog.result:
            # Mantieni campi non modificabili dal dialog (order, last_executed)
            dialog.result["order"] = script.get("order", 0)
            dialog.result["last_executed"] = script.get("last_executed", "Mai")
            
            self.scripts[index] = dialog.result
            self._save_json(DATA_FILE, self.scripts)
            self.refresh_script_list(rebuild_tabs_structure=False)

    def delete_script(self, index):
        if index in self.running_processes:
            messagebox.showwarning("Azione Negata", "Impossibile eliminare uno script mentre è in esecuzione.\nArrestalo prima.")
            return

        if messagebox.askyesno("Conferma", "Eliminare questo script?"):
            self.scripts.pop(index)
            self._save_json(DATA_FILE, self.scripts)
            self.refresh_script_list(rebuild_tabs_structure=False)

    def _show_notes(self, script):
        dialog = ctk.CTkToplevel(self)
        dialog.title(f"Note: {script['name']}")
        dialog.geometry("400x300")
        txt = ctk.CTkTextbox(dialog, wrap="word")
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.insert("1.0", script.get("notes", ""))
        txt.configure(state="disabled")

    def clear_log(self):
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

    # --- GESTIONE TAB ---

    def _add_tab(self):
        dialog = ctk.CTkInputDialog(text="Nome nuova scheda:", title="Nuova Scheda")
        new_name = dialog.get_input()
        if new_name:
            new_name = new_name.strip()
            if new_name in self.user_tab_names:
                messagebox.showerror("Errore", "Nome scheda già esistente.")
                return
            self.user_tab_names.append(new_name)
            self.config["tabs"] = self.user_tab_names
            self._save_json(CONFIG_FILE, self.config)
            self.refresh_script_list(rebuild_tabs_structure=True) # Refresh completo

    def _rename_tab(self):
        if not self.user_tab_names: return
        sel_dlg = SelectTabDialog(self, self.user_tab_names, "Rinomina", "Scegli scheda da rinominare:")
        self.wait_window(sel_dlg)
        old_name = sel_dlg.result
        if old_name:
            name_dlg = ctk.CTkInputDialog(text=f"Nuovo nome per '{old_name}':", title="Rinomina")
            new_name = name_dlg.get_input()
            if new_name:
                new_name = new_name.strip()
                if new_name in self.user_tab_names and new_name != old_name:
                    messagebox.showerror("Errore", "Nome già in uso.")
                    return
                
                # Aggiorna lista tab
                idx = self.user_tab_names.index(old_name)
                self.user_tab_names[idx] = new_name
                self.config["tabs"] = self.user_tab_names
                self._save_json(CONFIG_FILE, self.config)

                # Aggiorna script associati
                for s in self.scripts:
                    if s.get("tab") == old_name:
                        s["tab"] = new_name
                self._save_json(DATA_FILE, self.scripts)
                
                self.refresh_script_list(rebuild_tabs_structure=True)

    def _delete_tab(self):
        if not self.user_tab_names: return
        sel_dlg = SelectTabDialog(self, self.user_tab_names, "Elimina", "Scegli scheda da eliminare:")
        self.wait_window(sel_dlg)
        target = sel_dlg.result
        if target:
            if not messagebox.askyesno("Sicuro?", f"Eliminare '{target}'?\nGli script verranno spostati nella prima scheda disponibile."):
                return
            
            self.user_tab_names.remove(target)
            self.config["tabs"] = self.user_tab_names
            self._save_json(CONFIG_FILE, self.config)

            # Sposta script orfani
            fallback = self.user_tab_names[0] if self.user_tab_names else "Schede"
            for s in self.scripts:
                if s.get("tab") == target:
                    s["tab"] = fallback
            self._save_json(DATA_FILE, self.scripts)

            self.refresh_script_list(rebuild_tabs_structure=True)

if __name__ == "__main__":
    app = App()
    app.mainloop()
