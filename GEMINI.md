# ♾️ DASHBOARD AVVIO SCRIPT: ARCHITECTURAL CONTEXT

## 📌 PROJECT OVERVIEW
Questa è una **Dashboard GUI** sviluppata in Python per la gestione centralizzata e l'esecuzione di script di automazione (Batch, VBS, CMD, PowerShell, Python). Permette di organizzare gli script in schede e gruppi, monitorarne l'esecuzione tramite log in tempo reale e aprire file Excel correlati.

### 🛠️ TECH STACK
- **Language:** Python 3.10 - 3.14
- **Dependency Management:** [Poetry](https://python-poetry.org/)
- **GUI Framework:** [PySide6](https://doc.qt.io/qtforpython-6/) (Qt for Python)
- **Aesthetic:** Tema Chiaro nativo con stile "Fusion".
- **Execution:** `QProcess` (asincrono) con monitoraggio log in tempo reale.
- **Persistence:** File JSON (`config.json`, `data.json`).

---

## 🚀 BUILDING AND RUNNING

### 📋 Prerequisiti
- Python 3.10+ installato (compatibile con PySide6).
- Poetry installato (`pip install poetry`).

### ⚙️ Installazione Dipendenze
```powershell
poetry install
```

### ▶️ Avvio Applicazione
È possibile avviare la dashboard nei seguenti modi:
1. Tramite Poetry (comando installato):
   ```powershell
   poetry run dashboard
   ```
2. Tramite il file batch aggiornato:
   ```powershell
   .\start_dashboard.bat
   ```

---

## 📂 PROJECT STRUCTURE
- `src/`: Cartella contenente i sorgenti.
  - `dashboard_app/`: Pacchetto principale dell'applicazione.
    - `__init__.py`: Punto di ingresso del pacchetto.
    - `__main__.py`: Supporto per l'esecuzione come modulo (`python -m`).
    - `config.py`: Gestione configurazioni e costanti.
    - `models.py`: Modelli dati (ScriptModel).
    - `storage.py`: Persistenza JSON e Repository.
    - `process.py`: Logica di esecuzione processi (ProcessManager).
    - `main_window.py`: Controller della finestra principale.
    - `ui/`: Componenti dell'interfaccia utente.
      - `dialogs.py`: Finestre di dialogo.
      - `widgets.py`: Widget personalizzati (Card, LogArea).
- `pyproject.toml`: Configurazione di Poetry.
- `start_dashboard.bat`: Launcher ottimizzato per Poetry (avvio silenzioso senza console CMD).

---

## 🛠️ DEVELOPMENT CONVENTIONS

### 🏗️ Architettura SRP (Single Responsibility Principle)
Il progetto è stato scomposto in moduli indipendenti per garantire scalabilità e manutenibilità:
1. **Configurazione:** Isolata in `config.py`.
2. **Dati:** Modelli tipizzati in `models.py` e repository in `storage.py`.
3. **Logica di Processo:** Separata dalla UI in `process.py`, utilizza `QProcess` per l'asincronia.
4. **UI:** Suddivisa tra controller principale (`main_window.py`) e componenti riutilizzabili (`ui/`).

### 📝 Logging
I log vengono visualizzati in un widget `QTextEdit` dedicato. L'output standard degli script viene catturato riga per riga e inviato alla GUI tramite segnali.

### ⚠️ Error Handling
- Verifica visiva dei percorsi (bordo rosso sulle card se lo script non esiste).
- Gestione robusta dell'arresto dei processi tramite `taskkill /F /T`.
