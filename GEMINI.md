# Dashboard Avvio Script

## Panoramica del Progetto
**Dashboard Avvio Script** è un'applicazione desktop sviluppata in Python che funge da launcher centralizzato per script di automazione (Batch, VBS) e file Excel. Utilizza **CustomTkinter** per fornire un'interfaccia grafica moderna e intuitiva, permettendo all'utente di organizzare i propri strumenti in schede (Tab) e gruppi personalizzabili.

### Caratteristiche Principali
- **Organizzazione a Schede:** Creazione, rinomina ed eliminazione dinamica di schede per categorizzare gli script (es. Contabilità, Ufficio).
- **Esecuzione Monitorata:** Lancia script `.bat` e `.vbs` catturando l'output (stdout/stderr) in un pannello di log integrato in tempo reale.
- **Gestione Excel:** Collegamenti rapidi per aprire file `.xlsx` o `.xls`.
- **Persistenza Dati:** Salvataggio automatico di configurazioni e script su file JSON.
- **Note e Metadati:** Possibilità di aggiungere note dettagliate per ogni script e tracciamento dell'ultimo avvio.

## Architettura e File Chiave

La struttura del progetto è piatta e contenuta nella root directory.

- **`dashboard.py`**: Il cuore dell'applicazione. Contiene:
  - `App`: La classe principale che gestisce la GUI, il loop degli eventi e la logica di business.
  - `ScriptDialog`: Finestra modale per aggiungere/modificare script.
  - `SelectTabDialog`: Finestra di utility per la gestione delle schede.
  - Logica di esecuzione processi tramite `subprocess`.
- **`start_dashboard.bat`**: Script di avvio per Windows. Installa silenziosamente le dipendenze mancanti ed esegue l'applicazione.
- **`config.json`**: Memorizza la configurazione globale, principalmente l'elenco delle schede (Tabs) attive.
- **`data.json`**: Database JSON che contiene l'array di oggetti "script" con le relative proprietà (percorso, nome, descrizione, gruppo, note, timestamp esecuzione).
- **`requirements.txt`**: Elenco delle dipendenze Python (`customtkinter`, `pyautogui`).
- **`test.db`**: File SQLite presente ma **non utilizzato** attivamente nel codice principale (`dashboard.py`).

## Istruzioni per l'Uso

### Prerequisiti
- Python 3.x installato e aggiunto al PATH di sistema.
- Sistema Operativo Windows (per il supporto `.bat`/`.vbs` e `os.startfile`).

### Avvio
Eseguire il file batch:
```cmd
start_dashboard.bat
```
Questo comando verificherà e installerà le dipendenze necessarie prima di lanciare la GUI.

### Sviluppo e Manutenzione
- **Librerie UI:** Il progetto usa `customtkinter`. Assicurarsi di seguire le convenzioni di questo framework per modifiche alla UI.
- **Dipendenze:** `pyautogui` è listato in `requirements.txt` ma non importato in `dashboard.py`. Verificare se è necessario per gli script esterni o se può essere rimosso.
- **Logging:** L'output dei processi figli viene letto in un thread separato (`_read_process_output`) per non bloccare la UI principale.

## Convenzioni di Codice
- **Stile:** Codice Python standard.
- **Encoding:** I file JSON vengono letti/scritti con encoding `utf-8`.
- **Paths:** Si raccomanda l'uso di percorsi assoluti per gli script configurati per evitare errori di "file not found" se la working directory cambia.
