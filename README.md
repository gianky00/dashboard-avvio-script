# Automation Script Control Dashboard

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=flat-square&logo=python&logoColor=white)
![PySide6](https://img.shields.io/badge/GUI-PySide6%20(Qt6)-41CD52?style=flat-square&logo=qt&logoColor=white)
![Build](https://img.shields.io/badge/Packaging-Poetry-60A5FA?style=flat-square)

Applicazione desktop sviluppata in **Python** e **PySide6 (Qt6)** che funge da pannello di controllo unificato per catalogare, configurare e avviare gli script e gli applicativi di automazione utilizzati nelle attività tecniche di cantiere.

---

## Funzionalità Principali

- **Organizzazione Multi-Dominio a Tab:** Categorizzazione ordinata degli strumenti operativi per area funzionale (*Schede*, *Contabilità*, *Programmazione*, *Report Giornaliere*, *Strumenti Campione*).
- **Gestione e Monitoraggio Processi (`process.py`):** Avvio asincrono degli script Python/Batch in sottoprocessi dedicati senza bloccare la GUI principale, con tracciamento dello stato di esecuzione.
- **Configurazione Dinamica da Interfaccia (`dialogs.py` & `storage.py`):** Aggiunta, modifica, riordino e rimozione dei collegamenti agli script direttamente tramite finestre di dialogo Qt, con persistenza automatica su file JSON (`data.json` / `config.json`).
- **Design Moderno Fusion:** Interfaccia pulita, responsiva e ad alto contrasto progettata per un utilizzo quotidiano rapido in ufficio tecnico.

---

## Struttura del Progetto

```text
dashboard-avvio-script/
├── src/dashboard_app/
│   ├── __init__.py         # Inizializzazione QApplication e stile Fusion
│   ├── __main__.py         # Entry point del modulo
│   ├── config.py           # Definizione parametri applicativi e tab predefinite
│   ├── models.py           # Modelli dati per gli script e le categorie
│   ├── process.py          # Gestione esecuzione sottoprocessi esterni
│   ├── storage.py          # Persistenza e caricamento configurazioni JSON
│   ├── main_window.py      # Finestra principale e orchestrazione della dashboard
│   └── ui/
│       ├── dialogs.py      # Finestre di dialogo per aggiunta/modifica script
│       └── widgets.py      # Card e pulsanti personalizzati per l'avvio rapido
├── pyproject.toml          # Configurazione Poetry e dipendenze
└── start_dashboard.bat     # Script di avvio rapido per ambienti Windows
```

---

## Installazione e Avvio

```powershell
# Installazione delle dipendenze con Poetry
poetry install

# Avvio della dashboard
poetry run dashboard
# oppure tramite modulo Python:
poetry run python -m dashboard_app
```

Su Windows è possibile avviare direttamente l'applicazione tramite `start_dashboard.bat`.

---

## Autore
Sviluppato da **Giancarlo Allegretti** — *Industrial Automation, Instrumentation & AI Software Engineering*.
