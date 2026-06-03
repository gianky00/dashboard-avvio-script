from dataclasses import dataclass
from typing import List

@dataclass
class AppConfig:
    APP_NAME: str = "Dashboard Avvio Script"
    VERSION: str = "1.3.0"
    DATA_FILE: str = "data.json"
    CONFIG_FILE: str = "config.json"
    DEFAULT_TABS: List[str] = None

    def __post_init__(self):
        if self.DEFAULT_TABS is None:
            self.DEFAULT_TABS = ["Schede", "Contabilità", "Programmazione", "Report Giornaliere", "Strumenti Campione"]

config_env = AppConfig()
