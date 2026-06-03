import os
import json
import tempfile
import shutil
from typing import List
from .models import ScriptModel

class Storage:
    """Gestisce la persistenza dei dati su file JSON."""
    @staticmethod
    def load_json(filepath: str, default: dict) -> dict:
        if not os.path.exists(filepath):
            return default
        try:
            with open(filepath, "r", encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return default

    @staticmethod
    def save_json(filepath: str, data: any):
        try:
            dir_name = os.path.dirname(filepath) or "."
            with tempfile.NamedTemporaryFile("w", delete=False, dir=dir_name, encoding='utf-8', suffix=".tmp") as tmp:
                json.dump(data, tmp, indent=4, ensure_ascii=False)
                tmp_name = tmp.name
            shutil.move(tmp_name, filepath)
        except Exception as e:
            print(f"Errore salvataggio {filepath}: {e}")

class ScriptRepository:
    """Repository per la gestione della lista degli script."""
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.scripts: List[ScriptModel] = []
        self.load()

    def load(self):
        data = Storage.load_json(self.filepath, [])
        self.scripts = [ScriptModel.from_dict(d) for d in data]
        self.scripts.sort(key=lambda x: x.order)

    def save(self):
        Storage.save_json(self.filepath, [s.to_dict() for s in self.scripts])

    def add(self, script: ScriptModel):
        script.order = len(self.scripts)
        self.scripts.append(script)
        self.save()

    def update(self, index: int, script: ScriptModel):
        if 0 <= index < len(self.scripts):
            self.scripts[index] = script
            self.save()

    def remove(self, index: int):
        if 0 <= index < len(self.scripts):
            self.scripts.pop(index)
            # Ricalcola ordini
            for i, s in enumerate(self.scripts):
                s.order = i
            self.save()
