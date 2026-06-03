import subprocess
from typing import Dict
from PySide6.QtCore import QObject, Signal, QProcess
from .models import ScriptModel

class ScriptProcess(QProcess):
    """Gestisce l'esecuzione di un singolo processo script."""
    output_ready = Signal(str, str) # name, text
    finished_with_code = Signal(str, int) # name, exit_code

    def __init__(self, script: ScriptModel):
        super().__init__()
        self.script = script
        self.setProcessChannelMode(QProcess.MergedChannels)
        self.readyReadStandardOutput.connect(self._handle_output)
        self.finished.connect(self._handle_finished)

    def start_script(self):
        self.start("cmd", ["/c", self.script.path])

    def _handle_output(self):
        data = self.readAllStandardOutput().data().decode('cp1252', errors='replace')
        self.output_ready.emit(self.script.name, data.strip())

    def _handle_finished(self, exit_code, exit_status):
        self.finished_with_code.emit(self.script.name, exit_code)

class ProcessManager(QObject):
    """Gestisce tutti i processi attivi."""
    log_requested = Signal(str)

    def __init__(self):
        super().__init__()
        self.active_processes: Dict[int, ScriptProcess] = {}

    def is_running(self, index: int) -> bool:
        return index in self.active_processes

    def launch(self, index: int, script: ScriptModel, on_finished_callback):
        if self.is_running(index):
            return

        process = ScriptProcess(script)
        process.output_ready.connect(lambda name, text: self.log_requested.emit(f"[{name}] {text}"))
        process.finished_with_code.connect(lambda name, code: self._on_process_finished(index, name, code, on_finished_callback))
        
        self.active_processes[index] = process
        process.start_script()
        self.log_requested.emit(f"🚀 Avvio: {script.name}...")

    def stop(self, index: int):
        if self.is_running(index):
            process = self.active_processes[index]
            self.log_requested.emit(f"🛑 Arresto forzato: {process.script.name}...")
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(process.processId())], 
                         creationflags=subprocess.CREATE_NO_WINDOW)

    def _on_process_finished(self, index: int, name: str, code: int, callback):
        if index in self.active_processes:
            del self.active_processes[index]
        self.log_requested.emit(f"✅ Terminato: {name} (Codice: {code})")
        callback()
