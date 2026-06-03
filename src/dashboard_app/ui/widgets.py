import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFrame, QLabel, 
    QPushButton, QTextEdit, QMessageBox
)
from PySide6.QtGui import QFont
from ..models import ScriptModel

class ScriptCard(QFrame):
    def __init__(self, script: ScriptModel, index: int, is_running: bool, controller):
        super().__init__()
        self.script = script
        self.index = index
        self.controller = controller
        self._setup_ui(is_running)

    def _setup_ui(self, is_running):
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet("QFrame { background-color: white; border: 1px solid #ddd; border-radius: 4px; } QLabel { border: none; }")
        
        layout = QHBoxLayout(self)
        info_layout = QVBoxLayout()
        
        exists = os.path.exists(self.script.path)
        title = QLabel(f"<b>{self.script.name}</b>" + (" ⚠️" if not exists else ""))
        if not exists: title.setStyleSheet("color: red;")
        info_layout.addWidget(title)
        
        if self.script.description:
            info_layout.addWidget(QLabel(self.script.description, styleSheet="color: #666; font-size: 11px;"))
        
        info_layout.addWidget(QLabel(f"Ultimo avvio: {self.script.last_executed}", styleSheet="color: #999; font-size: 10px;"))
        layout.addLayout(info_layout, stretch=1)

        actions = QHBoxLayout()
        if self.script.notes:
            btn_notes = QPushButton("📝")
            btn_notes.setFixedWidth(30)
            btn_notes.clicked.connect(lambda: QMessageBox.information(self, "Note", self.script.notes))
            actions.addWidget(btn_notes)

        if self.script.excel_path:
            btn_xl = QPushButton("📊 Excel")
            btn_xl.setStyleSheet("background-color: #217346; color: white;")
            btn_xl.clicked.connect(lambda: os.startfile(self.script.excel_path) if os.path.exists(self.script.excel_path) else None)
            actions.addWidget(btn_xl)

        btn_run = QPushButton("⏹ Stop" if is_running else "▶ Avvia")
        btn_run.setStyleSheet(f"background-color: {'#d32f2f' if is_running else '#1976d2'}; color: white; font-weight: bold; min-width: 60px;")
        btn_run.clicked.connect(lambda: self.controller.toggle_script(self.index))
        actions.addWidget(btn_run)

        btn_edit = QPushButton("⚙")
        btn_edit.setFixedWidth(30)
        btn_edit.clicked.connect(lambda: self.controller.edit_script(self.index))
        actions.addWidget(btn_edit)

        btn_del = QPushButton("🗑")
        btn_del.setFixedWidth(30)
        btn_del.setStyleSheet("color: red;")
        btn_del.clicked.connect(lambda: self.controller.delete_script(self.index))
        actions.addWidget(btn_del)

        layout.addLayout(actions)

class LogArea(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 5, 0, 0)
        
        hdr = QHBoxLayout()
        hdr.addWidget(QLabel("<b>Log Esecuzione</b>"))
        btn_clear = QPushButton("Pulisci")
        btn_clear.setFixedWidth(60)
        btn_clear.clicked.connect(lambda: self.text.clear())
        hdr.addStretch()
        hdr.addWidget(btn_clear)
        layout.addLayout(hdr)

        self.text = QTextEdit()
        self.text.setReadOnly(True)
        self.text.setFont(QFont("Consolas", 10))
        self.text.setStyleSheet("background-color: #f9f9f9; border: 1px solid #ccc;")
        self.text.setMaximumHeight(150)
        layout.addWidget(self.text)

    def append(self, message: str):
        self.text.append(message)
        self.text.ensureCursorVisible()
