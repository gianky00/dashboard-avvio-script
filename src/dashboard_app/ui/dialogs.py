import os
from typing import Optional, List
from PySide6.QtWidgets import (
    QDialog, QFormLayout, QLineEdit, QComboBox, QHBoxLayout, 
    QPushButton, QTextEdit, QFileDialog
)
from ..models import ScriptModel

class ScriptDialog(QDialog):
    def __init__(self, parent, tab_names: List[str], script: Optional[ScriptModel] = None):
        super().__init__(parent)
        self.setWindowTitle("Modifica Script" if script else "Nuovo Script")
        self.setMinimumWidth(500)
        self.tab_names = tab_names
        self.script = script or ScriptModel(name="", path="")
        self._setup_ui()

    def _setup_ui(self):
        layout = QFormLayout(self)
        self.name_edit = QLineEdit(self.script.name)
        self.desc_edit = QLineEdit(self.script.description)
        self.tab_combo = QComboBox()
        self.tab_combo.addItems(self.tab_names)
        self.tab_combo.setCurrentText(self.script.tab)
        self.group_edit = QLineEdit(self.script.group)
        self.path_edit = QLineEdit(self.script.path)
        self.excel_edit = QLineEdit(self.script.excel_path)
        self.notes_edit = QTextEdit(self.script.notes)
        self.notes_edit.setMaximumHeight(80)

        layout.addRow("Nome:", self.name_edit)
        layout.addRow("Descrizione:", self.desc_edit)
        layout.addRow("Scheda:", self.tab_combo)
        layout.addRow("Gruppo:", self.group_edit)
        
        path_box = QHBoxLayout()
        path_box.addWidget(self.path_edit)
        btn_path = QPushButton("...")
        btn_path.setFixedWidth(30)
        btn_path.clicked.connect(self._browse_path)
        path_box.addWidget(btn_path)
        layout.addRow("Percorso Script:", path_box)

        excel_box = QHBoxLayout()
        excel_box.addWidget(self.excel_edit)
        btn_excel = QPushButton("...")
        btn_excel.setFixedWidth(30)
        btn_excel.clicked.connect(self._browse_excel)
        excel_box.addWidget(btn_excel)
        layout.addRow("File Excel:", excel_box)
        
        layout.addRow("Note:", self.notes_edit)

        btns = QHBoxLayout()
        btn_ok = QPushButton("Salva")
        btn_ok.clicked.connect(self.accept)
        btn_ok.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        btn_cancel = QPushButton("Annulla")
        btn_cancel.clicked.connect(self.reject)
        btns.addStretch()
        btns.addWidget(btn_cancel)
        btns.addWidget(btn_ok)
        layout.addRow(btns)

    def _browse_path(self):
        file, _ = QFileDialog.getOpenFileName(self, "Seleziona Script")
        if file: self.path_edit.setText(os.path.normpath(file))

    def _browse_excel(self):
        file, _ = QFileDialog.getOpenFileName(self, "Seleziona Excel")
        if file: self.excel_edit.setText(os.path.normpath(file))

    def get_model(self) -> ScriptModel:
        return ScriptModel(
            name=self.name_edit.text().strip(),
            path=self.path_edit.text().strip(),
            description=self.desc_edit.text().strip(),
            excel_path=self.excel_edit.text().strip(),
            tab=self.tab_combo.currentText(),
            group=self.group_edit.text().strip(),
            notes=self.notes_edit.toPlainText().strip(),
            last_executed=self.script.last_executed,
            order=self.script.order
        )
