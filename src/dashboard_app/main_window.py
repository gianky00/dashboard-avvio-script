import sys
from datetime import datetime
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLineEdit, QPushButton, QTabWidget, QScrollArea, QMessageBox
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPalette, QColor

from .config import config_env
from .storage import Storage, ScriptRepository
from .process import ProcessManager
from .ui.dialogs import ScriptDialog
from .ui.widgets import ScriptCard, LogArea

class DashboardWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{config_env.APP_NAME} v{config_env.VERSION}")
        self.resize(1000, 750)
        
        # Inizializza componenti SRP
        self.repo = ScriptRepository(config_env.DATA_FILE)
        self.pm = ProcessManager()
        self.user_tabs = Storage.load_json(config_env.CONFIG_FILE, {"tabs": config_env.DEFAULT_TABS})["tabs"]
        
        self._setup_ui()
        self._connect_signals()
        self.refresh()

    def _setup_ui(self):
        self.setPalette(self._light_palette())
        
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # Search & Add
        top = QHBoxLayout()
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔍 Cerca script...")
        self.search_edit.textChanged.connect(self.refresh)
        top.addWidget(self.search_edit)

        btn_add = QPushButton("+ Nuovo Script")
        btn_add.clicked.connect(self.add_script)
        btn_add.setStyleSheet("background-color: #0078d4; color: white; padding: 6px 15px;")
        top.addWidget(btn_add)
        layout.addLayout(top)

        # Tabs
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        # Log
        self.log_area = LogArea()
        layout.addWidget(self.log_area)

        # Menu
        self._setup_menu()

    def _light_palette(self):
        p = QPalette()
        p.setColor(QPalette.Window, QColor(245, 245, 245))
        p.setColor(QPalette.Base, Qt.white)
        p.setColor(QPalette.Text, Qt.black)
        return p

    def _setup_menu(self):
        mb = self.menuBar()
        f = mb.addMenu("File")
        f.addAction("Salva", self.repo.save)
        f.addSeparator()
        f.addAction("Esci", self.close)

    def _connect_signals(self):
        self.pm.log_requested.connect(self.log_area.append)

    def refresh(self):
        search = self.search_edit.text().lower().strip()
        current_idx = self.tabs.currentIndex()
        current_name = self.tabs.tabText(current_idx) if current_idx >= 0 else None
        
        self.tabs.clear()
        all_tabs = ["Generale"] + self.user_tabs
        layouts = {}

        for name in all_tabs:
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            content = QWidget()
            content_layout = QVBoxLayout(content)
            content_layout.setAlignment(Qt.AlignTop)
            scroll.setWidget(content)
            self.tabs.addTab(scroll, name)
            layouts[name] = content_layout

        for i, script in enumerate(self.repo.scripts):
            if search and not (search in script.name.lower() or search in script.description.lower()):
                continue
            
            # Generale
            layouts["Generale"].addWidget(ScriptCard(script, i, self.pm.is_running(i), self))
            
            # Specifica
            if script.tab in layouts and script.tab != "Generale":
                layouts[script.tab].addWidget(ScriptCard(script, i, self.pm.is_running(i), self))

        # Ripristina tab
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == current_name:
                self.tabs.setCurrentIndex(i)
                break

    # Controller Actions
    def toggle_script(self, index: int):
        if self.pm.is_running(index):
            self.pm.stop(index)
        else:
            script = self.repo.scripts[index]
            script.last_executed = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.repo.save()
            self.pm.launch(index, script, self.refresh)
        self.refresh()

    def add_script(self):
        dlg = ScriptDialog(self, self.user_tabs)
        if dlg.exec():
            self.repo.add(dlg.get_model())
            self.refresh()

    def edit_script(self, index: int):
        dlg = ScriptDialog(self, self.user_tabs, self.repo.scripts[index])
        if dlg.exec():
            self.repo.update(index, dlg.get_model())
            self.refresh()

    def delete_script(self, index: int):
        if QMessageBox.question(self, "Elimina", "Sei sicuro?") == QMessageBox.Yes:
            self.repo.remove(index)
            self.refresh()
