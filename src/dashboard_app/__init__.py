import sys
from PySide6.QtWidgets import QApplication
from .main_window import DashboardWindow

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = DashboardWindow()
    win.show()
    sys.exit(app.exec())
