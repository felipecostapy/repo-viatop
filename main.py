"""
Entry point do executável.
Roda o updater primeiro, depois a interface.
"""
import updater
import sys
from PySide6.QtWidgets import QApplication
from interface import UI

if __name__ == "__main__":
    updater.verificar_e_atualizar()

    app = QApplication(sys.argv)
    win = UI()
    win.show()
    sys.exit(app.exec())
