"""
Entry point do executável.
Roda o updater primeiro, depois carrega a interface externa se disponível.
"""
import sys
import os
from pathlib import Path

# Garante que o PyInstaller encontra os plugins do Qt ao extrair
if getattr(sys, "frozen", False):
    os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = os.path.join(
        sys._MEIPASS, "PySide6", "plugins", "platforms"
    )

import updater
updater.verificar_e_atualizar()

# Após atualização, carrega interface.py da pasta do exe (versão atualizada)
# e não a versão embutida no .exe
pasta = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent

if (pasta / "interface.py").exists():
    sys.path.insert(0, str(pasta))
    # Remove módulos em cache para forçar reload dos arquivos externos
    for mod in ["interface", "gerador"]:
        if mod in sys.modules:
            del sys.modules[mod]

from PySide6.QtWidgets import QApplication
from interface import UI

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = UI()
    win.show()
    sys.exit(app.exec())