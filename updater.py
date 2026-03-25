import sys
import os
import json
import shutil
import subprocess
import urllib.request
import urllib.error
from pathlib import Path

GITHUB_USER    = "felipecostapy"
GITHUB_REPO    = "repo-viatop"
GITHUB_BRANCH  = "master"

CREDENTIALS_DRIVE_ID  = "1ANUx62cswnhOpUmwOlyb-NwF4ji7VI-b"
CREDENTIALS_FILE      = "credentials.json"

ARQUIVOS = [
    "interface.py",
    "gerador.py",
    "planilha.py",
    "logo_agro.png",
    "logo_top.png",
    "ordemagroviav2.xlsx",
    "ordemtopv2.xlsx",
]

VERSION_LOCAL_FILE = "version.txt"
VERSION_URL = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/version.txt"
BASE_URL    = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}"

def _pasta_app():
                                                             
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent

def _ler_versao_local():
    path = _pasta_app() / VERSION_LOCAL_FILE
    if path.exists():
        return path.read_text(encoding="utf-8").strip()
    return "0.0.0"

def _ler_versao_remota():
    try:
        with urllib.request.urlopen(VERSION_URL, timeout=5) as r:
            return r.read().decode("utf-8").strip()
    except Exception:
        return None

def _versao_maior(remota, local):
                                                                      
    def parse(v):
        try:
            return tuple(int(x) for x in v.split("."))
        except Exception:
            return (0, 0, 0)
    return parse(remota) > parse(local)

def _baixar_arquivo(nome):
    url  = f"{BASE_URL}/{nome}"
    dest = _pasta_app() / nome
    tmp  = dest.with_suffix(dest.suffix + ".tmp")
    try:
        urllib.request.urlretrieve(url, tmp)
        shutil.move(str(tmp), str(dest))
        return True
    except Exception as e:
        if tmp.exists():
            tmp.unlink()
        raise Exception(f"Falha ao baixar {nome}: {e}")

def _salvar_versao(versao):
    path = _pasta_app() / VERSION_LOCAL_FILE
    path.write_text(versao, encoding="utf-8")

def _baixar_credentials():
\
\
       
    dest = _pasta_app() / CREDENTIALS_FILE
    if dest.exists():
        return

    tmp = dest.with_suffix(".tmp")
    log = _pasta_app() / "credentials_log.txt"
    try:
        url = f"https://drive.usercontent.google.com/download?id={CREDENTIALS_DRIVE_ID}&export=download&authuser=0"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            conteudo = resp.read()

        conteudo_str = conteudo.decode("utf-8", errors="ignore").strip()

        with open(log, "w", encoding="utf-8") as f:
            f.write(f"Tamanho: {len(conteudo)} bytes\n")
            f.write(f"Primeiros 200 chars:\n{conteudo_str[:200]}\n")

        if conteudo_str.startswith("{"):
            with open(tmp, "w", encoding="utf-8") as f:
                f.write(conteudo_str)
            shutil.move(str(tmp), str(dest))
        else:
            with open(log, "a", encoding="utf-8") as f:
                f.write("\nNao era JSON — nao salvou.\n")

    except Exception as e:
        with open(log, "w", encoding="utf-8") as f:
            f.write(f"Erro: {e}\n")
        if tmp.exists():
            tmp.unlink()

def _perguntar_usuario(versao_local, versao_remota):
\
\
\
       
    from PySide6.QtWidgets import QApplication, QMessageBox
    from PySide6.QtCore import Qt

    app = QApplication.instance() or QApplication(sys.argv)

    msg = QMessageBox()
    msg.setWindowTitle("Atualização disponível")
    msg.setIcon(QMessageBox.Information)
    msg.setText(
        f"Uma nova versão do sistema está disponível!\n\n"
        f"  Versão atual:  {versao_local}\n"
        f"  Nova versão:   {versao_remota}\n\n"
        f"Deseja atualizar agora?"
    )
    msg.setStyleSheet("""
        QMessageBox { background-color: #0f1115; }
        QLabel { color: #e5e7eb; font-size: 13px; }
        QPushButton {
            border-radius: 6px;
            padding: 8px 18px;
            font-weight: bold;
            min-width: 80px;
        }
    """)

    btn_sim = msg.addButton("ATUALIZAR", QMessageBox.AcceptRole)
    btn_sim.setStyleSheet("background-color: #2e7d32; color: white;")

    btn_nao = msg.addButton("AGORA NÃO", QMessageBox.RejectRole)
    btn_nao.setStyleSheet("background-color: #2a2f3a; color: #e5e7eb;")

    msg.exec()
    return msg.clickedButton() == btn_sim

def _mostrar_progresso(total):
                                                 
    from PySide6.QtWidgets import QApplication, QDialog, QVBoxLayout, QLabel, QProgressBar
    from PySide6.QtCore import Qt

    app = QApplication.instance() or QApplication(sys.argv)

    dialog = QDialog()
    dialog.setWindowTitle("Atualizando...")
    dialog.setFixedSize(360, 110)
    dialog.setWindowFlag(Qt.WindowCloseButtonHint, False)
    dialog.setStyleSheet("QDialog { background-color: #0f1115; } QLabel { color: #e5e7eb; font-size: 12px; }")

    layout = QVBoxLayout(dialog)
    layout.setContentsMargins(20, 20, 20, 20)
    layout.setSpacing(10)

    lbl = QLabel("Baixando atualização...")
    bar = QProgressBar()
    bar.setMaximum(total)
    bar.setValue(0)
    bar.setStyleSheet("""
        QProgressBar {
            border: 1px solid #2a2f3a;
            border-radius: 6px;
            background: #1a1d23;
            height: 18px;
        }
        QProgressBar::chunk {
            background-color: #2e7d32;
            border-radius: 6px;
        }
    """)

    layout.addWidget(lbl)
    layout.addWidget(bar)
    dialog.show()
    app.processEvents()

    return dialog, lbl, bar

def _mostrar_erro(mensagem):
    from PySide6.QtWidgets import QApplication, QMessageBox
    app = QApplication.instance() or QApplication(sys.argv)
    msg = QMessageBox()
    msg.setWindowTitle("Erro na atualização")
    msg.setIcon(QMessageBox.Warning)
    msg.setText(f"Não foi possível atualizar:\n\n{mensagem}\n\nO sistema será iniciado normalmente.")
    msg.setStyleSheet("QMessageBox { background-color: #0f1115; } QLabel { color: #e5e7eb; }")
    msg.exec()

def verificar_e_atualizar():
                                                   
    _baixar_credentials()

    versao_local  = _ler_versao_local()
    versao_remota = _ler_versao_remota()

    if not versao_remota:
        return
    if not _versao_maior(versao_remota, versao_local):
        return

    if not _perguntar_usuario(versao_local, versao_remota):
        return

    from PySide6.QtWidgets import QApplication
    app = QApplication.instance() or QApplication(sys.argv)

    dialog, lbl, bar = _mostrar_progresso(len(ARQUIVOS))

    try:
        for i, arquivo in enumerate(ARQUIVOS, 1):
            lbl.setText(f"Baixando: {arquivo}")
            app.processEvents()
            _baixar_arquivo(arquivo)
            bar.setValue(i)
            app.processEvents()

        _salvar_versao(versao_remota)
        dialog.close()

        from PySide6.QtWidgets import QMessageBox
        msg = QMessageBox()
        msg.setWindowTitle("Atualização concluída")
        msg.setText(f"Versão {versao_remota} instalada com sucesso!\n\nO sistema será reiniciado.")
        msg.setStyleSheet("QMessageBox { background-color: #0f1115; } QLabel { color: #e5e7eb; }")
        msg.setIcon(QMessageBox.NoIcon)
        msg.exec()

        exe = sys.executable
        os.execv(exe, [exe] + sys.argv)

    except Exception as e:
        dialog.close()
        _mostrar_erro(str(e))

if __name__ == "__main__":
    verificar_e_atualizar()

    import interface