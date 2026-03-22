import sys
import os
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, QThread, Signal, QTimer
from PySide6.QtGui import QFont, QPainter, QColor, QPixmap
from PySide6.QtWidgets import QCompleter, QDateEdit
from gerador import gerar_ordem, _listar_contas_gmail, adicionar_conta_gmail


STYLE = """
QWidget {
    background-color: #0f1115;
    color: #e5e7eb;
    font-family: Arial;
    font-size: 13px;
}

QLineEdit, QComboBox {
    background-color: #1a1d23;
    border: 1px solid #2a2f3a;
    border-radius: 8px;
    padding: 8px;
    text-transform: uppercase;
}

QLineEdit:focus, QComboBox:focus {
    border: 1px solid #3b82f6;
}

QLabel {
    color: #9ca3af;
}

QPushButton {
    border-radius: 8px;
    padding: 12px;
    font-weight: bold;
}

QPushButton:hover {
    opacity: 0.85;
    filter: brightness(1.15);
}

#btn_gerar {
    background-color: #2e7d32;
    color: white;
}

#btn_gerar:hover {
    background-color: #388e3c;
}

#btn_email {
    background-color: #2a2f3a;
    color: #e5e7eb;
}

#btn_email:hover {
    background-color: #353b4a;
}

#btn_nova {
    background-color: transparent;
    border: 1px solid #2a2f3a;
    color: #e5e7eb;
}

#btn_nova:hover {
    background-color: #1a1d23;
    border: 1px solid #3b82f6;
    color: #3b82f6;
}

#btn_wpp {
    background-color: #1a3a2a;
    color: #4ade80;
    border: 1px solid #166534;
}

#btn_wpp:hover {
    background-color: #1f4a33;
    border: 1px solid #4ade80;
}
"""


def criar_card(titulo):
    box = QGroupBox(titulo.upper())
    box.setStyleSheet("""
        QGroupBox {
            border: 1px solid #2a2f3a;
            border-radius: 12px;
            margin-top: 10px;
            padding: 15px;
            font-weight: bold;
        }
    """)

    layout = QGridLayout()
    layout.setHorizontalSpacing(15)
    layout.setVerticalSpacing(10)
    layout.setColumnStretch(0, 1)
    layout.setColumnStretch(1, 1)
    layout.setAlignment(Qt.AlignTop)

    box.setLayout(layout)
    return box, layout


# =========================
# PARSER WHATSAPP
# =========================
def parsear_mensagem_whatsapp(texto):
    """
    Extrai campos do modelo de mensagem WhatsApp e retorna
    um dict com os nomes dos campos do formulário.
    """
    import re

    resultado = {}

    def extrair(chave):
        match = re.search(rf"^{chave}\s*:\s*(.+)", texto, re.IGNORECASE | re.MULTILINE)
        return match.group(1).strip() if match else ""

    # TRANSPORTADORA → empresa
    transp = extrair("TRANSPORTADORA").upper()
    if "AGRO" in transp:
        resultado["empresa"] = "Agrovia"
    else:
        resultado["empresa"] = "TopBrasil"

    # PAGADOR → Fábrica (nome curto, ex: TIMAC)
    pagador = extrair("PAGADOR").upper().split()[0] if extrair("PAGADOR") else ""
    resultado["Fábrica"] = pagador

    # CLIENTE (opcional)
    cliente = extrair("CLIENTE")
    if cliente:
        resultado["Cliente"] = cliente

    # PEDIDO
    resultado["Pedido"] = extrair("PEDIDO")

    # PRODUTO
    resultado["Produto"] = extrair("PRODUTO")

    # PESO
    resultado["Peso"] = extrair("PESO")

    # MOTORISTA
    resultado["Motorista"] = extrair("MOTORISTA")

    # DESTINO → separar "BARREIRAS-BA FAZ COSMOS" em Destino + Fazenda
    destino_raw = extrair("DESTINO")
    if destino_raw:
        # Tenta separar pela palavra FAZ, FAZENDA, SITIO, SITÍO, CHACARA
        sep = re.split(r"\s+(FAZ\.?\s|FAZENDA\s|SITIO\s|SÍTIO\s|CHACARA\s)", destino_raw, flags=re.IGNORECASE)
        if len(sep) >= 3:
            resultado["Destino"] = sep[0].strip()
            resultado["Fazenda"] = (sep[1] + sep[2]).strip()
        else:
            resultado["Destino"] = destino_raw.strip()
            resultado["Fazenda"] = ""
    else:
        resultado["Destino"] = ""
        resultado["Fazenda"] = ""

    return resultado


class GeradorThread(QThread):
    sucesso = Signal()
    erro = Signal(str)

    def __init__(self, dados, pasta, email, conta_gmail=None):
        super().__init__()
        self.dados = dados
        self.pasta = pasta
        self.email = email
        self.conta_gmail = conta_gmail

    def run(self):
        try:
            gerar_ordem(self.dados, self.pasta, self.email, self.conta_gmail)
            self.sucesso.emit()
        except Exception as e:
            import traceback
            with open("erro_log.txt", "w") as f:
                f.write(traceback.format_exc())
            self.erro.emit(str(e))
        except BaseException as e:
            import traceback
            with open("erro_log.txt", "w") as f:
                f.write(traceback.format_exc())
            self.erro.emit(f"Erro inesperado: {e}")


class LoadingOverlay(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, False)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self._dots = 0
        self._label = QLabel("Gerando ordem", self)
        self._label.setAlignment(Qt.AlignCenter)
        self._label.setStyleSheet("color: #e5e7eb; font-size: 15px; font-weight: bold; background: transparent;")

        self._timer = QTimer(self)
        self._timer.timeout.connect(self._animar)
        self.hide()

    def showEvent(self, event):
        self._resize()
        self._timer.start(400)
        super().showEvent(event)

    def hideEvent(self, event):
        self._timer.stop()
        super().hideEvent(event)

    def _resize(self):
        if self.parent():
            self.setGeometry(self.parent().rect())
            self._label.setGeometry(self.parent().rect())

    def resizeEvent(self, event):
        self._resize()
        super().resizeEvent(event)

    def _animar(self):
        self._dots = (self._dots + 1) % 4
        self._label.setText("Gerando ordem" + "." * self._dots)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.fillRect(self.rect(), QColor(15, 17, 21, 210))


class UI(QWidget):
    def __init__(self):
        super().__init__()

        self.empresa = None
        self.setWindowTitle("Sistema de Ordens")

        # FUNDO COM LOGO
        self._bg_label = QLabel(self)
        self._bg_label.setScaledContents(True)
        self._bg_label.setGeometry(self.rect())
        self._bg_label.lower()

        self.entradas = {}

        main = QVBoxLayout(self)

        grid = QGridLayout()
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        grid.setHorizontalSpacing(20)
        grid.setVerticalSpacing(20)

        main.addLayout(grid)

        # CABEÇALHO
        card, lay = criar_card("Cabeçalho")
        self.add_input(lay, "Data Apresentação", 0, 0)
        self.add_input(lay, "Fábrica", 0, 1)
        self.add_input(lay, "Solicitante", 0, 2)
        grid.addWidget(card, 0, 0)
        
        self.entradas["Data Apresentação"].setMaximumWidth(120)
        self.entradas["Fábrica"].setMinimumWidth(150)
        self.entradas["Solicitante"].setMinimumWidth(200)

        # MOTORISTA
        card, lay = criar_card("Motorista")
        self.add_input(lay, "Motorista", 0, 0)
        self.add_input(lay, "CPF", 0, 1)
        self.add_input(lay, "Contato", 0, 2)
        grid.addWidget(card, 0, 1)

        # CARGA
        card, lay = criar_card("Carga")
        campos = ["Destino", "Fazenda", "Produto", "Embalagem", "Cliente", "Pedido", "Peso"]

        for i, campo in enumerate(campos):
            par   = i // 2          # qual par (0,1,2,3)
            col   = i % 2           # coluna 0 ou 1
            row   = par * 2         # linha do label: 0, 2, 4, 6
            self.add_input(lay, campo, row, col)

        lay.setVerticalSpacing(4)

        grid.addWidget(card, 1, 0)

        # VEÍCULO
        card, lay = criar_card("Veículo")

        self.add_input(lay, "Carroceria", 0, 0, combo=[
            "Graneleiro", "Basculante", "Baú", "Sider", "Tanque"
        ])
        self.add_input(lay, "Cavalo", 0, 1)
        self.add_input(lay, "Carreta 1", 2, 0)
        self.add_input(lay, "Carreta 2", 2, 1)
        self.add_input(lay, "Carreta 3", 4, 0)

        grid.addWidget(card, 1, 1)

        # ASSINATURA
        card, lay = criar_card("Assinatura")
        entrada = QLineEdit()
        entrada.setMinimumHeight(36)
        lay.addWidget(entrada, 0, 0, 1, 2)
        self.entradas["Assinatura"] = entrada

        dev_label = QLabel("© 2026 dev by Felipe")
        dev_label.setAlignment(Qt.AlignCenter)
        dev_label.setStyleSheet("color: #3b4252; font-size: 11px; padding-top: 4px;")
        lay.addWidget(dev_label, 1, 0, 1, 2)

        grid.addWidget(card, 2, 0)

        # BOTÕES
        botoes = QVBoxLayout()

        self.btn1 = QPushButton("GERAR ORDEM")
        self.btn1.setObjectName("btn_gerar")

        self.btn2 = QPushButton("GERAR + EMAIL")
        self.btn2.setObjectName("btn_email")

        self.btn3 = QPushButton("NOVA ORDEM")
        self.btn3.setObjectName("btn_nova")

        self.btn_wpp = QPushButton("📋  IMPORTAR WHATSAPP")
        self.btn_wpp.setObjectName("btn_wpp")

        self.btn1.clicked.connect(lambda: self.executar(False))
        self.btn2.clicked.connect(lambda: self.executar(True))
        self.btn3.clicked.connect(self.nova_ordem)
        self.btn_wpp.clicked.connect(self.importar_whatsapp)

        botoes.addWidget(self.btn_wpp)
        botoes.addSpacing(8)
        botoes.addWidget(self.btn1)
        botoes.addWidget(self.btn2)
        botoes.addWidget(self.btn3)
        botoes.addStretch()

        grid.addLayout(botoes, 2, 1, alignment=Qt.AlignTop)

        self.setStyleSheet(STYLE)

        self.overlay = LoadingOverlay(self)

        self.showMaximized()
        self.escolher_empresa()
        self.setar_data_hoje()

    def _atualizar_fundo(self, empresa):
        # Resolve o caminho base: dentro do .exe usa sys._MEIPASS, fora usa a pasta do script
        if getattr(sys, "frozen", False):
            from pathlib import Path
            base = Path(sys._MEIPASS)
        else:
            from pathlib import Path
            base = Path(__file__).parent

        nomes = {"Agrovia": "logo_agro.png", "TopBrasil": "logo_top.png"}
        arquivo = nomes.get(empresa, "")
        caminho = str(base / arquivo) if arquivo else ""

        if caminho and os.path.exists(caminho):
            # Blur simulado: escala pra baixo e volta pra cima
            original = QPixmap(caminho)
            pequeno = original.scaled(
                original.width() // 8,
                original.height() // 8,
                Qt.KeepAspectRatioByExpanding,
                Qt.SmoothTransformation
            )
            borrado = pequeno.scaled(
                original.width(),
                original.height(),
                Qt.KeepAspectRatioByExpanding,
                Qt.SmoothTransformation
            )
            # Overlay escuro desenhado sobre o pixmap borrado
            painter = QPainter(borrado)
            painter.fillRect(borrado.rect(), QColor(15, 17, 21, 185))
            painter.end()
            self._bg_label.setPixmap(borrado)
        else:
            self._bg_label.setPixmap(QPixmap())
        self._bg_label.setGeometry(self.rect())
        self._bg_label.lower()

    def resizeEvent(self, event):
        self._bg_label.setGeometry(self.rect())
        super().resizeEvent(event)

    def setar_data_hoje(self):
        if isinstance(self.entradas["Data Apresentação"], QDateEdit):
            self.entradas["Data Apresentação"].setDate(QDate.currentDate())
        else:
            hoje = QDate.currentDate().toString("dd/MM/yyyy")
            self.entradas["Data Apresentação"].setText(hoje)

    def escolher_empresa(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Escolher Empresa")
        dialog.setFixedSize(320, 220)

        # ✅ RESTAURADO
        dialog.setStyleSheet("""
            QDialog { background-color: #0f1115; }

            QLabel {
                color: #e5e7eb;
                font-size: 14px;
            }

            QPushButton {
                border-radius: 8px;
                padding: 12px;
                font-weight: bold;
            }

            #agro { background-color: #2e7d32; color: white; }
            #top { background-color: #c62828; color: white; }
        """)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        titulo = QLabel("ESCOLHA A EMPRESA")
        titulo.setAlignment(Qt.AlignCenter)

        btn_agro = QPushButton("AGROVIA")
        btn_agro.setObjectName("agro")

        btn_top = QPushButton("TOPBRASIL")
        btn_top.setObjectName("top")

        def selecionar(nome):
            self.empresa = nome
            if nome == "Agrovia":
                self.btn1.setStyleSheet("background-color: #2e7d32; color: white;")
            else:
                self.btn1.setStyleSheet("background-color: #c62828; color: white;")
            self._atualizar_fundo(nome)
            dialog.accept()

        btn_agro.clicked.connect(lambda: selecionar("Agrovia"))
        btn_top.clicked.connect(lambda: selecionar("TopBrasil"))

        layout.addWidget(titulo)
        layout.addWidget(btn_agro)
        layout.addWidget(btn_top)

        dialog.exec()

    def add_input(self, layout, label, row, col, combo=None):
        lbl = QLabel(label)
        layout.addWidget(lbl, row, col)

        if label == "Data Apresentação":
            inp = QDateEdit()
            inp.setDisplayFormat("dd/MM/yyyy")
            inp.setDate(QDate.currentDate())
            inp.setCalendarPopup(True)

        elif combo:
            inp = QComboBox()
            inp.setEditable(True)
            inp.addItems(combo)

            completer = QCompleter(combo)
            completer.setCaseSensitivity(Qt.CaseInsensitive)
            inp.setCompleter(completer)
        else:
            inp = QLineEdit()
            if label == "Cavalo":
                inp.setMaxLength(8)
                inp.textChanged.connect(lambda texto, i=inp: self._formatar_placa(i, texto))
            elif label not in ("CPF", "Contato"):
                inp.textChanged.connect(lambda texto, i=inp: self._forcar_maiusculo(i, texto))

        inp.setMinimumHeight(36)
        inp.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout.addWidget(inp, row + 1, col, 1, 1)
        self.entradas[label] = inp

    def _forcar_maiusculo(self, inp, texto):
        if texto != texto.upper():
            inp.blockSignals(True)
            cursor = inp.cursorPosition()
            inp.setText(texto.upper())
            inp.setCursorPosition(cursor)
            inp.blockSignals(False)

    def _formatar_placa(self, inp, texto):
        """Formata a placa no modelo XXX-XXXX em maiúsculas enquanto digita."""
        import re
        # Remove tudo que não for alfanumérico e converte para maiúsculas
        limpo = re.sub(r"[^A-Za-z0-9]", "", texto).upper()
        # Aplica máscara XXX-XXXX
        if len(limpo) > 3:
            formatado = limpo[:3] + "-" + limpo[3:7]
        else:
            formatado = limpo
        # Bloqueia o sinal para não entrar em loop
        inp.blockSignals(True)
        cursor = inp.cursorPosition()
        inp.setText(formatado)
        # Mantém o cursor na posição correta
        inp.setCursorPosition(min(cursor, len(formatado)))
        inp.blockSignals(False)

    def coletar(self):
        dados = {"empresa": self.empresa}

        for k, v in self.entradas.items():
            if isinstance(v, QComboBox):
                dados[k] = v.currentText()
            elif isinstance(v, QDateEdit):
                dados[k] = v.date().toString("dd/MM/yyyy")
            else:
                dados[k] = v.text()

        return dados

    def executar(self, email):
        pasta = QFileDialog.getExistingDirectory(self, "Salvar em")
        if not pasta:
            return

        dados = self.coletar()

        if not dados.get("Motorista") or not dados.get("Cavalo"):
            QMessageBox.warning(self, "Atenção", "Preencha Motorista e Cavalo")
            return

        conta_gmail = None
        preview_email = None
        if email:
            conta_gmail = self._dialog_escolher_conta()
            if conta_gmail is None:
                return

            from gerador import obter_email_fabrica, montar_email
            destinatario = obter_email_fabrica(dados.get("Fábrica"))
            assunto, corpo = montar_email(dados)
            preview_email = self._dialog_preview_email(destinatario, assunto, corpo)
            if preview_email is None:
                return
            # Usa os valores (possivelmente editados) do preview
            dados["_email_destinatario"] = preview_email["destinatario"]
            dados["_email_assunto"]      = preview_email["assunto"]
            dados["_email_corpo"]        = preview_email["corpo"]

        self.btn1.setEnabled(False)
        self.btn2.setEnabled(False)
        self.btn3.setEnabled(False)

        self.overlay.show()
        self.overlay.raise_()

        self._thread = GeradorThread(dados, pasta, email, conta_gmail)
        self._thread.sucesso.connect(self._on_sucesso)
        self._thread.erro.connect(self._on_erro)
        self._thread.start()

    def _dialog_escolher_conta(self):
        """
        Abre um dialog para escolher (ou adicionar) conta Gmail.
        Retorna o email selecionado, ou None se cancelado.
        """
        dialog = QDialog(self)
        dialog.setWindowTitle("Enviar por Gmail")
        dialog.setFixedSize(400, 220)
        dialog.setStyleSheet("""
            QDialog { background-color: #0f1115; }
            QLabel { color: #e5e7eb; font-size: 13px; }
            QComboBox {
                background-color: #1a1d23;
                border: 1px solid #2a2f3a;
                border-radius: 8px;
                padding: 8px;
                color: #e5e7eb;
                min-height: 36px;
            }
            QPushButton {
                border-radius: 8px;
                padding: 10px;
                font-weight: bold;
            }
            #btn_ok     { background-color: #166534; color: #4ade80; }
            #btn_cancel { background-color: #2a2f3a; color: #e5e7eb; }
            #btn_add    { background-color: transparent; border: 1px solid #3b82f6; color: #3b82f6; padding: 6px; }
        """)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        layout.addWidget(QLabel("Escolha a conta remetente:"))

        combo = QComboBox()
        contas = _listar_contas_gmail()
        if contas:
            combo.addItems(contas)
        else:
            combo.addItem("(nenhuma conta configurada)")
        layout.addWidget(combo)

        btn_add = QPushButton("+ Adicionar conta Gmail")
        btn_add.setObjectName("btn_add")
        layout.addWidget(btn_add)

        btns = QHBoxLayout()
        btn_cancel = QPushButton("CANCELAR")
        btn_cancel.setObjectName("btn_cancel")
        btn_ok = QPushButton("ENVIAR")
        btn_ok.setObjectName("btn_ok")
        btns.addWidget(btn_cancel)
        btns.addWidget(btn_ok)
        layout.addLayout(btns)

        resultado = [None]

        def adicionar():
            try:
                email_novo = adicionar_conta_gmail()
                combo.clear()
                combo.addItems(_listar_contas_gmail())
                idx = combo.findText(email_novo)
                if idx >= 0:
                    combo.setCurrentIndex(idx)
            except Exception as e:
                QMessageBox.critical(dialog, "Erro", str(e))

        def confirmar():
            conta = combo.currentText()
            if conta == "(nenhuma conta configurada)":
                QMessageBox.warning(dialog, "Atenção", "Adicione uma conta Gmail primeiro.")
                return
            resultado[0] = conta
            dialog.accept()

        btn_add.clicked.connect(adicionar)
        btn_cancel.clicked.connect(dialog.reject)
        btn_ok.clicked.connect(confirmar)

        dialog.exec()
        return resultado[0]

    def _dialog_preview_email(self, destinatario, assunto, corpo):
        """
        Mostra prévia do email antes de enviar. Campos editáveis.
        Retorna dict com destinatario/assunto/corpo, ou None se cancelado.
        """
        DIALOG_STYLE = """
            QDialog { background-color: #0f1115; }
            QLabel { color: #6b7280; font-size: 11px; }
            QLineEdit, QTextEdit {
                background-color: #1a1d23;
                border: 1px solid #2a2f3a;
                border-radius: 6px;
                padding: 8px;
                color: #e5e7eb;
                font-size: 13px;
            }
            QLineEdit:focus, QTextEdit:focus { border: 1px solid #3b82f6; }
            QPushButton { border-radius: 8px; padding: 10px; font-weight: bold; }
            #btn_ok     { background-color: #166534; color: #4ade80; }
            #btn_cancel { background-color: #2a2f3a; color: #e5e7eb; }
        """

        dialog = QDialog(self)
        dialog.setWindowTitle("Prévia do email")
        dialog.setMinimumWidth(520)
        dialog.setStyleSheet(DIALOG_STYLE)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(8)

        layout.addWidget(QLabel("DESTINATÁRIO"))
        inp_dest = QLineEdit(destinatario)
        inp_dest.setMinimumHeight(36)
        layout.addWidget(inp_dest)

        layout.addWidget(QLabel("ASSUNTO"))
        inp_assunto = QLineEdit(assunto)
        inp_assunto.setMinimumHeight(36)
        layout.addWidget(inp_assunto)

        layout.addWidget(QLabel("CORPO"))
        inp_corpo = QTextEdit()
        inp_corpo.setPlainText(corpo)
        inp_corpo.setMinimumHeight(120)
        layout.addWidget(inp_corpo)

        btns = QHBoxLayout()
        btn_cancel = QPushButton("CANCELAR")
        btn_cancel.setObjectName("btn_cancel")
        btn_ok = QPushButton("CONFIRMAR ENVIO")
        btn_ok.setObjectName("btn_ok")
        btns.addWidget(btn_cancel)
        btns.addWidget(btn_ok)
        layout.addLayout(btns)

        resultado = [None]

        def confirmar():
            resultado[0] = {
                "destinatario": inp_dest.text().strip(),
                "assunto":      inp_assunto.text().strip(),
                "corpo":        inp_corpo.toPlainText().strip(),
            }
            dialog.accept()

        btn_cancel.clicked.connect(dialog.reject)
        btn_ok.clicked.connect(confirmar)

        dialog.exec()
        return resultado[0]
        self.overlay.hide()
        self.btn1.setEnabled(True)
        self.btn2.setEnabled(True)
        self.btn3.setEnabled(True)

        msg = QMessageBox(self)
        msg.setWindowTitle("Sucesso")
        msg.setText("✔ Ordem gerada com sucesso")
        msg.setIcon(QMessageBox.NoIcon)

        msg.setStyleSheet("""
        QMessageBox {
            background-color: #0f1115;
        }

        QLabel {
            color: #e5e7eb;
            font-size: 13px;
        }

        QPushButton {
            background-color: #2e7d32;
            color: white;
            border-radius: 6px;
            padding: 6px 15px;
        }
        """)

        msg.exec()

    def _on_sucesso(self):
        self.overlay.hide()
        self.btn1.setEnabled(True)
        self.btn2.setEnabled(True)
        self.btn3.setEnabled(True)

        msg = QMessageBox(self)
        msg.setWindowTitle("Sucesso")
        msg.setText("✔ Ordem gerada com sucesso")
        msg.setIcon(QMessageBox.NoIcon)

        msg.setStyleSheet("""
        QMessageBox { background-color: #0f1115; }
        QLabel { color: #e5e7eb; font-size: 13px; }
        QPushButton {
            background-color: #2e7d32;
            color: white;
            border-radius: 6px;
            padding: 6px 15px;
        }
        """)

        msg.exec()

    def _on_erro(self, mensagem):
        self.overlay.hide()
        self.btn1.setEnabled(True)
        self.btn2.setEnabled(True)
        self.btn3.setEnabled(True)
        QMessageBox.critical(self, "Erro", mensagem)

    def importar_whatsapp(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Importar mensagem do WhatsApp")
        dialog.setFixedSize(480, 360)
        dialog.setStyleSheet("""
            QDialog { background-color: #0f1115; }
            QLabel { color: #e5e7eb; font-size: 13px; }
            QTextEdit {
                background-color: #1a1d23;
                border: 1px solid #2a2f3a;
                border-radius: 8px;
                padding: 8px;
                color: #e5e7eb;
                font-size: 12px;
            }
            QPushButton {
                border-radius: 8px;
                padding: 10px;
                font-weight: bold;
            }
            #btn_ok  { background-color: #166534; color: #4ade80; }
            #btn_cancel { background-color: #2a2f3a; color: #e5e7eb; }
        """)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        lbl = QLabel("Cole a mensagem do WhatsApp abaixo:")
        caixa = QTextEdit()
        caixa.setPlaceholderText("🗒️ TAG\nTRANSPORTADORA: TOP BRASIL\n...")

        btns = QHBoxLayout()
        btn_ok = QPushButton("PREENCHER")
        btn_ok.setObjectName("btn_ok")
        btn_cancel = QPushButton("CANCELAR")
        btn_cancel.setObjectName("btn_cancel")
        btns.addWidget(btn_cancel)
        btns.addWidget(btn_ok)

        layout.addWidget(lbl)
        layout.addWidget(caixa)
        layout.addLayout(btns)

        btn_cancel.clicked.connect(dialog.reject)

        def confirmar():
            texto = caixa.toPlainText().strip()
            if not texto:
                return
            dados = parsear_mensagem_whatsapp(texto)
            self._preencher_campos(dados)
            dialog.accept()

        btn_ok.clicked.connect(confirmar)
        dialog.exec()

    def _preencher_campos(self, dados):
        mapa = {
            "Fábrica":    "Fábrica",
            "Pedido":     "Pedido",
            "Produto":    "Produto",
            "Peso":       "Peso",
            "Motorista":  "Motorista",
            "Destino":    "Destino",
            "Fazenda":    "Fazenda",
            "Cliente":    "Cliente",
        }

        for chave, campo in mapa.items():
            valor = dados.get(chave, "")
            if not valor:
                continue
            widget = self.entradas.get(campo)
            if isinstance(widget, QLineEdit):
                widget.setText(valor)
            elif isinstance(widget, QComboBox):
                widget.setEditText(valor)

        # Atualiza empresa se veio na mensagem
        empresa = dados.get("empresa")
        if empresa:
            self.empresa = empresa
            if empresa == "Agrovia":
                self.btn1.setStyleSheet("background-color: #2e7d32; color: white;")
            else:
                self.btn1.setStyleSheet("background-color: #c62828; color: white;")

    def nova_ordem(self):
        for v in self.entradas.values():
            if isinstance(v, QLineEdit):
                v.clear()
            elif isinstance(v, QComboBox):
                v.setCurrentIndex(0)
            elif isinstance(v, QDateEdit):
                v.setDate(QDate.currentDate())

        self.escolher_empresa()
        self.setar_data_hoje()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = UI()
    win.show()
    sys.exit(app.exec())