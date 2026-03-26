import sys
import os
from pathlib import Path
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, QThread, Signal, QTimer, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QFont, QPainter, QColor, QPixmap, QIcon, QPalette
from PySide6.QtWidgets import QCompleter, QDateEdit
from gerador import gerar_ordem, _listar_contas_gmail, adicionar_conta_gmail
from planilha import carregar_blocos, carregar_blocos_dados, gravar_carregamento, carregar_base

BG       = "#0d1117"
SURFACE  = "#161b22"
BORDER   = "#21262d"
BORDER2  = "#30363d"
TEXT     = "#e6edf3"
MUTED    = "#8b949e"
ACCENT   = "#238636"
ACCENT_H = "#2ea043"
ACCENT_L = "#1a7f37"
DANGER   = "#da3633"
DANGER_H = "#f85149"

DIALOG_SS = f"""
    QDialog {{ background-color: {BG}; }}
    QLabel {{ color: {TEXT}; font-family: "Segoe UI"; font-size: 13px; background: transparent; }}
    QLineEdit, QTextEdit, QComboBox {{
        background-color: {SURFACE};
        border: 1px solid {BORDER2};
        border-radius: 6px;
        padding: 8px 10px;
        color: {TEXT};
        font-size: 13px;
    }}
    QLineEdit:focus, QTextEdit:focus {{ border-color: {ACCENT}; }}
    QPushButton {{
        border-radius: 6px;
        padding: 9px 18px;
        font-weight: 700;
        font-size: 13px;
        font-family: "Segoe UI";
    }}
    #btn_ok   {{ background-color: {ACCENT}; color: white; border: none; }}
    #btn_ok:hover {{ background-color: {ACCENT_H}; }}
    #btn_cancel {{ background-color: transparent; border: 1px solid {BORDER2}; color: {MUTED}; }}
    #btn_cancel:hover {{ background-color: {SURFACE}; color: {TEXT}; }}
    #btn_add  {{ background-color: transparent; border: 1px solid {ACCENT}; color: {ACCENT}; }}
    #btn_add:hover {{ background-color: {ACCENT}18; }}
    #btn_agro {{ background-color: {ACCENT}; color: white; border: none; }}
    #btn_agro:hover {{ background-color: {ACCENT_H}; }}
    #btn_top  {{ background-color: {DANGER}; color: white; border: none; }}
    #btn_top:hover {{ background-color: {DANGER_H}; }}
"""

def make_field(label_text, widget):
                                                            
    w = QWidget()
    w.setStyleSheet("background: transparent;")
    v = QVBoxLayout(w)
    v.setContentsMargins(0, 0, 0, 0)
    v.setSpacing(4)
    lbl = QLabel(label_text.upper())
    lbl.setStyleSheet(f"color: {MUTED}; font-size: 10px; font-weight: 700; letter-spacing: 0.8px; background: transparent;")
    v.addWidget(lbl)
    v.addWidget(widget)
    return w

def make_input(placeholder="", maiusculo=True, max_len=None):
    inp = QLineEdit()
    inp.setMinimumHeight(36)
    inp.setPlaceholderText(placeholder)
    if max_len:
        inp.setMaxLength(max_len)
    if maiusculo:
        inp.textChanged.connect(lambda t, i=inp: _forcar_maiusculo(i, t))
    return inp

def make_combo(items):
    cb = QComboBox()
    cb.setEditable(True)
    cb.addItems(items)
    cb.setMinimumHeight(36)
    cb.setCompleter(QCompleter(items))
    cb.completer().setCaseSensitivity(Qt.CaseInsensitive)
    return cb

def make_date():
    d = QDateEdit()
    d.setDisplayFormat("dd/MM/yyyy")
    d.setDate(QDate.currentDate())
    d.setCalendarPopup(True)
    d.setMinimumHeight(36)
    return d

def _forcar_maiusculo(inp, texto):
    if texto != texto.upper():
        inp.blockSignals(True)
        c = inp.cursorPosition()
        inp.setText(texto.upper())
        inp.setCursorPosition(c)
        inp.blockSignals(False)

def _formatar_placa(inp, texto):
    import re
    limpo = re.sub(r"[^A-Za-z0-9]", "", texto).upper()
    formatado = limpo[:3] + "-" + limpo[3:7] if len(limpo) > 3 else limpo
    inp.blockSignals(True)
    c = inp.cursorPosition()
    inp.setText(formatado)
    inp.setCursorPosition(min(c, len(formatado)))
    inp.blockSignals(False)

def make_card(title):
                                                             
    frame = QFrame()
    frame.setObjectName("card")
    frame.setStyleSheet(f"""
        QFrame#card {{
            background-color: {SURFACE};
            border: 1px solid {BORDER};
            border-radius: 10px;
        }}
    """)
    vbox = QVBoxLayout(frame)
    vbox.setContentsMargins(16, 14, 16, 16)
    vbox.setSpacing(12)

    lbl = QLabel(title.upper())
    lbl.setStyleSheet(f"""
        color: {MUTED};
        font-size: 10px;
        font-weight: 700;
        letter-spacing: 1.5px;
        background: transparent;
        padding-bottom: 4px;
        border-bottom: 1px solid {BORDER};
    """)
    vbox.addWidget(lbl)

    content = QWidget()
    content.setStyleSheet("background: transparent;")
    vbox.addWidget(content)

    return frame, content

import json
import datetime

def _historico_path():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / "historico.json"
    return Path(__file__).parent / "historico.json"

def salvar_historico(dados, caminho_arquivo):
    path = _historico_path()
    try:
        historico = json.loads(path.read_text(encoding="utf-8")) if path.exists() else []
    except Exception:
        historico = []

    registro = {
        "data_hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "motorista": dados.get("Motorista", ""),
        "placa":     dados.get("Cavalo", ""),
        "empresa":   dados.get("empresa", ""),
        "arquivo":   caminho_arquivo,
    }
    historico.insert(0, registro)
    path.write_text(json.dumps(historico, ensure_ascii=False, indent=2), encoding="utf-8")

def carregar_historico():
    path = _historico_path()
    if not path.exists():
        return []
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []


class HistoricoWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background: transparent;")
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        topo = QHBoxLayout()
        titulo = QLabel("HISTÓRICO DE ORDENS")
        titulo.setStyleSheet(f"color: {TEXT}; font-size: 14px; font-weight: 700; letter-spacing: 1px; background: transparent;")
        btn_atualizar = QPushButton("↺  ATUALIZAR")
        btn_atualizar.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: 1px solid {BORDER2};
                border-radius: 6px;
                color: {MUTED};
                padding: 6px 12px;
                font-size: 12px;
            }}
            QPushButton:hover {{ color: {TEXT}; border-color: {ACCENT}; }}
        """)
        btn_atualizar.clicked.connect(self.recarregar)
        topo.addWidget(titulo)
        topo.addStretch()
        topo.addWidget(btn_atualizar)
        root.addLayout(topo)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")

        self._container = QWidget()
        self._container.setStyleSheet("background: transparent;")
        self._vbox = QVBoxLayout(self._container)
        self._vbox.setSpacing(8)
        self._vbox.setContentsMargins(0, 0, 0, 0)
        self._vbox.addStretch()

        scroll.setWidget(self._container)
        root.addWidget(scroll)

        self.recarregar()

    def recarregar(self):
        while self._vbox.count() > 1:
            item = self._vbox.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        registros = carregar_historico()

        if not registros:
            vazio = QLabel("Nenhuma ordem gerada ainda.")
            vazio.setAlignment(Qt.AlignCenter)
            vazio.setStyleSheet(f"color: {MUTED}; font-size: 13px; background: transparent;")
            self._vbox.insertWidget(0, vazio)
            return

        grupos = {}
        for r in registros:
            data = r.get("data_hora", "")[:10]
            grupos.setdefault(data, []).append(r)

        for i, (data, items) in enumerate(grupos.items()):
            lbl_data = QLabel(data)
            lbl_data.setStyleSheet(f"color: {MUTED}; font-size: 11px; font-weight: 700; letter-spacing: 1px; background: transparent; padding-top: 4px;")
            self._vbox.insertWidget(self._vbox.count() - 1, lbl_data)

            for r in items:
                card = self._make_card(r)
                self._vbox.insertWidget(self._vbox.count() - 1, card)

    def _make_card(self, r):
        empresa = r.get("empresa", "")
        cor_borda = ACCENT if empresa == "Agrovia" else DANGER if empresa == "TopBrasil" else BORDER
        frame = QFrame()
        frame.setStyleSheet(f"""
            QFrame {{
                background-color: {SURFACE};
                border: 1px solid {cor_borda};
                border-radius: 8px;
            }}
        """)
        h = QHBoxLayout(frame)
        h.setContentsMargins(14, 10, 14, 10)
        h.setSpacing(12)

        hora = r.get("data_hora", "")[-5:]
        lbl_hora = QLabel(hora)
        lbl_hora.setStyleSheet(f"color: {MUTED}; font-size: 11px; background: transparent; min-width: 40px;")

        info = QVBoxLayout()
        info.setSpacing(2)
        motorista = QLabel(r.get("motorista", "—").title())
        motorista.setStyleSheet(f"color: {TEXT}; font-size: 13px; font-weight: 600; background: transparent;")
        placa = QLabel(r.get("placa", "—"))
        placa.setStyleSheet(f"color: {MUTED}; font-size: 11px; background: transparent;")
        info.addWidget(motorista)
        info.addWidget(placa)

        empresa = r.get("empresa", "")
        badge = QLabel(empresa.upper() if empresa else "—")
        cor_badge = ACCENT if empresa == "Agrovia" else DANGER if empresa == "TopBrasil" else MUTED
        badge.setStyleSheet(f"""
            color: {cor_badge};
            background: {cor_badge}18;
            border: 1px solid {cor_badge}44;
            border-radius: 4px;
            font-size: 10px;
            font-weight: 700;
            padding: 2px 7px;
            letter-spacing: 0.5px;
        """)

        arquivo = r.get("arquivo", "")
        arquivo_xlsx = arquivo if arquivo.endswith(".xlsx") else arquivo.replace(".pdf", ".xlsx")

        btn_editar = QPushButton("EDITAR")
        btn_editar.setToolTip(f"Abrir: {arquivo_xlsx}")
        btn_editar.setFixedHeight(28)
        btn_editar.setFont(QFont("Arial", 8, QFont.Bold))
        btn_editar.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: 1px solid {BORDER2};
                border-radius: 6px;
                color: {MUTED};
                font-size: 9px;
                padding: 0px 8px;
            }}
            QPushButton:hover {{ border-color: {ACCENT}; color: {ACCENT}; }}
        """)
        btn_editar.clicked.connect(lambda _, a=arquivo_xlsx: self._abrir_arquivo(a))

        h.addWidget(lbl_hora)
        h.addLayout(info, 1)
        h.addWidget(badge)
        h.addWidget(btn_editar)
        return frame

    def _abrir_arquivo(self, caminho):
        import subprocess
        try:
            os.startfile(caminho)
        except Exception:
            try:
                subprocess.run(["explorer", "/select,", caminho])
            except Exception:
                pass


class PlanilhaWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background: transparent;")
        self._conta = None

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        topo = QHBoxLayout()
        titulo = QLabel("CONTROLE DE PEDIDOS")
        titulo.setStyleSheet(f"color: {TEXT}; font-size: 14px; font-weight: 700; letter-spacing: 1px; background: transparent;")

        self._combo_conta = QComboBox()
        self._combo_conta.setFixedWidth(220)
        self._combo_conta.setStyleSheet(f"""
            QComboBox {{
                background: {SURFACE}; border: 1px solid {BORDER2};
                border-radius: 6px; padding: 6px 10px; color: {TEXT}; font-size: 12px;
            }}
        """)

        btn_carregar = QPushButton("↺  CARREGAR")
        btn_carregar.setStyleSheet(f"""
            QPushButton {{
                background: {ACCENT}; color: white; border: none;
                border-radius: 6px; padding: 7px 14px; font-weight: 700; font-size: 12px;
            }}
            QPushButton:hover {{ background: {ACCENT_H}; }}
        """)
        btn_carregar.clicked.connect(self._carregar)

        btn_novo = QPushButton("+  NOVO PEDIDO")
        btn_novo.setStyleSheet(f"""
            QPushButton {{
                background: transparent; color: {ACCENT}; border: 1px solid {ACCENT};
                border-radius: 6px; padding: 7px 14px; font-weight: 700; font-size: 12px;
            }}
            QPushButton:hover {{ background: {ACCENT}18; }}
        """)
        btn_novo.clicked.connect(self._novo_pedido)

        topo.addWidget(titulo)
        topo.addStretch()
        topo.addWidget(self._combo_conta)
        topo.addWidget(btn_novo)
        topo.addWidget(btn_carregar)
        root.addLayout(topo)

        # Busca
        self._inp_busca_pedidos = QLineEdit()
        self._inp_busca_pedidos.setPlaceholderText("Buscar por cliente, pedido, produto, destino...")
        self._inp_busca_pedidos.setStyleSheet(f"""
            QLineEdit {{
                background: {SURFACE}; border: 1px solid {BORDER2};
                border-radius: 6px; padding: 6px 12px; color: {TEXT}; font-size: 12px;
            }}
            QLineEdit:focus {{ border-color: {ACCENT}; }}
        """)
        self._inp_busca_pedidos.textChanged.connect(self._filtrar_pedidos)
        root.addWidget(self._inp_busca_pedidos)

        self._lbl_status = QLabel("")
        self._lbl_status.setStyleSheet(f"color: {MUTED}; font-size: 11px; background: transparent;")
        root.addWidget(self._lbl_status)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")

        self._container = QWidget()
        self._container.setStyleSheet("background: transparent;")
        self._grid = QGridLayout(self._container)
        self._grid.setSpacing(10)
        self._grid.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        scroll.setWidget(self._container)
        root.addWidget(scroll)

        self._detalhes    = []
        self._expand_btns = []
        self._atualizar_contas()

    def _ajustar_saldo(self, bloco):
        conta = self._combo_conta.currentText()
        if conta == "(nenhuma conta)":
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Ajustar Saldo")
        dlg.setFixedSize(360, 240)
        dlg.setStyleSheet(DIALOG_SS)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(12)

        info = QLabel(f"{bloco['cliente']} — Pedido {bloco['pedido']}\nSaldo atual: {bloco['saldo_total']:.0f} t")
        info.setStyleSheet(f"color: {TEXT}; font-size: 12px; font-weight: 600; background: transparent;")
        lay.addWidget(info)

        tipo_lay = QHBoxLayout()
        rb_add = QRadioButton("Adicionar")
        rb_sub = QRadioButton("Diminuir")
        rb_add.setChecked(True)
        for rb in [rb_add, rb_sub]:
            rb.setFont(QFont("Arial", 11))
            rb.setStyleSheet("""
                QRadioButton { color: #e6edf3; background: transparent; spacing: 8px; }
                QRadioButton::indicator { width: 16px; height: 16px; border-radius: 8px; border: 2px solid #30363d; background: #161b22; }
                QRadioButton::indicator:checked { border-color: #238636; background: #238636; }
                QRadioButton::indicator:hover { border-color: #58a6ff; }
            """)
        tipo_lay.addWidget(rb_add)
        tipo_lay.addWidget(rb_sub)
        lay.addLayout(tipo_lay)

        lbl_v = QLabel("Quantidade (t):")
        lbl_v.setStyleSheet(f"color: {MUTED}; font-size: 10px; font-weight: 700; background: transparent;")
        inp_v = QLineEdit()
        inp_v.setPlaceholderText("Ex: 100")
        inp_v.setMinimumHeight(34)
        lay.addWidget(lbl_v)
        lay.addWidget(inp_v)

        btns = QHBoxLayout()
        bc = QPushButton("CANCELAR"); bc.setObjectName("btn_cancel")
        bo = QPushButton("CONFIRMAR"); bo.setObjectName("btn_ok")
        btns.addWidget(bc); btns.addWidget(bo)
        lay.addLayout(btns)

        def confirmar():
            try:
                qtd = float(inp_v.text().replace(",", "."))
            except Exception:
                QMessageBox.warning(dlg, "Atenção", "Informe um valor numérico.")
                return

            novo_saldo = bloco["saldo_total"] + qtd if rb_add.isChecked() else bloco["saldo_total"] - qtd
            if novo_saldo < 0:
                QMessageBox.warning(dlg, "Atenção", f"Saldo resultante seria negativo: {novo_saldo:.0f} t")
                return

            try:
                from planilha import atualizar_saldo_dados
                atualizar_saldo_dados(conta, bloco["cliente"], bloco["pedido"], bloco["produto"], novo_saldo)
                dlg.accept()
                self._lbl_status.setText(f"Saldo atualizado para {novo_saldo:.0f} t")
                self._carregar()
            except Exception as e:
                QMessageBox.critical(dlg, "Erro", str(e))

        bc.clicked.connect(dlg.reject)
        bo.clicked.connect(confirmar)
        dlg.exec()

    def _novo_pedido(self):
        conta = self._combo_conta.currentText()
        if conta == "(nenhuma conta)":
            self._lbl_status.setText("Configure uma conta Gmail primeiro.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Novo Pedido")
        dlg.setFixedSize(420, 420)
        dlg.setStyleSheet(DIALOG_SS)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(12)

        titulo = QLabel("NOVO PEDIDO")
        titulo.setStyleSheet(f"color: {TEXT}; font-size: 14px; font-weight: 700; background: transparent; margin-bottom: 6px;")
        lay.addWidget(titulo)

        campos = [
            ("Destino / Cidade", "Ex: CAMPO ALEGRE - GO"),
            ("Cliente",          "Ex: JOSE FAVA NETO"),
            ("Pedido",           "Ex: 31441"),
            ("Produto",          "Ex: UREIA"),
            ("Saldo Total (t)",  "Ex: 600"),
        ]

        inputs = []
        for label, placeholder in campos:
            grp = QVBoxLayout()
            grp.setSpacing(4)
            lbl = QLabel(label + ":")
            lbl.setStyleSheet(f"color: {MUTED}; font-size: 10px; font-weight: 700; letter-spacing: 0.5px; background: transparent;")
            inp = QLineEdit()
            inp.setPlaceholderText(placeholder)
            inp.setMinimumHeight(32)
            inp.textChanged.connect(lambda t, i=inp: _forcar_maiusculo(i, t))
            grp.addWidget(lbl)
            grp.addWidget(inp)
            lay.addLayout(grp)
            inputs.append(inp)

        btns = QHBoxLayout()
        bc = QPushButton("CANCELAR"); bc.setObjectName("btn_cancel")
        bo = QPushButton("CRIAR");    bo.setObjectName("btn_ok")
        btns.addWidget(bc); btns.addWidget(bo)
        lay.addLayout(btns)

        def criar():
            vals = [inp.text().strip().upper() for inp in inputs]
            destino, cliente, pedido, produto, saldo_str = vals
            if not all([destino, cliente, pedido, produto, saldo_str]):
                QMessageBox.warning(dlg, "Atenção", "Preencha todos os campos.")
                return
            try:
                saldo = float(saldo_str.replace(",", "."))
            except Exception:
                QMessageBox.warning(dlg, "Atenção", "Saldo total deve ser um número.")
                return
            try:
                from planilha import criar_pedido_dados
                criar_pedido_dados(conta, destino, cliente, pedido, produto, saldo)
                dlg.accept()
                self._lbl_status.setText(f"Pedido {pedido} criado com sucesso!")
                self._carregar()
            except Exception as e:
                QMessageBox.critical(dlg, "Erro", str(e))

        bc.clicked.connect(dlg.reject)
        bo.clicked.connect(criar)
        dlg.exec()

    def _atualizar_contas(self):
        self._combo_conta.clear()
        contas = _listar_contas_gmail()
        if contas:
            self._combo_conta.addItems(contas)
        else:
            self._combo_conta.addItem("(nenhuma conta)")

    def _carregar(self):
        conta = self._combo_conta.currentText()
        if conta == "(nenhuma conta)":
            self._lbl_status.setText("Configure uma conta Gmail primeiro.")
            return

        self._lbl_status.setText("Carregando...")
        QApplication.processEvents()

        try:
            blocos = carregar_blocos_dados(conta)
            self._todos_blocos = blocos
            self._renderizar(blocos)
            self._lbl_status.setText(f"{len(blocos)} pedido(s) encontrado(s)")
        except Exception as e:
            self._lbl_status.setText(f"Erro: {e}")

    def _filtrar_pedidos(self, texto):
        if not hasattr(self, '_todos_blocos'):
            return
        if not texto:
            self._renderizar(self._todos_blocos)
            return
        txt = texto.upper()
        filtrado = [
            b for b in self._todos_blocos
            if any(txt in str(v).upper() for v in [
                b.get("cliente",""), b.get("pedido",""),
                b.get("produto",""), b.get("destino",""),
                b.get("cidade",""), b.get("fabrica","")
            ])
        ]
        self._renderizar(filtrado)
        self._lbl_status.setText(f"{len(filtrado)} resultado(s)")

    def _renderizar(self, blocos):
        while self._grid.count():
            item = self._grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        self._detalhes    = []
        self._expand_btns = []

        COLS = 3
        cols_widgets = []
        for i in range(COLS):
            cw = QWidget()
            cw.setStyleSheet("background: transparent;")
            vb = QVBoxLayout(cw)
            vb.setSpacing(10)
            vb.setContentsMargins(0, 0, 0, 0)
            vb.setAlignment(Qt.AlignTop)
            cols_widgets.append((cw, vb))
            self._grid.addWidget(cw, 0, i)

        for i, b in enumerate(blocos):
            card, detalhe = self._make_bloco_card(b)
            self._detalhes.append(detalhe)
            cols_widgets[i % COLS][1].addWidget(card)

    def _make_bloco_card(self, b):
        outer = QFrame()
        outer.setStyleSheet(f"""
            QFrame {{
                background-color: {SURFACE};
                border: 1px solid {BORDER};
                border-radius: 10px;
            }}
        """)
        outer.setMinimumWidth(280)

        v_outer = QVBoxLayout(outer)
        v_outer.setContentsMargins(0, 0, 0, 0)
        v_outer.setSpacing(0)

        saldo     = b["saldo_restante"]
        total     = b["saldo_total"]
        carregado = b["total_carregado"]
        pct_saldo   = (saldo / total * 100) if total > 0 else 0
        pct_fill    = (carregado / total * 100) if total > 0 else 100
        cor_saldo   = DANGER if pct_saldo > 30 else "#e3b341" if pct_saldo > 10 else ACCENT

        # ── CABEÇALHO CLICÁVEL ─────────────────
        header = QWidget()
        header.setStyleSheet("background: transparent;")
        header.setCursor(Qt.PointingHandCursor)
        h_lay = QVBoxLayout(header)
        h_lay.setContentsMargins(14, 12, 14, 10)
        h_lay.setSpacing(6)

        top = QHBoxLayout()
        lbl_dest = QLabel(b["cliente"].upper())
        lbl_dest.setStyleSheet(f"color: {TEXT}; font-size: 12px; font-weight: 700; background: transparent;")
        lbl_dest.setWordWrap(True)
        lbl_saldo = QLabel(f"Saldo: {saldo:.0f} t")
        lbl_saldo.setStyleSheet(f"""
            color: {cor_saldo};
            background: {cor_saldo}18;
            border: 1px solid {cor_saldo}44;
            border-radius: 4px;
            font-size: 11px;
            font-weight: 700;
            padding: 2px 8px;
        """)

        btn_saldo = QPushButton("+/-")
        btn_saldo.setFixedHeight(22)
        btn_saldo.setMinimumWidth(32)
        btn_saldo.setFont(QFont("Arial", 8, QFont.Bold))
        btn_saldo.setToolTip("Ajustar saldo total")
        btn_saldo.setStyleSheet("QPushButton { background-color: #1a2a3a; border: 1px solid #58a6ff; color: #58a6ff; font-size: 9px; font-weight: 700; border-radius: 4px; padding: 0px 4px; } QPushButton:hover { background-color: #2a3a4a; }")
        p_s = QPalette(); p_s.setColor(QPalette.ButtonText, QColor("#58a6ff")); btn_saldo.setPalette(p_s)
        btn_saldo.clicked.connect(lambda _, bloco=b: self._ajustar_saldo(bloco))

        btn_expand = QPushButton("▶")
        btn_expand.setFixedSize(22, 22)
        btn_expand.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: none;
                color: {MUTED}; font-size: 10px;
            }}
            QPushButton:hover {{ color: {TEXT}; }}
        """)

        self._expand_btns.append(btn_expand)
        idx = len(self._expand_btns) - 1

        top.addWidget(lbl_dest, 1)
        top.addWidget(lbl_saldo)
        top.addWidget(btn_saldo)
        top.addWidget(btn_expand)
        h_lay.addLayout(top)

        def info_row(label, valor):
            h = QHBoxLayout()
            l = QLabel(label)
            l.setStyleSheet(f"color: {MUTED}; font-size: 10px; font-weight: 600; letter-spacing: 0.5px; background: transparent;")
            r = QLabel(str(valor).upper())
            r.setStyleSheet(f"color: {TEXT}; font-size: 11px; background: transparent;")
            r.setWordWrap(True)
            h.addWidget(l, 1)
            h.addWidget(r, 2)
            return h

        h_lay.addLayout(info_row("CIDADE",  b.get("cidade",  "")))
        h_lay.addLayout(info_row("FAZENDA", b.get("fazenda", "")))
        h_lay.addLayout(info_row("FÁBRICA", b.get("fabrica", "")))
        h_lay.addLayout(info_row("PEDIDO",  b["pedido"]))
        h_lay.addLayout(info_row("PRODUTO", b["produto"]))

        bar = QProgressBar()
        bar.setFixedHeight(5)
        bar.setRange(0, 100)
        bar.setValue(int(pct_fill))
        bar.setTextVisible(False)
        bar.setStyleSheet(f"""
            QProgressBar {{
                background: {BORDER2};
                border-radius: 2px;
                border: none;
            }}
            QProgressBar::chunk {{
                background: {cor_saldo};
                border-radius: 2px;
            }}
        """)
        h_lay.addWidget(bar)

        rodape = QHBoxLayout()
        lbl_t = QLabel(f"Total: {total:.0f} t")
        lbl_t.setStyleSheet(f"color: {MUTED}; font-size: 10px; background: transparent;")
        lbl_c = QLabel(f"Carregado: {carregado:.0f} t")
        lbl_c.setStyleSheet(f"color: {MUTED}; font-size: 10px; background: transparent;")
        rodape.addWidget(lbl_t)
        rodape.addStretch()
        rodape.addWidget(lbl_c)
        h_lay.addLayout(rodape)

        v_outer.addWidget(header)

        # ── DETALHE EXPANSÍVEL ─────────────────
        detalhe = QWidget()
        detalhe.setStyleSheet(f"background: transparent;")
        detalhe.setVisible(False)
        d_lay = QVBoxLayout(detalhe)
        d_lay.setContentsMargins(14, 0, 14, 12)
        d_lay.setSpacing(4)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"background: {BORDER}; border: none; max-height: 1px; margin-bottom: 6px;")
        d_lay.addWidget(sep)

        cab_linha = QHBoxLayout()
        for txt, w in [("DATA", 2), ("NOTA", 2), ("PLACA", 2), ("PESO", 1), ("STATUS", 3)]:
            l = QLabel(txt)
            l.setStyleSheet(f"color: #444d56; font-size: 10px; font-weight: 700; background: transparent;")
            cab_linha.addWidget(l, w)
        d_lay.addLayout(cab_linha)

        for linha in b.get("linhas", []):
            row_w = QWidget()
            row_w.setStyleSheet(f"background: transparent;")
            row_h = QHBoxLayout(row_w)
            row_h.setContentsMargins(0, 2, 0, 2)
            row_h.setSpacing(4)

            status = str(linha.get("status", "")).upper()
            if "NÃO" in status or "NAO" in status:
                cor_st = DANGER
            elif "CARREGADO" in status or "PAGO" in status:
                cor_st = ACCENT
            else:
                cor_st = MUTED

            for val, w in [
                (linha.get("data",   ""), 2),
                (linha.get("nota",   ""), 2),
                (linha.get("placa",  ""), 2),
                (linha.get("peso",   ""), 1),
            ]:
                l = QLabel(str(val))
                l.setStyleSheet(f"color: {TEXT}; font-size: 11px; background: transparent;")
                row_h.addWidget(l, w)

            lbl_st = QLabel(status)
            lbl_st.setStyleSheet(f"color: {cor_st}; font-size: 10px; font-weight: 600; background: transparent;")
            row_h.addWidget(lbl_st, 3)

            d_lay.addWidget(row_w)

        v_outer.addWidget(detalhe)

        def toggle(my_idx=idx):
            expanded = detalhe.isVisible()
            for j, d in enumerate(self._detalhes):
                try:
                    d.setVisible(False)
                except RuntimeError:
                    pass
            for j, btn in enumerate(self._expand_btns):
                try:
                    btn.setText("▶")
                except RuntimeError:
                    pass
            if not expanded:
                try:
                    detalhe.setVisible(True)
                    self._expand_btns[my_idx].setText("▼")
                except RuntimeError:
                    pass

        btn_expand.clicked.connect(toggle)
        header.mousePressEvent = lambda e: toggle()

        return outer, detalhe


class BaseWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background: transparent;")

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        topo = QHBoxLayout()
        titulo = QLabel("CONTROLE DE ORDENS")
        titulo.setStyleSheet(f"color: {TEXT}; font-size: 14px; font-weight: 700; letter-spacing: 1px; background: transparent;")

        self._combo_conta = QComboBox()
        self._combo_conta.setFixedWidth(220)
        self._combo_conta.setStyleSheet(f"""
            QComboBox {{
                background: {SURFACE}; border: 1px solid {BORDER2};
                border-radius: 6px; padding: 6px 10px; color: {TEXT}; font-size: 12px;
            }}
        """)

        self._inp_busca = QLineEdit()
        self._inp_busca.setPlaceholderText("Buscar por pagador, pedido, produto...")
        self._inp_busca.setFixedWidth(260)
        self._inp_busca.setStyleSheet(f"""
            QLineEdit {{
                background: {SURFACE}; border: 1px solid {BORDER2};
                border-radius: 6px; padding: 6px 10px; color: {TEXT}; font-size: 12px;
            }}
            QLineEdit:focus {{ border-color: {ACCENT}; }}
        """)
        self._inp_busca.textChanged.connect(self._filtrar)

        btn_carregar = QPushButton("↺  CARREGAR")
        btn_carregar.setStyleSheet(f"""
            QPushButton {{
                background: {ACCENT}; color: white; border: none;
                border-radius: 6px; padding: 7px 14px; font-weight: 700; font-size: 12px;
            }}
            QPushButton:hover {{ background: {ACCENT_H}; }}
        """)
        btn_carregar.clicked.connect(self._carregar)

        topo.addWidget(titulo)
        topo.addStretch()
        topo.addWidget(self._inp_busca)
        topo.addWidget(self._combo_conta)
        topo.addWidget(btn_carregar)
        root.addLayout(topo)

        self._lbl_status = QLabel("")
        self._lbl_status.setStyleSheet(f"color: {MUTED}; font-size: 11px; background: transparent;")
        root.addWidget(self._lbl_status)

        # Tabela
        self._tabela = QTableWidget()
        self._tabela.setStyleSheet(f"""
            QTableWidget {{
                background: {SURFACE};
                border: 1px solid {BORDER};
                border-radius: 8px;
                gridline-color: {BORDER};
                color: {TEXT};
                font-size: 11px;
            }}
            QTableWidget::item {{ padding: 4px 8px; }}
            QTableWidget::item:selected {{
                background: {ACCENT}33;
                color: {TEXT};
            }}
            QHeaderView::section {{
                background: #1c2128;
                color: {MUTED};
                font-size: 10px;
                font-weight: 700;
                letter-spacing: 0.5px;
                padding: 6px 8px;
                border: none;
                border-bottom: 1px solid {BORDER};
                border-right: 1px solid {BORDER};
            }}
            QScrollBar:vertical {{
                background: {SURFACE}; width: 8px;
            }}
            QScrollBar::handle:vertical {{
                background: {BORDER2}; border-radius: 4px;
            }}
        """)
        self._tabela.setEditTriggers(QTableWidget.NoEditTriggers)
        self._tabela.setSelectionBehavior(QTableWidget.SelectRows)
        self._tabela.setAlternatingRowColors(True)
        self._tabela.verticalHeader().setVisible(False)
        self._tabela.horizontalHeader().setStretchLastSection(True)
        self._tabela.setSortingEnabled(True)

        COLUNAS = ["DATA", "PAGADOR", "PEDIDO", "PRODUTO", "EMBALAGEM",
                   "PESO", "FRETE/EMP", "PLACA", "ORIGEM", "DESTINO", "UF", "STATUS", "", ""]
        self._tabela.setColumnCount(len(COLUNAS))
        self._tabela.setHorizontalHeaderLabels(COLUNAS)

        larguras = [90, 130, 80, 130, 90, 60, 80, 90, 120, 120, 40, 120, 50, 50]
        for i, w in enumerate(larguras):
            self._tabela.setColumnWidth(i, w)
        self._tabela.horizontalHeader().setStretchLastSection(False)
        self._tabela.setColumnWidth(12, 50)
        self._tabela.setColumnWidth(13, 50)

        root.addWidget(self._tabela)

        self._todos_dados = []
        self._linhas_editando = {}
        self._row_to_linha_planilha = {}
        self._tabela.itemClicked.connect(self._on_item_click)
        self._atualizar_contas()

    def _atualizar_contas(self):
        self._combo_conta.clear()
        contas = _listar_contas_gmail()
        self._combo_conta.addItems(contas if contas else ["(nenhuma conta)"])

    def _carregar(self):
        conta = self._combo_conta.currentText()
        if conta == "(nenhuma conta)":
            self._lbl_status.setText("Configure uma conta Gmail primeiro.")
            return

        self._lbl_status.setText("Carregando...")
        QApplication.processEvents()

        try:
            dados = carregar_base(conta)
            self._todos_dados = dados
            self._renderizar(dados)
            self._lbl_status.setText(f"{len(dados)} ordem(ns) encontrada(s)")
        except Exception as e:
            self._lbl_status.setText(f"Erro: {e}")

    def _filtrar(self, texto):
        if not texto:
            self._renderizar(self._todos_dados)
            return
        txt = texto.upper()
        filtrado = [
            d for d in self._todos_dados
            if any(txt in str(v).upper() for v in d)
        ]
        self._renderizar(filtrado)
        self._lbl_status.setText(f"{len(filtrado)} resultado(s)")

    def _renderizar(self, dados):
        self._tabela.setRowCount(0)
        self._tabela.setSortingEnabled(False)
        self._linhas_editando   = {}
        self._row_to_linha_planilha = {}

        COR_CARR = "#1a3a1a"
        COR_NAO  = "#3a1a1a"
        COR_PAGO = "#1a2a3a"
        COR_MARC = "#3a2a00"

        for linha in dados:
            r = self._tabela.rowCount()
            self._tabela.insertRow(r)
            self._tabela.setRowHeight(r, 28)

            if len(linha) > 12:
                self._row_to_linha_planilha[r] = linha[12]

            status = str(linha[11] if len(linha) > 11 else "").upper()
            if "CARREGADO" in status and "NÃO" not in status:
                bg = QColor(COR_CARR)
            elif "NÃO" in status or "NAO" in status:
                bg = QColor(COR_NAO)
            elif "PAGO" in status:
                bg = QColor(COR_PAGO)
            elif "MARCADO" in status:
                bg = QColor(COR_MARC)
            else:
                bg = None

            for c, val in enumerate(linha[:12]):
                item = QTableWidgetItem(str(val) if val is not None else "")
                item.setTextAlignment(Qt.AlignCenter)
                if bg:
                    item.setBackground(bg)
                self._tabela.setItem(r, c, item)

            self._tabela.setItem(r, 12, self._make_btn_item("EDIT", "#e3b341"))
            self._tabela.setItem(r, 13, self._make_btn_item("DEL",  "#da3633"))

        self._tabela.setSortingEnabled(True)
        try:
            self._tabela.itemClicked.disconnect()
        except Exception:
            pass
        self._tabela.itemClicked.connect(self._on_item_click)

    def _make_btn_item(self, texto, cor):
        item = QTableWidgetItem(texto)
        item.setTextAlignment(Qt.AlignCenter)
        item.setForeground(QColor(cor))
        item.setBackground(QColor("#161b22"))
        item.setFont(QFont("Segoe UI", 8, QFont.Bold))
        item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        item.setToolTip({"EDIT": "Editar linha", "DEL": "Remover linha", "OK": "Salvar"}.get(texto, ""))
        return item

    def _on_item_click(self, item):
        col = item.column()
        row = item.row()
        if col == 12:
            dados = [self._tabela.item(row, c).text() if self._tabela.item(row, c) else "" for c in range(12)]
            self._toggle_edicao(row, dados)
        elif col == 13:
            dados = [self._tabela.item(row, c).text() if self._tabela.item(row, c) else "" for c in range(12)]
            self._deletar_linha(row, dados)

    def _toggle_edicao(self, row, dados_orig):
        if row in self._linhas_editando:
            self._salvar_edicao(row, dados_orig)
            return

        self._linhas_editando[row] = dados_orig
        self._tabela.setRowHeight(row, 36)

        STATUS_OPTS = ["MARCADO", "CHEGA", "CARREGADO", "AGUARDANDO", "DESCARGA"]

        for c in range(11):
            val = str(dados_orig[c]) if c < len(dados_orig) else ""
            inp = QLineEdit(val)
            inp.setAlignment(Qt.AlignCenter)
            inp.setFrame(False)
            inp.setStyleSheet(f"""
                QLineEdit {{
                    background: #0d1117;
                    border: none;
                    border-bottom: 2px solid {ACCENT};
                    color: {TEXT};
                    font-size: 11px;
                    padding: 0px 2px;
                }}
            """)
            self._tabela.setCellWidget(row, c, inp)

        combo = QComboBox()
        combo.addItems(STATUS_OPTS)
        status_atual = str(dados_orig[11]) if len(dados_orig) > 11 else ""
        idx = combo.findText(status_atual, Qt.MatchFixedString)
        if idx >= 0:
            combo.setCurrentIndex(idx)
        combo.setStyleSheet(f"""
            QComboBox {{
                background: #0d1117;
                border: none;
                border-bottom: 2px solid {ACCENT};
                color: {TEXT};
                font-size: 11px;
                padding: 0px 2px;
            }}
            QComboBox::drop-down {{ border: none; width: 16px; }}
            QComboBox::down-arrow {{
                border-left: 3px solid transparent;
                border-right: 3px solid transparent;
                border-top: 4px solid {MUTED};
                width: 0; height: 0; margin-right: 4px;
            }}
        """)
        self._tabela.setCellWidget(row, 11, combo)

        self._tabela.setItem(row, 12, self._make_btn_item("OK", "#2ea043"))

    def _salvar_edicao(self, row, dados_orig):
        conta = self._combo_conta.currentText()
        novos = []
        for c in range(11):
            w = self._tabela.cellWidget(row, c)
            novos.append(w.text() if isinstance(w, QLineEdit) else (self._tabela.item(row, c).text() if self._tabela.item(row, c) else ""))
        combo = self._tabela.cellWidget(row, 11)
        novos.append(combo.currentText() if isinstance(combo, QComboBox) else "")

        try:
            from planilha import atualizar_linha_base
            num_linha = self._row_to_linha_planilha.get(row) or self._encontrar_linha_base(dados_orig, conta)
            if num_linha:
                atualizar_linha_base(conta, num_linha, novos)
            else:
                QMessageBox.warning(self, "Aviso", "Linha não encontrada na planilha.")
                return

            STATUS_OPTS_COR = {
                "CARREGADO": "#1a3a1a", "NÃO CARREGADO": "#3a1a1a",
                "PAGO": "#1a2a3a", "MARCADO": "#3a2a00"
            }
            status = novos[11].upper()
            bg = QColor(STATUS_OPTS_COR.get(status, "#161b22"))

            for c in range(11):
                self._tabela.removeCellWidget(row, c)
            self._tabela.removeCellWidget(row, 11)

            for c in range(12):
                item = QTableWidgetItem(novos[c])
                item.setTextAlignment(Qt.AlignCenter)
                item.setBackground(bg)
                self._tabela.setItem(row, c, item)

            self._tabela.setItem(row, 12, self._make_btn_item("EDIT", "#e3b341"))
            self._tabela.setItem(row, 13, self._make_btn_item("DEL",  "#da3633"))

            self._linhas_editando.pop(row, None)
            self._tabela.setRowHeight(row, 28)
            self._lbl_status.setText("Linha atualizada com sucesso.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))

    def _deletar_linha(self, row, dados):
        conta = self._combo_conta.currentText()
        resp = QMessageBox.question(
            self, "Confirmar", "Deseja remover esta linha da planilha?",
            QMessageBox.Yes | QMessageBox.No
        )
        if resp != QMessageBox.Yes:
            return
        try:
            from planilha import deletar_linha_base
            num_linha = self._row_to_linha_planilha.get(row) or self._encontrar_linha_base(dados, conta)
            if num_linha:
                deletar_linha_base(conta, num_linha)
            self._tabela.removeRow(row)
            self._lbl_status.setText("Linha removida com sucesso.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))

    def _encontrar_linha_base(self, dados, conta):
        from planilha import carregar_base_com_linhas
        try:
            linhas = carregar_base_com_linhas(conta)
            chave = [str(dados[i]) if i < len(dados) else "" for i in range(4)]
            for num_linha, linha in linhas:
                if [str(linha[i]) if i < len(linha) else "" for i in range(4)] == chave:
                    return num_linha
        except Exception:
            pass
        return None


def parsear_mensagem_whatsapp(texto):
    import re
    resultado = {}

    def extrair(chave):
        match = re.search(rf"^{chave}\s*:\s*(.+)", texto, re.IGNORECASE | re.MULTILINE)
        return match.group(1).strip() if match else ""

    filial = extrair("FILIAL").upper()
    resultado["empresa"] = "Agrovia" if "AGRO" in filial else "TopBrasil"

    # PAGADOR → Solicitante
    pagador = extrair("PAGADOR")
    if pagador:
        resultado["Solicitante"] = pagador

    # FABRICA → Fábrica (nome) e Origem
    fabrica = extrair("FABRICA")
    if fabrica:
        resultado["Fábrica"] = fabrica.upper().split()[0]
        resultado["Origem"]  = fabrica

    cliente = extrair("CLIENTE")
    if cliente:
        resultado["Cliente"] = cliente

    resultado["Motorista"] = extrair("MOTORISTA")
    resultado["Cavalo"]    = extrair("PLACA")
    resultado["Peso"]      = extrair("PESO")

    produto = extrair("PRODUTO")
    resultado["Produto"] = produto

    # Detecta embalagem no texto do produto
    embalagem = ""
    prod_upper = produto.upper()
    if "BIG BAG" in prod_upper:
        embalagem = "BIG BAG"
    elif "GRANEL" in prod_upper:
        embalagem = "GRANEL"
    elif "PALETIZADO" in prod_upper:
        embalagem = "PALETIZADO"
    elif "SACO 50" in prod_upper:
        embalagem = "SACO 50KG"
    elif "SACO 25" in prod_upper:
        embalagem = "SACO 25KG"
    elif "SACO 40" in prod_upper:
        embalagem = "SACO 40KG"
    if embalagem:
        resultado["Embalagem"] = embalagem

    destino_raw = extrair("DESTINO")
    uf = extrair("UF").upper()
    if destino_raw and uf:
        destino_completo = f"{destino_raw.strip()} - {uf}"
    elif destino_raw:
        destino_completo = destino_raw.strip()
    else:
        destino_completo = ""

    if destino_completo:
        sep = re.split(r"\s+(FAZ\.?\s|FAZENDA\s|SITIO\s|SÍTIO\s|CHACARA\s)", destino_completo, flags=re.IGNORECASE)
        if len(sep) >= 3:
            resultado["Destino"] = sep[0].strip()
            resultado["Fazenda"] = (sep[1] + sep[2]).strip()
        else:
            resultado["Destino"] = destino_completo
            resultado["Fazenda"] = ""
    else:
        resultado["Destino"] = resultado["Fazenda"] = ""

    fazenda = extrair("FAZENDA")
    if fazenda:
        resultado["Fazenda"] = fazenda

    pedidos = []
    for i in range(1, 5):
        p = extrair(f"PEDIDO {i}")
        if p:
            pedidos.append(p)

    if not pedidos:
        p = extrair("PEDIDO")
        if p:
            pedidos.append(p)

    produto = resultado.get("Produto", "")
    peso    = resultado.get("Peso", "")

    for idx, ped in enumerate(pedidos):
        sufixo = f" {idx + 1}" if idx > 0 else ""
        resultado[f"Pedido{sufixo}"]   = ped
        resultado[f"Produto{sufixo}"]  = produto
        resultado[f"Peso{sufixo}"]     = peso

    resultado["_num_pedidos"] = len(pedidos)

    return resultado

class GeradorThread(QThread):
    sucesso = Signal(str)
    erro    = Signal(str)

    def __init__(self, dados, pasta, email, conta_gmail=None):
        super().__init__()
        self.dados       = dados
        self.pasta       = pasta
        self.email       = email
        self.conta_gmail = conta_gmail

    def run(self):
        try:
            caminho = gerar_ordem(self.dados, self.pasta, self.email, self.conta_gmail)
            self.sucesso.emit(caminho)
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
        self._dots  = 0
        self._label = QLabel("Gerando ordem", self)
        self._label.setAlignment(Qt.AlignCenter)
        self._label.setStyleSheet(f"color: {TEXT}; font-size: 15px; font-weight: bold; background: transparent;")
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._animar)
        self.setGeometry(0, 0, 0, 0)
        self.hide()

    def showEvent(self, e):
        self._resize(); self._timer.start(400); super().showEvent(e)

    def hideEvent(self, e):
        self._timer.stop(); super().hideEvent(e)

    def _resize(self):
        if self.parent():
            self.setGeometry(self.parent().rect())
            self._label.setGeometry(self.parent().rect())

    def resizeEvent(self, e):
        self._resize(); super().resizeEvent(e)

    def _animar(self):
        self._dots = (self._dots + 1) % 4
        self._label.setText("Gerando ordem" + "." * self._dots)

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(self.rect(), QColor(13, 17, 23, 220))

class UI(QWidget):
    def __init__(self):
        super().__init__()
        self.empresa = None
        self.entradas = {}
        self._pedido_linhas = []

        self.setWindowTitle("Sistema de Ordens")
        self._setup_icon()
        self._setup_bg()
        self._build_ui()
        self._apply_style()

        self.overlay = LoadingOverlay(self)
        self.showMaximized()
        self.escolher_empresa()
        self.setar_data_hoje()

    def _setup_icon(self):
        base = Path(sys._MEIPASS) if getattr(sys, "frozen", False) else Path(__file__).parent
        ico = base / "icone.ico"
        if ico.exists():
            self.setWindowIcon(QIcon(str(ico)))

    def _setup_bg(self):
        self._bg_label = QLabel(self)
        self._bg_label.setScaledContents(True)
        self._bg_label.setGeometry(self.rect())
        self._bg_label.lower()

    def _apply_style(self):
        self.setStyleSheet(f"""
            QWidget {{
                background-color: {BG};
                color: {TEXT};
                font-family: "Segoe UI", sans-serif;
                font-size: 13px;
            }}
            QLineEdit, QComboBox, QDateEdit {{
                background-color: {SURFACE};
                border: 1px solid {BORDER2};
                border-radius: 6px;
                padding: 8px 10px;
                color: {TEXT};
                min-height: 18px;
            }}
            QLineEdit:focus, QComboBox:focus, QDateEdit:focus {{
                border-color: {ACCENT};
                background-color: #1c2128;
            }}
            QLineEdit:hover, QComboBox:hover, QDateEdit:hover {{
                border-color: {BORDER2};
            }}
            QComboBox::drop-down {{ border: none; width: 22px; }}
            QComboBox::down-arrow {{
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 5px solid {MUTED};
                width: 0;
                height: 0;
                margin-right: 8px;
            }}
            QComboBox::down-arrow:hover {{ border-top-color: {TEXT}; }}
            QComboBox QAbstractItemView {{
                background-color: {SURFACE};
                border: 1px solid {BORDER2};
                selection-background-color: {ACCENT}22;
                color: {TEXT};
                outline: none;
            }}
            QLabel {{
                color: {MUTED};
                background: transparent;
            }}
            QScrollBar:vertical {{
                background: transparent; width: 4px; border-radius: 2px;
            }}
            QScrollBar::handle:vertical {{
                background: {BORDER2}; border-radius: 2px; min-height: 20px;
            }}
            QScrollBar::handle:vertical:hover {{ background: {ACCENT}; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
            QPushButton {{
                border-radius: 6px;
                padding: 11px 16px;
                font-weight: 700;
                font-size: 13px;
                font-family: "Segoe UI", sans-serif;
            }}
            #btn_gerar {{
                background-color: {ACCENT}; color: white; border: none;
            }}
            #btn_gerar:hover {{ background-color: {ACCENT_H}; }}
            #btn_gerar:pressed {{ background-color: {ACCENT_L}; }}
            #btn_email {{
                background-color: transparent;
                border: 1px solid {ACCENT};
                color: {ACCENT};
            }}
            #btn_email:hover {{ background-color: {ACCENT}18; }}
            #btn_nova {{
                background-color: transparent;
                border: 1px solid {BORDER2};
                color: {MUTED};
            }}
            #btn_nova:hover {{
                background-color: {SURFACE};
                border-color: {BORDER2};
                color: {TEXT};
            }}
            #btn_wpp {{
                background-color: #1a3a24;
                color: #4ade80;
                border: 1px solid #238636;
            }}
            #btn_wpp:hover {{ background-color: #1f4a2e; border-color: #4ade80; }}
            #btn_add_pedido {{
                background-color: transparent;
                border: 1px dashed {BORDER2};
                color: {ACCENT};
                font-size: 12px;
                padding: 7px;
            }}
            #btn_add_pedido:hover {{
                border-color: {ACCENT};
                background-color: {ACCENT}0a;
            }}
        """)

    def _build_ui(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── SIDEBAR ──────────────────────────────
        sidebar = QWidget()
        sidebar.setFixedWidth(56)
        sidebar.setStyleSheet(f"background-color: {SURFACE}; border-right: 1px solid {BORDER};")
        sb_lay = QVBoxLayout(sidebar)
        sb_lay.setContentsMargins(0, 16, 0, 16)
        sb_lay.setSpacing(4)

        def make_nav_btn(icon, tooltip, idx):
            btn = QPushButton(icon)
            btn.setToolTip(tooltip)
            btn.setFixedSize(56, 48)
            btn.setCheckable(True)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background: transparent;
                    border: none;
                    border-left: 3px solid transparent;
                    color: {MUTED};
                    font-size: 18px;
                    border-radius: 0;
                }}
                QPushButton:hover {{ color: {TEXT}; background: {BORDER}22; }}
                QPushButton:checked {{
                    color: {ACCENT};
                    border-left: 3px solid {ACCENT};
                    background: {ACCENT}18;
                }}
            """)
            btn.clicked.connect(lambda _, i=idx: self._nav(i))
            return btn

        self._nav_btns = []
        self._nav_btns.append(make_nav_btn("📋", "Gerar Ordem", 0))
        self._nav_btns.append(make_nav_btn("🕐", "Histórico", 1))
        self._nav_btns.append(make_nav_btn("📦", "Controle de Pedidos", 2))
        self._nav_btns.append(make_nav_btn("📊", "Controle de Ordens", 3))
        self._nav_btns[0].setChecked(True)

        for b in self._nav_btns:
            sb_lay.addWidget(b)
        sb_lay.addStretch()

        root.addWidget(sidebar)

        # ── STACKED ───────────────────────────────
        self._stack = QStackedWidget()
        self._stack.setStyleSheet("background: transparent;")

        self._stack.addWidget(self._build_pagina_ordem())
        self._historico_widget = HistoricoWidget()
        self._stack.addWidget(self._historico_widget)
        self._planilha_widget = PlanilhaWidget()
        self._stack.addWidget(self._planilha_widget)
        self._base_widget = BaseWidget()
        self._stack.addWidget(self._base_widget)

        root.addWidget(self._stack, 1)

    def _nav(self, idx):
        self._stack.setCurrentIndex(idx)
        for i, b in enumerate(self._nav_btns):
            b.setChecked(i == idx)
        if idx == 1:
            self._historico_widget.recarregar()
        if idx == 2:
            self._planilha_widget._atualizar_contas()
        if idx == 3:
            self._base_widget._atualizar_contas()

    def _build_pagina_ordem(self):
        pagina = QWidget()
        pagina.setStyleSheet("background: transparent;")
        root = QVBoxLayout(pagina)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")

        container = QWidget()
        container.setStyleSheet("background: transparent;")
        inner = QVBoxLayout(container)
        inner.setContentsMargins(16, 16, 16, 16)
        inner.setSpacing(10)

        row1 = QHBoxLayout()
        row1.setSpacing(10)
        row1.setAlignment(Qt.AlignTop)

        cab = self._build_cabecalho()
        cab.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        mot = self._build_motorista()
        mot.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)

        row1.addWidget(cab, 3)
        row1.addWidget(mot, 2)
        inner.addLayout(row1)

        row2 = QHBoxLayout()
        row2.setSpacing(10)
        row2.setAlignment(Qt.AlignTop)

        car = self._build_carga()
        car.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        vei = self._build_veiculo()
        vei.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)

        row2.addWidget(car, 3)
        row2.addWidget(vei, 2)
        inner.addLayout(row2)

        row3 = QHBoxLayout()
        row3.setSpacing(10)
        row3.setAlignment(Qt.AlignTop)

        ass = self._build_assinatura()
        ass.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        row3.addWidget(ass, 3)
        row3.addLayout(self._build_botoes(), 2)
        inner.addLayout(row3)

        inner.addStretch()

        scroll.setWidget(container)
        root.addWidget(scroll)
        return pagina

    def _build_cabecalho(self):
        frame, content = make_card("Cabeçalho")
        v = QVBoxLayout(content)
        v.setSpacing(8)
        v.setContentsMargins(0, 0, 0, 0)

        r1 = QHBoxLayout(); r1.setSpacing(8)
        self.entradas["Data Apresentação"] = make_date()
        self.entradas["Fábrica"]    = make_input()
        self.entradas["Solicitante"] = make_input()
        r1.addWidget(make_field("Data", self.entradas["Data Apresentação"]), 1)
        r1.addWidget(make_field("Fábrica", self.entradas["Fábrica"]), 2)
        r1.addWidget(make_field("Solicitante", self.entradas["Solicitante"]), 2)
        v.addLayout(r1)

        def _atualizar_origem(texto):
            t = texto.upper().strip()
            origem = self.entradas["Origem"]
            if "FERTIMAXI" in t:
                origem.setText("FEIRA DE SANTANA - BA")
            elif "INTERMARITIMA" in t or ("TIMAC" in t and "CAMACARI" not in t and "CAMAÇARI" not in t):
                origem.setText("CANDEIAS - BA")
            elif "TIMAC" in t and ("CAMACARI" in t or "CAMAÇARI" in t):
                origem.setText("CAMAÇARI - BA")

        self.entradas["Fábrica"].textChanged.connect(_atualizar_origem)

        r2 = QHBoxLayout(); r2.setSpacing(8)
        self.entradas["Origem"]  = make_input()
        self.entradas["Destino"] = make_input()
        self.entradas["Cliente"] = make_input()
        r2.addWidget(make_field("Origem", self.entradas["Origem"]), 1)
        r2.addWidget(make_field("Destino", self.entradas["Destino"]), 1)
        r2.addWidget(make_field("Cliente", self.entradas["Cliente"]), 1)
        v.addLayout(r2)

        self.entradas["Fazenda"] = make_input()
        v.addWidget(make_field("Fazenda", self.entradas["Fazenda"]))

        return frame

    def _build_motorista(self):
        frame, content = make_card("Motorista")
        v = QVBoxLayout(content)
        v.setSpacing(8)
        v.setContentsMargins(0, 0, 0, 0)

        self.entradas["Motorista"] = make_input()
        v.addWidget(make_field("Nome", self.entradas["Motorista"]))

        r = QHBoxLayout(); r.setSpacing(8)
        self.entradas["CPF"]     = make_input(maiusculo=False)
        self.entradas["Contato"] = make_input(maiusculo=False)
        r.addWidget(make_field("CPF", self.entradas["CPF"]), 1)
        r.addWidget(make_field("Contato", self.entradas["Contato"]), 1)
        v.addLayout(r)
        v.addStretch()

        return frame

    def _build_carga(self):
        frame, content = make_card("Carga")
        v = QVBoxLayout(content)
        v.setSpacing(6)
        v.setContentsMargins(0, 0, 0, 0)

        header = QHBoxLayout(); header.setSpacing(8)
        for txt, stretch in [("Pedido", 3), ("Produto", 3), ("Peso", 1), ("Embalagem", 2)]:
            lbl = QLabel(txt.upper())
            lbl.setStyleSheet(f"color: #444d56; font-size: 10px; font-weight: 700; letter-spacing: 0.8px; background: transparent;")
            header.addWidget(lbl, stretch)
        v.addLayout(header)

        self._carga_container = QWidget()
        self._carga_container.setStyleSheet("background: transparent;")
        self._carga_vbox = QVBoxLayout(self._carga_container)
        self._carga_vbox.setSpacing(6)
        self._carga_vbox.setContentsMargins(0, 0, 0, 0)
        v.addWidget(self._carga_container)

        self.btn_add_pedido = QPushButton("＋  ADICIONAR PEDIDO")
        self.btn_add_pedido.setObjectName("btn_add_pedido")
        self.btn_add_pedido.setMinimumHeight(32)
        self.btn_add_pedido.clicked.connect(self._adicionar_linha_pedido)
        v.addWidget(self.btn_add_pedido)

        self._adicionar_linha_pedido()
        return frame

    def _build_veiculo(self):
        frame, content = make_card("Veículo")
        v = QVBoxLayout(content)
        v.setSpacing(8)
        v.setContentsMargins(0, 0, 0, 0)

        r1 = QHBoxLayout(); r1.setSpacing(8)
        self.entradas["Carroceria"] = make_combo(["Graneleiro", "Basculante", "Baú", "Sider", "Tanque"])
        self.entradas["Carroceria"].setCurrentIndex(-1)
        self.entradas["Carroceria"].lineEdit().setPlaceholderText("Selecione...")
        self.entradas["Cavalo"] = make_input()
        self.entradas["Cavalo"].textChanged.connect(lambda t, i=self.entradas["Cavalo"]: _forcar_maiusculo(i, t))
        r1.addWidget(make_field("Carroceria", self.entradas["Carroceria"]), 1)
        r1.addWidget(make_field("Cavalo", self.entradas["Cavalo"]), 1)
        v.addLayout(r1)

        r2 = QHBoxLayout(); r2.setSpacing(8)
        self.entradas["Carreta 1"] = make_input()
        self.entradas["Carreta 2"] = make_input()
        r2.addWidget(make_field("Carreta 1", self.entradas["Carreta 1"]))
        r2.addWidget(make_field("Carreta 2", self.entradas["Carreta 2"]))
        v.addLayout(r2)

        self.entradas["Carreta 3"] = make_input()
        v.addWidget(make_field("Carreta 3", self.entradas["Carreta 3"]))
        v.addStretch()

        return frame

    def _build_assinatura(self):
        frame, content = make_card("Assinatura")
        v = QVBoxLayout(content)
        v.setSpacing(6)
        v.setContentsMargins(0, 0, 0, 0)

        self.entradas["Assinatura"] = make_input(maiusculo=False)
        self.entradas["Assinatura"].setMinimumHeight(40)
        v.addWidget(self.entradas["Assinatura"])

        dev = QLabel("© 2026 dev by Felipe")
        dev.setAlignment(Qt.AlignCenter)
        dev.setStyleSheet(f"color: #21262d; font-size: 10px; letter-spacing: 1px; background: transparent;")
        v.addWidget(dev)
        v.addStretch()

        return frame

    def _build_botoes(self):
        v = QVBoxLayout()
        v.setSpacing(8)

        self.btn_wpp = QPushButton("📋  IMPORTAR WHATSAPP")
        self.btn_wpp.setObjectName("btn_wpp")

        self.btn1 = QPushButton("GERAR ORDEM")
        self.btn1.setObjectName("btn_gerar")

        self.btn2 = QPushButton("GERAR + EMAIL")
        self.btn2.setObjectName("btn_email")

        self.btn3 = QPushButton("NOVA ORDEM")
        self.btn3.setObjectName("btn_nova")

        for btn in [self.btn_wpp, self.btn1, self.btn2, self.btn3]:
            btn.setMinimumHeight(44)
            v.addWidget(btn)

        v.addStretch()

        self.btn_wpp.clicked.connect(self.importar_whatsapp)
        self.btn1.clicked.connect(lambda: self.executar(False))
        self.btn2.clicked.connect(lambda: self.executar(True))
        self.btn3.clicked.connect(self.nova_ordem)

        return v

    def _adicionar_linha_pedido(self):
        MAX = 4
        if len(self._pedido_linhas) >= MAX:
            return

        idx    = len(self._pedido_linhas)
        sufixo = f" {idx + 1}" if idx > 0 else ""

        row_w = QWidget()
        row_w.setStyleSheet("background: transparent;")
        row_h = QHBoxLayout(row_w)
        row_h.setContentsMargins(0, 0, 0, 0)
        row_h.setSpacing(8)

        EMBALAGENS = ["BIG BAG", "SACO 50KG", "SACO 25KG", "SACO 40KG", "GRANEL", "PALETIZADO"]

        linha = {}
        for chave, stretch in [("Pedido", 3), ("Produto", 3), ("Peso", 1), ("Embalagem", 2)]:
            if chave == "Embalagem":
                inp = QComboBox()
                inp.setEditable(True)
                inp.addItems(EMBALAGENS)
                inp.setCurrentIndex(-1)
                inp.lineEdit().setPlaceholderText("EMBALAGEM")
                comp = QCompleter(EMBALAGENS)
                comp.setCaseSensitivity(Qt.CaseInsensitive)
                inp.setCompleter(comp)
                inp.setMinimumHeight(34)
            else:
                inp = QLineEdit()
                inp.setMinimumHeight(34)
                inp.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                inp.textChanged.connect(lambda t, i=inp: _forcar_maiusculo(i, t))
            row_h.addWidget(inp, stretch)
            nome = chave + sufixo
            linha[nome] = inp
            self.entradas[nome] = inp

        btn_del = QPushButton("×")
        btn_del.setFixedSize(28, 34)
        btn_del.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                border: 1px solid #30363d;
                border-radius: 6px;
                color: #8b949e;
                font-size: 16px;
                font-weight: bold;
                padding: 0;
            }}
            QPushButton:hover {{
                background-color: #da363320;
                border-color: #da3633;
                color: #f85149;
            }}
        """)
        if idx == 0:
            btn_del.setVisible(False)
        btn_del.clicked.connect(lambda _, rw=row_w, ln=linha: self._deletar_linha_pedido(rw, ln))
        row_h.addWidget(btn_del)

        self._pedido_linhas.append((row_w, linha))
        self._carga_vbox.addWidget(row_w)

        if len(self._pedido_linhas) >= MAX:
            self.btn_add_pedido.hide()

    def _atualizar_fundo(self, empresa):
        base = Path(sys._MEIPASS) if getattr(sys, "frozen", False) else Path(__file__).parent
        nomes = {"Agrovia": "logo_agro.png", "TopBrasil": "logo_top.png"}
        arquivo = nomes.get(empresa, "")
        caminho = str(base / arquivo)
        if arquivo and os.path.exists(caminho):
            original = QPixmap(caminho)
            pequeno  = original.scaled(original.width() // 8, original.height() // 8,
                                       Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)
            borrado  = pequeno.scaled(original.width(), original.height(),
                                      Qt.KeepAspectRatioByExpanding, Qt.SmoothTransformation)
            p = QPainter(borrado)
            p.fillRect(borrado.rect(), QColor(13, 17, 23, 210))
            p.end()
            self._bg_label.setPixmap(borrado)
        else:
            self._bg_label.setPixmap(QPixmap())
        self._bg_label.setGeometry(self.rect())
        self._bg_label.lower()

    def resizeEvent(self, e):
        self._bg_label.setGeometry(self.rect())
        super().resizeEvent(e)

    def setar_data_hoje(self):
        self.entradas["Data Apresentação"].setDate(QDate.currentDate())

    def escolher_empresa(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Sistema de Ordens")
        dlg.setFixedSize(340, 230)
        dlg.setStyleSheet(DIALOG_SS)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(24, 24, 24, 24)
        lay.setSpacing(10)

        t1 = QLabel("SELECIONE A EMPRESA")
        t1.setAlignment(Qt.AlignCenter)
        t1.setStyleSheet(f"color: {TEXT}; font-size: 14px; font-weight: 700; letter-spacing: 1px;")

        t2 = QLabel("Sistema de Ordens de Carregamento")
        t2.setAlignment(Qt.AlignCenter)
        t2.setStyleSheet(f"color: {MUTED}; font-size: 11px;")

        btn_a = QPushButton("AGROVIA")
        btn_a.setObjectName("btn_agro")
        btn_a.setMinimumHeight(42)

        btn_t = QPushButton("TOPBRASIL")
        btn_t.setObjectName("btn_top")
        btn_t.setMinimumHeight(42)

        def sel(nome):
            self.empresa = nome
            cor = ACCENT if nome == "Agrovia" else DANGER
            self.btn1.setStyleSheet(f"background-color: {cor}; color: white; border: none;")
            self._atualizar_fundo(nome)
            dlg.accept()

        btn_a.clicked.connect(lambda: sel("Agrovia"))
        btn_t.clicked.connect(lambda: sel("TopBrasil"))

        lay.addWidget(t1)
        lay.addWidget(t2)
        lay.addSpacing(6)
        lay.addWidget(btn_a)
        lay.addWidget(btn_t)
        dlg.exec()

    def _deletar_linha_pedido(self, row_w, linha):
        self._pedido_linhas = [(rw, ln) for rw, ln in self._pedido_linhas if rw is not row_w]
        self._carga_vbox.removeWidget(row_w)
        row_w.deleteLater()
        for chave in linha:
            self.entradas.pop(chave, None)
        self.btn_add_pedido.show()

        for idx, (rw, ln) in enumerate(self._pedido_linhas):
            novo_sufixo = f" {idx + 1}" if idx > 0 else ""
            nova_linha = {}
            for chave_antiga, widget in list(ln.items()):
                base = chave_antiga.rsplit(" ", 1)[0] if " " in chave_antiga else chave_antiga
                novo_nome = base + novo_sufixo
                nova_linha[novo_nome] = widget
                for k in list(self.entradas.keys()):
                    if self.entradas[k] is widget:
                        del self.entradas[k]
                        break
                self.entradas[novo_nome] = widget
            self._pedido_linhas[idx] = (rw, nova_linha)

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
        if email:
            conta_gmail = self._dialog_escolher_conta()
            if conta_gmail is None:
                return
            from gerador import obter_email_fabrica, montar_email
            prev = self._dialog_preview_email(
                obter_email_fabrica(dados.get("Fábrica")),
                *montar_email(dados)
            )
            if prev is None:
                return
            dados["_email_destinatario"] = prev["destinatario"]
            dados["_email_assunto"]      = prev["assunto"]
            dados["_email_corpo"]        = prev["corpo"]

        for b in [self.btn1, self.btn2, self.btn3]:
            b.setEnabled(False)
        self.overlay.show()
        self.overlay.raise_()

        self._thread = GeradorThread(dados, pasta, email, conta_gmail)
        self._thread.sucesso.connect(self._on_sucesso)
        self._thread.erro.connect(self._on_erro)
        self._thread.start()

    def _dialog_escolher_conta(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Enviar por Gmail")
        dlg.setFixedSize(400, 230)
        dlg.setStyleSheet(DIALOG_SS)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(10)

        lay.addWidget(QLabel("Conta remetente:"))

        combo = QComboBox()
        combo.setMinimumHeight(36)
        contas = _listar_contas_gmail()
        combo.addItems(contas if contas else ["(nenhuma conta configurada)"])
        lay.addWidget(combo)

        btn_add = QPushButton("+ Adicionar conta Gmail")
        btn_add.setObjectName("btn_add")
        lay.addWidget(btn_add)

        btns = QHBoxLayout()
        bc = QPushButton("CANCELAR"); bc.setObjectName("btn_cancel")
        bo = QPushButton("ENVIAR");   bo.setObjectName("btn_ok")
        btns.addWidget(bc); btns.addWidget(bo)
        lay.addLayout(btns)

        resultado = [None]

        def adicionar():
            try:
                novo = adicionar_conta_gmail()
                combo.clear()
                combo.addItems(_listar_contas_gmail())
                idx = combo.findText(novo)
                if idx >= 0:
                    combo.setCurrentIndex(idx)
            except Exception as e:
                QMessageBox.critical(dlg, "Erro", str(e))

        def confirmar():
            c = combo.currentText()
            if c == "(nenhuma conta configurada)":
                QMessageBox.warning(dlg, "Atenção", "Adicione uma conta Gmail primeiro.")
                return
            resultado[0] = c
            dlg.accept()

        btn_add.clicked.connect(adicionar)
        bc.clicked.connect(dlg.reject)
        bo.clicked.connect(confirmar)
        dlg.exec()
        return resultado[0]

    def _dialog_preview_email(self, destinatario, assunto, corpo):
        dlg = QDialog(self)
        dlg.setWindowTitle("Prévia do email")
        dlg.setMinimumWidth(520)
        dlg.setStyleSheet(DIALOG_SS)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(8)

        lay.addWidget(QLabel("DESTINATÁRIO"))
        inp_d = QLineEdit(destinatario); inp_d.setMinimumHeight(36)
        lay.addWidget(inp_d)

        lay.addWidget(QLabel("ASSUNTO"))
        inp_a = QLineEdit(assunto); inp_a.setMinimumHeight(36)
        lay.addWidget(inp_a)

        lay.addWidget(QLabel("CORPO"))
        inp_c = QTextEdit(); inp_c.setPlainText(corpo); inp_c.setMinimumHeight(120)
        lay.addWidget(inp_c)

        btns = QHBoxLayout()
        bc = QPushButton("CANCELAR"); bc.setObjectName("btn_cancel")
        bo = QPushButton("CONFIRMAR ENVIO"); bo.setObjectName("btn_ok")
        btns.addWidget(bc); btns.addWidget(bo)
        lay.addLayout(btns)

        resultado = [None]

        def confirmar():
            resultado[0] = {
                "destinatario": inp_d.text().strip(),
                "assunto":      inp_a.text().strip(),
                "corpo":        inp_c.toPlainText().strip(),
            }
            dlg.accept()

        bc.clicked.connect(dlg.reject)
        bo.clicked.connect(confirmar)
        dlg.exec()
        return resultado[0]

    def _on_sucesso(self, caminho):
        self.overlay.hide()
        for b in [self.btn1, self.btn2, self.btn3]:
            b.setEnabled(True)

        salvar_historico(self._thread.dados, caminho)

        msg = QMessageBox(self)
        msg.setWindowTitle("Sucesso")
        msg.setText("✔  Ordem gerada com sucesso")
        msg.setIcon(QMessageBox.NoIcon)
        msg.setStyleSheet(f"""
            QMessageBox {{ background-color: {BG}; }}
            QLabel {{ color: {TEXT}; font-size: 13px; }}
            QPushButton {{
                background-color: {ACCENT}; color: white;
                border-radius: 6px; padding: 6px 18px; font-weight: 700;
            }}
        """)
        msg.exec()

        self._dialog_gravar_planilha(self._thread.dados)

    def _dialog_gravar_planilha(self, dados):
        conta = self._planilha_widget._combo_conta.currentText()
        if not conta or conta == "(nenhuma conta)":
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Gravar na Planilha?")
        dlg.setFixedSize(400, 300)
        dlg.setStyleSheet(DIALOG_SS)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(12)

        lay.addWidget(QLabel("Deseja registrar esta ordem na planilha de controle?"))

        STATUS_OPTS = ["DESCARGA", "CARREGADO", "MARCADO", "CHEGA", "AGUARDANDO"]
        lbl_st = QLabel("STATUS:")
        combo_st = QComboBox()
        combo_st.addItems(STATUS_OPTS)
        combo_st.setStyleSheet(f"background:{SURFACE}; border:1px solid {BORDER2}; border-radius:6px; padding:6px; color:{TEXT};")
        lay.addWidget(lbl_st)
        lay.addWidget(combo_st)

        pedido  = dados.get("Pedido", "")
        cliente = dados.get("Cliente", "")
        produto = dados.get("Produto", "")

        info = QLabel(f"Cliente: {cliente}\nPedido: {pedido}\nProduto: {produto}")
        info.setStyleSheet(f"color: {MUTED}; font-size: 11px;")
        lay.addWidget(info)

        btns = QHBoxLayout()
        bc = QPushButton("CANCELAR"); bc.setObjectName("btn_cancel")
        bo = QPushButton("GRAVAR");   bo.setObjectName("btn_ok")
        btns.addWidget(bc); btns.addWidget(bo)
        lay.addLayout(btns)

        def gravar():
            try:
                from planilha import gravar_carregamento_dados
                data    = str(dados.get("Data Apresentação", "")).upper()
                placa   = str(dados.get("Cavalo", "")).upper()
                peso    = str(dados.get("Peso", "")).upper()
                frete   = str(dados.get("Frete/EMP", "")).upper()
                st      = combo_st.currentText().upper()
                destino = str(dados.get("Destino", "")).upper()
                produto_val = str(dados.get("Produto", "")).upper()

                gravar_carregamento_dados(
                    conta, destino, cliente, pedido, produto_val,
                    data, placa, peso, frete, st
                )
                dlg.accept()
                QMessageBox.information(self, "Sucesso", "Registrado na planilha com sucesso!")

            except Exception as e:
                QMessageBox.critical(dlg, "Erro", str(e))

        bc.clicked.connect(dlg.reject)
        bo.clicked.connect(gravar)
        dlg.exec()

    def _on_erro(self, mensagem):
        self.overlay.hide()
        for b in [self.btn1, self.btn2, self.btn3]:
            b.setEnabled(True)
        QMessageBox.critical(self, "Erro", mensagem)

    def importar_whatsapp(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Importar WhatsApp")
        dlg.setFixedSize(480, 360)
        dlg.setStyleSheet(DIALOG_SS)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(10)

        lay.addWidget(QLabel("Cole a mensagem do WhatsApp:"))
        caixa = QTextEdit()
        caixa.setPlaceholderText("🗒️ TAG\nTRANSPORTADORA: TOP BRASIL\n...")
        lay.addWidget(caixa)

        btns = QHBoxLayout()
        bc = QPushButton("CANCELAR"); bc.setObjectName("btn_cancel")
        bo = QPushButton("PREENCHER"); bo.setObjectName("btn_ok")
        btns.addWidget(bc); btns.addWidget(bo)
        lay.addLayout(btns)

        bc.clicked.connect(dlg.reject)

        def confirmar():
            texto = caixa.toPlainText().strip()
            if not texto:
                return
            self.nova_ordem()
            self._preencher_campos(parsear_mensagem_whatsapp(texto))
            dlg.accept()

        bo.clicked.connect(confirmar)
        dlg.exec()

    def _preencher_campos(self, dados):
                                               
        num_pedidos = dados.get("_num_pedidos", 1)
        while len(self._pedido_linhas) < num_pedidos:
            self._adicionar_linha_pedido()

        campos_simples = ["Fábrica", "Cliente", "Fazenda", "Origem",
                          "Destino", "Motorista", "Cavalo", "Solicitante"]
        for campo in campos_simples:
            valor = dados.get(campo, "")
            if not valor:
                continue
            w = self.entradas.get(campo)
            if isinstance(w, QLineEdit):
                w.setText(valor)
            elif isinstance(w, QComboBox):
                w.setEditText(valor)

        for idx in range(num_pedidos):
            sufixo = f" {idx + 1}" if idx > 0 else ""
            for chave in ["Pedido", "Produto", "Peso", "Embalagem"]:
                # Embalagem detectada do produto (sem sufixo) aplica em todas as linhas
                valor = dados.get(f"{chave}{sufixo}", "") or (dados.get("Embalagem", "") if chave == "Embalagem" else "")
                if not valor:
                    continue
                w = self.entradas.get(f"{chave}{sufixo}")
                if isinstance(w, QLineEdit):
                    w.setText(valor)
                elif isinstance(w, QComboBox):
                    w.setEditText(valor)

        emp = dados.get("empresa")
        if emp:
            self.empresa = emp
            cor = ACCENT if emp == "Agrovia" else DANGER
            self.btn1.setStyleSheet(f"background-color: {cor}; color: white; border: none;")
            self._atualizar_fundo(emp)

    def nova_ordem(self):
        for v in self.entradas.values():
            if isinstance(v, QLineEdit):
                v.clear()
            elif isinstance(v, QComboBox):
                v.setCurrentIndex(0)
            elif isinstance(v, QDateEdit):
                v.setDate(QDate.currentDate())

        while len(self._pedido_linhas) > 1:
            row_w, linha = self._pedido_linhas.pop()
            self._carga_vbox.removeWidget(row_w)
            row_w.deleteLater()
            for chave in linha:
                self.entradas.pop(chave, None)

        self.btn_add_pedido.show()
        self.escolher_empresa()
        self.setar_data_hoje()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = UI()
    win.show()
    sys.exit(app.exec())