import sys
import os
import json
import datetime
import re
from pathlib import Path
from PySide6.QtWidgets import *
from PySide6.QtCore import Qt, QDate, QThread, Signal, QTimer, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QFont, QPainter, QColor, QPixmap, QIcon, QPalette
from PySide6.QtWidgets import QCompleter, QDateEdit
from gerador import gerar_ordem, _listar_contas_gmail, adicionar_conta_gmail

# ── Supabase — BASE e Controle de Pedidos ─────────────────────────
SUPABASE_URL = "https://xlirwzkmvkzldrssmhxg.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhsaXJ3emttdmt6bGRyc3NtaHhnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzU2MDgwMTksImV4cCI6MjA5MTE4NDAxOX0.ofTAEn628a-7JzF3REPj-tBcQJUrlXdfaFSbU5Ysfx4"

def _carregar_blocos_supabase():
    import urllib.request
    import urllib.parse
    import json as _json

    req = urllib.request.Request(
        f"{SUPABASE_URL}/rest/v1/carregamentos"
        f"?select=pedido,produto,pagador,destino,peso,status,data,placa"
        f"&peso=gt.0&order=pedido.asc&limit=5000",
        headers={
            "apikey":        SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
        }
    )
    with urllib.request.urlopen(req, timeout=10) as r:
        linhas = _json.loads(r.read().decode("utf-8"))

    # Agrupa por pedido + produto + cliente
    grupos = {}
    ordem  = []
    for l in linhas:
        ped  = str(l.get("pedido", "") or "").strip().upper()
        prod = str(l.get("produto", "") or "").strip().upper()
        cli  = str(l.get("pagador", "") or "").strip().upper()
        if not ped and not cli:
            continue
        key = f"{ped}||{prod}||{cli}"
        if key not in grupos:
            grupos[key] = {
                "pedido":          ped,
                "produto":         prod,
                "cliente":         cli,
                "cidade":          str(l.get("destino", "") or "").strip().upper(),
                "destino":         str(l.get("destino", "") or "").strip().upper(),
                "saldo_total":     0,
                "total_carregado": 0,
                "saldo_restante":  0,
                "pct":             0,
                "linhas":          [],
            }
            ordem.append(key)
        try:
            peso = float(l.get("peso") or 0)
        except Exception:
            peso = 0.0
        grupos[key]["total_carregado"] += peso
        grupos[key]["linhas"].append({
            "dataStr": str(l.get("data") or "")[:10],
            "placa":   str(l.get("placa") or "").strip().upper(),
            "peso":    peso,
            "status":  str(l.get("status") or "").strip().upper(),
            "nota":    "",
            "frete":   "",
        })

    # Busca saldo_total da planilha DADOS (aba PEDIDOS) se disponível,
    # senão usa total_carregado como base (saldo_restante = 0 até atualizar)
    for key in ordem:
        g = grupos[key]
        g["saldo_restante"] = g["saldo_total"] - g["total_carregado"]
        g["pct"] = (g["total_carregado"] / g["saldo_total"] * 100) if g["saldo_total"] > 0 else 0

    return [grupos[k] for k in ordem]

def _sb_request(endpoint, method="GET", body=None):
    import urllib.request, urllib.parse, json as _json
    req = urllib.request.Request(
        f"{SUPABASE_URL}/rest/v1/{endpoint}",
        data=json.dumps(body).encode() if body else None,
        headers={
            "apikey":        SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
            "Content-Type":  "application/json",
            "Prefer":        "return=representation",
        },
        method=method
    )
    with urllib.request.urlopen(req, timeout=10) as r:
        return json.loads(r.read().decode("utf-8"))

def _carregar_base_supabase(mes=None):
    """Carrega carregamentos do Supabase. mes = 'MM/YYYY' para filtrar."""
    import urllib.parse
    filtros = ["peso=gt.0", "order=data.asc,id.asc", "limit=2000"]
    if mes:
        try:
            m, y = mes.split("/")
            import datetime as _dt
            inicio = f"{y}-{m}-01"
            ultimo = (_dt.date(int(y), int(m), 1).replace(day=28) +
                      _dt.timedelta(days=4)).replace(day=1) - _dt.timedelta(days=1)
            fim    = ultimo.strftime("%Y-%m-%d")
            filtros += [f"data=gte.{inicio}", f"data=lte.{fim}"]
        except Exception:
            pass
    qs   = "&".join(filtros)
    rows = _sb_request(f"carregamentos?{qs}")
    # Converte para lista de listas compatível com o _renderizar existente
    # colunas visíveis: DATA(0) FILIAL(1) PAGADOR(2) MOTORISTA(3) PLACA(4)
    #                   DESTINO(5) UF(6) PESO(7) STATUS(8) COLOCADOR(9) [id=19]
    result = []
    for r in rows:
        data = str(r.get("data") or "")
        if len(data) == 10:
            data = f"{data[8:10]}/{data[5:7]}/{data[:4]}"
        linha = [
            data,                                    # 0  DATA
            str(r.get("filial","") or ""),           # 1  FILIAL
            str(r.get("pagador","") or ""),          # 2  PAGADOR
            str(r.get("agencia","") or ""),          # 3  AGENCIA
            str(r.get("motorista","") or ""),        # 4  MOTORISTA
            str(r.get("placa","") or ""),            # 5  PLACA
            str(r.get("fabrica","") or ""),          # 6  FABRICA
            str(r.get("destino","") or ""),          # 7  DESTINO
            str(r.get("uf","") or ""),               # 8  UF
            str(r.get("peso","") or ""),             # 9  PESO
            str(r.get("frete_emp","") or ""),        # 10 FRETE/E
            str(r.get("frete_mot","") or ""),        # 11 FRETE/M
            str(r.get("rota","") or ""),             # 12 ROTA
            str(r.get("agenciamento","") or ""),     # 13 AGENCIAMENTO
            str(r.get("status","") or ""),           # 14 STATUS
            str(r.get("pedido","") or ""),           # 15 PEDIDO
            str(r.get("produto","") or ""),          # 16 PRODUTO
            str(r.get("embalagem","") or ""),        # 17 EMBALAGEM
            str(r.get("colocador","") or ""),        # 18 COLOCADOR
            r.get("id"),                             # 19 ID (chave Supabase)
            str(r.get("pagamento","") or ""),        # 20 PAGAMENTO
        ]
        result.append(linha)
    return result

def _meses_supabase():
    """Retorna lista de meses disponíveis no Supabase."""
    rows = _sb_request("carregamentos?select=data&peso=gt.0&order=data.asc&limit=5000")
    meses = set()
    for r in rows:
        d = str(r.get("data") or "")
        if len(d) >= 7:
            meses.add(f"{d[5:7]}/{d[:4]}")
    return sorted(meses, key=lambda x: (x[3:], x[:2]))

def _atualizar_supabase_linha(sb_id, campos):
    """Atualiza campos de um registro no Supabase pelo id."""
    _sb_request(f"carregamentos?id=eq.{sb_id}", method="PATCH", body=campos)

def _deletar_supabase_linha(sb_id):
    """Deleta um registro do Supabase pelo id."""
    import urllib.request
    req = urllib.request.Request(
        f"{SUPABASE_URL}/rest/v1/carregamentos?id=eq.{sb_id}",
        headers={
            "apikey":        SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
        },
        method="DELETE"
    )
    urllib.request.urlopen(req, timeout=10)
from planilha import carregar_blocos_dados

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
    QComboBox:focus {{ border-color: {ACCENT}; }}
    QComboBox::drop-down {{ border: none; width: 20px; }}
    QComboBox::down-arrow {{
        border-left: 4px solid transparent;
        border-right: 4px solid transparent;
        border-top: 5px solid {MUTED};
        width: 0; height: 0; margin-right: 6px;
    }}
    QComboBox QAbstractItemView {{
        background-color: {SURFACE};
        border: 1px solid {BORDER2};
        color: {TEXT};
        selection-background-color: {ACCENT}33;
        selection-color: {TEXT};
        outline: none;
    }}
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
    v.setSpacing(1)
    lbl = QLabel(label_text.upper())
    lbl.setStyleSheet(f"color: {MUTED}; font-size: 9px; font-weight: 700; letter-spacing: 0.8px; background: transparent;")
    v.addWidget(lbl)
    v.addWidget(widget)
    return w

def make_input(placeholder="", maiusculo=True, max_len=None):
    inp = QLineEdit()
    inp.setMinimumHeight(32)
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
    cb.setMinimumHeight(32)
    cb.setCompleter(QCompleter(items))
    cb.completer().setCaseSensitivity(Qt.CaseInsensitive)
    cb.setStyleSheet(f"""
        QComboBox {{
            background: {SURFACE}; border: 1px solid {BORDER2};
            border-radius: 6px; padding: 6px 10px;
            color: {TEXT}; font-size: 12px;
        }}
        QComboBox:focus {{ border-color: {ACCENT}; }}
        QComboBox::drop-down {{ border: none; width: 20px; }}
        QComboBox::down-arrow {{
            border-left: 4px solid transparent;
            border-right: 4px solid transparent;
            border-top: 5px solid {MUTED};
            width: 0; height: 0; margin-right: 6px;
        }}
        QComboBox QAbstractItemView {{
            background-color: {SURFACE};
            border: 1px solid {BORDER2};
            color: {TEXT};
            selection-background-color: {ACCENT}33;
            selection-color: {TEXT};
            outline: none;
        }}
    """)
    return cb

def make_date():
    d = QDateEdit()
    d.setDisplayFormat("dd/MM/yyyy")
    d.setDate(QDate.currentDate())
    d.setCalendarPopup(True)
    d.setMinimumHeight(32)
    return d

def _forcar_maiusculo(inp, texto):
    if texto != texto.upper():
        inp.blockSignals(True)
        c = inp.cursorPosition()
        inp.setText(texto.upper())
        inp.setCursorPosition(c)
        inp.blockSignals(False)

def _formatar_placa(inp, texto):
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
    vbox.setContentsMargins(10, 8, 10, 10)
    vbox.setSpacing(4)

    lbl = QLabel(title.upper())
    lbl.setStyleSheet(f"""
        color: {MUTED};
        font-size: 9px;
        font-weight: 700;
        letter-spacing: 1.5px;
        background: transparent;
        padding-bottom: 3px;
        border-bottom: 1px solid {BORDER};
    """)
    vbox.addWidget(lbl)

    content = QWidget()
    content.setStyleSheet("background: transparent;")
    vbox.addWidget(content, 1)

    return frame, content

def _historico_path():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / "historico.json"
    return Path(__file__).parent / "historico.json"

def salvar_historico(dados, caminho_arquivo, conta=None, usuario=None, supabase_id=None):
    path = _historico_path()
    try:
        historico = json.loads(path.read_text(encoding="utf-8")) if path.exists() else []
    except Exception:
        historico = []

    # Normaliza empresa para formato canônico
    _emp_raw = str(dados.get("empresa", "") or "")
    _emp_norm = "Agrovia" if "AGRO" in _emp_raw.upper() else ("TopBrasil" if _emp_raw else "")

    registro = {
        "data_hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "usuario":   usuario or "",
        "motorista": dados.get("Motorista", ""),
        "placa":     dados.get("Cavalo", ""),
        "empresa":   _emp_norm,
        "arquivo":   caminho_arquivo,
        "dados":       {k: v for k, v in dados.items() if not k.startswith("_")},
        "supabase_id": supabase_id,
    }
    historico.insert(0, registro)
    historico = historico[:200]
    path.write_text(json.dumps(historico, ensure_ascii=False, indent=2), encoding="utf-8")

    # gravar_historico_planilha removido — histórico mantido localmente

def carregar_historico():
    path = _historico_path()
    if not path.exists():
        return []
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []

def carregar_historico_supabase(limite=200):
    """Carrega histórico de ordens direto do Supabase, ordenado do mais recente."""
    try:
        req = __import__("urllib.request", fromlist=["Request"]).Request(
            f"{SUPABASE_URL}/rest/v1/carregamentos"
            f"?select=id,criado_em,data,filial,pagador,motorista,placa,fabrica,"
            f"destino,uf,peso,status,pedido,produto,colocador,usuario,ativo,observacao"
            f"&order=id.desc&limit={limite}",
            headers={
                "apikey":        SUPABASE_KEY,
                "Authorization": f"Bearer {SUPABASE_KEY}",
            }
        )
        import urllib.request as _ureq
        with _ureq.urlopen(req, timeout=10) as r:
            return json.loads(r.read().decode("utf-8"))
    except Exception:
        return []

USUARIOS = {
    "FELIPE":   "Felipe Costa",
    "MARCOS":   "Marcos Silva",
    "ANA":      "Ana Souza",
    "RAFAEL":   "Rafael Lima",
}

def _contas_empresa_path():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / "contas_empresa.json"
    return Path(__file__).parent / "contas_empresa.json"

def carregar_contas_empresa():
    """Retorna dict {empresa: conta_gmail}. Ex: {'Agrovia': 'agro@gmail.com'}"""
    path = _contas_empresa_path()
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}

def salvar_contas_empresa(mapa):
    path = _contas_empresa_path()
    path.write_text(json.dumps(mapa, ensure_ascii=False, indent=2), encoding="utf-8")


def _usuarios_path():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / "usuarios.json"
    return Path(__file__).parent / "usuarios.json"

def carregar_usuarios():
    """Carrega usuários do arquivo JSON. Se não existir, usa o dicionário padrão."""
    path = _usuarios_path()
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            pass
    # Cria o arquivo com os usuários padrão na primeira execução
    path.write_text(json.dumps(USUARIOS, ensure_ascii=False, indent=2), encoding="utf-8")
    return dict(USUARIOS)




class HistoricoWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background: transparent;")
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(10)

        # ── Topo ──
        topo = QHBoxLayout()
        titulo = QLabel("HISTÓRICO DE ORDENS")
        titulo.setStyleSheet(f"color: {TEXT}; font-size: 14px; font-weight: 700; letter-spacing: 1px; background: transparent;")

        self._inp_busca = QLineEdit()
        self._inp_busca.setPlaceholderText("Buscar motorista, placa, pedido, cliente...")
        self._inp_busca.setFixedWidth(280)
        self._inp_busca.setStyleSheet(f"""
            QLineEdit {{
                background: {SURFACE}; border: 1px solid {BORDER2};
                border-radius: 6px; padding: 6px 10px; color: {TEXT}; font-size: 12px;
            }}
            QLineEdit:focus {{ border-color: {ACCENT}; }}
        """)
        self._inp_busca.textChanged.connect(self._filtrar)

        btn_atualizar = QPushButton("↺  ATUALIZAR")
        btn_atualizar.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: 1px solid {BORDER2};
                border-radius: 6px; color: {MUTED}; padding: 6px 12px; font-size: 12px;
            }}
            QPushButton:hover {{ color: {TEXT}; border-color: {ACCENT}; }}
        """)
        btn_atualizar.clicked.connect(self.recarregar)
        topo.addWidget(titulo)
        topo.addStretch()
        topo.addWidget(self._inp_busca)
        topo.addWidget(btn_atualizar)
        root.addLayout(topo)

        # ── Tabela ──
        self._tabela = QTableWidget()
        self._tabela.setStyleSheet(f"""
            QTableWidget {{
                background: {SURFACE}; border: 1px solid {BORDER};
                border-radius: 8px; gridline-color: {BORDER};
                color: {TEXT}; font-size: 11px;
            }}
            QTableWidget::item {{ padding: 6px 10px; }}
            QTableWidget::item:selected {{ background: {ACCENT}22; color: {TEXT}; }}
            QHeaderView::section {{
                background: #1c2128; color: {MUTED}; font-size: 10px;
                font-weight: 700; padding: 8px 10px; border: none;
                border-bottom: 1px solid {BORDER}; border-right: 1px solid {BORDER};
            }}
            QScrollBar:vertical {{ background: {SURFACE}; width: 8px; }}
            QScrollBar::handle:vertical {{ background: {BORDER2}; border-radius: 4px; }}
        """)
        self._tabela.setEditTriggers(QTableWidget.NoEditTriggers)
        self._tabela.setSelectionBehavior(QTableWidget.SelectRows)
        self._tabela.setAlternatingRowColors(False)
        self._tabela.verticalHeader().setVisible(False)
        self._tabela.horizontalHeader().setStretchLastSection(True)
        self._tabela.setSortingEnabled(True)

        COLUNAS = ["#", "DATA/HORA", "EMPRESA", "MOTORISTA", "PLACA",
                   "CLIENTE", "PEDIDO", "PRODUTO", "PESO", "DESTINO",
                   "COLOCADOR", "STATUS", "USUÁRIO", "EDITAR"]
        self._tabela.setColumnCount(len(COLUNAS))
        self._tabela.setHorizontalHeaderLabels(COLUNAS)
        larguras = [45, 115, 75, 140, 80, 130, 70, 150, 55, 110, 90, 100, 70, 60]
        for i, w in enumerate(larguras):
            self._tabela.setColumnWidth(i, w)
        self._tabela.horizontalHeader().setStretchLastSection(False)

        root.addWidget(self._tabela)
        self._todos = []
        self.recarregar()

    def recarregar(self):
        self._tabela.setSortingEnabled(False)
        self._tabela.setRowCount(0)
        # Tenta Supabase primeiro, fallback para JSON local
        registros = carregar_historico_supabase()
        if registros:
            self._todos = registros
            self._renderizar_supabase(registros)
        else:
            local = carregar_historico()
            self._todos = local
            self._renderizar_local(local)
        self._tabela.setSortingEnabled(True)

    def _renderizar_supabase(self, registros):
        STATUS_COR = {
            "CARREGADO": "#1a3a1a", "PAGO": "#1a2a3a", "MARCADO": "#3a2a00",
            "AGUARDANDO": "#1e2200", "DESCARGA": "#1a0a2a",
            "CHEGA": "#0a1a2a", "A CAMINHO": "#0a1a2a",
            "DESISTIU": "#3a0a0a", "ALTERADO": "#3a0a0a",
        }
        for r in registros:
            row = self._tabela.rowCount()
            self._tabela.insertRow(row)
            self._tabela.setRowHeight(row, 26)

            sb_id   = r.get("id", "")
            filial  = str(r.get("filial", "") or "").upper()
            empresa = "Agrovia" if "AGRO" in filial else "TopBrasil"
            criado  = str(r.get("criado_em", "") or "")
            if len(criado) >= 16:
                data_hora = criado[8:10]+"/"+criado[5:7]+"/"+criado[:4]+" "+criado[11:16]
            else:
                data_hora = str(r.get("data", "") or "")
            status  = str(r.get("status", "") or "").upper()
            ativo   = r.get("ativo", True)
            inativo = (not ativo) or status in ("DESISTIU", "ALTERADO")
            obs     = str(r.get("observacao", "") or "")
            peso    = r.get("peso", "")
            try: peso = f"{float(peso):.0f} t"
            except: peso = str(peso)

            cor_emp = ACCENT if empresa == "Agrovia" else DANGER
            bg = QColor(STATUS_COR.get(status, "#161b22"))
            fg = QColor("#f85149") if inativo else QColor(TEXT)

            vals = [
                str(sb_id), data_hora, empresa,
                str(r.get("motorista","") or "").title(),
                str(r.get("placa","") or ""),
                str(r.get("pagador","") or ""),
                str(r.get("pedido","") or ""),
                str(r.get("produto","") or ""),
                peso,
                str(r.get("destino","") or ""),
                str(r.get("colocador","") or ""),
                status + (f" → {obs}" if obs and inativo else ""),
                str(r.get("usuario","") or ""),
                "EDITAR",
            ]

            for col, val in enumerate(vals):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignCenter if col != 3 else Qt.AlignLeft | Qt.AlignVCenter)
                item.setBackground(bg)
                item.setForeground(fg)
                if col == 2:  # empresa
                    item.setForeground(QColor(cor_emp))
                if col == 13:  # EDITAR
                    item.setForeground(QColor("#e3b341"))
                    item.setFont(QFont("Segoe UI", 8, QFont.Bold))
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                self._tabela.setItem(row, col, item)

            # Guarda dados completos para edição
            self._tabela.item(row, 0).setData(Qt.UserRole, r)

        try:
            self._tabela.itemClicked.disconnect()
        except Exception:
            pass
        self._tabela.itemClicked.connect(self._on_click)

    def _renderizar_local(self, registros):
        for r in registros:
            row = self._tabela.rowCount()
            self._tabela.insertRow(row)
            self._tabela.setRowHeight(row, 26)
            empresa = r.get("empresa", "")
            cor_emp = ACCENT if empresa == "Agrovia" else DANGER
            dados   = r.get("dados", {})
            vals = [
                str(r.get("supabase_id","") or ""),
                r.get("data_hora",""),
                empresa,
                r.get("motorista","—").title(),
                r.get("placa","—"),
                dados.get("Cliente","") or dados.get("Pagador",""),
                dados.get("Pedido",""),
                dados.get("Produto",""),
                dados.get("Peso Total","") or dados.get("Peso",""),
                dados.get("Destino",""),
                dados.get("Colocador",""),
                "",
                r.get("usuario",""),
                "EDITAR",
            ]
            for col, val in enumerate(vals):
                item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignCenter if col != 3 else Qt.AlignLeft | Qt.AlignVCenter)
                if col == 2: item.setForeground(QColor(cor_emp))
                if col == 13:
                    item.setForeground(QColor("#e3b341"))
                    item.setFont(QFont("Segoe UI", 8, QFont.Bold))
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                self._tabela.setItem(row, col, item)
            self._tabela.item(row, 0).setData(Qt.UserRole, r)
        try:
            self._tabela.itemClicked.disconnect()
        except Exception:
            pass
        self._tabela.itemClicked.connect(self._on_click)

    def _filtrar(self, texto):
        txt = texto.upper()
        for row in range(self._tabela.rowCount()):
            visivel = not txt or any(
                txt in (self._tabela.item(row, c).text().upper() if self._tabela.item(row, c) else "")
                for c in range(self._tabela.columnCount())
            )
            self._tabela.setRowHidden(row, not visivel)

    def _on_click(self, item):
        if item.column() == 13:
            r = self._tabela.item(item.row(), 0).data(Qt.UserRole)
            if r:
                self._reeditar_ordem(r)

    # _make_card substituído por tabela — não mais necessário

    def _reeditar_ordem(self, registro):
        """Carrega dados do histórico/Supabase no formulário para reeditar."""
        # Suporte a registro do Supabase (tem 'motorista' direto) ou local (tem 'dados')
        dados = registro.get("dados")
        supabase_id = registro.get("supabase_id") or registro.get("id")

        if not dados and not registro.get("motorista"):
            QMessageBox.warning(self, "Aviso",
                "Este registro nao possui dados suficientes para reeditar.")
            return

        # Se veio do Supabase sem dados completos, monta dados básicos
        if not dados and registro.get("motorista"):
            dados = {
                "empresa":          "Agrovia" if "AGRO" in str(registro.get("filial","")).upper() else "TopBrasil",
                "Motorista":        str(registro.get("motorista","") or ""),
                "Cavalo":           str(registro.get("placa","") or ""),
                "Pagador":          str(registro.get("pagador","") or ""),
                "Cliente":          str(registro.get("pagador","") or ""),
                "Fábrica":          str(registro.get("fabrica","") or ""),
                "Destino":          str(registro.get("destino","") or ""),
                "UF":               str(registro.get("uf","") or ""),
                "Pedido":           str(registro.get("pedido","") or ""),
                "Produto":          str(registro.get("produto","") or ""),
                "Embalagem":        str(registro.get("embalagem","") or ""),
                "Colocador":        str(registro.get("colocador","") or ""),
                "Pagamento":        str(registro.get("pagamento","") or ""),
                "Peso Total":       str(registro.get("peso","") or ""),
            }

        arquivo_antigo_xlsx = registro.get("arquivo", "")
        arquivo_antigo_pdf  = arquivo_antigo_xlsx.replace(".xlsx", ".pdf") if arquivo_antigo_xlsx else ""

        dlg = QDialog(self)
        dlg.setWindowTitle("Reeditar Ordem")
        dlg.setFixedSize(420, 200)
        dlg.setStyleSheet(DIALOG_SS)
        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(12)

        motorista = dados.get("Motorista", "")
        placa     = dados.get("Cavalo", "")
        lbl = QLabel(f"Reeditar ordem de <b>{motorista}</b> — placa <b>{placa}</b>")
        lbl.setWordWrap(True)
        lay.addWidget(lbl)

        lbl2 = QLabel("Os arquivos antigos (xlsx e pdf) serão deletados e novos gerados.")
        lbl2.setStyleSheet(f"color: #e3b341; font-size: 11px; background: transparent;")
        lbl2.setWordWrap(True)
        lay.addWidget(lbl2)

        btns = QHBoxLayout()
        bc = QPushButton("CANCELAR"); bc.setObjectName("btn_cancel")
        bo = QPushButton("CONTINUAR"); bo.setObjectName("btn_ok")
        btns.addWidget(bc); btns.addWidget(bo)
        lay.addLayout(btns)
        bc.clicked.connect(dlg.reject)

        def continuar():
            dlg.accept()
            # Encontra a janela principal (UI) para navegar e preencher
            ui = None
            w = self.parent()
            while w:
                if hasattr(w, "_nav") and hasattr(w, "_preencher_campos"):
                    ui = w
                    break
                w = w.parent() if hasattr(w, "parent") else None

            if ui is None:
                QMessageBox.warning(self, "Aviso", "Nao foi possivel navegar para o formulario.")
                return

            # Define empresa correta da ordem
            empresa_ordem = dados.get("empresa", "")
            if empresa_ordem and hasattr(ui, "_aplicar_empresa"):
                ui._aplicar_empresa(empresa_ordem)
            # Preenche formulário com dados da ordem antiga
            ui._preencher_campos(dados)
            # Navega para aba de geração de ordem
            ui._nav(0)
            # Guarda id anterior para marcar como ALTERADO após gerar nova ordem
            id_anterior = supabase_id or registro.get("supabase_id") or registro.get("id")
            if id_anterior:
                ui._reeditando_id_anterior = id_anterior
            # Guarda caminhos antigos para deletar após gerar
            if arquivo_antigo_xlsx:
                ui._arquivos_para_deletar = [arquivo_antigo_xlsx, arquivo_antigo_pdf]

        bo.clicked.connect(continuar)
        dlg.exec()

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

        btn_migrar = QPushButton("⇄  MIGRAR DA BASE")
        btn_migrar.setToolTip("Importa carregamentos da planilha BASE para as abas PEDIDOS e DADOS")
        btn_migrar.setStyleSheet(f"""
            QPushButton {{
                background: transparent; color: #58a6ff; border: 1px solid #58a6ff;
                border-radius: 6px; padding: 7px 14px; font-weight: 700; font-size: 12px;
            }}
            QPushButton:hover {{ background: #58a6ff18; }}
        """)
        btn_migrar.clicked.connect(self._migrar_da_base)

        topo.addWidget(titulo)
        topo.addStretch()
        topo.addWidget(self._combo_conta)
        topo.addWidget(btn_migrar)
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

        saldo_restante = bloco.get("saldo_restante", 0)
        total_carregado = bloco.get("total_carregado", 0)

        dlg = QDialog(self)
        dlg.setWindowTitle("Ajustar Saldo")
        dlg.setFixedSize(380, 310)
        dlg.setStyleSheet(DIALOG_SS)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(10)

        # ── Info do pedido ──
        titulo = QLabel(f"{bloco['cliente']}")
        titulo.setStyleSheet(f"color: {TEXT}; font-size: 13px; font-weight: 700; background: transparent;")
        subtitulo = QLabel(f"Pedido {bloco['pedido']}  ·  {bloco['produto']}")
        subtitulo.setStyleSheet(f"color: {MUTED}; font-size: 11px; background: transparent;")
        lay.addWidget(titulo)
        lay.addWidget(subtitulo)

        # ── Card com saldos atuais ──
        card = QFrame()
        card.setStyleSheet(f"QFrame {{ background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 8px; }}")
        card_lay = QHBoxLayout(card)
        card_lay.setContentsMargins(16, 10, 16, 10)
        card_lay.setSpacing(24)
        for label, valor, cor in [
            ("JÁ CARREGADO",  f"{total_carregado:.0f} t",  MUTED),
            ("SALDO RESTANTE", f"{saldo_restante:.0f} t",  ACCENT if saldo_restante > 0 else DANGER),
        ]:
            col = QVBoxLayout()
            l = QLabel(label)
            l.setStyleSheet(f"color: {MUTED}; font-size: 9px; font-weight: 700; letter-spacing: 0.5px; background: transparent;")
            v = QLabel(valor)
            v.setStyleSheet(f"color: {cor}; font-size: 16px; font-weight: 700; background: transparent;")
            col.addWidget(l)
            col.addWidget(v)
            card_lay.addLayout(col)
        card_lay.addStretch()
        lay.addWidget(card)

        # ── Input do novo saldo restante ──
        lbl_v = QLabel("NOVO SALDO RESTANTE (t):")
        lbl_v.setStyleSheet(f"color: {MUTED}; font-size: 9px; font-weight: 700; letter-spacing: 0.5px; background: transparent;")
        inp_v = QLineEdit()
        inp_v.setPlaceholderText(f"Ex: {saldo_restante:.0f}")
        inp_v.setMinimumHeight(36)
        lay.addWidget(lbl_v)
        lay.addWidget(inp_v)

        lbl_aviso = QLabel("")
        lbl_aviso.setStyleSheet(f"color: {MUTED}; font-size: 10px; background: transparent;")
        lay.addWidget(lbl_aviso)

        def _preview(texto):
            try:
                novo = float(texto.replace(",", "."))
                lbl_aviso.setText(
                    f"saldo_total será gravado como {novo + total_carregado:.0f} t  "
                    f"({novo:.0f} restante + {total_carregado:.0f} já carregado)"
                )
            except Exception:
                lbl_aviso.setText("")

        inp_v.textChanged.connect(_preview)

        btns = QHBoxLayout()
        bc = QPushButton("CANCELAR"); bc.setObjectName("btn_cancel")
        bo = QPushButton("CONFIRMAR"); bo.setObjectName("btn_ok")
        btns.addWidget(bc); btns.addWidget(bo)
        lay.addLayout(btns)

        def confirmar():
            try:
                novo_restante = float(inp_v.text().replace(",", "."))
            except Exception:
                QMessageBox.warning(dlg, "Atenção", "Informe um valor numérico.")
                return
            if novo_restante < 0:
                QMessageBox.warning(dlg, "Atenção", "Saldo restante não pode ser negativo.")
                return
            try:
                from planilha import atualizar_saldo_dados
                saldo_total_gravado, carregado = atualizar_saldo_dados(
                    conta, bloco["cliente"], bloco["pedido"], bloco["produto"], novo_restante
                )
                dlg.accept()
                self._lbl_status.setText(
                    f"Saldo de {bloco['pedido']} atualizado — restante: {novo_restante:.0f} t"
                )
                self._carregar()
            except Exception as e:
                QMessageBox.critical(dlg, "Erro", str(e))

        bc.clicked.connect(dlg.reject)
        bo.clicked.connect(confirmar)
        dlg.exec()

    def _remover_pedido(self, bloco):
        conta = self._combo_conta.currentText()
        if conta == "(nenhuma conta)":
            return

        cliente  = bloco.get("cliente", "")
        pedido   = bloco.get("pedido", "")
        produto  = bloco.get("produto", "")
        n_carrs  = len(bloco.get("linhas", []))

        msg = QMessageBox(self)
        msg.setWindowTitle("Remover Pedido")
        msg.setIcon(QMessageBox.NoIcon)
        msg.setText(
            f"Deseja remover este pedido permanentemente?\n\n"
            f"Cliente: {cliente}\n"
            f"Pedido: {pedido}\n"
            f"Produto: {produto}\n\n"
            f"Isso também removerá {n_carrs} carregamento(s) vinculado(s)."
        )
        msg.setStyleSheet(f"""
            QMessageBox {{ background-color: {BG}; }}
            QLabel {{ color: {TEXT}; font-size: 13px; }}
            QPushButton {{
                border-radius: 6px; padding: 7px 18px;
                font-weight: 700; font-size: 12px; min-width: 80px;
            }}
        """)
        btn_sim = msg.addButton("REMOVER", QMessageBox.AcceptRole)
        btn_sim.setStyleSheet(f"background-color: {DANGER}; color: white; border: none;")
        btn_nao = msg.addButton("CANCELAR", QMessageBox.RejectRole)
        btn_nao.setStyleSheet(f"background-color: transparent; border: 1px solid {BORDER2}; color: {MUTED};")
        msg.exec()

        if msg.clickedButton() != btn_sim:
            return

        try:
            from planilha import remover_pedido_dados
            resultado = remover_pedido_dados(conta, cliente, pedido, produto)
            self._lbl_status.setText(
                f"Pedido {pedido} removido — "
                f"{resultado['carregamentos_removidos']} carregamento(s) excluído(s)"
            )
            self._carregar()
        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))

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

    def _migrar_da_base(self):
        conta = self._combo_conta.currentText()
        if conta == "(nenhuma conta)":
            self._lbl_status.setText("Configure uma conta Gmail primeiro.")
            return

        # ── 1. Busca abas disponíveis ──────────────────────────────
        self._lbl_status.setText("Buscando abas disponíveis...")
        QApplication.processEvents()
        try:
            from planilha import listar_abas_base, migrar_base_para_dados
            import re as _re
            todas = listar_abas_base(conta)
            abas_base = [a for a in todas if _re.match(r"BASE\s+\d{2}/?\d{4}", a, _re.IGNORECASE)]
            if not abas_base:
                QMessageBox.warning(self, "Aviso", "Nenhuma aba BASE MM/AAAA encontrada na planilha.")
                return
        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))
            return

        # ── 2. Diálogo de seleção de abas ─────────────────────────
        dlg = QDialog(self)
        dlg.setWindowTitle("Selecionar abas para migrar")
        dlg.setFixedWidth(340)
        dlg.setStyleSheet(DIALOG_SS)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(10)

        lbl = QLabel("Selecione as abas que deseja migrar:")
        lbl.setStyleSheet(f"color: {MUTED}; font-size: 11px; background: transparent;")
        lay.addWidget(lbl)

        # Checkboxes — marca a mais recente por padrão
        checks = []
        for i, aba in enumerate(abas_base):
            chk = QCheckBox(aba)
            chk.setStyleSheet(f"color: {TEXT}; font-size: 12px; background: transparent;")
            chk.setChecked(i == len(abas_base) - 1)  # marca a última (mais recente)
            lay.addWidget(chk)
            checks.append((chk, aba))

        # Botões selecionar tudo / nenhum
        sel_row = QHBoxLayout()
        btn_todos   = QPushButton("Selecionar todas")
        btn_nenhum  = QPushButton("Desmarcar todas")
        for b in [btn_todos, btn_nenhum]:
            b.setStyleSheet(f"QPushButton {{ background: transparent; border: 1px solid {BORDER2}; color: {MUTED}; border-radius: 4px; padding: 4px 8px; font-size: 10px; }} QPushButton:hover {{ color: {TEXT}; }}")
        btn_todos.clicked.connect(lambda: [c.setChecked(True)  for c, _ in checks])
        btn_nenhum.clicked.connect(lambda: [c.setChecked(False) for c, _ in checks])
        sel_row.addWidget(btn_todos)
        sel_row.addWidget(btn_nenhum)
        lay.addLayout(sel_row)

        btns = QHBoxLayout()
        bc = QPushButton("CANCELAR"); bc.setObjectName("btn_cancel")
        bo = QPushButton("MIGRAR");   bo.setObjectName("btn_ok")
        btns.addWidget(bc); btns.addWidget(bo)
        lay.addLayout(btns)

        abas_selecionadas = []

        def confirmar():
            selecionadas = [aba for chk, aba in checks if chk.isChecked()]
            if not selecionadas:
                QMessageBox.warning(dlg, "Atenção", "Selecione pelo menos uma aba.")
                return
            abas_selecionadas.extend(selecionadas)
            dlg.accept()

        bc.clicked.connect(dlg.reject)
        bo.clicked.connect(confirmar)
        if dlg.exec() != QDialog.Accepted or not abas_selecionadas:
            return

        # ── 3. Executa migração ────────────────────────────────────
        self._lbl_status.setText("Migrando... aguarde.")
        QApplication.processEvents()

        try:
            def progresso(msg):
                self._lbl_status.setText(msg)
                QApplication.processEvents()

            resultado = migrar_base_para_dados(
                conta,
                abas_filtro=abas_selecionadas,
                callback_progresso=progresso
            )

            abas_str = ", ".join(resultado["abas_lidas"])
            msg = (
                f"Migração concluída!\n\n"
                f"Abas lidas: {abas_str}\n"
                f"Total lido: {resultado['total_lido']} carregamento(s)\n"
                f"Pedidos criados: {resultado['pedidos_criados']}\n"
                f"Carregamentos migrados: {resultado['carregamentos_migrados']}\n\n"
                f"Ajuste o SALDO_TOTAL de cada pedido usando o botão +/-."
            )
            QMessageBox.information(self, "Migração concluída", msg)
            self._lbl_status.setText(
                f"Migração OK — {resultado['pedidos_criados']} pedido(s), "
                f"{resultado['carregamentos_migrados']} carregamento(s) importado(s)"
            )
            self._carregar()
        except Exception as e:
            QMessageBox.critical(self, "Erro na migração", str(e))
            self._lbl_status.setText(f"Erro: {e}")

    def _atualizar_contas(self):
        self._combo_conta.clear()
        contas = _listar_contas_gmail()
        if contas:
            self._combo_conta.addItems(contas)
        else:
            self._combo_conta.addItem("(nenhuma conta)")

    def _carregar(self):
        self._lbl_status.setText("Carregando do banco de dados...")
        QApplication.processEvents()
        try:
            blocos = _carregar_blocos_supabase()
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
        top.setSpacing(4)
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

        # QLabel clicável — renderiza corretamente com background transparent no PySide6
        class IconLabel(QLabel):
            def __init__(self, texto, cor_normal, cor_hover, callback, parent=None):
                super().__init__(texto, parent)
                self._cor    = cor_normal
                self._hover  = cor_hover
                self._cb     = callback
                self.setFixedSize(22, 22)
                self.setAlignment(Qt.AlignCenter)
                self.setCursor(Qt.PointingHandCursor)
                self._aplicar(False)
            def _aplicar(self, hover):
                cor = self._hover if hover else self._cor
                self.setStyleSheet(
                    f"QLabel {{ background: transparent; border: none;"
                    f" color: {cor}; font-size: 13px; }}"
                )
            def enterEvent(self, e):  self._aplicar(True)
            def leaveEvent(self, e):  self._aplicar(False)
            def mousePressEvent(self, e):
                if e.button() == Qt.LeftButton:
                    self._cb()

        bloco_ref = b
        btn_remover = IconLabel("🗑", MUTED, DANGER,
                                lambda b=bloco_ref: self._remover_pedido(b))
        btn_remover.setToolTip("Remover pedido e todos os carregamentos")

        btn_expand = IconLabel("▶", MUTED, TEXT, lambda: None)
        self._expand_btns.append(btn_expand)
        idx = len(self._expand_btns) - 1

        top.addWidget(lbl_dest, 1)
        top.addWidget(lbl_saldo)
        top.addWidget(btn_saldo)
        top.addWidget(btn_remover)
        top.addWidget(btn_expand)
        h_lay.addLayout(top)

        # callback preenchido depois que toggle() é definido
        _expand_callbacks = [None]

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
                    btn._aplicar(False)
                except RuntimeError:
                    pass
            if not expanded:
                try:
                    detalhe.setVisible(True)
                    self._expand_btns[my_idx].setText("▼")
                    self._expand_btns[my_idx]._aplicar(False)
                except RuntimeError:
                    pass

        btn_expand._cb = toggle
        header.mousePressEvent = lambda e: toggle()

        return outer, detalhe


class BaseWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background: transparent;")

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        COMBO_SS = f"""
            QComboBox {{
                background: {SURFACE}; border: 1px solid {BORDER2};
                border-radius: 6px; padding: 6px 10px; color: {TEXT}; font-size: 12px;
            }}
            QComboBox::drop-down {{ border: none; width: 18px; }}
            QComboBox::down-arrow {{
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 5px solid {MUTED};
                width: 0; height: 0; margin-right: 5px;
            }}
        """

        topo = QHBoxLayout()
        titulo = QLabel("CONTROLE DE ORDENS")
        titulo.setStyleSheet(f"color: {TEXT}; font-size: 14px; font-weight: 700; letter-spacing: 1px; background: transparent;")

        # combo_conta mantido para compatibilidade com _on_sucesso
        self._combo_conta = QComboBox()
        self._combo_conta.setFixedWidth(210)
        self._combo_conta.setStyleSheet(COMBO_SS)

        # Select de mês
        self._combo_mes = QComboBox()
        self._combo_mes.setFixedWidth(150)
        self._combo_mes.setStyleSheet(COMBO_SS)
        self._combo_mes.setPlaceholderText("Mês...")
        self._combo_mes.setToolTip("Selecione o mês para visualizar")

        btn_abas = QPushButton("↺")
        btn_abas.setFixedSize(32, 32)
        btn_abas.setToolTip("Atualizar lista de meses")
        btn_abas.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: 1px solid {BORDER2};
                border-radius: 6px; color: {MUTED}; font-size: 14px;
            }}
            QPushButton:hover {{ border-color: {ACCENT}; color: {ACCENT}; }}
        """)
        btn_abas.clicked.connect(self._atualizar_abas)

        self._inp_busca = QLineEdit()
        self._inp_busca.setPlaceholderText("Buscar por pagador, pedido, produto...")
        self._inp_busca.setFixedWidth(240)
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
        topo.addWidget(self._combo_mes)
        topo.addWidget(btn_abas)
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
        self._tabela.setAlternatingRowColors(False)
        self._tabela.verticalHeader().setVisible(False)
        self._tabela.horizontalHeader().setStretchLastSection(True)
        self._tabela.setSortingEnabled(True)

        # Ordem das colunas: STATUS primeiro, depois os demais
        # índices na lista do Supabase:
        # STATUS(14) DATA(0) FILIAL(1) PAGADOR(2) PEDIDO(15) PRODUTO(16)
        # MOTORISTA(4) PLACA(5) DESTINO(7) UF(8) PESO(9) COLOCADOR(18)
        # STATUS(14) DATA(0) FILIAL(1) PAGADOR(2) PEDIDO(15) PRODUTO(16)
        # MOTORISTA(4) PLACA(5) DESTINO(7) UF(8) PESO(9)
        # FRETE/E(10) FRETE/M(11) ROTA(12) AGENCIAMENTO(13) PAGAMENTO(20*) COLOCADOR(18)
        # *PAGAMENTO não está na lista atual — adicionamos no índice 20 via _carregar_base_supabase
        # índice 19 = ID Supabase (usado como número de ordem)
        self._colunas_visiveis = [19, 14, 0, 1, 2, 15, 16, 4, 5, 7, 8, 9, 10, 11, 12, 13, 20, 18]
        COLUNAS = ["#", "STATUS", "DATA", "FILIAL", "CLIENTE", "PEDIDO", "PRODUTO",
                   "MOTORISTA", "PLACA", "DESTINO", "UF", "PESO",
                   "FRT/E", "FRT/M", "ROTA", "AGENC.", "PAGAMENTO", "COLOCADOR", "", ""]
        self._tabela.setColumnCount(len(COLUNAS))
        self._tabela.setHorizontalHeaderLabels(COLUNAS)

        larguras = [40, 110, 82, 55, 110, 60, 130, 100, 75, 100, 35, 50,
                    55, 55, 80, 70, 90, 80, 46, 46]
        for i, w in enumerate(larguras):
            self._tabela.setColumnWidth(i, w)
        self._tabela.horizontalHeader().setStretchLastSection(False)

        root.addWidget(self._tabela)

        self._todos_dados = []
        self._linhas_editando = {}
        self._row_to_linha_planilha = {}
        self._aba_selecionada = None
        self._tabela.itemClicked.connect(self._on_item_click)
        self._atualizar_contas()

    def _atualizar_contas(self):
        self._combo_conta.clear()
        contas = _listar_contas_gmail()
        self._combo_conta.addItems(contas if contas else ["(nenhuma conta)"])

    def _atualizar_abas(self):
        try:
            meses = _meses_supabase()
            self._combo_mes.clear()
            self._combo_mes.addItems(meses)
            if meses:
                self._combo_mes.setCurrentIndex(len(meses) - 1)
            self._aba_selecionada = self._combo_mes.currentText()
            self._lbl_status.setText(f"{len(meses)} mes(es) disponível(is) no banco de dados")
        except Exception as e:
            self._lbl_status.setText(f"Erro ao listar meses: {e}")

    def _carregar(self):
        mes = self._combo_mes.currentText().strip()
        if not mes:
            self._atualizar_abas()
            mes = self._combo_mes.currentText().strip()
        self._aba_selecionada = mes
        self._lbl_status.setText(f"Carregando {mes} do banco de dados...")
        QApplication.processEvents()
        try:
            dados = _carregar_base_supabase(mes=mes)
            self._todos_dados = dados
            self._renderizar(dados)
            self._lbl_status.setText(f"{len(dados)} ordem(ns) em '{mes}'")
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

            if len(linha) > 19:
                self._row_to_linha_planilha[r] = linha[19]

            status = str(linha[14] if len(linha) > 14 else "").upper()
            if "CARREGADO" in status and "NÃO" not in status:
                bg = QColor(COR_CARR)
            elif "NÃO" in status or "NAO" in status:
                bg = QColor(COR_NAO)
            elif "PAGO" in status:
                bg = QColor(COR_PAGO)
            elif "MARCADO" in status:
                bg = QColor(COR_MARC)
            elif "AGUARDANDO" in status:
                bg = QColor("#1e2200")
            elif "DESCARGA" in status:
                bg = QColor("#1a0a2a")
            elif "CHEGA" in status or "CAMINHO" in status:
                bg = QColor("#0a1a2a")
            else:
                bg = None

            for col_tabela, col_dados in enumerate(self._colunas_visiveis):
                val = linha[col_dados] if col_dados < len(linha) else ""
                item = QTableWidgetItem(str(val) if val is not None else "")
                item.setTextAlignment(Qt.AlignCenter)
                if bg:
                    item.setBackground(bg)
                self._tabela.setItem(r, col_tabela, item)

            self._tabela.setItem(r, 18, self._make_btn_item("EDIT", "#e3b341"))
            self._tabela.setItem(r, 19, self._make_btn_item("DEL",  "#da3633"))

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
        txt = item.text() if item else ""
        if col == 18:
            if txt == "OK":
                self._salvar_edicao(row)
            else:  # EDIT
                dados = [self._tabela.item(row, c).text() if self._tabela.item(row, c) else "" for c in range(18)]
                self._toggle_edicao(row, dados)
        elif col == 19:
            if txt == "✕":  # cancelar edição
                self._linhas_editando.pop(row, None)
                self._carregar()
            else:  # DEL
                dados = [self._tabela.item(row, c).text() if self._tabela.item(row, c) else "" for c in range(18)]
                self._deletar_linha(row, dados)

    def _toggle_edicao(self, row, dados_orig):
        if row in self._linhas_editando:
            self._salvar_edicao(row)
            return

        self._linhas_editando[row] = dados_orig
        self._tabela.setRowHeight(row, 36)

        STATUS_OPTS = ["AGUARDANDO", "MARCADO", "CHEGA", "A CAMINHO",
                       "CARREGADO", "DESCARGA", "PAGO"]

        INP_SS = f"""
            QLineEdit {{
                background: #0d1117; border: none;
                border-bottom: 2px solid {ACCENT};
                color: {TEXT}; font-size: 11px; padding: 0px 4px;
            }}
        """
        COMBO_SS = f"""
            QComboBox {{
                background: #0d1117; border: none;
                border-bottom: 2px solid {ACCENT};
                color: {TEXT}; font-size: 11px; padding: 0px 2px;
            }}
            QComboBox::drop-down {{ border: none; width: 16px; }}
            QComboBox::down-arrow {{
                border-left: 3px solid transparent;
                border-right: 3px solid transparent;
                border-top: 4px solid {MUTED};
                width: 0; height: 0; margin-right: 4px;
            }}
            QComboBox QAbstractItemView {{
                background: {SURFACE}; border: 1px solid {BORDER2};
                color: {TEXT}; selection-background-color: {ACCENT}33;
            }}
        """

        # col 0=# (não editável), col 1=STATUS(combo), cols 2-17=QLineEdit
        # #(0) STATUS(1) DATA(2) FILIAL(3) CLIENTE(4) PEDIDO(5) PRODUTO(6)
        # MOTORISTA(7) PLACA(8) DESTINO(9) UF(10) PESO(11)
        # FRT/E(12) FRT/M(13) ROTA(14) AGENC.(15) PAGAMENTO(16) COLOCADOR(17)
        for c in range(18):
            val = self._tabela.item(row, c).text() if self._tabela.item(row, c) else ""
            if c == 0:  # # — não editável, mostra só o número
                lbl = QLabel(val)
                lbl.setAlignment(Qt.AlignCenter)
                lbl.setStyleSheet(f"color: {MUTED}; font-size: 10px; background: transparent;")
                self._tabela.setCellWidget(row, c, lbl)
                continue
            if c == 1:  # STATUS — combo
                combo = QComboBox()
                combo.addItems(STATUS_OPTS)
                combo.setStyleSheet(COMBO_SS)
                idx = combo.findText(val, Qt.MatchFixedString)
                combo.setCurrentIndex(idx if idx >= 0 else 0)
                self._tabela.setCellWidget(row, c, combo)
            else:
                inp = QLineEdit(val)
                inp.setAlignment(Qt.AlignCenter)
                inp.setFrame(False)
                inp.setStyleSheet(INP_SS)
                if c != 2:  # não força maiúsculo na data (col 2)
                    inp.textChanged.connect(lambda t, i=inp: _forcar_maiusculo(i, t))
                self._tabela.setCellWidget(row, c, inp)

        # Botão OK na coluna 18, ✕ na 19
        self._tabela.setItem(row, 18, self._make_btn_item("OK",  "#2ea043"))
        self._tabela.setItem(row, 19, self._make_btn_item("✕", "#da3633"))

    def _salvar_edicao(self, row):
        import datetime as _dt

        sb_id = self._row_to_linha_planilha.get(row)
        if not sb_id:
            QMessageBox.warning(self, "Aviso", "ID do registro não encontrado no banco.")
            return

        # Coleta valores de todos os widgets editáveis
        # #(0-skip) STATUS(1) DATA(2) FILIAL(3) CLIENTE(4) PEDIDO(5) PRODUTO(6)
        # MOTORISTA(7) PLACA(8) DESTINO(9) UF(10) PESO(11)
        # FRT/E(12) FRT/M(13) ROTA(14) AGENC.(15) PAGAMENTO(16) COLOCADOR(17)
        vals = {}
        for c in range(18):
            w = self._tabela.cellWidget(row, c)
            if isinstance(w, QComboBox):
                vals[c] = w.currentText()
            elif isinstance(w, QLineEdit):
                vals[c] = w.text().strip()
            else:
                item = self._tabela.item(row, c)
                vals[c] = item.text() if item else ""

        STATUS_OPTS_COR = {
            "CARREGADO": "#1a3a1a", "PAGO": "#1a2a3a",
            "MARCADO": "#3a2a00",   "CHEGA": "#0a1a2a",
            "AGUARDANDO": "#1e2200","DESCARGA": "#1a0a2a",
            "A CAMINHO": "#0a1a2a",
        }

        try:
            # #(0-skip) STATUS(1) DATA(2) FILIAL(3) CLIENTE(4) PEDIDO(5) PRODUTO(6)
            # MOTORISTA(7) PLACA(8) DESTINO(9) UF(10) PESO(11)
            # FRT/E(12) FRT/M(13) ROTA(14) AGENC.(15) PAGAMENTO(16) COLOCADOR(17)
            campos = {
                "status":       vals[1].upper(),
                "filial":       vals[3].upper(),
                "pagador":      vals[4].upper(),
                "pedido":       vals[5].upper(),
                "produto":      vals[6].upper(),
                "motorista":    vals[7].upper(),
                "placa":        vals[8].upper(),
                "destino":      vals[9].upper(),
                "uf":           vals[10].upper(),
                "rota":         vals[14].upper(),
                "agenciamento": vals[15].upper(),
                "pagamento":    vals[16].upper(),
                "colocador":    vals[17].upper(),
            }
            # Data: converte DD/MM/AAAA → AAAA-MM-DD
            try:
                campos["data"] = _dt.datetime.strptime(
                    vals[2].strip(), "%d/%m/%Y").strftime("%Y-%m-%d")
            except Exception:
                pass
            # Peso: float
            try:
                campos["peso"]      = float(vals[11].replace(",", "."))
            except Exception:
                pass
            try:
                campos["frete_emp"] = float(vals[12].replace(",", "."))
            except Exception:
                pass
            try:
                campos["frete_mot"] = float(vals[13].replace(",", "."))
            except Exception:
                pass
            campos = {k: v for k, v in campos.items() if v not in ("", None)}

            _atualizar_supabase_linha(sb_id, campos)

            status  = campos.get("status", "")
            bg_cor  = STATUS_OPTS_COR.get(status, "#161b22")
            bg      = QColor(bg_cor)
            novos   = [vals.get(c, "") for c in range(18)]

            for c in range(18):
                self._tabela.removeCellWidget(row, c)
                item = QTableWidgetItem(str(novos[c]))
                item.setTextAlignment(Qt.AlignCenter)
                item.setBackground(bg)
                self._tabela.setItem(row, c, item)

            self._tabela.setItem(row, 18, self._make_btn_item("EDIT", "#e3b341"))
            self._tabela.setItem(row, 19, self._make_btn_item("DEL",  "#da3633"))
            self._linhas_editando.pop(row, None)
            self._tabela.setRowHeight(row, 28)
            self._lbl_status.setText("Registro atualizado no banco de dados.")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao salvar", str(e))

    def _deletar_linha(self, row, dados):
        resp = QMessageBox.question(
            self, "Confirmar",
            "Deseja remover este registro do banco de dados?",
            QMessageBox.Yes | QMessageBox.No
        )
        if resp != QMessageBox.Yes:
            return
        try:
            sb_id = self._row_to_linha_planilha.get(row)
            if sb_id:
                _deletar_supabase_linha(sb_id)
            self._tabela.removeRow(row)
            self._lbl_status.setText("Registro removido do banco de dados.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))

    def _encontrar_linha_base(self, dados, conta):
        # Não mais necessário — BASE usa Supabase com id direto
        return None


# Mapa de fábricas → origem fixa
# Chave: substring que aparece no nome da fábrica (uppercase)
# Valor: origem a preencher
ORIGENS_FABRICA = {
    "FERTIMAXI":        "FEIRA DE SANTANA - BA",
    "INTERMARITIMA":    "CANDEIAS - BA",
    "ARMAZEM VITORIA":  "CANDEIAS - BA",
    "ARMAZEM VITÓRIA":  "CANDEIAS - BA",
    "ARMAZÉM VITORIA":  "CANDEIAS - BA",
    "ARMAZÉM VITÓRIA":  "CANDEIAS - BA",
    "AZ VITORIA":       "CANDEIAS - BA",
    "AZ VITÓRIA":       "CANDEIAS - BA",
    "YARA":             "CANDEIAS - BA",
    "UNIGEL":           "CAMAÇARI - BA",
    # TIMAC tem duas unidades — tratado em _origem_por_fabrica
    "TIMAC CANDEIAS":   "CANDEIAS - BA",
    "TIMAC CAMACARI":   "CAMAÇARI - BA",
    "TIMAC CAMAÇARI":   "CAMAÇARI - BA",
    "TIMAC":            "CAMAÇARI - BA",  # fallback: sem localidade = Camaçari
}

def _origem_por_fabrica(fabrica):
    """Retorna a origem fixa para uma fábrica, ou '' se não reconhecida.
    Normaliza acentos para garantir match independente de grafia.
    Para TIMAC: verifica se contém CANDEIAS antes de usar o fallback."""
    import unicodedata
    def _norm(s):
        return unicodedata.normalize("NFD", s.upper().strip())                           .encode("ascii", "ignore").decode("ascii")
    t = _norm(fabrica)

    # TIMAC: verifica localidade específica primeiro
    if "TIMAC" in t:
        if "CANDEIAS" in t:
            return "CANDEIAS - BA"
        return "CAMAÇARI - BA"

    for chave, origem in ORIGENS_FABRICA.items():
        if "TIMAC" in _norm(chave):
            continue  # já tratado acima
        if _norm(chave) in t:
            return origem
    return ""


def parsear_mensagem_whatsapp(texto):
    resultado = {}

    def extrair(chave):
        match = re.search(rf"^{chave}\s*:\s*([^\n\r]*)", texto, re.IGNORECASE | re.MULTILINE)
        val = match.group(1).strip() if match else ""
        # Se o valor capturado começa com outra chave de campo (ex: "PAGAMENTO: 90/10"),
        # significa que o campo estava vazio — descarta
        if val and re.match(r"^[A-ZÀ-ÿ/]+\s*:", val, re.IGNORECASE):
            val = ""
        return val

    filial = extrair("FILIAL").upper()
    resultado["empresa"] = "Agrovia" if "AGRO" in filial else "TopBrasil"

    # ── PAGADOR / SOLICITANTE / CLIENTE ──────────────────────────────
    # Regras:
    # - PAGADOR → campo Pagador
    # - SOLICITANTE → campo Solicitante
    # - CLIENTE → campo Cliente
    # - PAGADOR/CLIENTE na mesma linha → preenche Pagador E Cliente com mesmo valor
    # - Cada campo preenchido apenas com o que vier na tag — sem cópias automáticas

    pagador     = extrair("PAGADOR")
    solicitante = extrair("SOLICITANTE")
    cliente     = extrair("CLIENTE")

    # Verifica se PAGADOR e CLIENTE estão na mesma linha (ex: PAGADOR/CLIENTE: NOME)
    match_mesmo = re.search(
        r"^(PAGADOR|CLIENTE)\s*/\s*(PAGADOR|CLIENTE)\s*:\s*([^\n\r]*)",
        texto, re.IGNORECASE | re.MULTILINE
    )
    if match_mesmo:
        valor_comum = match_mesmo.group(3).strip()
        # Verifica se o valor é outra chave de campo e descarta se for
        if not re.match(r"^[A-Z\u00C0-\u00FF/]+\s*:", valor_comum, re.IGNORECASE):
            resultado["Pagador"] = valor_comum
            resultado["Cliente"] = valor_comum
    else:
        if pagador:
            resultado["Pagador"] = pagador
        if cliente:
            resultado["Cliente"] = cliente

    if solicitante:
        resultado["Solicitante"] = solicitante

    # ── FÁBRICA e ORIGEM ─────────────────────────────────────────────
    # Origem é sempre definida pela fábrica (fixa por regra)
    # Nunca sobrescreve com o texto livre da tag
    fabrica = extrair("FABRICA")
    if fabrica:
        resultado["Fábrica"] = fabrica.strip().upper()
        origem_fixa = _origem_por_fabrica(fabrica)
        if origem_fixa:
            resultado["Origem"] = origem_fixa
        # Se fábrica não reconhecida, deixa Origem em branco (usuário preenche)

    resultado["Motorista"] = extrair("MOTORISTA")
    resultado["Cavalo"]    = extrair("PLACA")
    resultado["Peso"]      = extrair("PESO")

    # Campos extras para planilha
    agencia = extrair("AGÊNCIA") or extrair("AGENCIA")
    if agencia:
        resultado["Agência"] = agencia

    frete_emp = extrair("FRETE/EMP")
    if frete_emp:
        resultado["Frete/Emp"] = frete_emp

    frete_mot = extrair("FRETE/MOT")
    if frete_mot:
        resultado["Frete/Mot"] = frete_mot

    rota = extrair("ROTA")
    if rota:
        resultado["Rota"] = rota

    # Agenciamento: sempre define no resultado (pode ser vazio) para limpar campo
    agenciamento = extrair("AGENCIAMENTO")
    resultado["Agenciamento"] = agenciamento  # vazio se não preenchido na tag

    colocador = extrair("COLOCADOR")
    if colocador:
        resultado["Colocador"] = colocador

    pagamento = extrair("PAGAMENTO")
    if pagamento:
        resultado["Pagamento"] = pagamento

    # UF já extraído junto com Destino acima

    produto = extrair("PRODUTO")

    # Tenta extrair EMBALAGEM como campo separado da tag
    embalagem_tag = extrair("EMBALAGEM")
    if embalagem_tag:
        resultado["Produto"]   = produto.strip()
        resultado["Embalagem"] = embalagem_tag.strip().upper()
    else:
        # Fallback: detecta embalagem dentro do texto do produto
        embalagem = ""
        prod_upper = produto.upper()
        EMBALAGENS_MAP = [
            ("BIG BAG",    "BIG BAG"),
            ("GRANEL",     "GRANEL"),
            ("PALETIZADO", "PALETIZADO"),
            ("SACO 50KG",  "SACO 50KG"),
            ("SACO 50",    "SACO 50KG"),
            ("SACO 25KG",  "SACO 25KG"),
            ("SACO 25",    "SACO 25KG"),
            ("SACO 40KG",  "SACO 40KG"),
            ("SACO 40",    "SACO 40KG"),
        ]
        produto_limpo = produto
        for token, emb_label in EMBALAGENS_MAP:
            if token in prod_upper:
                embalagem = emb_label
                import re as _re
                produto_limpo = _re.sub(
                    r"[-–\s]*" + _re.escape(token) + r"[-–\s]*",
                    " ", produto, flags=_re.IGNORECASE
                ).strip(" -–")
                break
        resultado["Produto"] = produto_limpo
        if embalagem:
            resultado["Embalagem"] = embalagem

    destino_raw = extrair("DESTINO")
    uf = extrair("UF").upper()

    # UF sempre separado — vai para campo próprio
    if uf:
        resultado["UF"] = uf

    # Destino: só a cidade, sem UF concatenado
    destino_limpo = destino_raw.strip() if destino_raw else ""

    if destino_limpo:
        sep = re.split(r"\s+(FAZ\.?\s|FAZENDA\s|SITIO\s|SÍTIO\s|CHACARA\s)", destino_limpo, flags=re.IGNORECASE)
        if len(sep) >= 3:
            resultado["Destino"] = sep[0].strip()
            resultado["Fazenda"] = (sep[1] + sep[2]).strip()
        else:
            resultado["Destino"] = destino_limpo
            resultado["Fazenda"] = ""
    else:
        resultado["Destino"] = resultado.get("Destino", "")
        resultado["Fazenda"] = resultado.get("Fazenda", "")

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

class ConfigWidget(QWidget):
    """Aba de configurações — gerenciador de contas Gmail e mapeamento empresa→conta."""

    def __init__(self):
        super().__init__()
        self.setStyleSheet("background: transparent;")
        root = QVBoxLayout(self)
        root.setContentsMargins(24, 20, 24, 20)
        root.setSpacing(16)

        titulo = QLabel("CONFIGURAÇÕES")
        titulo.setStyleSheet(f"color: {TEXT}; font-size: 16px; font-weight: 700; letter-spacing: 1px; background: transparent;")
        root.addWidget(titulo)

        # ── Card: Contas Gmail ───────────────────────────────────
        card_contas, body_contas = make_card("Contas Gmail")
        v_contas = QVBoxLayout(body_contas)
        v_contas.setSpacing(8)
        v_contas.setContentsMargins(0, 0, 0, 0)

        self._list_contas = QListWidget()
        self._list_contas.setMinimumHeight(140)
        self._list_contas.setStyleSheet(f"""
            QListWidget {{
                background: {BG}; border: 1px solid {BORDER};
                border-radius: 6px; color: {TEXT};
                font-size: 12px; padding: 4px;
            }}
            QListWidget::item {{ padding: 6px 8px; border-radius: 4px; }}
            QListWidget::item:selected {{ background: {ACCENT}22; color: {ACCENT}; }}
            QListWidget::item:hover {{ background: {BORDER}44; }}
        """)
        v_contas.addWidget(self._list_contas)

        btn_row = QHBoxLayout()
        btn_add = QPushButton("+ Adicionar conta Gmail")
        btn_add.setStyleSheet(f"""
            QPushButton {{
                background: {ACCENT}; color: white; border: none;
                border-radius: 6px; padding: 7px 14px; font-weight: 700; font-size: 11px;
            }}
            QPushButton:hover {{ background: {ACCENT_H}; }}
        """)
        btn_rem = QPushButton("🗑  Remover selecionada")
        btn_rem.setStyleSheet(f"""
            QPushButton {{
                background: transparent; color: {DANGER};
                border: 1px solid {DANGER}; border-radius: 6px;
                padding: 7px 14px; font-weight: 700; font-size: 11px;
            }}
            QPushButton:hover {{ background: {DANGER}18; }}
        """)
        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_rem)
        btn_row.addStretch()
        v_contas.addLayout(btn_row)

        self._lbl_contas = QLabel("")
        self._lbl_contas.setStyleSheet(f"color: {MUTED}; font-size: 10px; background: transparent;")
        v_contas.addWidget(self._lbl_contas)
        root.addWidget(card_contas)

        # ── Card: Conta padrão por empresa ───────────────────────
        card_emp, body_emp = make_card("Conta padrão por empresa")
        v_emp = QVBoxLayout(body_emp)
        v_emp.setSpacing(10)
        v_emp.setContentsMargins(0, 0, 0, 0)

        descr = QLabel("Define qual conta Gmail é usada automaticamente ao gerar ordem para cada empresa.")
        descr.setStyleSheet(f"color: {MUTED}; font-size: 10px; background: transparent;")
        descr.setWordWrap(True)
        v_emp.addWidget(descr)

        self._combos_empresa = {}
        for empresa in ["Agrovia", "TopBrasil"]:
            row = QHBoxLayout()
            lbl = QLabel(empresa)
            lbl.setFixedWidth(100)
            lbl.setStyleSheet(f"color: {TEXT}; font-size: 12px; font-weight: 600; background: transparent;")
            combo = QComboBox()
            combo.setMinimumHeight(32)
            combo.setStyleSheet(f"""
                QComboBox {{
                    background: {BG}; border: 1px solid {BORDER};
                    border-radius: 6px; color: {TEXT}; padding: 4px 8px; font-size: 11px;
                }}
                QComboBox:focus {{ border-color: {ACCENT}; }}
                QComboBox QAbstractItemView {{
                    background: {SURFACE}; color: {TEXT};
                    border: 1px solid {BORDER}; selection-background-color: {ACCENT}22;
                }}
            """)
            self._combos_empresa[empresa] = combo
            btn_limpar = QPushButton("✕")
            btn_limpar.setFixedSize(28, 28)
            btn_limpar.setToolTip(f"Remover vínculo de {empresa}")
            btn_limpar.setStyleSheet(f"""
                QPushButton {{ background: transparent; border: none; color: {MUTED}; font-size: 12px; }}
                QPushButton:hover {{ color: {DANGER}; }}
            """)
            btn_limpar.clicked.connect(lambda _, e=empresa: self._limpar_empresa(e))
            row.addWidget(lbl)
            row.addWidget(combo, 1)
            row.addWidget(btn_limpar)
            v_emp.addLayout(row)

        btn_salvar_emp = QPushButton("SALVAR VÍNCULOS")
        btn_salvar_emp.setStyleSheet(f"""
            QPushButton {{
                background: {ACCENT}; color: white; border: none;
                border-radius: 6px; padding: 7px 18px; font-weight: 700; font-size: 11px;
            }}
            QPushButton:hover {{ background: {ACCENT_H}; }}
        """)
        btn_salvar_emp.clicked.connect(self._salvar_vinculos)
        v_emp.addWidget(btn_salvar_emp, alignment=Qt.AlignLeft)

        self._lbl_emp = QLabel("")
        self._lbl_emp.setStyleSheet(f"color: {MUTED}; font-size: 10px; background: transparent;")
        v_emp.addWidget(self._lbl_emp)
        root.addWidget(card_emp)
        root.addStretch()

        btn_add.clicked.connect(self._adicionar_conta)
        btn_rem.clicked.connect(self._remover_conta)

    def _carregar(self):
        contas = _listar_contas_gmail()
        self._list_contas.clear()
        for c in contas:
            item = QListWidgetItem(c)
            self._list_contas.addItem(item)
        self._lbl_contas.setText(f"{len(contas)} conta(s) configurada(s)")

        mapa = carregar_contas_empresa()
        for empresa, combo in self._combos_empresa.items():
            combo.clear()
            combo.addItem("(nenhuma — perguntar sempre)")
            combo.addItems(contas)
            conta_atual = mapa.get(empresa, "")
            if conta_atual:
                idx = combo.findText(conta_atual)
                if idx >= 0:
                    combo.setCurrentIndex(idx)

    def _adicionar_conta(self):
        try:
            novo = adicionar_conta_gmail()
            self._lbl_contas.setText(f"✓  Conta {novo} adicionada com sucesso.")
            self._carregar()
        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))

    def _remover_conta(self):
        item = self._list_contas.currentItem()
        if not item:
            QMessageBox.warning(self, "Atenção", "Selecione uma conta para remover.")
            return
        conta = item.text()
        resp = QMessageBox.question(
            self, "Remover conta",
            f"Remover a conta {conta}?\n\nO arquivo de token será excluído.",
            QMessageBox.Yes | QMessageBox.No
        )
        if resp != QMessageBox.Yes:
            return
        try:
            import os
            token = os.path.join("gmail_tokens", f"token_{conta}.json")
            if os.path.exists(token):
                os.remove(token)
            mapa = carregar_contas_empresa()
            for emp in list(mapa.keys()):
                if mapa[emp] == conta:
                    del mapa[emp]
            salvar_contas_empresa(mapa)
            self._lbl_contas.setText(f"Conta {conta} removida.")
            self._carregar()
        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))

    def _salvar_vinculos(self):
        mapa = carregar_contas_empresa()
        for empresa, combo in self._combos_empresa.items():
            val = combo.currentText()
            if val.startswith("(nenhuma"):
                mapa.pop(empresa, None)
            else:
                mapa[empresa] = val
        salvar_contas_empresa(mapa)
        self._lbl_emp.setText("✓  Vínculos salvos com sucesso.")

    def _limpar_empresa(self, empresa):
        self._combos_empresa[empresa].setCurrentIndex(0)
        mapa = carregar_contas_empresa()
        mapa.pop(empresa, None)
        salvar_contas_empresa(mapa)
        self._lbl_emp.setText(f"Vínculo de {empresa} removido.")


class UI(QWidget):
    def __init__(self):
        super().__init__()
        self.empresa = None
        self.entradas = {}
        self._pedido_linhas = []
        self.usuario_logado  = ""
        self.assinatura_usuario = ""

        self.setWindowTitle("Sistema de Ordens")
        self._setup_icon()
        self._setup_bg()
        self._build_ui()
        self._apply_style()

        self.overlay = LoadingOverlay(self)
        self.showMaximized()
        self.escolher_empresa()
        self.setar_data_hoje()
        # Restaura assinatura do usuário logado após limpar campos
        if self.assinatura_usuario and "Assinatura" in self.entradas:
            self.entradas["Assinatura"].setText(self.assinatura_usuario)

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
        self._nav_btns.append(make_nav_btn("📊", "Controle de Ordens", 2))
        self._nav_btns.append(make_nav_btn("⚙", "Configurações", 3))
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
        self._planilha_widget = None  # removido
        self._base_widget = BaseWidget()
        self._stack.addWidget(self._base_widget)
        self._config_widget = ConfigWidget()
        self._stack.addWidget(self._config_widget)

        root.addWidget(self._stack, 1)

    def _nav(self, idx):
        self._stack.setCurrentIndex(idx)
        for i, b in enumerate(self._nav_btns):
            b.setChecked(i == idx)
        if idx == 1:
            self._historico_widget.recarregar()
        if idx == 2:
            self._base_widget._atualizar_contas()
        if idx == 3:
            self._config_widget._carregar()

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
        inner.setContentsMargins(8, 8, 8, 8)
        inner.setSpacing(6)

        # ── LINHA 1: Cabeçalho | Motorista | Veículo ─────
        row1 = QHBoxLayout()
        row1.setSpacing(8)
        row1.setAlignment(Qt.AlignTop)

        cab = self._build_cabecalho()
        cab.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        mot = self._build_motorista()
        mot.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        vei = self._build_veiculo()
        vei.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)

        row1.addWidget(cab, 4)
        row1.addWidget(mot, 3)
        row1.addWidget(vei, 3)
        inner.addLayout(row1)

        # ── LINHA 2: Carga | Dados Planilha | Assinatura+Botões ─
        row2 = QHBoxLayout()
        row2.setSpacing(8)
        row2.setAlignment(Qt.AlignTop)

        car = self._build_carga()
        car.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        extra = self._build_dados_planilha()
        extra.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)

        col3_w = QWidget()
        col3_w.setStyleSheet("background: transparent;")
        col3_w.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        col3 = QVBoxLayout(col3_w)
        col3.setSpacing(8)
        col3.setContentsMargins(0, 0, 0, 0)
        ass = self._build_assinatura()
        ass.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        col3.addWidget(ass)
        col3.addLayout(self._build_botoes())
        col3.addStretch()

        row2.addWidget(car, 4)
        row2.addWidget(extra, 3)
        row2.addWidget(col3_w, 3)
        inner.addLayout(row2)

        inner.addStretch()

        scroll.setWidget(container)
        root.addWidget(scroll)
        return pagina

    def _build_cabecalho(self):
        frame, content = make_card("Cabeçalho")
        v = QVBoxLayout(content)
        v.setSpacing(5)
        v.setContentsMargins(0, 0, 0, 0)
        v.setAlignment(Qt.AlignTop)

        r1 = QHBoxLayout(); r1.setSpacing(6)
        self.entradas["Data Apresentação"] = make_date()
        self.entradas["Fábrica"]     = make_input()
        self.entradas["Pagador"]     = make_input()
        self.entradas["Solicitante"] = make_input()
        r1.addWidget(make_field("Data",        self.entradas["Data Apresentação"]), 1)
        r1.addWidget(make_field("Fábrica",     self.entradas["Fábrica"]),           2)
        r1.addWidget(make_field("Pagador",     self.entradas["Pagador"]),           2)
        r1.addWidget(make_field("Solicitante", self.entradas["Solicitante"]),       2)
        v.addLayout(r1)

        def _atualizar_origem(texto):
            origem_fixa = _origem_por_fabrica(texto)
            if origem_fixa:
                self.entradas["Origem"].setText(origem_fixa)

        self.entradas["Fábrica"].textChanged.connect(_atualizar_origem)

        r2 = QHBoxLayout(); r2.setSpacing(6)
        self.entradas["Origem"]  = make_input()
        self.entradas["Destino"] = make_input()
        self._uf_cab             = make_input()  # campo visual no cabeçalho
        self.entradas["Cliente"] = make_input()
        r2.addWidget(make_field("Origem",  self.entradas["Origem"]),  1)
        r2.addWidget(make_field("Destino", self.entradas["Destino"]), 2)
        r2.addWidget(make_field("UF",      self._uf_cab),             1)
        r2.addWidget(make_field("Cliente", self.entradas["Cliente"]), 2)
        v.addLayout(r2)

        # Sincroniza UF do cabeçalho com o campo UF de Dados da Planilha
        def _sync_uf_cab(txt):
            w = self.entradas.get("UF")
            if w and isinstance(w, QLineEdit):
                w.blockSignals(True)
                w.setText(txt.upper())
                w.blockSignals(False)
        self._uf_cab.textChanged.connect(lambda t: _sync_uf_cab(t))

        self.entradas["Fazenda"] = make_input()
        v.addWidget(make_field("Fazenda", self.entradas["Fazenda"]))

        return frame

    def _build_motorista(self):
        frame, content = make_card("Motorista")
        v = QVBoxLayout(content)
        v.setSpacing(5)
        v.setContentsMargins(0, 0, 0, 0)
        v.setAlignment(Qt.AlignTop)

        self.entradas["Motorista"] = make_input()
        v.addWidget(make_field("Nome", self.entradas["Motorista"]))

        r = QHBoxLayout(); r.setSpacing(6)
        self.entradas["CPF"]     = make_input(maiusculo=False)
        self.entradas["Contato"] = make_input(maiusculo=False)
        r.addWidget(make_field("CPF", self.entradas["CPF"]), 1)
        r.addWidget(make_field("Contato", self.entradas["Contato"]), 1)
        v.addLayout(r)
        v.addStretch(1)

        return frame

    def _build_carga(self):
        frame, content = make_card("Carga")
        v = QVBoxLayout(content)
        v.setSpacing(3)
        v.setContentsMargins(0, 0, 0, 0)

        hdr = QHBoxLayout()
        hdr.setSpacing(4)
        hdr.setContentsMargins(0, 0, 0, 0)
        for txt, stretch in [("Pedido", 3), ("Produto", 3), ("Peso", 1), ("Embalagem", 2)]:
            lbl = QLabel(txt.upper())
            lbl.setStyleSheet("color: #555d66; font-size: 9px; font-weight: 700; background: transparent;")
            hdr.addWidget(lbl, stretch)
        hdr.addSpacing(26)
        v.addLayout(hdr)

        self._carga_container = QWidget()
        self._carga_container.setStyleSheet("background: transparent;")
        self._carga_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self._carga_vbox = QVBoxLayout(self._carga_container)
        self._carga_vbox.setSpacing(6)
        self._carga_vbox.setContentsMargins(0, 0, 0, 0)
        v.addWidget(self._carga_container)
        v.addStretch()

        # Cria 4 linhas: primeira ativa, demais como botão
        for i in range(4):
            self._adicionar_linha_pedido(ativa=(i == 0))

        return frame

    def _build_veiculo(self):
        frame, content = make_card("Veículo")
        v = QVBoxLayout(content)
        v.setSpacing(5)
        v.setContentsMargins(0, 0, 0, 0)
        v.setAlignment(Qt.AlignTop)

        r1 = QHBoxLayout(); r1.setSpacing(6)
        self.entradas["Carroceria"] = make_combo(["Graneleiro", "Basculante", "Baú", "Sider", "Tanque"])
        self.entradas["Carroceria"].setCurrentIndex(-1)
        self.entradas["Carroceria"].lineEdit().setPlaceholderText("Selecione...")
        self.entradas["Cavalo"] = make_input()
        r1.addWidget(make_field("Carroceria", self.entradas["Carroceria"]), 1)
        r1.addWidget(make_field("Cavalo", self.entradas["Cavalo"]), 1)
        v.addLayout(r1)

        r2 = QHBoxLayout(); r2.setSpacing(6)
        self.entradas["Carreta 1"] = make_input()
        self.entradas["Carreta 2"] = make_input()
        r2.addWidget(make_field("Carreta 1", self.entradas["Carreta 1"]))
        r2.addWidget(make_field("Carreta 2", self.entradas["Carreta 2"]))
        v.addLayout(r2)

        self.entradas["Carreta 3"] = make_input()
        v.addWidget(make_field("Carreta 3", self.entradas["Carreta 3"]))
        v.addStretch()

        return frame

    def _build_dados_planilha(self):
        frame, content = make_card("Dados da Planilha")
        v = QVBoxLayout(content)
        v.setSpacing(8)
        v.setContentsMargins(0, 0, 0, 0)

        r1 = QHBoxLayout(); r1.setSpacing(6)
        self.entradas["Agência"]   = make_input()
        self.entradas["UF"]        = make_input()
        self.entradas["UF"].textChanged.connect(
            lambda t: self._uf_cab.setText(t.upper()) if hasattr(self, "_uf_cab") else None)
        self.entradas["Frete/Emp"] = make_input()
        self.entradas["Frete/Mot"] = make_input()
        r1.addWidget(make_field("Agência",   self.entradas["Agência"]),   3)
        r1.addWidget(make_field("UF",        self.entradas["UF"]),        1)
        r1.addWidget(make_field("Frete/Emp", self.entradas["Frete/Emp"]), 2)
        r1.addWidget(make_field("Frete/Mot", self.entradas["Frete/Mot"]), 2)
        v.addLayout(r1)

        r2 = QHBoxLayout(); r2.setSpacing(6)
        self.entradas["Rota"]         = make_input()
        self.entradas["Agenciamento"] = make_input()
        self.entradas["Colocador"]    = make_input()
        r2.addWidget(make_field("Rota",         self.entradas["Rota"]),         2)
        r2.addWidget(make_field("Agenciamento", self.entradas["Agenciamento"]), 2)
        r2.addWidget(make_field("Colocador",    self.entradas["Colocador"]),    2)
        v.addLayout(r2)

        r3 = QHBoxLayout(); r3.setSpacing(6)
        self.entradas["Pagamento"]   = make_input()
        self.entradas["Peso Total"]  = make_input()
        self.entradas["Peso Total"].setReadOnly(True)
        self.entradas["Peso Total"].setPlaceholderText("Auto")
        self.entradas["Peso Total"].setStyleSheet(f"""
            QLineEdit {{
                background: #1c2128; border: 1px solid {BORDER2};
                border-radius: 6px; padding: 6px 10px;
                color: {ACCENT}; font-size: 12px; font-weight: 700;
            }}
        """)
        r3.addWidget(make_field("Pagamento",  self.entradas["Pagamento"]),  2)
        r3.addWidget(make_field("Peso Total", self.entradas["Peso Total"]), 1)
        v.addLayout(r3)

        v.addStretch()
        return frame

    def _build_assinatura(self):
        frame, content = make_card("Assinatura")
        v = QVBoxLayout(content)
        v.setSpacing(4)
        v.setContentsMargins(0, 0, 0, 0)
        v.setAlignment(Qt.AlignTop)

        self.entradas["Assinatura"] = make_input(maiusculo=False)
        self.entradas["Assinatura"].setMinimumHeight(32)
        v.addWidget(self.entradas["Assinatura"])

        dev = QLabel("© 2026 dev by Felipe")
        dev.setAlignment(Qt.AlignCenter)
        dev.setStyleSheet(f"color: #21262d; font-size: 10px; letter-spacing: 1px; background: transparent;")
        v.addWidget(dev)
        v.addStretch()

        return frame

    def _build_botoes(self):
        v = QVBoxLayout()
        v.setSpacing(5)

        self.btn_wpp = QPushButton("📋  IMPORTAR WHATSAPP")
        self.btn_wpp.setObjectName("btn_wpp")

        self.btn1 = QPushButton("GERAR ORDEM")
        self.btn1.setObjectName("btn_gerar")

        self.btn2 = QPushButton("GERAR + EMAIL")
        self.btn2.setObjectName("btn_email")

        self.btn3 = QPushButton("NOVA ORDEM")
        self.btn3.setObjectName("btn_nova")

        for btn in [self.btn_wpp, self.btn1, self.btn2, self.btn3]:
            btn.setMinimumHeight(40)
            v.addWidget(btn)

        v.addStretch()

        self.btn_wpp.clicked.connect(self.importar_whatsapp)
        self.btn1.clicked.connect(lambda: self.executar(False))
        self.btn2.clicked.connect(lambda: self.executar(True))
        self.btn3.clicked.connect(self.nova_ordem)

        return v

    def _atualizar_peso_total(self):
        """Soma todos os campos Peso ativos e atualiza o campo Peso Total."""
        total = 0.0
        for i in range(4):
            sufixo = f" {i + 1}" if i > 0 else ""
            chave  = f"Peso{sufixo}"
            w = self.entradas.get(chave)
            if w and isinstance(w, QLineEdit):
                try:
                    total += float(str(w.text()).replace(",", "."))
                except Exception:
                    pass
        w_total = self.entradas.get("Peso Total")
        if w_total:
            w_total.setText(str(int(total)) if total == int(total) else f"{total:.1f}")

    def _adicionar_linha_pedido(self, ativa=True):
        MAX = 4
        if len(self._pedido_linhas) >= MAX:
            return

        idx    = len(self._pedido_linhas)
        sufixo = f" {idx + 1}" if idx > 0 else ""

        EMBALAGENS = ["BIG BAG", "SACO 50KG", "SACO 25KG", "SACO 40KG", "GRANEL", "PALETIZADO"]

        # QStackedWidget: página 0 = botão, página 1 = inputs
        stack = QStackedWidget()
        stack.setFixedHeight(36)
        stack.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        # ── Página 0: Botão inativo ──
        btn_ativar = QPushButton(f"＋  Pedido {idx + 1}")
        btn_ativar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        btn_ativar.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: 1px dashed {BORDER2};
                border-radius: 6px;
                color: {MUTED};
                font-size: 11px;
                font-weight: 600;
            }}
            QPushButton:hover {{ border-color: {ACCENT}; color: {ACCENT}; }}
        """)
        stack.addWidget(btn_ativar)  # índice 0

        # ── Página 1: Inputs ──
        inp_w = QWidget()
        inp_w.setStyleSheet("background: transparent;")
        inp_w.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        inp_lay = QHBoxLayout(inp_w)
        inp_lay.setContentsMargins(0, 0, 0, 0)
        inp_lay.setSpacing(4)

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
                inp.setFixedHeight(32)
                inp.setStyleSheet("QComboBox { padding: 2px 6px; font-size: 12px; }")
            else:
                inp = QLineEdit()
                inp.setFixedHeight(32)
                inp.setStyleSheet("QLineEdit { padding: 2px 6px; font-size: 12px; }")
                inp.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
                inp.textChanged.connect(lambda t, i=inp: _forcar_maiusculo(i, t))
                if chave == "Peso":
                    inp.textChanged.connect(lambda _: self._atualizar_peso_total())
            inp_lay.addWidget(inp, stretch)
            nome = chave + sufixo
            linha[nome] = inp
            self.entradas[nome] = inp

        btn_del = QPushButton("×")
        btn_del.setFixedSize(24, 32)
        btn_del.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: 1px solid #30363d;
                border-radius: 4px; color: #8b949e;
                font-size: 14px; font-weight: bold; padding: 0;
            }}
            QPushButton:hover {{ background: #da363320; border-color: #da3633; color: #f85149; }}
        """)
        inp_lay.addWidget(btn_del)
        stack.addWidget(inp_w)  # índice 1

        # Captura referências finais para os closures (linha já completa aqui)
        linha_ref = linha
        stack_ref = stack

        # Conecta botões — usando referências explícitas para evitar bug de closure
        btn_ativar.clicked.connect(lambda checked=False, s=stack_ref: s.setCurrentIndex(1))

        def _desativar(checked=False, s=stack_ref, ln=linha_ref):
            for widget in ln.values():
                if isinstance(widget, QLineEdit):
                    widget.clear()
                elif isinstance(widget, QComboBox):
                    widget.setCurrentIndex(-1)
            s.setCurrentIndex(0)

        btn_del.clicked.connect(_desativar)

        # Estado inicial
        stack.setCurrentIndex(1 if ativa else 0)

        self._pedido_linhas.append((stack, linha))
        self._carga_vbox.addWidget(stack)

    def _desativar_linha_pedido(self, row_w, row_stack, linha):
        for inp in linha.values():
            if isinstance(inp, QLineEdit):
                inp.clear()
            elif isinstance(inp, QComboBox):
                inp.setCurrentIndex(-1)
        row_stack.setCurrentIndex(0)

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
        usuarios = carregar_usuarios()
        is_primeiro_login = not self.usuario_logado

        dlg = QDialog(self)
        dlg.setWindowTitle("Sistema de Ordens")
        dlg.setFixedSize(360, is_primeiro_login and 360 or 240)
        dlg.setStyleSheet(DIALOG_SS)
        dlg.setWindowFlag(Qt.WindowCloseButtonHint, is_primeiro_login is False)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(10)

        # ── Título ──
        titulo = QLabel("SISTEMA DE ORDENS")
        titulo.setAlignment(Qt.AlignCenter)
        titulo.setStyleSheet(f"color: {TEXT}; font-size: 15px; font-weight: 700; letter-spacing: 1.5px; background: transparent;")
        lay.addWidget(titulo)

        subtitulo = QLabel("Sistema de Ordens de Carregamento")
        subtitulo.setAlignment(Qt.AlignCenter)
        subtitulo.setStyleSheet(f"color: {MUTED}; font-size: 11px; background: transparent;")
        lay.addWidget(subtitulo)
        lay.addSpacing(4)

        # ── Seção de login (só na primeira vez) ──
        if is_primeiro_login:
            sep1 = QFrame(); sep1.setFrameShape(QFrame.HLine)
            sep1.setStyleSheet(f"color: {BORDER}; background: {BORDER}; max-height: 1px;")
            lay.addWidget(sep1)

            lbl_usuario = QLabel("USUÁRIO")
            lbl_usuario.setStyleSheet(f"color: {MUTED}; font-size: 9px; font-weight: 700; letter-spacing: 0.8px; background: transparent;")
            lay.addWidget(lbl_usuario)

            combo_usuario = QComboBox()
            combo_usuario.setEditable(True)
            combo_usuario.addItems(sorted(usuarios.keys()))
            combo_usuario.setCurrentIndex(-1)
            combo_usuario.lineEdit().setPlaceholderText("Selecione ou digite seu nome...")
            combo_usuario.setMinimumHeight(36)
            comp = QCompleter(sorted(usuarios.keys()))
            comp.setCaseSensitivity(Qt.CaseInsensitive)
            combo_usuario.setCompleter(comp)
            lay.addWidget(combo_usuario)

            lbl_erro = QLabel("")
            lbl_erro.setAlignment(Qt.AlignCenter)
            lbl_erro.setStyleSheet(f"color: {DANGER}; font-size: 11px; background: transparent;")
            lay.addWidget(lbl_erro)
            lay.addSpacing(4)

        # ── Separador empresa ──
        sep2 = QFrame(); sep2.setFrameShape(QFrame.HLine)
        sep2.setStyleSheet(f"color: {BORDER}; background: {BORDER}; max-height: 1px;")
        lay.addWidget(sep2)

        lbl_emp = QLabel("EMPRESA")
        lbl_emp.setStyleSheet(f"color: {MUTED}; font-size: 9px; font-weight: 700; letter-spacing: 0.8px; background: transparent;")
        lay.addWidget(lbl_emp)

        btn_a = QPushButton("AGROVIA")
        btn_a.setObjectName("btn_agro")
        btn_a.setMinimumHeight(42)

        btn_t = QPushButton("TOPBRASIL")
        btn_t.setObjectName("btn_top")
        btn_t.setMinimumHeight(42)

        def sel(nome_empresa):
            if is_primeiro_login:
                login = combo_usuario.currentText().strip().upper()
                if not login or login not in [k.upper() for k in usuarios.keys()]:
                    lbl_erro.setText("⚠  Selecione um usuário válido.")
                    combo_usuario.setFocus()
                    return
                # Encontra a chave case-insensitive
                chave_real = next(k for k in usuarios if k.upper() == login)
                self.usuario_logado      = chave_real
                self.assinatura_usuario  = usuarios[chave_real]
                # Preenche o campo assinatura
                if "Assinatura" in self.entradas:
                    self.entradas["Assinatura"].setText(self.assinatura_usuario)

            self._aplicar_empresa(nome_empresa)
            dlg.accept()

        btn_a.clicked.connect(lambda: sel("Agrovia"))
        btn_t.clicked.connect(lambda: sel("TopBrasil"))

        lay.addWidget(btn_a)
        lay.addWidget(btn_t)

        # Se já logado, mostra quem está logado no rodapé
        if not is_primeiro_login:
            lbl_logado = QLabel(f"👤  {self.assinatura_usuario}")
            lbl_logado.setAlignment(Qt.AlignCenter)
            lbl_logado.setStyleSheet(f"color: {MUTED}; font-size: 11px; background: transparent; margin-top: 4px;")
            lay.addWidget(lbl_logado)

        dlg.exec()

    def _deletar_linha_pedido(self, stack, linha):
        self._pedido_linhas = [(s, ln) for s, ln in self._pedido_linhas if s is not stack]
        self._carga_vbox.removeWidget(stack)
        stack.deleteLater()
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

        # ── Validações obrigatórias ──────────────────────────
        erros = []

        if not dados.get("Motorista"):
            erros.append("• Motorista")
        if not dados.get("Cavalo"):
            erros.append("• Cavalo (Placa)")

        # Carroceria obrigatória
        carroceria = dados.get("Carroceria", "").strip()
        if not carroceria:
            erros.append("• Carroceria")

        # Embalagem obrigatória em cada linha de pedido ativa
        for i, (stack, linha) in enumerate(self._pedido_linhas):
            if stack.currentIndex() == 1:   # linha ativa (mostrando inputs)
                sufixo = f" {i + 1}" if i > 0 else ""
                emb_key = f"Embalagem{sufixo}"
                ped_key = f"Pedido{sufixo}"
                # Só exige embalagem se o pedido estiver preenchido
                if dados.get(ped_key, "").strip():
                    emb = dados.get(emb_key, "").strip()
                    if not emb:
                        label = f"Pedido {i + 1}" if i > 0 else "Pedido"
                        erros.append(f"• Embalagem ({label})")

        if erros:
            QMessageBox.warning(
                self, "Campos obrigatórios",
                "Preencha os campos obrigatórios antes de gerar a ordem:\n\n" +
                "\n".join(erros)
            )
            return

        conta_gmail = None
        if email:
            conta_gmail = self._dialog_escolher_conta()
            if conta_gmail is None:
                return
            from gerador import montar_email
            prev = self._dialog_preview_email(
                "",
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

        # Injeta usuário logado nos dados para gravar no Supabase
        dados["_usuario"] = self.usuario_logado or ""
        # Reedição sempre gera novo registro — _reeditando_id_anterior tratado em _on_sucesso
        self._thread = GeradorThread(dados, pasta, email, conta_gmail)
        self._thread.sucesso.connect(self._on_sucesso)
        self._thread.erro.connect(self._on_erro)
        self._thread.start()

    def _aplicar_empresa(self, nome_empresa):
        """Define a empresa ativa programaticamente."""
        self.empresa = nome_empresa
        nome_empresa = "Agrovia" if "AGRO" in str(nome_empresa).upper() else "TopBrasil"
        self.empresa = nome_empresa
        cor = ACCENT if nome_empresa == "Agrovia" else DANGER
        self.btn1.setStyleSheet(f"background-color: {cor}; color: white; border: none;")
        self._atualizar_fundo(nome_empresa)

    def _dialog_escolher_conta(self):
        contas        = _listar_contas_gmail()
        mapa_empresa  = carregar_contas_empresa()
        empresa_atual = self.empresa or ""

        # Se há uma conta mapeada para esta empresa, usa direto sem diálogo
        conta_mapeada = mapa_empresa.get(empresa_atual)
        if conta_mapeada and conta_mapeada in contas:
            return conta_mapeada

        # Caso contrário, abre o diálogo para escolher e opcionalmente fixar
        dlg = QDialog(self)
        dlg.setWindowTitle("Enviar por Gmail")
        dlg.setFixedSize(420, 290)
        dlg.setStyleSheet(DIALOG_SS)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(10)

        # Info da empresa
        if empresa_atual:
            lbl_emp = QLabel(f"Empresa: {empresa_atual}")
            lbl_emp.setStyleSheet(f"color: {MUTED}; font-size: 11px; background: transparent;")
            lay.addWidget(lbl_emp)

        lay.addWidget(QLabel("Conta remetente:"))

        combo = QComboBox()
        combo.setMinimumHeight(36)
        combo.addItems(contas if contas else ["(nenhuma conta configurada)"])

        # Pré-seleciona a conta mapeada se existir mas não estiver disponível,
        # ou a primeira da lista
        if conta_mapeada:
            idx = combo.findText(conta_mapeada)
            if idx >= 0:
                combo.setCurrentIndex(idx)

        lay.addWidget(combo)

        # Checkbox para fixar esta conta para a empresa
        chk_fixar = QCheckBox(f"Usar sempre esta conta para {empresa_atual}")
        chk_fixar.setStyleSheet(f"color: {MUTED}; font-size: 11px; background: transparent;")
        chk_fixar.setChecked(bool(conta_mapeada))
        if empresa_atual:
            lay.addWidget(chk_fixar)

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
            # Salva mapeamento se checkbox marcado
            if empresa_atual and chk_fixar.isChecked():
                mapa = carregar_contas_empresa()
                mapa[empresa_atual] = c
                salvar_contas_empresa(mapa)
            elif empresa_atual and not chk_fixar.isChecked():
                # Remove mapeamento se desmarcou
                mapa = carregar_contas_empresa()
                mapa.pop(empresa_atual, None)
                salvar_contas_empresa(mapa)
            resultado[0] = c
            dlg.accept()

        btn_add.clicked.connect(adicionar)
        bc.clicked.connect(dlg.reject)
        bo.clicked.connect(confirmar)
        dlg.exec()
        return resultado[0]

    def _dialog_preview_email(self, destinatario, assunto, corpo):
        from gerador import REGRAS_EMAIL

        # Coleta todos os emails cadastrados em REGRAS_EMAIL
        todos_emails = []
        vistos = set()
        for valor in REGRAS_EMAIL.values():
            for em in valor.split(";"):
                em = em.strip()
                if em and "@" in em and em not in vistos:
                    todos_emails.append(em)
                    vistos.add(em)

        dlg = QDialog(self)
        dlg.setWindowTitle("Prévia do email")
        dlg.setMinimumWidth(540)
        dlg.setStyleSheet(DIALOG_SS)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(8)

        # ── Destinatário com botões de email ─────────────
        lbl_dest = QLabel("DESTINATÁRIO")
        lbl_dest.setStyleSheet(f"color: {MUTED}; font-size: 9px; font-weight: 700; letter-spacing: 0.8px; background: transparent;")
        lay.addWidget(lbl_dest)

        inp_d = QLineEdit(destinatario)
        inp_d.setMinimumHeight(34)
        inp_d.setPlaceholderText("email1@exemplo.com; email2@exemplo.com")
        lay.addWidget(inp_d)

        # Botões de seleção rápida — todos os emails cadastrados
        if todos_emails:
            lbl_opcoes = QLabel("SELECIONAR:")
            lbl_opcoes.setStyleSheet(f"color: {MUTED}; font-size: 9px; font-weight: 700; letter-spacing: 0.8px; background: transparent; margin-top: 4px;")
            lay.addWidget(lbl_opcoes)

            wrap = QWidget(); wrap.setStyleSheet("background: transparent;")
            flow = QHBoxLayout(wrap)
            flow.setContentsMargins(0, 0, 0, 4)
            flow.setSpacing(6)
            flow.setAlignment(Qt.AlignLeft)
            flow.setSpacing(6)

            email_btns = {}

            def _atualizar_btns():
                atual = set(e.strip() for e in inp_d.text().split(";") if e.strip())
                for em, b in email_btns.items():
                    b.blockSignals(True)
                    b.setChecked(em in atual)
                    b.blockSignals(False)

            def _toggle(email):
                atual = [e.strip() for e in inp_d.text().split(";") if e.strip()]
                if email in atual:
                    atual.remove(email)
                else:
                    atual.append(email)
                inp_d.setText("; ".join(atual))

            for em in todos_emails:
                btn = QPushButton(em)
                btn.setCheckable(True)
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background: transparent; border: 1px solid {BORDER2};
                        border-radius: 5px; color: {MUTED};
                        font-size: 10px; padding: 4px 8px;
                    }}
                    QPushButton:checked {{
                        background: {ACCENT}22; border-color: {ACCENT};
                        color: {ACCENT}; font-weight: 700;
                    }}
                    QPushButton:hover {{ border-color: {ACCENT}; color: {TEXT}; }}
                """)
                btn.clicked.connect(lambda checked, e=em: _toggle(e))
                email_btns[em] = btn
                flow.addWidget(btn)

            flow.addStretch()
            lay.addWidget(wrap)
            inp_d.textChanged.connect(lambda _: _atualizar_btns())
            _atualizar_btns()

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

        # Deleta arquivos antigos se for reeedição
        if hasattr(self, "_arquivos_para_deletar"):
            for arq in self._arquivos_para_deletar:
                try:
                    if arq and os.path.exists(arq) and arq != caminho:
                        os.remove(arq)
                except Exception:
                    pass
            del self._arquivos_para_deletar

        supabase_id = self._thread.dados.get("_supabase_id_resultado")
        salvar_historico(
            self._thread.dados, caminho,
            usuario     = self.usuario_logado,
            supabase_id = supabase_id,
        )

        # Se era reedição, marcar ordem anterior como ALTERADO no Supabase
        if hasattr(self, "_reeditando_id_anterior") and self._reeditando_id_anterior:
            try:
                import urllib.request as _ureq
                obs = f"Alterado para #{supabase_id}" if supabase_id else "Alterado"
                body = json.dumps({
                    "status":     "ALTERADO",
                    "ativo":      False,
                    "observacao": obs,
                }).encode()
                req = _ureq.Request(
                    f"{SUPABASE_URL}/rest/v1/carregamentos?id=eq.{self._reeditando_id_anterior}",
                    data=body,
                    headers={
                        "apikey":        SUPABASE_KEY,
                        "Authorization": f"Bearer {SUPABASE_KEY}",
                        "Content-Type":  "application/json",
                        "Prefer":        "return=minimal",
                    },
                    method="PATCH"
                )
                _ureq.urlopen(req, timeout=8)
            except Exception:
                pass
            del self._reeditando_id_anterior

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

        # Popup de gravação na planilha removido — dados já vão para o Supabase automaticamente

    def _dialog_gravar_planilha(self, dados):
        conta = self._planilha_widget._combo_conta.currentText()
        if not conta or conta == "(nenhuma conta)":
            # Sem conta configurada — avisa e pede para selecionar
            QMessageBox.warning(
                self, "Nenhuma conta configurada",
                "Configure uma conta Gmail na aba 'Controle de Pedidos' para gravar na planilha."
            )
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Gravar na Planilha?")
        dlg.setFixedSize(420, 360)
        dlg.setStyleSheet(DIALOG_SS)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(10)

        # ── Título ──
        t = QLabel("REGISTRAR ORDEM NA PLANILHA")
        t.setStyleSheet(f"color: {TEXT}; font-size: 13px; font-weight: 700; letter-spacing: 0.8px; background: transparent;")
        lay.addWidget(t)

        # ── Info do pedido ──
        pedido  = dados.get("Pedido", "—")
        cliente = dados.get("Cliente", "—")
        produto = dados.get("Produto", "—")
        peso    = dados.get("Peso", "—")
        destino = dados.get("Destino", "—")

        card = QFrame()
        card.setStyleSheet(f"QFrame {{ background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 8px; }}")
        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(14, 10, 14, 10)
        card_lay.setSpacing(4)
        for label, valor in [
            ("Cliente",  cliente),
            ("Pedido",   pedido),
            ("Produto",  produto),
            ("Destino",  destino),
            ("Peso",     f"{peso} t"),
        ]:
            row = QHBoxLayout()
            lbl = QLabel(label.upper())
            lbl.setStyleSheet(f"color: {MUTED}; font-size: 9px; font-weight: 700; letter-spacing: 0.5px; background: transparent; min-width: 60px;")
            val = QLabel(str(valor))
            val.setStyleSheet(f"color: {TEXT}; font-size: 12px; background: transparent;")
            row.addWidget(lbl)
            row.addWidget(val, 1)
            card_lay.addLayout(row)

        # Aviso de desconto de saldo
        lbl_saldo = QLabel(f"⚡  {peso} t serão descontadas do saldo do pedido {pedido}")
        lbl_saldo.setStyleSheet(f"color: #e3b341; font-size: 11px; font-weight: 600; background: transparent;")
        lbl_saldo.setWordWrap(True)
        card_lay.addWidget(lbl_saldo)
        lay.addWidget(card)

        # ── Status ──
        lbl_st = QLabel("STATUS:")
        lbl_st.setStyleSheet(f"color: {MUTED}; font-size: 9px; font-weight: 700; letter-spacing: 0.5px; background: transparent;")
        combo_st = QComboBox()
        combo_st.addItems(["AGUARDANDO", "DESCARGA", "CARREGADO", "MARCADO", "CHEGA"])
        combo_st.setStyleSheet(f"""
            QComboBox {{
                background: {SURFACE}; border: 1px solid {BORDER2};
                border-radius: 6px; padding: 7px 10px; color: {TEXT}; font-size: 12px;
            }}
        """)
        lay.addWidget(lbl_st)
        lay.addWidget(combo_st)

        # ── Aba de destino ──
        aba_base = self._base_widget._aba_selecionada or ""
        lbl_aba = QLabel(f"Gravando em:  {aba_base if aba_base else 'aba mais recente (automático)'}")
        lbl_aba.setStyleSheet(f"color: {MUTED}; font-size: 10px; background: transparent;")
        lay.addWidget(lbl_aba)

        btns = QHBoxLayout()
        bc = QPushButton("CANCELAR"); bc.setObjectName("btn_cancel")
        bo = QPushButton("GRAVAR E DESCONTAR SALDO"); bo.setObjectName("btn_ok")
        btns.addWidget(bc); btns.addWidget(bo)
        lay.addLayout(btns)

        def gravar():
            try:
                from planilha import gravar_ordem_dupla, _autenticar, _descontar_saldo_pedido
                filial = self.empresa if self.empresa else "AGROVIA"
                st     = combo_st.currentText().upper()
                aba    = self._base_widget._aba_selecionada or None
                gravou_saldo = gravar_ordem_dupla(conta, dados, filial, st, aba=aba)
                dlg.accept()

                if gravou_saldo:
                    QMessageBox.information(
                        self, "Sucesso",
                        f"Ordem gravada.\n{peso} t descontadas do saldo do pedido {pedido}."
                    )
                else:
                    QMessageBox.warning(
                        self, "Atenção",
                        f"Ordem gravada na BASE.\n\nPedido {pedido} não encontrado na planilha de saldo — "
                        f"o desconto não foi aplicado.\n\nCadastre o pedido na aba 'Controle de Pedidos' primeiro."
                    )
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
        caixa.setPlaceholderText("🗒️ TAG\nFILIAL: TOP BRASIL\n...")
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

            dados = parsear_mensagem_whatsapp(texto)

            # Verifica se há campos preenchidos e pede confirmação de limpeza
            tem_dados = any(
                (isinstance(v, QLineEdit) and v.text().strip()) or
                (isinstance(v, QComboBox) and v.currentText().strip())
                for v in self.entradas.values()
            )
            if tem_dados:
                msg = QMessageBox(self)
                msg.setWindowTitle("Importar WhatsApp")
                msg.setText("Os campos atuais serão limpos.\n\nDeseja continuar?")
                msg.setIcon(QMessageBox.NoIcon)
                msg.setStyleSheet(f"""
                    QMessageBox {{ background-color: {BG}; }}
                    QLabel {{ color: {TEXT}; font-size: 13px; }}
                    QPushButton {{
                        border-radius: 6px; padding: 7px 18px;
                        font-weight: 700; font-size: 12px; min-width: 80px;
                    }}
                """)
                btn_sim = msg.addButton("CONTINUAR", QMessageBox.AcceptRole)
                btn_sim.setStyleSheet(f"background-color: {DANGER}; color: white; border: none;")
                btn_nao = msg.addButton("CANCELAR", QMessageBox.RejectRole)
                btn_nao.setStyleSheet(f"background-color: transparent; border: 1px solid {BORDER2}; color: {MUTED};")
                msg.exec()
                if msg.clickedButton() != btn_sim:
                    return

            # Limpa campos sem chamar nova_ordem (que abriria diálogo de empresa)
            for v in self.entradas.values():
                if isinstance(v, QLineEdit):
                    v.clear()
                elif isinstance(v, QComboBox):
                    v.setCurrentIndex(-1)
                elif isinstance(v, QDateEdit):
                    v.setDate(QDate.currentDate())
            for i, (stack, linha) in enumerate(self._pedido_linhas):
                for inp in linha.values():
                    if isinstance(inp, QLineEdit):
                        inp.clear()
                    elif isinstance(inp, QComboBox):
                        inp.setCurrentIndex(-1)
                stack.setCurrentIndex(1 if i == 0 else 0)
            self.setar_data_hoje()

            # Define empresa direto pela tag — sem abrir diálogo
            empresa = dados.get("empresa", "")
            if empresa:
                self.empresa = empresa
                cor = ACCENT if "AGRO" in empresa.upper() else DANGER
                self.btn1.setStyleSheet(f"background-color: {cor}; color: white; border: none;")
                self._atualizar_fundo("Agrovia" if "AGRO" in empresa.upper() else "TopBrasil")
            else:
                self.escolher_empresa()

            # Restaura assinatura
            if self.assinatura_usuario and "Assinatura" in self.entradas:
                self.entradas["Assinatura"].setText(self.assinatura_usuario)

            self._preencher_campos(dados)
            dlg.accept()

        bo.clicked.connect(confirmar)
        dlg.exec()

    def _preencher_campos(self, dados):
                                               
        num_pedidos = dados.get("_num_pedidos", 1)
        while len(self._pedido_linhas) < num_pedidos:
            self._adicionar_linha_pedido()

        campos_simples = ["Fábrica", "Cliente", "Fazenda", "Origem",
                          "Destino", "Motorista", "Cavalo", "Pagador",
                          "Solicitante", "Agência", "UF", "Frete/Emp", "Frete/Mot",
                          "Rota", "Agenciamento", "Colocador", "Pagamento", "Peso Total"]
        for campo in campos_simples:
            valor = dados.get(campo, "")
            if not valor:
                continue
            w = self.entradas.get(campo)
            if isinstance(w, QLineEdit):
                w.setText(valor)
            elif isinstance(w, QComboBox):
                w.setEditText(valor)

        # Sincroniza campo UF do cabeçalho com o de Dados da Planilha
        uf_val = dados.get("UF", "")
        if hasattr(self, "_uf_cab") and uf_val:
            self._uf_cab.setText(uf_val)

        for idx in range(num_pedidos):
            sufixo = f" {idx + 1}" if idx > 0 else ""
            # Garante que a linha está visível (página 1 do stack)
            if idx < len(self._pedido_linhas):
                self._pedido_linhas[idx][0].setCurrentIndex(1)
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
            emp = "Agrovia" if "AGRO" in emp.upper() else "TopBrasil"
            self.empresa = emp
            cor = ACCENT if emp == "Agrovia" else DANGER
            self.btn1.setStyleSheet(f"background-color: {cor}; color: white; border: none;")
            self._atualizar_fundo(emp)

    def nova_ordem(self):
        # Verifica se há algum campo preenchido antes de perguntar
        tem_dados = any(
            (isinstance(v, QLineEdit) and v.text().strip()) or
            (isinstance(v, QComboBox) and v.currentText().strip())
            for v in self.entradas.values()
        )

        if tem_dados:
            msg = QMessageBox(self)
            msg.setWindowTitle("Nova Ordem")
            msg.setText("Todos os campos preenchidos serão perdidos.\n\nDeseja continuar?")
            msg.setIcon(QMessageBox.NoIcon)
            msg.setStyleSheet(f"""
                QMessageBox {{ background-color: {BG}; }}
                QLabel {{ color: {TEXT}; font-size: 13px; }}
                QPushButton {{
                    border-radius: 6px; padding: 7px 18px;
                    font-weight: 700; font-size: 12px; min-width: 80px;
                }}
            """)
            btn_sim = msg.addButton("CONTINUAR", QMessageBox.AcceptRole)
            btn_sim.setStyleSheet(f"background-color: {DANGER}; color: white; border: none;")
            btn_nao = msg.addButton("CANCELAR", QMessageBox.RejectRole)
            btn_nao.setStyleSheet(f"background-color: transparent; border: 1px solid {BORDER2}; color: {MUTED};")
            msg.exec()
            if msg.clickedButton() != btn_sim:
                return

        # Limpa todos os campos
        for v in self.entradas.values():
            if isinstance(v, QLineEdit):
                v.clear()
            elif isinstance(v, QComboBox):
                v.setCurrentIndex(-1)
            elif isinstance(v, QDateEdit):
                v.setDate(QDate.currentDate())

        # Reseta linhas: primeira ativa (página 1), demais volta para botão (página 0)
        for i, (stack, linha) in enumerate(self._pedido_linhas):
            for inp in linha.values():
                if isinstance(inp, QLineEdit):
                    inp.clear()
                elif isinstance(inp, QComboBox):
                    inp.setCurrentIndex(-1)
            stack.setCurrentIndex(1 if i == 0 else 0)

        self.escolher_empresa()
        self.setar_data_hoje()
        # Restaura assinatura do usuário logado após limpar campos
        if self.assinatura_usuario and "Assinatura" in self.entradas:
            self.entradas["Assinatura"].setText(self.assinatura_usuario)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = UI()
    win.show()
    sys.exit(app.exec())