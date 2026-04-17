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
from planilha import carregar_blocos_dados, carregar_base

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
    cb.setStyleSheet("""
        QComboBox {
            background: #161b22;
            border: 1px solid #30363d;
            border-radius: 6px;
            padding: 6px 10px;
            color: #e6edf3;
            font-size: 12px;
        }
        QComboBox:focus { border-color: #238636; }
        QComboBox::drop-down { border: none; width: 20px; }
        QComboBox::down-arrow {
            border-left: 4px solid transparent;
            border-right: 4px solid transparent;
            border-top: 5px solid #8b949e;
            width: 0; height: 0; margin-right: 6px;
        }
        QComboBox QAbstractItemView {
            background-color: #161b22;
            border: 1px solid #30363d;
            color: #e6edf3;
            selection-background-color: #23863633;
            selection-color: #e6edf3;
            outline: none;
        }
    """)
    return cb

def make_date():
    d = QDateEdit()
    d.setDisplayFormat("dd/MM/yyyy")
    d.setDate(QDate.currentDate())
    d.setCalendarPopup(True)
    d.setMinimumHeight(32)
    return d

def _to_float(v):
    """Converte string para float aceitando tanto ponto quanto vírgula decimal."""
    try:
        s = str(v or "0").strip()
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        return float(s)
    except:
        return 0.0

def _forcar_maiusculo(inp, texto):
    if texto != texto.upper():
        inp.blockSignals(True)
        c = inp.cursorPosition()
        inp.setText(texto.upper())
        inp.setCursorPosition(c)
        inp.blockSignals(False)

def _formatar_placa(inp, texto):
    # Remove tudo exceto letras e números (mantém hífen se já digitado)
    limpo = re.sub(r"[^A-Za-z0-9]", "", texto).upper()
    # Formato: XXX-XXXX + sufixo (ex: RCY5C45BA → RCY-5C45BA)
    # Os 3 primeiros são letras/prefixo, depois hífen, depois o resto
    if len(limpo) > 3:
        formatado = limpo[:3] + "-" + limpo[3:]
    else:
        formatado = limpo
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

    registro = {
        "data_hora": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "usuario":   usuario or "",
        "motorista": dados.get("Motorista", ""),
        "placa":     dados.get("Cavalo", ""),
        "empresa":   dados.get("empresa", ""),
        "arquivo":   caminho_arquivo,
    }
    historico.insert(0, registro)
    historico = historico[:200]
    path.write_text(json.dumps(historico, ensure_ascii=False, indent=2), encoding="utf-8")

    # Tenta gravar na planilha centralizada (não bloqueia se falhar)
    if conta and conta != "(nenhuma conta)":
        try:
            from planilha import gravar_historico_planilha
            gravar_historico_planilha(conta, registro)
        except Exception:
            pass

def carregar_historico():
    path = _historico_path()
    if not path.exists():
        return []
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []

# ── Supabase ───────────────────────────────────────────────────────
SUPABASE_URL = "https://xlirwzkmvkzldrssmhxg.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhsaXJ3emttdmt6bGRyc3NtaHhnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzU2MDgwMTksImV4cCI6MjA5MTE4NDAxOX0.ofTAEn628a-7JzF3REPj-tBcQJUrlXdfaFSbU5Ysfx4"

def _sb_request(endpoint, method="GET", body=None):
    import urllib.request
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

def carregar_historico_supabase(limite=300):
    try:
        import urllib.request
        req = urllib.request.Request(
            f"{SUPABASE_URL}/rest/v1/carregamentos"
            f"?select=id,criado_em,data,filial,pagador,motorista,placa,"
            f"fabrica,destino,uf,peso,status,pedido,produto,embalagem,colocador,cliente,"
            f"usuario,ativo,observacao,pagamento,frete_emp,frete_mot,rota,"
            f"agenciamento,agencia,origem,cpf,contato,carroceria,carreta1,carreta2,"
            f"carreta3,fazenda,solicitante,"
            f"peso1,peso2,peso3,peso4,"
            f"pedido2,produto2,embalagem2,pedido3,produto3,embalagem3,"
            f"pedido4,produto4,embalagem4"
            f"&order=id.desc&limit={limite}",
            headers={
                "apikey":        SUPABASE_KEY,
                "Authorization": f"Bearer {SUPABASE_KEY}",
            }
        )
        with urllib.request.urlopen(req, timeout=10) as r:
            return json.loads(r.read().decode("utf-8"))
    except Exception:
        return []

def _atualizar_supabase_linha(sb_id, campos):
    _sb_request(f"carregamentos?id=eq.{sb_id}", method="PATCH", body=campos)

def _deletar_supabase_linha(sb_id):
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

_assinaturas_supabase = {}

USUARIOS = {
    "FELIPE":   "Felipe Costa",
    "MARCOS":   "Marcos Silva",
    "ANA":      "Ana Souza",
    "RAFAEL":   "Rafael Lima",
}

def _usuarios_path():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent / "usuarios.json"
    return Path(__file__).parent / "usuarios.json"

def carregar_usuarios():
    """Carrega usuarios da tabela 'usuarios' no Supabase. Fallback para JSON local."""
    try:
        import urllib.request as _ureq
        req = _ureq.Request(
            f"{SUPABASE_URL}/rest/v1/usuarios?select=usuario,nome,assinatura&order=usuario.asc",
            headers={
                "apikey":        SUPABASE_KEY,
                "Authorization": f"Bearer {SUPABASE_KEY}",
            }
        )
        with _ureq.urlopen(req, timeout=5) as r:
            rows = json.loads(r.read().decode("utf-8"))
        if rows:
            # Retorna {USUARIO: nome} e guarda assinatura em _assinaturas_supabase
            global _assinaturas_supabase
            _assinaturas_supabase = {
                str(r["usuario"]).upper(): str(r.get("assinatura") or r.get("nome") or r["usuario"])
                for r in rows
            }
            return {str(r["usuario"]).upper(): str(r.get("nome") or r["usuario"]) for r in rows}
    except Exception:
        pass
    # Fallback: JSON local
    path = _usuarios_path()
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            pass
    return dict(USUARIOS)




class HistoricoWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background: transparent;")
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(10)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")

        # Topo com busca
        topo = QHBoxLayout()
        titulo = QLabel("HISTÓRICO DE ORDENS")
        titulo.setStyleSheet(f"color: {TEXT}; font-size: 14px; font-weight: 700; letter-spacing: 1px; background: transparent;")
        self._inp_busca = QLineEdit()
        self._inp_busca.setPlaceholderText("Buscar motorista, placa, pedido, cliente...")
        self._inp_busca.setFixedWidth(300)
        self._inp_busca.setStyleSheet(f"""
            QLineEdit {{
                background: {SURFACE}; border: 1px solid {BORDER2};
                border-radius: 6px; padding: 6px 10px; color: {TEXT}; font-size: 12px;
            }}
            QLineEdit:focus {{ border-color: {ACCENT}; }}
        """)
        self._inp_busca.textChanged.connect(self._filtrar)
        self._btn_att = QPushButton("↺  ATUALIZAR")
        self._btn_att.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: 1px solid {BORDER2};
                border-radius: 6px; color: {MUTED}; padding: 6px 12px; font-size: 12px;
            }}
            QPushButton:hover {{ color: {TEXT}; border-color: {ACCENT}; }}
            QPushButton:disabled {{ opacity: 0.4; }}
        """)
        self._btn_att.clicked.connect(self.recarregar)

        self._lbl_carregando = QLabel("")
        self._lbl_carregando.setStyleSheet(f"color: {ACCENT}; font-size: 11px; background: transparent;")

        topo.addWidget(titulo); topo.addStretch()
        topo.addWidget(self._lbl_carregando)
        topo.addWidget(self._inp_busca); topo.addWidget(self._btn_att)
        root.insertLayout(0, topo)

        # Stack: 0 = loading, 1 = cards
        self._stack_hist = QStackedWidget()
        self._stack_hist.setStyleSheet("background: transparent;")

        # Página 0 — loading
        pg_load = QWidget()
        pg_load.setStyleSheet("background: transparent;")
        load_lay = QVBoxLayout(pg_load)
        load_lay.setAlignment(Qt.AlignCenter)
        self._lbl_load_anim = QLabel("⟳")
        self._lbl_load_anim.setAlignment(Qt.AlignCenter)
        self._lbl_load_anim.setStyleSheet(f"color: {ACCENT}; font-size: 36px; background: transparent;")
        lbl_load_txt = QLabel("Sincronizando histórico...")
        lbl_load_txt.setAlignment(Qt.AlignCenter)
        lbl_load_txt.setStyleSheet(f"color: {MUTED}; font-size: 13px; background: transparent;")
        load_lay.addStretch()
        load_lay.addWidget(self._lbl_load_anim)
        load_lay.addSpacing(8)
        load_lay.addWidget(lbl_load_txt)
        load_lay.addStretch()

        # Timer para animar o ícone
        self._load_timer = QTimer()
        self._load_timer.setInterval(300)
        _frames = ["⟳", "↻", "⟳", "↺"]
        self._load_frame = [0]
        def _animar():
            self._load_frame[0] = (self._load_frame[0] + 1) % len(_frames)
            self._lbl_load_anim.setText(_frames[self._load_frame[0]])
        self._load_timer.timeout.connect(_animar)

        # Página 1 — cards
        self._container = QWidget()
        self._container.setStyleSheet("background: transparent;")
        self._vbox = QVBoxLayout(self._container)
        self._vbox.setSpacing(6)
        self._vbox.setContentsMargins(0, 0, 4, 0)
        self._vbox.addStretch()

        scroll.setWidget(self._container)

        self._stack_hist.addWidget(pg_load)    # índice 0
        self._stack_hist.addWidget(scroll)     # índice 1
        root.addWidget(self._stack_hist)

        self._todos_cards = []
        self.recarregar()

    def recarregar(self):
        # Mostrar tela de loading
        self._stack_hist.setCurrentIndex(0)
        self._load_timer.start()
        self._btn_att.setEnabled(False)
        self._btn_att.setText("Carregando...")
        self._lbl_carregando.setText("")
        QApplication.processEvents()

        # Limpar cards atuais
        while self._vbox.count() > 1:
            item = self._vbox.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._todos_cards = []

        # Rodar em thread para não travar UI
        class _Thread(QThread):
            concluido = Signal(list, str)
            def run(self_):
                registros = carregar_historico_supabase()
                if registros:
                    self_.concluido.emit(registros, "supabase")
                else:
                    self_.concluido.emit(carregar_historico(), "local")

        self._thread_hist = _Thread()
        self._thread_hist.concluido.connect(self._on_historico_carregado)
        self._thread_hist.start()

    def _on_historico_carregado(self, registros, fonte):
        self._load_timer.stop()
        self._stack_hist.setCurrentIndex(1)
        self._btn_att.setEnabled(True)
        self._btn_att.setText("↺  ATUALIZAR")
        self._lbl_carregando.setText("")

        if not registros:
            vazio = QLabel("Nenhuma ordem gerada ainda.")
            vazio.setAlignment(Qt.AlignCenter)
            vazio.setStyleSheet(f"color: {MUTED}; font-size: 13px; background: transparent;")
            self._vbox.insertWidget(0, vazio)
            return

        # Agrupar por data
        grupos = {}
        for r in registros:
            if fonte == "supabase":
                criado = str(r.get("criado_em","") or "")
                try:
                    import datetime as _dt
                    # Supabase pode retornar com T ou espaço: normalizar
                    criado_norm = criado[:19].replace("T", " ")
                    dt = _dt.datetime.strptime(criado_norm, "%Y-%m-%d %H:%M:%S") - _dt.timedelta(hours=3)
                    data = dt.strftime("%d/%m/%Y")
                    r["_data_fmt"] = data
                    r["_hora_fmt"] = dt.strftime("%H:%M")
                except Exception:
                    raw = str(r.get("data",""))
                    if len(raw) == 10 and raw[4] == "-":
                        data = raw[8:10]+"/"+raw[5:7]+"/"+raw[:4]
                    else:
                        data = raw
                    r["_data_fmt"] = data
                    r["_hora_fmt"] = ""
            else:
                data = r.get("data_hora","")[:10]
                r["_data_fmt"] = data
                r["_hora_fmt"] = r.get("data_hora","")[-5:]
            grupos.setdefault(data, []).append(r)

        for data, items in grupos.items():
            lbl = QLabel(data)
            lbl.setStyleSheet(f"color: {MUTED}; font-size: 10px; font-weight: 700; letter-spacing: 1px; background: transparent; padding: 6px 0 2px 0;")
            self._vbox.insertWidget(self._vbox.count() - 1, lbl)
            for r in items:
                card, frame = self._make_card(r, fonte)
                self._todos_cards.append((r, frame))
                self._vbox.insertWidget(self._vbox.count() - 1, frame)

    def _filtrar(self, texto):
        txt = texto.strip().upper()
        for r, frame in self._todos_cards:
            if not txt:
                frame.setVisible(True)
                continue
            campos = [
                str(r.get("motorista","") or r.get("Motorista","")),
                str(r.get("placa","") or r.get("Cavalo","")),
                str(r.get("pagador","") or r.get("Cliente","")),
                str(r.get("pedido","") or r.get("Pedido","")),
                str(r.get("destino","") or r.get("Destino","")),
                str(r.get("colocador","") or r.get("Colocador","")),
                str(r.get("id","") or r.get("supabase_id","")),
            ]
            frame.setVisible(any(txt in c.upper() for c in campos))

    def _make_card(self, r, fonte="supabase"):
        # Extrair campos compatíveis com Supabase e local
        sb_id    = r.get("id") or r.get("supabase_id", "")
        hora     = r.get("_hora_fmt", r.get("data_hora","")[-5:])
        filial   = str(r.get("filial","") or "").upper()
        empresa  = r.get("empresa","") or ("Agrovia" if "AGRO" in filial else "TopBrasil" if filial else "")
        motorista_txt = str(r.get("motorista","") or r.get("dados",{}).get("Motorista","—") or "—").title()
        placa_txt     = str(r.get("placa","")     or r.get("dados",{}).get("Cavalo","—")    or "—")
        cliente_txt   = str(r.get("pagador","")   or r.get("dados",{}).get("Cliente","")   or "").upper()
        pedido_txt    = str(r.get("pedido","")    or r.get("dados",{}).get("Pedido","")    or "")
        destino_txt   = str(r.get("destino","")   or r.get("dados",{}).get("Destino","")   or "").upper()
        colocador_txt = str(r.get("colocador","") or r.get("dados",{}).get("Colocador","") or "").upper()
        status_txt    = str(r.get("status","")    or "").upper()
        ativo         = r.get("ativo", True)
        inativo       = (not ativo) or status_txt in ("DESISTIU","ALTERADO")
        obs_txt       = str(r.get("observacao","") or "")

        cor_emp   = ACCENT if "AGRO" in empresa.upper() else DANGER
        cor_borda = BORDER2 if not inativo else "#f8514933"
        bg_cor    = "#2a08082a" if inativo else SURFACE

        frame = QFrame()
        frame.setMinimumHeight(52)
        frame.setStyleSheet(f"""
            QFrame {{
                background-color: {bg_cor};
                border: 1px solid {cor_borda};
                border-radius: 8px;
            }}
        """)

        h = QHBoxLayout(frame)
        h.setContentsMargins(12, 8, 12, 8)
        h.setSpacing(8)

        # Número da ordem + hora
        num_lbl = QLabel(f"#{sb_id}" if sb_id else "")
        num_lbl.setStyleSheet(f"color: {ACCENT}; font-size: 12px; font-weight: 700; background: transparent; min-width: 48px;")
        hora_lbl = QLabel(hora)
        hora_lbl.setStyleSheet(f"color: {MUTED}; font-size: 10px; background: transparent; min-width: 38px;")
        num_col = QVBoxLayout(); num_col.setSpacing(1)
        num_col.addWidget(num_lbl); num_col.addWidget(hora_lbl)

        # Motorista + placa
        mot_lbl = QLabel(motorista_txt)
        mot_lbl.setStyleSheet(f"color: {TEXT}; font-size: 12px; font-weight: 600; background: transparent;")
        pla_lbl = QLabel(placa_txt)
        pla_lbl.setStyleSheet(f"color: {MUTED}; font-size: 10px; background: transparent;")
        mot_col = QVBoxLayout(); mot_col.setSpacing(1)
        mot_col.addWidget(mot_lbl); mot_col.addWidget(pla_lbl)
        # mot_col substituído por mot_w abaixo

        def _info(label, valor, maxw=None):
            w = QWidget()
            w.setStyleSheet("background: transparent; border: none;")
            if maxw: w.setMaximumWidth(maxw)
            v = QVBoxLayout(w); v.setSpacing(0); v.setContentsMargins(0,0,0,0)
            l = QLabel(label)
            l.setStyleSheet(f"color: {MUTED}; font-size: 8px; font-weight: 700; letter-spacing: 0.5px; background: transparent;")
            t = QLabel(str(valor or "—"))
            t.setStyleSheet(f"color: {TEXT}; font-size: 11px; background: transparent;")
            if maxw: t.setMaximumWidth(maxw - 4)
            v.addWidget(l); v.addWidget(t)
            return w

        # Badge empresa
        badge = QLabel(empresa.upper() if empresa else "—")
        badge.setStyleSheet(f"""
            color: {cor_emp}; background: transparent;
            font-size: 10px; font-weight: 700; padding: 2px 4px;
        """)

        # Status badge
        STATUS_CORES = {
            "CARREGADO": "#3fb950","PAGO": "#58a6ff","MARCADO": "#e3b341",
            "AGUARDANDO": "#8b949e","DESCARGA": "#bc8cff",
            "CHEGA": "#79c0ff","A CAMINHO": "#79c0ff",
            "DESISTIU": "#f85149","ALTERADO": "#f85149",
        }
        st_cor = STATUS_CORES.get(status_txt, MUTED)
        if inativo and obs_txt:
            st_txt = f"{status_txt}: {obs_txt}"
        else:
            st_txt = status_txt or "—"
        st_badge = QLabel(st_txt)
        st_badge.setStyleSheet(f"""
            color: {st_cor}; background: transparent;
            font-size: 10px; font-weight: 700; padding: 2px 4px;
        """)
        st_badge.setMaximumWidth(180)

        # Botão editar
        btn = QPushButton("EDITAR")
        btn.setFixedHeight(28)
        btn.setFixedWidth(80)
        btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        btn.setFont(QFont("Segoe UI", 9, QFont.Bold))
        btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: 1px solid #e3b341;
                border-radius: 5px; color: #e3b341; font-weight: 700;
                padding: 0px 8px;
            }}
            QPushButton:hover {{ background: #e3b34122; }}
        """)
        btn.clicked.connect(lambda _, reg=r: self._reeditar_ordem(reg))

        usuario_txt = str(r.get("usuario","") or r.get("dados",{}).get("_usuario","") or "").upper()

        mot_w = QWidget(); mot_w.setStyleSheet("background:transparent;")
        mot_w.setMaximumWidth(180)
        mot_v = QVBoxLayout(mot_w); mot_v.setSpacing(1); mot_v.setContentsMargins(0,0,0,0)
        mot_v.addWidget(mot_lbl); mot_v.addWidget(pla_lbl)

        h.addLayout(num_col)
        h.addWidget(mot_w, 1)
        h.addWidget(_info("CLIENTE",  cliente_txt, 140), 0)
        h.addWidget(_info("PEDIDO",   pedido_txt,   70), 0)
        h.addWidget(_info("DESTINO",  destino_txt, 120), 0)
        h.addWidget(_info("USUARIO",  usuario_txt,  70), 0)
        h.addStretch(1)
        h.addWidget(badge)
        h.addWidget(st_badge)
        h.addSpacing(6)
        h.addWidget(btn)

        return r, frame

    def _reeditar_ordem(self, r):
        """Carrega dados no formulário para reeditar. Suporta Supabase e local."""
        dados = r.get("dados")
        sb_id = r.get("id") or r.get("supabase_id")

        # Se veio do Supabase sem dados completos, busca no histórico local pelo supabase_id
        if not dados and sb_id:
            historico_local = carregar_historico()
            for reg in historico_local:
                if str(reg.get("supabase_id","")) == str(sb_id) and reg.get("dados"):
                    dados = reg["dados"]
                    break

        # Se ainda não tem dados, monta o básico do que o Supabase tem
        if not dados and r.get("motorista"):
            filial = str(r.get("filial","") or "").upper()
            dados = {
                "empresa":    "Agrovia" if "AGRO" in filial else "TopBrasil",
                "Motorista":  str(r.get("motorista","") or ""),
                "Cavalo":     str(r.get("placa","") or ""),
                "Pagador":    str(r.get("pagador","") or ""),
                "Cliente":    str(r.get("cliente","") or r.get("pagador","") or ""),
                "Fábrica":    str(r.get("fabrica","") or ""),
                "Destino":    str(r.get("destino","") or ""),
                "UF":         str(r.get("uf","") or ""),
                "Pedido":     str(r.get("pedido","") or ""),
                "Produto":    str(r.get("produto","") or ""),
                "Embalagem":  str(r.get("embalagem","") or ""),
                "Colocador":    str(r.get("colocador","") or ""),
                "Pagamento":    str(r.get("pagamento","") or ""),
                "Frete/Emp":    str(r.get("frete_emp","") or ""),
                "Frete/Mot":    str(r.get("frete_mot","") or ""),
                "Rota":         str(r.get("rota","") or ""),
                "Agenciamento": str(r.get("agenciamento","") or ""),
                "Agência":      str(r.get("agencia","") or ""),
                "Origem":       str(r.get("origem","") or ""),
                "CPF":          str(r.get("cpf","") or ""),
                "Contato":      str(r.get("contato","") or ""),
                "Carroceria":   str(r.get("carroceria","") or ""),
                "Carreta 1":    str(r.get("carreta1","") or ""),
                "Carreta 2":    str(r.get("carreta2","") or ""),
                "Carreta 3":    str(r.get("carreta3","") or ""),
                "Fazenda":      str(r.get("fazenda","") or ""),
                "Solicitante":  str(r.get("solicitante","") or ""),
                "Peso Total":   str(r.get("peso","") or ""),
                "_num_pedidos":  sum(1 for k in ["pedido","pedido2","pedido3","pedido4"] if r.get(k)),
                "Peso":         str(r.get("peso1","") or r.get("peso","") or ""),
                "Peso 2":       str(r.get("peso2","") or ""),
                "Peso 3":       str(r.get("peso3","") or ""),
                "Peso 4":       str(r.get("peso4","") or ""),
                "Pedido 2":     str(r.get("pedido2","") or ""),
                "Produto 2":    str(r.get("produto2","") or ""),
                "Embalagem 2":  str(r.get("embalagem2","") or ""),
                "Pedido 3":     str(r.get("pedido3","") or ""),
                "Produto 3":    str(r.get("produto3","") or ""),
                "Embalagem 3":  str(r.get("embalagem3","") or ""),
                "Pedido 4":     str(r.get("pedido4","") or ""),
                "Produto 4":    str(r.get("produto4","") or ""),
                "Embalagem 4":  str(r.get("embalagem4","") or ""),
            }

        if not dados:
            QMessageBox.warning(self, "Aviso",
                "Este registro nao possui dados completos para reeditar.")
            return

        arquivo_antigo = r.get("arquivo","")
        arquivo_xlsx   = arquivo_antigo if arquivo_antigo and arquivo_antigo.endswith(".xlsx") else ""
        arquivo_pdf    = arquivo_xlsx.replace(".xlsx",".pdf") if arquivo_xlsx else ""

        motorista   = dados.get("Motorista","")
        placa       = dados.get("Cavalo","")
        empresa_orig = dados.get("empresa","")
        emp_norm     = "Agrovia" if "AGRO" in str(empresa_orig).upper() else "TopBrasil"
        cor_emp      = "#238636" if emp_norm == "Agrovia" else "#da3633"

        dlg = QDialog(self)
        dlg.setWindowTitle("Editar Ordem")
        dlg.setFixedSize(440, 240)
        dlg.setStyleSheet(DIALOG_SS)
        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(10)

        lbl = QLabel(f"Editar ordem de <b>{motorista.title()}</b> — placa <b>{placa}</b>")
        lbl.setWordWrap(True)
        lay.addWidget(lbl)

        lbl_emp = QLabel(f"Empresa atual: <b style='color:{cor_emp}'>{emp_norm.upper()}</b>")
        lbl_emp.setStyleSheet("background: transparent;")
        lbl_emp.setWordWrap(True)
        lay.addWidget(lbl_emp)

        lbl2 = QLabel("Uma nova ordem sera gerada e a anterior marcada como ALTERADO.")
        lbl2.setStyleSheet(f"color: #e3b341; font-size: 11px; background: transparent;")
        lbl2.setWordWrap(True)
        lay.addWidget(lbl2)

        # Seleção de empresa para nova ordem
        lbl3 = QLabel("Empresa da nova ordem:")
        lbl3.setStyleSheet(f"color: {MUTED}; font-size: 11px; background: transparent;")
        lay.addWidget(lbl3)

        emp_btns = QHBoxLayout(); emp_btns.setSpacing(8)
        btn_agro = QPushButton("AGROVIA")
        btn_top  = QPushButton("TOPBRASIL")
        for b in [btn_agro, btn_top]:
            b.setFixedHeight(32)
            b.setCheckable(True)
            b.setStyleSheet(f"""
                QPushButton {{
                    background: transparent; border: 1px solid {BORDER2};
                    border-radius: 6px; color: {MUTED}; font-size: 11px; font-weight: 700;
                }}
                QPushButton:checked {{ border-color: #238636; color: #238636; background: #23863622; }}
            """)
        btn_top.setStyleSheet(f"""
            QPushButton {{
                background: transparent; border: 1px solid {BORDER2};
                border-radius: 6px; color: {MUTED}; font-size: 11px; font-weight: 700;
            }}
            QPushButton:checked {{ border-color: #da3633; color: #da3633; background: #da363322; }}
        """)
        # Pré-selecionar empresa atual
        if emp_norm == "Agrovia":
            btn_agro.setChecked(True)
        else:
            btn_top.setChecked(True)

        def _sel_agro():
            btn_agro.setChecked(True)
            btn_top.setChecked(False)
        def _sel_top():
            btn_top.setChecked(True)
            btn_agro.setChecked(False)
        btn_agro.clicked.connect(_sel_agro)
        btn_top.clicked.connect(_sel_top)

        emp_btns.addWidget(btn_agro)
        emp_btns.addWidget(btn_top)
        lay.addLayout(emp_btns)

        btns = QHBoxLayout()
        bc = QPushButton("CANCELAR"); bc.setObjectName("btn_cancel")
        bo = QPushButton("CONTINUAR"); bo.setObjectName("btn_ok")
        btns.addWidget(bc); btns.addWidget(bo)
        lay.addLayout(btns)
        bc.clicked.connect(dlg.reject)

        def continuar():
            dlg.accept()
            empresa_nova = "Agrovia" if btn_agro.isChecked() else "TopBrasil"
            dados["empresa"] = empresa_nova

            ui = None
            w = self.parent()
            while w:
                if hasattr(w, "_nav") and hasattr(w, "_preencher_campos"):
                    ui = w; break
                w = w.parent() if hasattr(w, "parent") else None
            if ui is None:
                QMessageBox.warning(self, "Aviso", "Nao foi possivel navegar para o formulario.")
                return
            if hasattr(ui, "_aplicar_empresa"):
                ui._aplicar_empresa(empresa_nova)
            ui._preencher_campos(dados)
            ui._nav(0)
            if sb_id and hasattr(ui, "_entrar_modo_edicao"):
                ui._entrar_modo_edicao(sb_id)
            if arquivo_xlsx:
                ui._arquivos_para_deletar = [arquivo_xlsx, arquivo_pdf]

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

        resp = QMessageBox.question(
            self, "Migrar da BASE",
            "Isso irá importar todos os carregamentos com PEDIDO preenchido "
            "das abas BASE MM/AAAA para as abas PEDIDOS e DADOS.\n\n"
            "Registros já existentes não serão duplicados.\n\n"
            "Deseja continuar?",
            QMessageBox.Yes | QMessageBox.No
        )
        if resp != QMessageBox.Yes:
            return

        self._lbl_status.setText("Migrando... aguarde.")
        QApplication.processEvents()

        try:
            from planilha import migrar_base_para_dados

            def progresso(msg):
                self._lbl_status.setText(msg)
                QApplication.processEvents()

            resultado = migrar_base_para_dados(conta, callback_progresso=progresso)

            abas = ", ".join(resultado["abas_lidas"])
            msg = (
                f"Migração concluída!\n\n"
                f"Abas lidas: {abas}\n"
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

        self._combo_conta = QComboBox()
        self._combo_conta.setFixedWidth(210)
        self._combo_conta.setStyleSheet(COMBO_SS)

        # Select de mês/aba
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
        self._tabela.setAlternatingRowColors(False)
        self._tabela.verticalHeader().setVisible(False)
        self._tabela.horizontalHeader().setStretchLastSection(True)
        self._tabela.setSortingEnabled(True)

        # Colunas visíveis: DATA(0), FILIAL(1), PAGADOR(2), MOTORISTA(4), PLACA(5),
        #                   DESTINO(7), UF(8), PESO(9), STATUS(14)
        self._colunas_visiveis = [0, 1, 2, 4, 5, 7, 8, 9, 14]
        COLUNAS = ["DATA", "FILIAL", "PAGADOR", "MOTORISTA", "PLACA",
                   "DESTINO", "UF", "PESO", "STATUS", "", ""]
        self._tabela.setColumnCount(len(COLUNAS))
        self._tabela.setHorizontalHeaderLabels(COLUNAS)

        larguras = [90, 70, 130, 120, 90, 120, 40, 60, 110, 50, 50]
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
        conta = self._combo_conta.currentText()
        if conta == "(nenhuma conta)":
            return
        try:
            from planilha import listar_abas_base, _aba_mais_recente
            abas = listar_abas_base(conta)
            self._combo_mes.clear()
            self._combo_mes.addItems(abas)
            # Seleciona a aba mais recente por padrão
            mais_recente = _aba_mais_recente(abas)
            idx = self._combo_mes.findText(mais_recente)
            if idx >= 0:
                self._combo_mes.setCurrentIndex(idx)
            self._aba_selecionada = self._combo_mes.currentText()
            self._lbl_status.setText(f"{len(abas)} aba(s) encontrada(s) — selecionado: {self._aba_selecionada}")
        except Exception as e:
            self._lbl_status.setText(f"Erro ao listar abas: {e}")

    def _carregar(self):
        conta = self._combo_conta.currentText()
        if conta == "(nenhuma conta)":
            self._lbl_status.setText("Configure uma conta Gmail primeiro.")
            return

        # Usa aba do combo; se não carregou ainda, lista primeiro
        aba = self._combo_mes.currentText().strip()
        if not aba:
            self._atualizar_abas()
            aba = self._combo_mes.currentText().strip()

        self._aba_selecionada = aba
        self._lbl_status.setText(f"Carregando {aba}...")
        QApplication.processEvents()

        try:
            dados = carregar_base(conta, aba=aba)
            self._todos_dados = dados
            self._renderizar(dados)
            self._lbl_status.setText(f"{len(dados)} ordem(ns) em '{aba}'")
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

            if len(linha) > 17:
                self._row_to_linha_planilha[r] = linha[17]

            status = str(linha[14] if len(linha) > 14 else "").upper()
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

            for col_tabela, col_dados in enumerate(self._colunas_visiveis):
                val = linha[col_dados] if col_dados < len(linha) else ""
                item = QTableWidgetItem(str(val) if val is not None else "")
                item.setTextAlignment(Qt.AlignCenter)
                if bg:
                    item.setBackground(bg)
                self._tabela.setItem(r, col_tabela, item)

            self._tabela.setItem(r, 9,  self._make_btn_item("EDIT", "#e3b341"))
            self._tabela.setItem(r, 10, self._make_btn_item("DEL",  "#da3633"))

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
        if col == 9:
            dados = [self._tabela.item(row, c).text() if self._tabela.item(row, c) else "" for c in range(9)]
            self._toggle_edicao(row, dados)
        elif col == 10:
            dados = [self._tabela.item(row, c).text() if self._tabela.item(row, c) else "" for c in range(9)]
            self._deletar_linha(row, dados)

    def _toggle_edicao(self, row, dados_orig):
        if row in self._linhas_editando:
            self._salvar_edicao(row, dados_orig)
            return

        self._linhas_editando[row] = dados_orig
        self._tabela.setRowHeight(row, 36)

        STATUS_OPTS = ["MARCADO", "CHEGA", "CARREGADO", "AGUARDANDO", "DESCARGA"]

        # Edita colunas 0-7 (DATA até PESO), STATUS fica como combo na col 8
        for c in range(8):
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
            # Força maiúsculo ao digitar (exceto coluna de data)
            if c != 0:
                inp.textChanged.connect(lambda t, i=inp: _forcar_maiusculo(i, t))
            self._tabela.setCellWidget(row, c, inp)

        combo = QComboBox()
        combo.addItems(STATUS_OPTS)
        status_atual = str(dados_orig[8]) if len(dados_orig) > 8 else ""
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
        self._tabela.setCellWidget(row, 8, combo)

        self._tabela.setItem(row, 9, self._make_btn_item("OK", "#2ea043"))

    def _salvar_edicao(self, row, dados_orig):
        conta = self._combo_conta.currentText()
        novos_visiveis = []
        for c in range(8):
            w = self._tabela.cellWidget(row, c)
            novos_visiveis.append(w.text() if isinstance(w, QLineEdit) else (self._tabela.item(row, c).text() if self._tabela.item(row, c) else ""))
        combo = self._tabela.cellWidget(row, 8)
        novos_visiveis.append(combo.currentText() if isinstance(combo, QComboBox) else "")

        # Mapeamento: índice visível → índice real na planilha
        # visível: [0=DATA, 1=FILIAL, 2=PAGADOR, 3=MOTORISTA, 4=PLACA,
        #           5=DESTINO, 6=UF, 7=PESO, 8=STATUS]
        # planilha: [0=DATA,1=FILIAL,2=PAGADOR,3=AGENCIA,4=MOTORISTA,
        #            5=PLACA,6=FABRICA,7=DESTINO,8=UF,9=PESO,
        #            10=FRETE/E,11=FRETE/M,12=ROTA,13=AGENCIAMENTO,
        #            14=STATUS,15=PEDIDO,16=PRODUTO]
        VISIVEL_PARA_PLANILHA = {0: 0, 1: 1, 2: 2, 3: 4, 4: 5,
                                  5: 7, 6: 8, 7: 9, 8: 14}

        try:
            from planilha import atualizar_linha_base, carregar_base_com_linhas
            num_linha = self._row_to_linha_planilha.get(row) or self._encontrar_linha_base(dados_orig, conta)
            if not num_linha:
                QMessageBox.warning(self, "Aviso", "Linha não encontrada na planilha.")
                return

            # Lê a linha atual completa da planilha para não perder colunas não visíveis
            linhas = carregar_base_com_linhas(conta, aba=self._aba_selecionada)
            linha_atual = None
            for n, l in linhas:
                if n == num_linha:
                    linha_atual = list(l)
                    break

            if linha_atual is None:
                QMessageBox.warning(self, "Aviso", "Linha não encontrada na planilha.")
                return

            # Garante tamanho mínimo de 17 colunas
            while len(linha_atual) < 17:
                linha_atual.append("")

            # Aplica apenas as colunas editadas
            for idx_visivel, idx_planilha in VISIVEL_PARA_PLANILHA.items():
                linha_atual[idx_planilha] = novos_visiveis[idx_visivel]

            atualizar_linha_base(conta, num_linha, linha_atual, aba=self._aba_selecionada)

            STATUS_OPTS_COR = {
                "CARREGADO": "#1a3a1a", "NÃO CARREGADO": "#3a1a1a",
                "PAGO": "#1a2a3a", "MARCADO": "#3a2a00",
                "CHEGA": "#1a2a3a", "AGUARDANDO": "#2a2a00",
                "DESCARGA": "#2a1a3a",
            }
            status = novos_visiveis[8].upper()
            bg = QColor(STATUS_OPTS_COR.get(status, "#161b22"))

            for c in range(9):
                self._tabela.removeCellWidget(row, c)
                item = QTableWidgetItem(novos_visiveis[c] if c < len(novos_visiveis) else "")
                item.setTextAlignment(Qt.AlignCenter)
                item.setBackground(bg)
                self._tabela.setItem(row, c, item)

            self._tabela.setItem(row, 9,  self._make_btn_item("EDIT", "#e3b341"))
            self._tabela.setItem(row, 10, self._make_btn_item("DEL",  "#da3633"))

            self._linhas_editando.pop(row, None)
            self._tabela.setRowHeight(row, 28)
            self._lbl_status.setText("Linha atualizada com sucesso.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))

    def _deletar_linha(self, row, dados):
        conta = self._combo_conta.currentText()
        resp = QMessageBox.question(
            self, "Confirmar",
            "Deseja remover esta linha da planilha BASE?\n\n"
            "Se houver um carregamento correspondente na planilha de Pedidos (DADOS), "
            "ele também será removido.",
            QMessageBox.Yes | QMessageBox.No
        )
        if resp != QMessageBox.Yes:
            return
        try:
            from planilha import deletar_linha_base, carregar_base_com_linhas
            num_linha = self._row_to_linha_planilha.get(row) or self._encontrar_linha_base(dados, conta)
            if num_linha:
                # Busca a linha completa para identificar o carregamento em DADOS
                linha_completa = None
                try:
                    linhas = carregar_base_com_linhas(conta, aba=self._aba_selecionada)
                    for n, l in linhas:
                        if n == num_linha:
                            linha_completa = l
                            break
                except Exception:
                    pass
                deletar_linha_base(
                    conta, num_linha,
                    aba=self._aba_selecionada,
                    dados_linha=linha_completa
                )
            self._tabela.removeRow(row)
            self._lbl_status.setText("Linha removida da BASE e dos Pedidos.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", str(e))

    def _encontrar_linha_base(self, dados, conta):
        from planilha import carregar_base_com_linhas
        try:
            linhas = carregar_base_com_linhas(conta, aba=self._aba_selecionada)
            chave = [str(dados[i]) if i < len(dados) else "" for i in range(4)]
            for num_linha, linha in linhas:
                if [str(linha[i]) if i < len(linha) else "" for i in range(4)] == chave:
                    return num_linha
        except Exception:
            pass
        return None


# Mapa de fábricas → origem fixa
# Chave: substring que aparece no nome da fábrica (uppercase)
# Valor: origem a preencher
ORIGENS_FABRICA = {
    "FERTIMAXI":        "FEIRA DE SANTANA - BA",
    "INTERMARITIMA":    "CANDEIAS - BA",
    "ARMAZEM VITORIA":  "CANDEIAS - BA",
    "ARMAZÉM VITÓRIA":  "CANDEIAS - BA",
    "YARA":             "CANDEIAS - BA",
    "UNIGEL":           "CAMAÇARI - BA",
    # TIMAC tratado separadamente em _origem_por_fabrica
}

def _origem_por_fabrica(fabrica):
    """Retorna a origem fixa para uma fábrica, ou '' se não reconhecida.
    Normaliza acentos para garantir match independente de grafia."""
    import unicodedata
    def _norm(s):
        return unicodedata.normalize("NFD", s.upper().strip()).encode("ascii","ignore").decode()
    t = _norm(fabrica)

    # TIMAC: verifica localidade específica primeiro
    if "TIMAC" in t:
        if "CANDEIAS" in t:
            return "CANDEIAS - BA"
        return "CAMAÇARI - BA"

    for chave, origem in ORIGENS_FABRICA.items():
        if _norm(chave) in t:
            return origem
    return ""


def parsear_mensagem_whatsapp(texto):
    resultado = {}

    # Lista de todas as chaves conhecidas para detectar campo vazio
    CHAVES = ["FILIAL","PAGADOR","CLIENTE","AGENCIA","AGÊNCIA","MOTORISTA","PLACA",
              "FABRICA","FÁBRICA","DESTINO","UF","FAZENDA","PESO","FRETE/EMP",
              "FRETE/MOT","ROTA","AGENCIAMENTO","PAGAMENTO","PEDIDO","PRODUTO",
              "EMBALAGEM","COLOCADOR","SOLICITANTE","STATUS"]
    _CHAVES_RE = "|".join(CHAVES)

    def extrair(chave):
        match = re.search(rf"^{chave}\s*:\s*(.*)", texto, re.IGNORECASE | re.MULTILINE)
        if not match:
            return ""
        valor = match.group(1).strip()
        if re.match(rf"^(?:{_CHAVES_RE})\s*:", valor, re.IGNORECASE):
            return ""
        return valor

    filial = extrair("FILIAL").upper()
    resultado["empresa"] = "Agrovia" if "AGRO" in filial else "TopBrasil"

    # ── PAGADOR e CLIENTE ────────────────────────────────────────────
    # Regras:
    # - PAGADOR → Solicitante (quem paga o frete)
    # - CLIENTE → Cliente (destinatário da mercadoria)
    # - Se a tag tiver PAGADOR e CLIENTE na mesma linha (ex: "PAGADOR/CLIENTE: NOME"),
    #   ambos recebem o mesmo valor
    # - Se só PAGADOR existir (sem CLIENTE), PAGADOR vai para Solicitante apenas
    # - Se só CLIENTE existir (sem PAGADOR), vai para Cliente apenas

    pagador = extrair("PAGADOR")
    cliente = extrair("CLIENTE")

    # Verifica se PAGADOR e CLIENTE estão na mesma linha da tag
    match_mesmo = re.search(
        r"^(PAGADOR|CLIENTE)\s*/\s*(PAGADOR|CLIENTE)\s*:\s*(.+)",
        texto, re.IGNORECASE | re.MULTILINE
    )
    if match_mesmo:
        valor_comum = match_mesmo.group(3).strip()
        resultado["Solicitante"] = valor_comum
        resultado["Cliente"]     = valor_comum
        resultado["Pagador"]     = valor_comum
    else:
        if pagador:
            resultado["Solicitante"] = pagador
            resultado["Pagador"]     = pagador
        if cliente:
            resultado["Cliente"] = cliente
            if not pagador:
                resultado["Pagador"] = cliente

    # ── FÁBRICA e ORIGEM ─────────────────────────────────────────────
    # Origem é sempre definida pela fábrica (fixa por regra)
    # Nunca sobrescreve com o texto livre da tag
    fabrica = extrair("FABRICA") or extrair("FÁBRICA")
    if fabrica:
        resultado["Fábrica"] = fabrica.upper().strip()
        origem_fixa = _origem_por_fabrica(fabrica)
        if origem_fixa:
            resultado["Origem"] = origem_fixa

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

    agenciamento = extrair("AGENCIAMENTO")
    if agenciamento:
        resultado["Agenciamento"] = agenciamento

    uf_campo = extrair("UF").upper()
    if uf_campo:
        resultado["UF"] = uf_campo

    produto = extrair("PRODUTO")

    # Detecta embalagem no texto do produto e remove do nome do produto
    embalagem = ""
    prod_upper = produto.upper()
    EMBALAGENS_MAP = [
        ("BIG BAG",     "BIG BAG"),
        ("GRANEL",      "GRANEL"),
        ("PALETIZADO",  "PALETIZADO"),
        ("SACO 50KG",   "SACO 50KG"),
        ("SACO 50",     "SACO 50KG"),
        ("SACO 25KG",   "SACO 25KG"),
        ("SACO 25",     "SACO 25KG"),
        ("SACO 40KG",   "SACO 40KG"),
        ("SACO 40",     "SACO 40KG"),
    ]
    produto_limpo = produto
    for token, emb_label in EMBALAGENS_MAP:
        if token in prod_upper:
            embalagem = emb_label
            # Remove o token do nome do produto (case-insensitive) e limpa espaços/hífens sobrando
            produto_limpo = re.sub(
                rf"[-–\s]*{re.escape(token)}[-–\s]*",
                " ", produto, flags=re.IGNORECASE
            ).strip(" -–")
            break

    resultado["Produto"] = produto_limpo
    # Embalagem: primeiro tenta extração direta da tag, depois do produto
    embalagem_tag = extrair("EMBALAGEM")
    if embalagem_tag:
        resultado["Embalagem"] = embalagem_tag.upper()
    elif embalagem:
        resultado["Embalagem"] = embalagem

    destino_raw = extrair("DESTINO")
    uf = extrair("UF").upper()

    # UF sempre separado — o gerador concatena na célula O10
    if uf:
        resultado["UF"] = uf

    # Destino: só a cidade, sem UF
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

    colocador = extrair("COLOCADOR")
    if colocador:
        resultado["Colocador"] = colocador

    pagamento = extrair("PAGAMENTO")
    if pagamento:
        resultado["Pagamento"] = pagamento

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
        self.entradas["Fábrica"]    = make_input()
        self.entradas["Solicitante"] = make_input()
        r1.addWidget(make_field("Data", self.entradas["Data Apresentação"]), 1)
        r1.addWidget(make_field("Fábrica", self.entradas["Fábrica"]), 2)
        r1.addWidget(make_field("Solicitante", self.entradas["Solicitante"]), 2)
        v.addLayout(r1)

        SOLICITANTE_POR_FABRICA = {
            "ARMAZEM VITORIA":  "FERTIMAXI",
            "ARMAZÉM VITÓRIA":  "FERTIMAXI",
            "INTERMARITIMA":    "FERTIMAXI",
            "INTERMARÍTIMA":    "FERTIMAXI",
        }

        def _atualizar_origem(texto):
            import unicodedata as _ud
            t = _ud.normalize("NFD", texto.upper().strip()).encode("ascii","ignore").decode()
            origem_fixa = _origem_por_fabrica(texto)
            if origem_fixa:
                self.entradas["Origem"].setText(origem_fixa)
            # Solicitante automático
            sol = ""
            for chave, valor in SOLICITANTE_POR_FABRICA.items():
                chave_n = _ud.normalize("NFD", chave.upper()).encode("ascii","ignore").decode()
                if chave_n in t:
                    sol = valor
                    break
            w_sol = self.entradas.get("Solicitante")
            if w_sol and isinstance(w_sol, QLineEdit):
                if sol:
                    w_sol.setText(sol)
                elif not w_sol.text():
                    pass  # não limpa se já preenchido manualmente

        self.entradas["Fábrica"].textChanged.connect(_atualizar_origem)

        r2 = QHBoxLayout(); r2.setSpacing(6)
        self.entradas["Origem"]  = make_input()
        self.entradas["Destino"] = make_input()
        self.entradas["Cliente"] = make_input()
        r2.addWidget(make_field("Origem",  self.entradas["Origem"]),  1)
        r2.addWidget(make_field("Destino", self.entradas["Destino"]), 1)
        r2.addWidget(make_field("Cliente", self.entradas["Cliente"]), 1)
        v.addLayout(r2)

        r3 = QHBoxLayout(); r3.setSpacing(6)
        self.entradas["Pagador"] = make_input()
        r3.addWidget(make_field("Pagador", self.entradas["Pagador"]), 1)
        v.addLayout(r3)

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
        r2.addWidget(make_field("Rota",         self.entradas["Rota"]),         1)
        r2.addWidget(make_field("Agenciamento", self.entradas["Agenciamento"]), 1)
        v.addLayout(r2)

        r3 = QHBoxLayout(); r3.setSpacing(6)
        self.entradas["Colocador"]  = make_input()
        self.entradas["Pagamento"]  = make_input()
        self.entradas["Peso Total"] = make_input()
        self.entradas["Peso Total"].setReadOnly(True)
        self.entradas["Peso Total"].setPlaceholderText("Auto")
        self.entradas["Peso Total"].setStyleSheet(f"""
            QLineEdit {{
                background: #1c2128; border: 1px solid {BORDER2};
                border-radius: 6px; padding: 6px 10px;
                color: {ACCENT}; font-size: 12px; font-weight: 700;
            }}
        """)
        self.entradas["UF"].textChanged.connect(
            lambda t: self._uf_cab.setText(t.upper()) if hasattr(self, "_uf_cab") else None)
        r3.addWidget(make_field("Colocador",  self.entradas["Colocador"]),  2)
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

        self.btn2 = QPushButton("GRAVAR NO BANCO")
        self.btn2.setObjectName("btn_email")

        self.btn3 = QPushButton("NOVA ORDEM")
        self.btn3.setObjectName("btn_nova")

        # Banner modo edição — criado ANTES de ser adicionado ao layout
        self._banner_edicao = QLabel("")
        self._banner_edicao.setAlignment(Qt.AlignCenter)
        self._banner_edicao.setStyleSheet("""
            QLabel {
                background: #3a2a00;
                border: 1px solid #e3b341;
                border-radius: 6px;
                color: #e3b341;
                font-size: 11px;
                font-weight: 700;
                padding: 6px 12px;
                letter-spacing: 0.5px;
            }
        """)
        self._banner_edicao.hide()

        for btn in [self.btn_wpp, self.btn1, self.btn2, self.btn3]:
            btn.setMinimumHeight(40)
            v.addWidget(btn)

        v.addWidget(self._banner_edicao)
        v.addStretch()

        self.btn_wpp.clicked.connect(self.importar_whatsapp)
        self.btn1.clicked.connect(lambda: self.executar(False))
        self.btn2.clicked.connect(self._gravar_banco)
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
                    total += _to_float(w.text())
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

    def _entrar_modo_edicao(self, id_anterior):
        """Ativa o modo edição — mostra banner e muda textos dos botões."""
        self._reeditando_id_anterior = id_anterior
        self._banner_edicao.setText(f"✏  EDITANDO ORDEM #{id_anterior} — Gere ou grave para salvar a nova ordem")
        self._banner_edicao.show()
        self.btn1.setText("GERAR NOVA ORDEM")
        self.btn2.setText("GRAVAR NO BANCO")
        self.btn3.setText("CANCELAR EDIÇÃO")

    def _sair_modo_edicao(self):
        """Desativa o modo edição — restaura textos dos botões."""
        if hasattr(self, "_reeditando_id_anterior"):
            del self._reeditando_id_anterior
        self._banner_edicao.hide()
        self._banner_edicao.setText("")
        self.btn1.setText("GERAR ORDEM")
        self.btn2.setText("GRAVAR NO BANCO")
        self.btn3.setText("NOVA ORDEM")

    def _aplicar_empresa(self, nome_empresa):
        """Define a empresa ativa sem abrir diálogo."""
        if not nome_empresa:
            return
        nome = "Agrovia" if "AGRO" in str(nome_empresa).upper() else "TopBrasil"
        self.empresa = nome
        cor = ACCENT if nome == "Agrovia" else DANGER
        self.btn1.setStyleSheet(f"background-color: {cor}; color: white; border: none;")
        self._atualizar_fundo(nome)

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
                import unicodedata as _ud
                def _norm(s):
                    return _ud.normalize("NFD", s.upper().strip()).encode("ascii","ignore").decode()
                login = combo_usuario.currentText().strip()
                login_norm = _norm(login)
                if not login_norm or login_norm not in [_norm(k) for k in usuarios.keys()]:
                    lbl_erro.setText("⚠  Selecione um usuário válido.")
                    combo_usuario.setFocus()
                    return
                # Encontra a chave ignorando acentos e case
                chave_real = next(k for k in usuarios if _norm(k) == login_norm)
                self.usuario_logado     = chave_real
                # Usa assinatura do Supabase se disponível, senão usa o nome do dict local
                self.assinatura_usuario = _assinaturas_supabase.get(chave_real.upper()) or usuarios[chave_real]
                # Preenche o campo assinatura
                if "Assinatura" in self.entradas:
                    self.entradas["Assinatura"].setText(self.assinatura_usuario)

            self.empresa = nome_empresa
            cor = ACCENT if nome_empresa == "Agrovia" else DANGER
            self.btn1.setStyleSheet(f"background-color: {cor}; color: white; border: none;")
            self._atualizar_fundo(nome_empresa)
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
        dados["_usuario"] = self.usuario_logado or ""

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

    def _gravar_banco(self):
        """Grava os dados do formulário no Supabase sem gerar documento."""
        from gerador import gravar_supabase
        dados = self.coletar()
        erros = []
        if not dados.get("Motorista"): erros.append("Motorista")
        if not dados.get("Cavalo"):    erros.append("Placa")
        if erros:
            QMessageBox.warning(self, "Campos obrigatorios",
                "Preencha: " + " | ".join(erros))
            return
        dados["_usuario"] = self.usuario_logado or ""

        # Feedback imediato — bloquear botão
        self.btn2.setEnabled(False)
        self.btn2.setText("Gravando...")
        QApplication.processEvents()

        try:
            novo_id = gravar_supabase(dados, usuario=self.usuario_logado or "")
            dados["_supabase_id_resultado"] = novo_id
            salvar_historico(dados, "", usuario=self.usuario_logado)

            # Se era reedição, marcar anterior como ALTERADO
            id_anterior = getattr(self, "_reeditando_id_anterior", None)
            if id_anterior:
                try:
                    import urllib.request as _ureq
                    obs  = f"Alterado para #{novo_id}" if novo_id else "Alterado"
                    body = json.dumps({"status": "ALTERADO", "ativo": False, "observacao": obs}).encode()
                    req  = _ureq.Request(
                        f"{SUPABASE_URL}/rest/v1/carregamentos?id=eq.{id_anterior}",
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
                except Exception as e2:
                    try:
                        with open("supabase_log.txt", "a", encoding="utf-8") as f2:
                            import datetime as _dt2
                            f2.write(f"[{_dt2.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] Erro PATCH ALTERADO id={id_anterior}: {e2}\n")
                    except Exception:
                        pass
                del self._reeditando_id_anterior

            if hasattr(self, "_historico_widget") and self._historico_widget:
                self._historico_widget.recarregar()

            self._sair_modo_edicao()
            self.btn2.setEnabled(True)
            self.btn2.setText("GRAVAR NO BANCO")
            QMessageBox.information(self, "Gravado",
                f"Carregamento #{novo_id} gravado no banco de dados.")
        except Exception as e:
            self.btn2.setEnabled(True)
            self.btn2.setText("GRAVAR NO BANCO")
            QMessageBox.critical(self, "Erro", str(e))

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
        from gerador import REGRAS_EMAIL

        # Monta lista plana de emails únicos a partir de REGRAS_EMAIL
        emails_rapidos = []
        vistos = set()
        for valor in REGRAS_EMAIL.values():
            for email in valor.split(";"):
                email = email.strip()
                if email and email not in vistos and "@" in email:
                    emails_rapidos.append(email)
                    vistos.add(email)

        dlg = QDialog(self)
        dlg.setWindowTitle("Prévia do email")
        dlg.setMinimumWidth(540)
        dlg.setStyleSheet(DIALOG_SS)

        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(8)

        # ── Botões de email rápido ──────────────────────
        if emails_rapidos:
            lbl_rapido = QLabel("E-MAILS RÁPIDOS")
            lbl_rapido.setStyleSheet(f"color: {MUTED}; font-size: 9px; font-weight: 700; letter-spacing: 0.8px; background: transparent;")
            lay.addWidget(lbl_rapido)

            btn_wrap = QWidget()
            btn_wrap.setStyleSheet("background: transparent;")
            btn_flow = QHBoxLayout(btn_wrap)
            btn_flow.setContentsMargins(0, 0, 0, 4)
            btn_flow.setSpacing(6)
            btn_flow.setAlignment(Qt.AlignLeft)

            email_btns = {}

            def _make_btn_email(email):
                btn = QPushButton(email)
                btn.setCheckable(True)
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background: transparent;
                        border: 1px solid {BORDER2};
                        border-radius: 5px;
                        color: {MUTED};
                        font-size: 10px;
                        padding: 4px 8px;
                    }}
                    QPushButton:checked {{
                        background: {ACCENT}22;
                        border-color: {ACCENT};
                        color: {ACCENT};
                        font-weight: 700;
                    }}
                    QPushButton:hover {{ border-color: {ACCENT}; color: {TEXT}; }}
                """)
                return btn

            def _atualizar_botoes():
                atual = set(e.strip() for e in inp_d.text().split(";") if e.strip())
                for em, b in email_btns.items():
                    b.blockSignals(True)
                    b.setChecked(em in atual)
                    b.blockSignals(False)

            def _toggle_email(email):
                atual = [e.strip() for e in inp_d.text().split(";") if e.strip()]
                if email in atual:
                    atual.remove(email)
                else:
                    atual.append(email)
                inp_d.setText(";".join(atual))

            for email in emails_rapidos:
                b = _make_btn_email(email)
                email_btns[email] = b
                btn_flow.addWidget(b)
                b.clicked.connect(lambda checked, em=email: _toggle_email(em))

            btn_flow.addStretch()
            lay.addWidget(btn_wrap)

        # ── Campos ──────────────────────────────────────
        lay.addWidget(QLabel("DESTINATÁRIO"))
        inp_d = QLineEdit(destinatario)
        inp_d.setMinimumHeight(32)
        lay.addWidget(inp_d)

        # Atualiza botões ao editar o campo manualmente
        if emails_rapidos:
            inp_d.textChanged.connect(lambda _: _atualizar_botoes())
            # Estado inicial dos botões
            _atualizar_botoes()

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

        # Log de debug — gravar estado da reedição
        try:
            with open("supabase_log.txt", "a", encoding="utf-8") as _f:
                import datetime as _dtnow
                _id_ant = getattr(self, "_reeditando_id_anterior", "NÃO DEFINIDO")
                _f.write(f"[{_dtnow.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] _on_sucesso: id_anterior={_id_ant}\n")
        except Exception:
            pass

        # Deletar arquivos antigos se for reedição
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
            usuario = self.usuario_logado,
        )

        # Se era reedição, marcar ordem anterior como ALTERADO
        id_anterior = getattr(self, "_reeditando_id_anterior", None)
        if id_anterior:
            try:
                import urllib.request as _ureq
                obs  = f"Alterado para #{supabase_id}" if supabase_id else "Alterado"
                body = json.dumps({"status": "ALTERADO", "ativo": False, "observacao": obs}).encode()
                req  = _ureq.Request(
                    f"{SUPABASE_URL}/rest/v1/carregamentos?id=eq.{id_anterior}",
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
            except Exception as e:
                # Log do erro
                try:
                    with open("supabase_log.txt", "a", encoding="utf-8") as f:
                        import datetime as _dt2
                        f.write(f"[{_dt2.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] Erro PATCH ALTERADO id={id_anterior}: {e}\n")
                except Exception:
                    pass
            if hasattr(self, "_reeditando_id_anterior"):
                del self._reeditando_id_anterior

        self._sair_modo_edicao()

        # Recarrega histórico se estiver visível
        if hasattr(self, "_historico_widget") and self._historico_widget:
            self._historico_widget.recarregar()

        msg = QMessageBox(self)
        msg.setWindowTitle("Sucesso")
        msg.setText(f"Ordem #{supabase_id} gerada com sucesso." if supabase_id else "Ordem gerada com sucesso.")
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

    def _dialog_gravar_planilha(self, dados):
        conta = self._planilha_widget._combo_conta.currentText()
        if not conta or conta == "(nenhuma conta)":
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
        combo_st.addItems(["DESCARGA", "CARREGADO", "MARCADO", "CHEGA", "AGUARDANDO"])
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
            em_edicao = getattr(self, "_reeditando_id_anterior", None)

            if em_edicao:
                # Modo edição: só atualiza campos, preserva empresa e modo
                self._preencher_campos(dados)
                dlg.accept()
                return

            # Modo normal: pergunta se limpa campos existentes
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
                self._aplicar_empresa(empresa)
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
                          "Rota", "Agenciamento", "Colocador", "Pagamento",
                          "Peso Total", "CPF", "Contato", "Carroceria",
                          "Carreta 1", "Carreta 2", "Carreta 3",
                          "Peso", "Peso 2", "Peso 3", "Peso 4"]

        # Limpa todos os campos antes de preencher — evita dados da ordem anterior persistirem
        for campo in campos_simples:
            w = self.entradas.get(campo)
            if isinstance(w, QLineEdit):
                w.clear()
            elif isinstance(w, QComboBox):
                w.setCurrentIndex(-1)

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
            self._aplicar_empresa(emp)

    def nova_ordem(self):
        # Se em modo edição, cancelar edição sem confirmar
        if getattr(self, "_reeditando_id_anterior", None):
            self._sair_modo_edicao()
            for v in self.entradas.values():
                if isinstance(v, QLineEdit): v.clear()
                elif isinstance(v, QComboBox): v.setCurrentIndex(-1)
                elif isinstance(v, QDateEdit): v.setDate(QDate.currentDate())
            for i, (stack, linha) in enumerate(self._pedido_linhas):
                for inp in linha.values():
                    if isinstance(inp, QLineEdit): inp.clear()
                    elif isinstance(inp, QComboBox): inp.setCurrentIndex(-1)
                stack.setCurrentIndex(1 if i == 0 else 0)
            self.setar_data_hoje()
            return

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