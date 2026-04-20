import os
import re
import time
import json
import shutil
import base64
# win32com importado localmente em gerar_ordem para evitar falha na inicialização

CREDENTIALS_FILE = "credentials.json"
TOKENS_DIR = "gmail_tokens"

def _listar_contas_gmail():
                                                         
    if not os.path.exists(TOKENS_DIR):
        os.makedirs(TOKENS_DIR)
    contas = []
    for f in os.listdir(TOKENS_DIR):
        if f.startswith("token_") and f.endswith(".json"):
            conta = f[len("token_"):-len(".json")]
            contas.append(conta)
    return contas

def _autenticar_gmail(conta):
\
\
\
\
       
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    SCOPES = [
        "https://www.googleapis.com/auth/gmail.send",
        "https://www.googleapis.com/auth/userinfo.email",
        "https://www.googleapis.com/auth/spreadsheets",
        "openid",
    ]
    token_path = os.path.join(TOKENS_DIR, f"token_{conta}.json")

    creds = None

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                raise Exception(
                    f"Arquivo '{CREDENTIALS_FILE}' não encontrado.\n"
                    "Coloque o credentials.json na mesma pasta do sistema."
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_path, "w") as f:
            f.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)

def _montar_mensagem_gmail(remetente, destinatario, assunto, corpo, pdf_path):
                                                             
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders

    msg = MIMEMultipart()
    msg["From"] = remetente
    msg["To"] = destinatario
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo, "plain"))

    with open(pdf_path, "rb") as f:
        parte = MIMEBase("application", "octet-stream")
        parte.set_payload(f.read())
        encoders.encode_base64(parte)
        nome_pdf = os.path.basename(pdf_path)
        parte.add_header("Content-Disposition", f'attachment; filename="{nome_pdf}"')
        msg.attach(parte)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {"raw": raw}

def enviar_email_gmail(conta, destinatario, assunto, corpo, pdf_path):
                                                                  
    service = _autenticar_gmail(conta)
    mensagem = _montar_mensagem_gmail(conta, destinatario, assunto, corpo, pdf_path)
    service.users().messages().send(userId="me", body=mensagem).execute()

def adicionar_conta_gmail():
\
\
\
       
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    SCOPES = [
        "https://www.googleapis.com/auth/gmail.send",
        "https://www.googleapis.com/auth/userinfo.email",
        "https://www.googleapis.com/auth/spreadsheets",
        "openid",
    ]

    if not os.path.exists(CREDENTIALS_FILE):
        raise Exception(
            f"Arquivo '{CREDENTIALS_FILE}' não encontrado.\n"
            "Coloque o credentials.json na mesma pasta do sistema."
        )

    if not os.path.exists(TOKENS_DIR):
        os.makedirs(TOKENS_DIR)

    flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
    creds = flow.run_local_server(port=0)

    email = None
    try:
        import google.auth.transport.requests
        import google.oauth2.id_token
        request = google.auth.transport.requests.Request()
        id_info = google.oauth2.id_token.verify_oauth2_token(
            creds.id_token, request, clock_skew_in_seconds=10
        )
        email = id_info.get("email")
    except Exception:
        pass

    if not email:
                                     
        import urllib.request
        req = urllib.request.Request(
            "https://www.googleapis.com/oauth2/v1/userinfo?alt=json",
            headers={"Authorization": f"Bearer {creds.token}"}
        )
        with urllib.request.urlopen(req) as resp:
            info = json.loads(resp.read())
            email = info.get("email")

    if not email:
        raise Exception("Não foi possível obter o email da conta autorizada.")

    token_path = os.path.join(TOKENS_DIR, f"token_{email}.json")
    with open(token_path, "w") as f:
        f.write(creds.to_json())

    return email

# ── Supabase ───────────────────────────────────────────────────────
SUPABASE_URL = "https://xlirwzkmvkzldrssmhxg.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhsaXJ3emttdmt6bGRyc3NtaHhnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzU2MDgwMTksImV4cCI6MjA5MTE4NDAxOX0.ofTAEn628a-7JzF3REPj-tBcQJUrlXdfaFSbU5Ysfx4"

def gravar_supabase(dados, usuario=None):
    """Grava carregamento no Supabase via API REST. Silencioso em caso de erro."""
    import urllib.request
    import urllib.error
    import datetime as _dt

    def _data(v):
        if not v: return None
        for fmt in ["%d/%m/%Y", "%Y-%m-%d"]:
            try: return _dt.datetime.strptime(str(v).strip(), fmt).strftime("%Y-%m-%d")
            except: pass
        return None

    def _float(v):
        try:
            s = str(v or "0").strip()
            # Formato BR: 1.234,56 → tem ponto E vírgula → remove ponto, troca vírgula
            if "," in s and "." in s:
                s = s.replace(".", "").replace(",", ".")
            # Só vírgula: 49,08 → troca por ponto
            elif "," in s:
                s = s.replace(",", ".")
            # Só ponto: 49.08 → já é decimal, não mexe
            f = float(s)
            return int(f) if f == int(f) else f
        except: return 0

    registro = {
        "data":         _data(dados.get("Data Apresentação")),
        "filial":       str(dados.get("empresa", "")).upper(),
        "pagador":      str(dados.get("Pagador", "")).upper(),
        "cliente":      str(dados.get("Cliente", "")).upper(),
        "motorista":    str(dados.get("Motorista", "")).upper(),
        "agencia":      str(dados.get("Agência", "") or dados.get("Agencia", "")).upper(),
        "placa":        str(dados.get("Cavalo", "")).upper(),
        "fabrica":      str(dados.get("Fábrica", "")).upper(),
        "destino":      str(dados.get("Destino", "")).upper(),
        "uf":           str(dados.get("UF", "")).upper(),
        "peso":         _float(dados.get("Peso Total") or dados.get("Peso")),
        "frete_emp":    _float(dados.get("Frete/Emp")),
        "frete_mot":    _float(dados.get("Frete/Mot")),
        "rota":         str(dados.get("Rota", "")).upper(),
        "agenciamento": str(dados.get("Agenciamento", "")).upper(),
        "status":       "AGUARDANDO",
        "pedido":       str(dados.get("Pedido", "")).upper(),
        "produto":      str(dados.get("Produto", "")).upper(),
        "embalagem":    str(dados.get("Embalagem", "")).upper(),
        "peso1":        _float(dados.get("Peso")),
        "peso2":        _float(dados.get("Peso 2")),
        "peso3":        _float(dados.get("Peso 3")),
        "peso4":        _float(dados.get("Peso 4")),
        "pedido2":      str(dados.get("Pedido 2", "")).upper(),
        "produto2":     str(dados.get("Produto 2", "")).upper(),
        "embalagem2":   str(dados.get("Embalagem 2", "")).upper(),
        "pedido3":      str(dados.get("Pedido 3", "")).upper(),
        "produto3":     str(dados.get("Produto 3", "")).upper(),
        "embalagem3":   str(dados.get("Embalagem 3", "")).upper(),
        "pedido4":      str(dados.get("Pedido 4", "")).upper(),
        "produto4":     str(dados.get("Produto 4", "")).upper(),
        "embalagem4":   str(dados.get("Embalagem 4", "")).upper(),
        "colocador":    str(dados.get("Colocador", "")).upper(),
        "pagamento":    str(dados.get("Pagamento", "")).upper(),
        "usuario":      str(usuario or "").upper(),
        "origem":       str(dados.get("Origem", "")).upper(),
        "cpf":          str(dados.get("CPF", "")).upper(),
        "contato":      str(dados.get("Contato", "")).upper(),
        "carroceria":   str(dados.get("Carroceria", "")).upper(),
        "carreta1":     str(dados.get("Carreta 1", "")).upper(),
        "carreta2":     str(dados.get("Carreta 2", "")).upper(),
        "carreta3":     str(dados.get("Carreta 3", "")).upper(),
        "fazenda":      str(dados.get("Fazenda", "")).upper(),
        "solicitante":  str(dados.get("Solicitante", "")).upper(),
    }

    # Remove campos vazios
    registro = {k: v for k, v in registro.items() if v not in (None, "", 0, 0.0)}

    try:
        body = json.dumps(registro).encode("utf-8")
        req  = urllib.request.Request(
            f"{SUPABASE_URL}/rest/v1/carregamentos",
            data=body,
            headers={
                "apikey":        SUPABASE_KEY,
                "Authorization": f"Bearer {SUPABASE_KEY}",
                "Content-Type":  "application/json",
                "Prefer":        "return=representation",
            },
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            resultado = json.loads(resp.read().decode("utf-8"))
            if resultado and isinstance(resultado, list):
                return resultado[0].get("id")
    except Exception as e:
        try:
            log = os.path.join(os.path.dirname(os.path.abspath(__file__)), "supabase_log.txt")
            with open(log, "a", encoding="utf-8") as f:
                import datetime as _dt2
                f.write(f"[{_dt2.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] Erro INSERT: {e}\n")
        except:
            pass
    return None

def atualizar_supabase(supabase_id, dados, usuario=None):
    """Atualiza registro existente no Supabase via PATCH. Silencioso em caso de erro."""
    import urllib.request
    import urllib.error
    import datetime as _dt

    if not supabase_id:
        return

    def _data(v):
        if not v: return None
        for fmt in ["%d/%m/%Y", "%Y-%m-%d"]:
            try: return _dt.datetime.strptime(str(v).strip(), fmt).strftime("%Y-%m-%d")
            except: pass
        return None

    def _float(v):
        try:
            s = str(v or "0").strip()
            # Formato BR: 1.234,56 → tem ponto E vírgula → remove ponto, troca vírgula
            if "," in s and "." in s:
                s = s.replace(".", "").replace(",", ".")
            # Só vírgula: 49,08 → troca por ponto
            elif "," in s:
                s = s.replace(",", ".")
            # Só ponto: 49.08 → já é decimal, não mexe
            f = float(s)
            return int(f) if f == int(f) else f
        except: return 0

    registro = {
        "data":         _data(dados.get("Data Apresentação")),
        "filial":       str(dados.get("empresa", "")).upper(),
        "pagador":      str(dados.get("Pagador", "")).upper(),
        "cliente":      str(dados.get("Cliente", "")).upper(),
        "motorista":    str(dados.get("Motorista", "")).upper(),
        "agencia":      str(dados.get("Agência", "") or dados.get("Agencia", "")).upper(),
        "placa":        str(dados.get("Cavalo", "")).upper(),
        "fabrica":      str(dados.get("Fábrica", "")).upper(),
        "destino":      str(dados.get("Destino", "")).upper(),
        "uf":           str(dados.get("UF", "")).upper(),
        "peso":         _float(dados.get("Peso Total") or dados.get("Peso")),
        "frete_emp":    _float(dados.get("Frete/Emp")),
        "frete_mot":    _float(dados.get("Frete/Mot")),
        "rota":         str(dados.get("Rota", "")).upper(),
        "agenciamento": str(dados.get("Agenciamento", "")).upper(),
        "pedido":       str(dados.get("Pedido", "")).upper(),
        "produto":      str(dados.get("Produto", "")).upper(),
        "embalagem":    str(dados.get("Embalagem", "")).upper(),
        "peso1":        _float(dados.get("Peso")),
        "peso2":        _float(dados.get("Peso 2")),
        "peso3":        _float(dados.get("Peso 3")),
        "peso4":        _float(dados.get("Peso 4")),
        "pedido2":      str(dados.get("Pedido 2", "")).upper(),
        "produto2":     str(dados.get("Produto 2", "")).upper(),
        "embalagem2":   str(dados.get("Embalagem 2", "")).upper(),
        "pedido3":      str(dados.get("Pedido 3", "")).upper(),
        "produto3":     str(dados.get("Produto 3", "")).upper(),
        "embalagem3":   str(dados.get("Embalagem 3", "")).upper(),
        "pedido4":      str(dados.get("Pedido 4", "")).upper(),
        "produto4":     str(dados.get("Produto 4", "")).upper(),
        "embalagem4":   str(dados.get("Embalagem 4", "")).upper(),
        "colocador":    str(dados.get("Colocador", "")).upper(),
        "pagamento":    str(dados.get("Pagamento", "")).upper(),
        "usuario":      str(usuario or "").upper(),
        "origem":       str(dados.get("Origem", "")).upper(),
        "cpf":          str(dados.get("CPF", "")).upper(),
        "contato":      str(dados.get("Contato", "")).upper(),
        "carroceria":   str(dados.get("Carroceria", "")).upper(),
        "carreta1":     str(dados.get("Carreta 1", "")).upper(),
        "carreta2":     str(dados.get("Carreta 2", "")).upper(),
        "carreta3":     str(dados.get("Carreta 3", "")).upper(),
        "fazenda":      str(dados.get("Fazenda", "")).upper(),
        "solicitante":  str(dados.get("Solicitante", "")).upper(),
    }
    registro = {k: v for k, v in registro.items() if v not in (None, "", 0, 0.0)}

    try:
        body = json.dumps(registro).encode("utf-8")
        req  = urllib.request.Request(
            f"{SUPABASE_URL}/rest/v1/carregamentos?id=eq.{supabase_id}",
            data=body,
            headers={
                "apikey":        SUPABASE_KEY,
                "Authorization": f"Bearer {SUPABASE_KEY}",
                "Content-Type":  "application/json",
                "Prefer":        "return=minimal",
            },
            method="PATCH"
        )
        urllib.request.urlopen(req, timeout=8)
    except Exception as e:
        try:
            log = os.path.join(os.path.dirname(os.path.abspath(__file__)), "supabase_log.txt")
            with open(log, "a", encoding="utf-8") as f:
                import datetime as _dt2
                f.write(f"[{_dt2.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] Erro PATCH id={supabase_id}: {e}\n")
        except:
            pass

def formatar_telefone(telefone):
    if not telefone:
        return ""

    numeros = "".join(filter(str.isdigit, telefone))

    if len(numeros) == 11:
        return f"({numeros[:2]}) {numeros[2:7]}-{numeros[7:]}"
    elif len(numeros) == 10:
        return f"({numeros[:2]}) {numeros[2:6]}-{numeros[6:]}"
    return telefone

REGRAS_EMAIL = {
    "FERTIMAXI": "agendamento@fertimaxi.com;paulo@fertimaxi.com",
    "TIMAC": "EMAIL_AQUI_3",
    "INTERMARITIMA": "EMAIL_AQUI_4"
}

def obter_email_fabrica(fabrica):
    fabrica = (fabrica or "").upper()

    for chave in REGRAS_EMAIL:
        if chave in fabrica:
            return REGRAS_EMAIL[chave]

    return "SEU_EMAIL@gmail.com"

def montar_email(dados):
    cavalo = dados.get("Cavalo", "")
    pedido = dados.get("Pedido", "")
    produto = dados.get("Produto", "")
    data = dados.get("Data Apresentação", "")

    assunto = f"SEGUE AGENDAMENTO REFERENTE A PLACA {cavalo} PDD {pedido} PARA CARREGAMENTO NO DIA {data}"

    corpo = f"""PLACA {cavalo}
PDD {pedido}
PRODUTO {produto}
DIA {data}"""

    return assunto, corpo

def _encontrar_excel():
    caminhos = [
        r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
        r"C:\Program Files\Microsoft Office\Office16\EXCEL.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE",
        r"C:\Program Files\Microsoft Office\Office15\EXCEL.EXE",
        r"C:\Program Files\Microsoft Office\Office14\EXCEL.EXE",
    ]
    for c in caminhos:
        if os.path.exists(c):
            return c
    return None

def gerar_ordem(dados, pasta_destino, enviar_email=True, conta_gmail=None, imprimir=False):
    import xlwings as xw

    obrigatorios = ["Motorista", "Cavalo", "Pedido"]

    for campo in obrigatorios:
        if not dados.get(campo):
            raise Exception(f"Campo obrigatório não preenchido: {campo}")

    def formatar_cpf(cpf):
        cpf = cpf.replace(".", "").replace("-", "").strip()
        if len(cpf) == 11 and cpf.isdigit():
            return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
        return cpf

    def limpar_nome(texto):
        proibidos = r'\/:*?"<>|'
        for c in proibidos:
            texto = texto.replace(c, "")
        return texto.strip().replace(" ", "_").upper()

    def normalizar(valor):
        if isinstance(valor, str):
            return " ".join(valor.strip().upper().split())
        return valor

    if dados["empresa"] == "Agrovia":
        arquivo_modelo = "ordemagroviav2.xlsx"
    else:
        arquivo_modelo = "ordemtopv2.xlsx"

    if not os.path.exists(arquivo_modelo):
        raise Exception("Modelo não encontrado!")

    CAMPOS = {
        "Origem":        "J10",
        "Data Apresentação": "J11",
        # "Destino": "O10",  # montado manualmente abaixo (Destino + UF)
        # "Pagador": "O11",  # removido — Pagador não vai para o PDF por enquanto
        "Fábrica":       "D15",
        "Cliente":       "I15",
        "Fazenda":       "O15",
        "Motorista":     "D27",
        "CPF":           "D28",
        "Contato":       "M28",
        "Carroceria":    "I32",
        "Cavalo":        "I33",
        "Carreta 1":     "I34",
        "Carreta 2":     "I35",
        "Carreta 3":     "I36",
        "Data Geração":  "F39",
        "Assinatura":    "K39",
        "Numero Ordem":  "D39",
    }

    LINHAS_PEDIDO = [
        ("Pedido",   "B20", "Produto",   "E20", "Peso",   "N20", "Embalagem",   "O20"),
        ("Pedido 2", "B21", "Produto 2", "E21", "Peso 2", "N21", "Embalagem 2", "O21"),
        ("Pedido 3", "B22", "Produto 3", "E22", "Peso 3", "N22", "Embalagem 3", "O22"),
        ("Pedido 4", "B23", "Produto 4", "E23", "Peso 4", "N23", "Embalagem 4", "O23"),
    ]

    dados["CPF"] = formatar_cpf(dados.get("CPF", ""))
    dados["Contato"] = formatar_telefone(dados.get("Contato", ""))
    dados["Data Geração"] = time.strftime("%d/%m/%Y")

    # Grava no Supabase ANTES de gerar o PDF para ter o número da ordem
    usuario     = dados.get("_usuario", "")
    supabase_id = dados.get("_supabase_id")
    if supabase_id:
        atualizar_supabase(supabase_id, dados, usuario=usuario)
        _ordem_id = supabase_id
    else:
        _ordem_id = gravar_supabase(dados, usuario=usuario)
    dados["_supabase_id_resultado"] = _ordem_id

    # Define número da ordem para imprimir no PDF
    if _ordem_id:
        dados["Numero Ordem"] = str(_ordem_id) if _ordem_id else ""

    for chave in list(dados.keys()):
        if not chave.startswith("_"):
            dados[chave] = normalizar(dados[chave])

    motorista_completo = dados.get("Motorista", "SEM")
    primeiro_nome = limpar_nome(motorista_completo.split()[0] if motorista_completo else "SEM")

    def _fmt_placa(valor):
        """Formata placa para XXX-XXXX ou XXX-XXXXXUF. Preserva sufixo de UF."""
        if not valor or str(valor).strip() in ("", "SEM"):
            return valor
        limpo = re.sub(r"[^A-Za-z0-9]", "", str(valor)).upper()
        return limpo[:3] + "-" + limpo[3:] if len(limpo) > 3 else limpo

    cavalo_raw = dados.get("Cavalo", "SEM")
    cavalo_formatado = _fmt_placa(cavalo_raw)
    dados["Cavalo"]    = cavalo_formatado
    dados["Carreta 1"] = _fmt_placa(dados.get("Carreta 1", ""))
    dados["Carreta 2"] = _fmt_placa(dados.get("Carreta 2", ""))
    dados["Carreta 3"] = _fmt_placa(dados.get("Carreta 3", ""))

    cavalo = limpar_nome(cavalo_formatado)

    _num = str(_ordem_id) if _ordem_id else "0"
    nome_arquivo = f"OR{_num}_{primeiro_nome}_{cavalo}.xlsx"
    caminho = os.path.join(pasta_destino, nome_arquivo)

    contador = 1
    base = caminho
    while os.path.exists(caminho):
        caminho = base.replace(".xlsx", f"_{contador}.xlsx")
        contador += 1

    shutil.copy(arquivo_modelo, caminho)

    app = None
    wb = None
    xl_fallback = None  # instância win32com do fallback
    pdf_path = caminho.replace(".xlsx", ".pdf")

    def _fechar_tudo():
        """Garante fechamento de todas as instâncias Excel abertas."""
        nonlocal wb, app, xl_fallback
        for obj, metodo in [(wb, "close"), (app, "quit")]:
            if obj is not None:
                try:
                    getattr(obj, metodo)()
                except Exception:
                    pass
        wb = app = None
        if xl_fallback is not None:
            try:
                xl_fallback.Quit()
            except Exception:
                pass
            xl_fallback = None

    try:
        app = xw.App(visible=False)
        app.visible = False
        app.api.Visible = False
        app.api.ScreenUpdating = False
        app.api.DisplayAlerts = False
        app.display_alerts = False
        app.screen_updating = False

        wb = xw.Book(caminho)
        ws = wb.sheets[0]

        CAMPOS_DATA = {"Data Apresentação", "Data Geração"}

        for campo, celula in CAMPOS.items():
            valor = dados.get(campo, "")
            # Datas: converte DD/MM/AAAA para objeto date do Python
            # para que o Excel grave corretamente sem depender da localização
            if campo in CAMPOS_DATA and valor:
                try:
                    import datetime as _dt
                    partes = str(valor).strip().split("/")
                    if len(partes) == 3:
                        valor = _dt.date(int(partes[2]), int(partes[1]), int(partes[0]))
                except Exception:
                    pass  # mantém como string se falhar
            if isinstance(celula, list):
                for c in celula:
                    ws.range(c).value = valor
            else:
                ws.range(celula).value = valor

        for ped_key, ped_cel, prod_key, prod_cel, peso_key, peso_cel, emb_key, emb_cel in LINHAS_PEDIDO:
            ped  = dados.get(ped_key, "")
            prod = dados.get(prod_key, "")
            peso = dados.get(peso_key, "")
            emb  = dados.get(emb_key, "")

            if not ped and not prod and not peso and not emb:
                ws.range(ped_cel).value  = "x"
                ws.range(prod_cel).value = "x"
                ws.range(peso_cel).value = "x"
                ws.range(emb_cel).value  = "x"
            else:
                ws.range(ped_cel).value  = ped
                ws.range(prod_cel).value = prod
                ws.range(peso_cel).value = peso
                ws.range(emb_cel).value  = emb

        # Solicitante em O11 apenas para Intermarítima e Armazém Vitória
        import unicodedata as _ud
        def _norm_fab(s):
            return _ud.normalize("NFD", str(s).upper()).encode("ascii","ignore").decode()
        fabrica_norm = _norm_fab(dados.get("Fábrica","") or "")
        FABRICAS_SOL = ["INTERMARITIMA", "ARMAZEM VITORIA"]
        if any(f in fabrica_norm for f in FABRICAS_SOL):
            sol = str(dados.get("Solicitante","") or "").upper()
            if sol:
                ws.range("O11").value = sol

        # Destino no Excel = "CIDADE - UF"
        destino_val = dados.get("Destino", "")
        uf_val      = dados.get("UF", "")
        if destino_val and uf_val:
            ws.range("O10").value = f"{destino_val} - {uf_val}"
        elif destino_val:
            ws.range("O10").value = destino_val
        else:
            ws.range("O10").value = ""

        wb.save()
        app.api.CalculateFull()

        # Tentativa 1: exportar via xlwings
        try:
            wb.api.ExportAsFixedFormat(0, pdf_path)
        except Exception:
            pass

        # Tentativa 2: fallback via win32com se PDF não gerado
        if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) == 0:
            # Fecha xlwings antes de abrir nova instância
            try:
                wb.close(); wb = None
                app.quit();  app = None
            except Exception:
                pass
            time.sleep(1)

            try:
                import win32com.client as _win32
                xl_fallback = _win32.Dispatch("Excel.Application")
                xl_fallback.Visible = False
                xl_fallback.DisplayAlerts = False
                xl_fallback.ScreenUpdating = False
                wb2 = xl_fallback.Workbooks.Open(os.path.abspath(caminho))
                try:
                    wb2.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
                finally:
                    wb2.Close(False)
            except Exception as e_fallback:
                raise Exception(
                    f"Não foi possível gerar o PDF.\n"
                    f"O arquivo Excel foi salvo em:\n{caminho}\n\n"
                    f"Verifique se o 'Microsoft Print to PDF' está habilitado\n"
                    f"no Windows (Configurações > Impressoras).\n\nDetalhe: {e_fallback}"
                )

        if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) == 0:
            raise Exception(
                f"PDF não foi gerado.\n"
                f"O Excel foi salvo em:\n{caminho}\n\n"
                f"Verifique se o 'Microsoft Print to PDF' está habilitado\n"
                f"no Windows (Configurações > Impressoras e Scanners)."
            )

    finally:
        _fechar_tudo()


    # Impressão automática na impressora padrão via PowerShell (nativo Windows)
    if imprimir and os.path.exists(pdf_path):
        try:
            import subprocess
            pdf_abs = os.path.abspath(pdf_path)
            # PowerShell usa o app padrão do Windows para imprimir PDF
            cmd = (
                f'Start-Process -FilePath "{pdf_abs}" '
                f'-Verb Print -WindowStyle Hidden'
            )
            subprocess.Popen(
                ["powershell", "-WindowStyle", "Hidden", "-Command", cmd],
                creationflags=subprocess.CREATE_NO_WINDOW
            )
        except Exception:
            pass

    if enviar_email:
        if not conta_gmail:
            raise Exception("Nenhuma conta Gmail selecionada para envio.")
                                                        
        destinatario = dados.get("_email_destinatario") or obter_email_fabrica(dados.get("Fábrica"))
        assunto      = dados.get("_email_assunto")      or montar_email(dados)[0]
        corpo        = dados.get("_email_corpo")        or montar_email(dados)[1]
        enviar_email_gmail(conta_gmail, destinatario, assunto, corpo, pdf_path)

    return caminho