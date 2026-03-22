import os
import time


# =========================
# GMAIL — CREDENCIAIS
# =========================
CREDENTIALS_FILE = "credentials.json"
TOKENS_DIR = "gmail_tokens"


def _listar_contas_gmail():
    """Retorna lista de contas que já têm token salvo."""
    if not os.path.exists(TOKENS_DIR):
        os.makedirs(TOKENS_DIR)
    contas = []
    for f in os.listdir(TOKENS_DIR):
        if f.startswith("token_") and f.endswith(".json"):
            conta = f[len("token_"):-len(".json")]
            contas.append(conta)
    return contas


def _autenticar_gmail(conta):
    """
    Autentica uma conta Gmail via OAuth2.
    Se não houver token salvo, abre o navegador para autorização.
    Retorna um serviço autenticado da Gmail API.
    """
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    SCOPES = [
        "https://www.googleapis.com/auth/gmail.send",
        "https://www.googleapis.com/auth/userinfo.email",
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
    """Monta o objeto MIMEMultipart com corpo e anexo PDF."""
    import base64
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
    """Autentica e envia o email com PDF anexado via Gmail API."""
    service = _autenticar_gmail(conta)
    mensagem = _montar_mensagem_gmail(conta, destinatario, assunto, corpo, pdf_path)
    service.users().messages().send(userId="me", body=mensagem).execute()


def adicionar_conta_gmail():
    """
    Força o fluxo de autorização para uma nova conta Gmail.
    Retorna o email da conta autorizada.
    """
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    import json

    SCOPES = [
        "https://www.googleapis.com/auth/gmail.send",
        "https://www.googleapis.com/auth/userinfo.email",
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

    # Obtém o email via id_token ou userinfo
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
        # Fallback: userinfo endpoint
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


# =========================
# FORMATAR TELEFONE
# =========================
def formatar_telefone(telefone):
    if not telefone:
        return ""

    numeros = "".join(filter(str.isdigit, telefone))

    if len(numeros) == 11:
        return f"({numeros[:2]}) {numeros[2:7]}-{numeros[7:]}"
    elif len(numeros) == 10:
        return f"({numeros[:2]}) {numeros[2:6]}-{numeros[6:]}"
    return telefone


# =========================
# EMAIL POR FÁBRICA
# =========================
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


# =========================
# MONTAR EMAIL
# =========================
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


# =========================
# LOCALIZAR EXCEL.EXE
# =========================
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


# =========================
# FUNÇÃO PRINCIPAL
# =========================
def gerar_ordem(dados, pasta_destino, enviar_email=True, conta_gmail=None):
    import xlwings as xw
    import shutil

    # =========================
    # VALIDAÇÃO
    # =========================
    obrigatorios = ["Motorista", "Cavalo", "Pedido"]

    for campo in obrigatorios:
        if not dados.get(campo):
            raise Exception(f"Campo obrigatório não preenchido: {campo}")

    # =========================
    # AUXILIARES
    # =========================
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

    # =========================
    # MODELO
    # =========================
    if dados["empresa"] == "Agrovia":
        arquivo_modelo = "ordempadraoagro.xlsx"
    else:
        arquivo_modelo = "ordempadraotop.xlsx"

    if not os.path.exists(arquivo_modelo):
        raise Exception("Modelo não encontrado!")

    # =========================
    # CAMPOS
    # =========================
    CAMPOS = {
        "Data Apresentação": "M5",
        "Fábrica": ["M6", "M7"],
        "Solicitante": ["M8", "M9"],
        "Motorista": "F13",
        "CPF": "N13",
        "Contato": "F15",
        "Destino": "F21",
        "Fazenda": "N21",
        "Produto": "F22",
        "Embalagem": "N22",
        "Cliente": "F23",
        "Pedido": "M23",
        "Peso": "F24",
        "Carroceria": "I37",
        "Cavalo": "I38",
        "Carreta 1": "I39",
        "Carreta 2": "I40",
        "Carreta 3": "I41",
        "Data Geração": "C44",
        "Assinatura": "J44"
    }

    # =========================
    # TRATAR DADOS
    # =========================
    dados["CPF"] = formatar_cpf(dados.get("CPF", ""))
    dados["Contato"] = formatar_telefone(dados.get("Contato", ""))
    dados["Data Geração"] = time.strftime("%d/%m/%Y")

    for chave in list(dados.keys()):
        if not chave.startswith("_"):
            dados[chave] = normalizar(dados[chave])

    motorista_completo = dados.get("Motorista", "SEM")
    primeiro_nome = limpar_nome(motorista_completo.split()[0] if motorista_completo else "SEM")

    cavalo_raw = dados.get("Cavalo", "SEM")
    # Formata placa no modelo XXX-XXXX para o arquivo/pdf/email
    import re as _re
    cavalo_limpo = _re.sub(r"[^A-Za-z0-9]", "", cavalo_raw).upper()
    if len(cavalo_limpo) > 3:
        cavalo_formatado = cavalo_limpo[:3] + "-" + cavalo_limpo[3:7]
    else:
        cavalo_formatado = cavalo_limpo
    dados["Cavalo"] = cavalo_formatado

    cavalo = limpar_nome(cavalo_formatado)

    nome_arquivo = f"ORDEM_{primeiro_nome}_{cavalo}.xlsx"
    caminho = os.path.join(pasta_destino, nome_arquivo)

    contador = 1
    base = caminho
    while os.path.exists(caminho):
        caminho = base.replace(".xlsx", f"_{contador}.xlsx")
        contador += 1

    shutil.copy(arquivo_modelo, caminho)

    # =========================
    # EXCEL OTIMIZADO
    # =========================
    app = None
    wb = None

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

        for campo, celula in CAMPOS.items():
            valor = dados.get(campo, "")

            if isinstance(celula, list):
                for c in celula:
                    ws.range(c).value = valor
            else:
                ws.range(celula).value = valor

        wb.save()
        app.api.CalculateFull()

        pdf_path = caminho.replace(".xlsx", ".pdf")

        # =========================
        # EXPORTAR PDF COM FALLBACK
        # =========================
        try:
            wb.api.ExportAsFixedFormat(0, pdf_path)
        except Exception:
            pass

        if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) == 0:
            try:
                wb.close()
                wb = None
                app.quit()
                app = None

                import subprocess
                excel_exe = _encontrar_excel()
                if excel_exe:
                    subprocess.run([excel_exe, "/x", caminho], timeout=30)
                    time.sleep(3)

                app2 = xw.App(visible=False)
                app2.api.Visible = False
                app2.api.ScreenUpdating = False
                app2.api.DisplayAlerts = False
                wb2 = xw.Book(caminho)
                try:
                    wb2.api.ExportAsFixedFormat(0, pdf_path)
                finally:
                    wb2.close()
                    app2.quit()

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

        if wb:
            try:
                wb.close()
                wb = None
            except Exception:
                pass
        if app:
            try:
                app.quit()
                app = None
            except Exception:
                pass

    finally:
        try:
            if wb:
                wb.close()
        except:
            pass
        try:
            if app:
                app.quit()
        except:
            pass

    # =========================
    # EMAIL
    # =========================
    if enviar_email:
        if not conta_gmail:
            raise Exception("Nenhuma conta Gmail selecionada para envio.")
        # Usa valores editados no preview se disponíveis
        destinatario = dados.get("_email_destinatario") or obter_email_fabrica(dados.get("Fábrica"))
        assunto      = dados.get("_email_assunto")      or montar_email(dados)[0]
        corpo        = dados.get("_email_corpo")        or montar_email(dados)[1]
        enviar_email_gmail(conta_gmail, destinatario, assunto, corpo, pdf_path)

    return caminho