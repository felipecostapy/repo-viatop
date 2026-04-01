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

def gerar_ordem(dados, pasta_destino, enviar_email=True, conta_gmail=None):
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
        "Destino":       "O10",
        "Pagador":       "O11",
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

    for chave in list(dados.keys()):
        if not chave.startswith("_"):
            dados[chave] = normalizar(dados[chave])

    motorista_completo = dados.get("Motorista", "SEM")
    primeiro_nome = limpar_nome(motorista_completo.split()[0] if motorista_completo else "SEM")

    def _fmt_placa(valor):
        """Formata placa para XXX-XXXX. Aplica apenas na gravação."""
        if not valor or str(valor).strip() in ("", "SEM"):
            return valor
        limpo = re.sub(r"[^A-Za-z0-9]", "", str(valor)).upper()
        return limpo[:3] + "-" + limpo[3:7] if len(limpo) > 3 else limpo

    cavalo_raw = dados.get("Cavalo", "SEM")
    cavalo_formatado = _fmt_placa(cavalo_raw)
    dados["Cavalo"]    = cavalo_formatado
    dados["Carreta 1"] = _fmt_placa(dados.get("Carreta 1", ""))
    dados["Carreta 2"] = _fmt_placa(dados.get("Carreta 2", ""))
    dados["Carreta 3"] = _fmt_placa(dados.get("Carreta 3", ""))

    cavalo = limpar_nome(cavalo_formatado)

    nome_arquivo = f"ORDEM_{primeiro_nome}_{cavalo}.xlsx"
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


    if enviar_email:
        if not conta_gmail:
            raise Exception("Nenhuma conta Gmail selecionada para envio.")
                                                        
        destinatario = dados.get("_email_destinatario") or obter_email_fabrica(dados.get("Fábrica"))
        assunto      = dados.get("_email_assunto")      or montar_email(dados)[0]
        corpo        = dados.get("_email_corpo")        or montar_email(dados)[1]
        enviar_email_gmail(conta_gmail, destinatario, assunto, corpo, pdf_path)

    return caminho