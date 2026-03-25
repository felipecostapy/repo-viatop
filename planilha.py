import os
from pathlib import Path

SPREADSHEET_ID   = "1yyS5POpPyv6xanTLgJ8Xu-jy8enCjDIIXAJRNkF7Rl8"
ABA              = "teste pedidos"
TOKENS_DIR       = "gmail_tokens"
CREDENTIALS_FILE = "credentials.json"
COL_STARTS       = [0, 7, 14]
COL_WIDTH        = 6

SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/spreadsheets",
    "openid",
]


def _autenticar(conta):
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    token_path = os.path.join(TOKENS_DIR, f"token_{conta}.json")
    creds = None

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            with open(token_path, "w") as f:
                f.write(creds.to_json())

    return build("sheets", "v4", credentials=creds)


def carregar_blocos(conta):
    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    resp = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{ABA}'",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    valores = resp.get("values", [])
    if not valores:
        return []

    max_col    = max(len(r) for r in valores)
    linhas     = [r + [""] * (max_col - len(r)) for r in valores]
    num_linhas = len(linhas)

    def cell(r, c):
        if r < num_linhas and c < len(linhas[r]):
            return str(linhas[r][c]).strip()
        return ""

    def is_cabecalho(texto):
        return texto.count("/") >= 2

    blocos = []

    for col in COL_STARTS:
        row = 0
        while row < num_linhas:
            cab = cell(row, col)
            if not is_cabecalho(cab):
                row += 1
                continue

            partes  = [p.strip() for p in cab.split("/")]
            destino = partes[0] if len(partes) > 0 else ""
            cliente = partes[1] if len(partes) > 1 else ""
            fabrica = partes[2] if len(partes) > 2 else ""
            pedido  = partes[3] if len(partes) > 3 else ""
            produto = partes[4] if len(partes) > 4 else ""

            try:
                saldo_total = float(str(cell(row, col + COL_WIDTH)).replace(",", "."))
            except Exception:
                saldo_total = 0

            row += 2

            linhas_dados = []
            while row < num_linhas:
                c0 = cell(row, col).upper()
                if "TOTAL" in c0 or "SALDO" in c0:
                    row += 1
                    break

                data   = cell(row, col)
                nota   = cell(row, col + 1)
                placa  = cell(row, col + 2)
                peso   = cell(row, col + 3)
                frete  = cell(row, col + 4)
                status = cell(row, col + 5)

                if any([data, nota, placa, peso]):
                    linhas_dados.append({
                        "data": data, "nota": nota, "placa": placa,
                        "peso": peso, "frete": frete, "status": status,
                    })
                else:
                    row += 1
                    break

                row += 1

            try:
                total_carregado = sum(
                    float(str(l["peso"]).replace(",", "."))
                    for l in linhas_dados if l["peso"]
                )
                saldo_restante = saldo_total - total_carregado
            except Exception:
                total_carregado = 0
                saldo_restante  = saldo_total

            blocos.append({
                "destino":         destino,
                "cliente":         cliente,
                "fabrica":         fabrica,
                "pedido":          pedido,
                "produto":         produto,
                "saldo_total":     saldo_total,
                "total_carregado": total_carregado,
                "saldo_restante":  saldo_restante,
                "linhas":          linhas_dados,
            })

    return blocos