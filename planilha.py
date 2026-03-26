import os
from pathlib import Path

SPREADSHEET_ID   = "1gQCZjGW0nWJcRBMj9hajCQgocBN1QZxQ2Oxotmh-dmA"
ABA              = "Pagina01"
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
            cliente = partes[0] if len(partes) > 0 else ""
            cidade  = partes[1] if len(partes) > 1 else ""
            fazenda = partes[2] if len(partes) > 2 else ""
            fabrica = partes[3] if len(partes) > 3 else ""
            pedido  = partes[4] if len(partes) > 4 else ""
            produto = partes[5] if len(partes) > 5 else ""

            try:
                saldo_total = float(str(cell(row, col + 5)).replace(",", "."))
            except Exception:
                saldo_total = 0

            row_cab = row
            row += 2

            linhas_dados = []
            linha_total  = None
            while row < num_linhas:
                # TOTAL aparece na coluna C ou D do bloco (índices col+2 ou col+3)
                c2 = cell(row, col + 2).upper()
                c3 = cell(row, col + 3).upper()
                if ("TOTAL" in c2 or "TOTAL" in c3):
                    linha_total = row
                    row += 1
                    break

                data   = cell(row, col)
                nota   = cell(row, col + 1)
                placa  = cell(row, col + 2)
                peso   = cell(row, col + 3)
                frete  = cell(row, col + 4)
                status = cell(row, col + 5)

                if any([data, nota, placa, peso, frete, status]):
                    linhas_dados.append({
                        "data": data, "nota": nota, "placa": placa,
                        "peso": peso, "frete": frete, "status": status,
                    })

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
                "cliente":         cliente,
                "cidade":          cidade,
                "fazenda":         fazenda,
                "fabrica":         fabrica,
                "pedido":          pedido,
                "produto":         produto,
                "saldo_total":     saldo_total,
                "total_carregado": total_carregado,
                "saldo_restante":  saldo_restante,
                "linhas":          linhas_dados,
                "col":             col,
                "linha_total":     linha_total,
                "num_linhas_dados": len(linhas_dados),
            })

    return blocos


def _col_letra(n):
    resultado = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        resultado = chr(65 + r) + resultado
    return resultado


def gravar_carregamento(conta, bloco, data, placa, peso, status):
    service     = _autenticar(conta)
    sheet       = service.spreadsheets()
    col         = bloco["col"]
    linha_total = bloco["linha_total"]

    if linha_total is None:
        raise Exception("Linha de TOTAL não encontrada no bloco.")

    # Converte linha (1-based) para índice (0-based) para a API
    linha_idx = linha_total - 1

    # 1. Insere uma linha vazia antes do TOTAL
    sheet.batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={
            "requests": [{
                "insertDimension": {
                    "range": {
                        "sheetId": _get_sheet_id(service),
                        "dimension": "ROWS",
                        "startIndex": linha_idx,
                        "endIndex": linha_idx + 1,
                    },
                    "inheritFromBefore": True,
                }
            }]
        }
    ).execute()

    # 2. Escreve os dados na nova linha
    col_ini = _col_letra(col + 1)
    col_fim = _col_letra(col + 6)
    rng     = f"'{ABA}'!{col_ini}{linha_total}:{col_fim}{linha_total}"

    nova_linha = [[data, "", placa, str(peso), "", status]]

    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": nova_linha},
    ).execute()


def _get_sheet_id(service):
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == ABA:
            return s["properties"]["sheetId"]
    raise Exception(f"Aba '{ABA}' não encontrada.")