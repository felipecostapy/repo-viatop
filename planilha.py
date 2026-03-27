import os
from pathlib import Path

SPREADSHEET_ID   = "16WPNHpjECjcPN5F_PcV3jbXr0OUzkBlbMxn_1p7QXio"
ABA              = "Página01"
TOKENS_DIR       = "gmail_tokens"
CREDENTIALS_FILE = "credentials.json"
COL_STARTS       = [0, 7, 14]
COL_WIDTH        = 6

# Nova planilha de dados (tabela plana)
DADOS_ID  = "1Uh_CFR2C3YypSIVQhGuekRTk_Lvy7I-0G5aFzIr39As"
DADOS_ABA = "DADOS"

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


def carregar_blocos_dados(conta):
    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    resp = sheet.values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{DADOS_ABA}'",
        valueRenderOption="FORMATTED_VALUE",
    ).execute()

    valores = resp.get("values", [])
    if not valores:
        return []

    pedidos = {}
    ordem   = []

    for i, linha in enumerate(valores[1:], 2):
        while len(linha) < 11:
            linha.append("")

        destino = str(linha[0]).strip()
        cliente = str(linha[1]).strip()
        pedido  = str(linha[2]).strip()
        produto = str(linha[3]).strip()

        if not any([destino, cliente, pedido]):
            continue

        key = f"{cliente}||{pedido}||{produto}"

        if key not in pedidos:
            try:
                saldo_total = float(str(linha[4]).replace(",", "."))
            except Exception:
                saldo_total = 0
            pedidos[key] = {
                "destino":         destino,
                "cliente":         cliente,
                "cidade":          destino,
                "fazenda":         "",
                "fabrica":         "",
                "pedido":          pedido,
                "produto":         produto,
                "saldo_total":     saldo_total,
                "total_carregado": 0,
                "saldo_restante":  saldo_total,
                "linhas":          [],
                "col":             None,
                "linha_total":     None,
            }
            ordem.append(key)

        data   = str(linha[5]).strip()
        nota   = str(linha[6]).strip()
        placa  = str(linha[7]).strip()
        peso   = str(linha[8]).strip()
        frete  = str(linha[9]).strip()
        status = str(linha[10]).strip()

        if any([data, nota, placa, peso]):
            pedidos[key]["linhas"].append({
                "data": data, "nota": nota, "placa": placa,
                "peso": peso, "frete": frete, "status": status,
            })

    for key in ordem:
        b = pedidos[key]
        try:
            b["total_carregado"] = sum(
                float(str(l["peso"]).replace(",", "."))
                for l in b["linhas"] if l["peso"]
            )
            b["saldo_restante"] = b["saldo_total"] - b["total_carregado"]
        except Exception:
            pass

    return [pedidos[k] for k in ordem]


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


BASE_ID  = "1y9blepnFkYoVrUNnUhdwxybUIlhaSw_ADaGGbT4RWPs"
BASE_ABA = "BASE 03/2026"

def carregar_base(conta):
    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    resp = sheet.values().get(
        spreadsheetId=BASE_ID,
        range=f"'{BASE_ABA}'",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    valores = resp.get("values", [])
    if not valores:
        return []

    dados = []
    header_found = False
    for i, linha in enumerate(valores[1:], 2):
        while len(linha) < 17:
            linha.append("")
        if not header_found:
            if str(linha[0]).strip().upper() == "DATA":
                header_found = True
            continue
        data    = str(linha[0]).strip()
        pagador = str(linha[2]).strip()
        if not data or not pagador:
            continue
        row = list(linha[:17]) + [i]
        dados.append(row)

    return dados


def carregar_base_com_linhas(conta):
    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    resp = sheet.values().get(
        spreadsheetId=BASE_ID,
        range=f"'{BASE_ABA}'",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    valores = resp.get("values", [])
    resultado = []
    for i, linha in enumerate(valores[1:], 2):
        while len(linha) < 17:
            linha.append("")
        resultado.append((i, linha[:17]))
    return resultado


def atualizar_linha_base(conta, num_linha, novos_dados):
    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    col_fim = chr(64 + len(novos_dados)) if len(novos_dados) <= 26 else "Q"
    rng = f"'{BASE_ABA}'!A{num_linha}:{col_fim}{num_linha}"

    sheet.values().update(
        spreadsheetId=BASE_ID,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": [novos_dados]},
    ).execute()


def deletar_linha_base(conta, num_linha):
    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    meta = service.spreadsheets().get(spreadsheetId=BASE_ID).execute()
    sheet_id = None
    for s in meta["sheets"]:
        if s["properties"]["title"] == BASE_ABA:
            sheet_id = s["properties"]["sheetId"]
            break

    if sheet_id is None:
        raise Exception(f"Aba '{BASE_ABA}' não encontrada.")

    sheet.batchUpdate(
        spreadsheetId=BASE_ID,
        body={
            "requests": [{
                "deleteDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": num_linha - 1,
                        "endIndex": num_linha,
                    }
                }
            }]
        }
    ).execute()


def _get_dados_sheet_id(service):
    meta = service.spreadsheets().get(spreadsheetId=DADOS_ID).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == DADOS_ABA:
            return s["properties"]["sheetId"]
    raise Exception(f"Aba '{DADOS_ABA}' não encontrada.")


def _ultima_linha_dados(service):
    resp = service.spreadsheets().values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{DADOS_ABA}'!A:A",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    valores = resp.get("values", [])
    # Encontra última linha com conteúdo
    ultima = len(valores)
    # Retorna índice 0-based para inserir APÓS a última linha com dado
    return ultima


def _inserir_linha_dados(service, sheet_id, linha_idx, valores):
    service.spreadsheets().batchUpdate(
        spreadsheetId=DADOS_ID,
        body={"requests": [{
            "insertDimension": {
                "range": {
                    "sheetId":    sheet_id,
                    "dimension":  "ROWS",
                    "startIndex": linha_idx,
                    "endIndex":   linha_idx + 1,
                },
                "inheritFromBefore": True,
            }
        }]}
    ).execute()

    col_fim = _col_letra(len(valores))
    rng = f"'{DADOS_ABA}'!A{linha_idx + 1}:{col_fim}{linha_idx + 1}"
    service.spreadsheets().values().update(
        spreadsheetId=DADOS_ID,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": [valores]},
    ).execute()


def criar_pedido_dados(conta, destino, cliente, pedido, produto, saldo_total):
    service  = _autenticar(conta)
    sheet_id = _get_dados_sheet_id(service)
    ultima   = _ultima_linha_dados(service)

    valores = [destino, cliente, pedido, produto, saldo_total,
               "", "", "", "", "", "NÃO CARREGADO"]
    _inserir_linha_dados(service, sheet_id, ultima, valores)


def gravar_carregamento_dados(conta, destino, cliente, pedido, produto,
                               data, placa, peso, frete, status):
    service  = _autenticar(conta)
    sheet_id = _get_dados_sheet_id(service)
    ultima   = _ultima_linha_dados(service)

    try:
        saldo_total = _buscar_saldo_total_dados(service, pedido)
    except Exception:
        saldo_total = ""

    valores = [destino, cliente, pedido, produto, saldo_total,
               data, "", placa, peso, frete, status]
    _inserir_linha_dados(service, sheet_id, ultima, valores)


def _buscar_saldo_total_dados(service, pedido):
    resp = service.spreadsheets().values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{DADOS_ABA}'!A:E",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    for linha in resp.get("values", [])[1:]:
        if len(linha) > 2 and str(linha[2]).strip() == str(pedido).strip():
            try:
                return float(str(linha[4]).replace(",", "."))
            except Exception:
                return ""
    return ""


def atualizar_saldo_dados(conta, cliente, pedido, produto, novo_saldo):
    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    resp = sheet.values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{DADOS_ABA}'!A:E",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    valores = resp.get("values", [])
    linhas_atualizar = []

    for i, linha in enumerate(valores[1:], 2):
        if len(linha) < 3:
            continue
        if (str(linha[1]).strip().upper() == str(cliente).strip().upper() and
            str(linha[2]).strip() == str(pedido).strip() and
            str(linha[3]).strip().upper() == str(produto).strip().upper() if len(linha) > 3 else True):
            linhas_atualizar.append(i)

    if not linhas_atualizar:
        raise Exception(f"Pedido {pedido} do cliente {cliente} não encontrado.")

    requests = []
    for num_linha in linhas_atualizar:
        requests.append({
            "range": f"'{DADOS_ABA}'!E{num_linha}",
            "values": [[novo_saldo]],
        })

    sheet.values().batchUpdate(
        spreadsheetId=DADOS_ID,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": requests,
        }
    ).execute()


def atualizar_status_base(conta, num_linha, novo_status):
    service = _autenticar(conta)
    sheet   = service.spreadsheets()
    # STATUS é a coluna O (15ª coluna = índice 14)
    rng = f"'{BASE_ABA}'!O{num_linha}"
    sheet.values().update(
        spreadsheetId=BASE_ID,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": [[novo_status]]},
    ).execute()


def gravar_ordem_dupla(conta, dados_ordem, filial, status="CONFERIDO"):
    service = _autenticar(conta)

    # ── 1. FRETES (BASE) ──────────────────────────────
    sheet_id_base = _get_base_sheet_id(service)
    ultima_base   = _ultima_linha_base(service)

    data         = str(dados_ordem.get("Data Apresentação", "")).upper()
    motorista    = str(dados_ordem.get("Motorista", "")).upper()
    placa        = str(dados_ordem.get("Cavalo", "")).upper()
    fabrica      = str(dados_ordem.get("Fábrica", "")).upper()
    destino      = str(dados_ordem.get("Destino", "")).upper()
    pedido       = str(dados_ordem.get("Pedido", "")).upper()
    produto      = str(dados_ordem.get("Produto", "")).upper()
    peso         = str(dados_ordem.get("Peso", "")).upper()
    pagador      = str(dados_ordem.get("Cliente", "")).upper()
    agencia      = str(dados_ordem.get("Agência", filial)).upper()
    uf           = str(dados_ordem.get("UF", "")).upper()
    frete_emp    = str(dados_ordem.get("Frete/Emp", "")).upper()
    frete_mot    = str(dados_ordem.get("Frete/Mot", "")).upper()
    rota         = str(dados_ordem.get("Rota", "")).upper()
    agenciamento = str(dados_ordem.get("Agenciamento", "")).upper()

    linha_fretes = [
        data, filial, pagador, agencia, motorista,
        placa, fabrica, destino, uf, peso,
        frete_emp, frete_mot, rota, agenciamento, status, pedido, produto
    ]

    _inserir_linha_base(service, sheet_id_base, ultima_base, linha_fretes)

    # ── 2. DADOS (saldo) ──────────────────────────────
    cliente = pagador
    try:
        peso_num = float(peso.replace(",", "."))
    except Exception:
        peso_num = 0

    if peso_num > 0:
        gravar_carregamento_dados(
            conta, destino, cliente, pedido, produto,
            data, placa, peso, frete_emp, status
        )


def _get_base_sheet_id(service):
    meta = service.spreadsheets().get(spreadsheetId=BASE_ID).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == BASE_ABA:
            return s["properties"]["sheetId"]
    raise Exception(f"Aba '{BASE_ABA}' não encontrada.")


def _ultima_linha_base(service):
    resp = service.spreadsheets().values().get(
        spreadsheetId=BASE_ID,
        range=f"'{BASE_ABA}'!A:A",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    return len(resp.get("values", []))


def _inserir_linha_base(service, sheet_id, linha_idx, valores):
    service.spreadsheets().batchUpdate(
        spreadsheetId=BASE_ID,
        body={"requests": [{
            "insertDimension": {
                "range": {
                    "sheetId":    sheet_id,
                    "dimension":  "ROWS",
                    "startIndex": linha_idx,
                    "endIndex":   linha_idx + 1,
                },
                "inheritFromBefore": True,
            }
        }]}
    ).execute()

    col_fim = _col_letra(len(valores))
    rng = f"'{BASE_ABA}'!A{linha_idx + 1}:{col_fim}{linha_idx + 1}"
    service.spreadsheets().values().update(
        spreadsheetId=BASE_ID,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": [valores]},
    ).execute()