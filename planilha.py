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


def _formatar_data(valor):
    """
    Normaliza datas vindas do Sheets para DD/MM ou DD/MM/AAAA.
    Trata: '2026-03-02 00:00:00', '2026-03-02', '02/03/2026', '28/01', etc.
    """
    import re
    v = str(valor).strip()
    if not v:
        return v

    # Formato datetime completo: 2026-03-02 00:00:00
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})(?:\s.*)?$", v)
    if m:
        ano, mes, dia = m.group(1), m.group(2), m.group(3)
        return f"{dia}/{mes}/{ano}"

    # Formato ISO sem hora: 2026-03-02
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})$", v)
    if m:
        ano, mes, dia = m.group(1), m.group(2), m.group(3)
        return f"{dia}/{mes}/{ano}"

    # Já está no formato DD/MM ou DD/MM/AAAA — retorna como está
    return v


def _formatar_nota(valor):
    """
    Remove o .0 de notas gravadas como float (ex: 134594.0 → 134594).
    """
    v = str(valor).strip()
    if not v:
        return v
    # Se termina em .0 (ou .00, etc.) e é numérico, remove a parte decimal
    try:
        f = float(v)
        if f == int(f):
            return str(int(f))
    except Exception:
        pass
    return v


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

        data   = _formatar_data(str(linha[5]).strip())
        nota   = _formatar_nota(str(linha[6]).strip())
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
BASE_ABA = "BASE 03/2026"   # fallback — sobrescrito dinamicamente


def listar_abas_base(conta):
    """Retorna lista de nomes de abas da planilha BASE, ordenadas."""
    service = _autenticar(conta)
    meta    = service.spreadsheets().get(spreadsheetId=BASE_ID).execute()
    abas    = [s["properties"]["title"] for s in meta["sheets"]]
    return abas


def _aba_mais_recente(abas):
    """
    Detecta a aba mais recente com padrão 'BASE MM/AAAA'.
    Fallback: última aba da lista.
    """
    import re
    candidatas = []
    for nome in abas:
        m = re.match(r"BASE\s+(\d{2})/(\d{4})", nome.strip(), re.IGNORECASE)
        if m:
            mes, ano = int(m.group(1)), int(m.group(2))
            candidatas.append((ano, mes, nome))
    if candidatas:
        candidatas.sort(reverse=True)
        return candidatas[0][2]
    return abas[-1] if abas else BASE_ABA


def carregar_base(conta, aba=None):
    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    aba_usar = aba or BASE_ABA

    resp = sheet.values().get(
        spreadsheetId=BASE_ID,
        range=f"'{aba_usar}'",
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


def carregar_base_com_linhas(conta, aba=None):
    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    aba_usar = aba or BASE_ABA

    resp = sheet.values().get(
        spreadsheetId=BASE_ID,
        range=f"'{aba_usar}'",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    valores = resp.get("values", [])
    resultado = []
    for i, linha in enumerate(valores[1:], 2):
        while len(linha) < 17:
            linha.append("")
        resultado.append((i, linha[:17]))
    return resultado


def atualizar_linha_base(conta, num_linha, novos_dados, aba=None):
    service  = _autenticar(conta)
    sheet    = service.spreadsheets()
    aba_usar = aba or BASE_ABA

    col_fim = _col_letra(len(novos_dados))
    rng = f"'{aba_usar}'!A{num_linha}:{col_fim}{num_linha}"

    sheet.values().update(
        spreadsheetId=BASE_ID,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": [novos_dados]},
    ).execute()


def deletar_linha_base(conta, num_linha, aba=None):
    service  = _autenticar(conta)
    sheet    = service.spreadsheets()
    aba_usar = aba or BASE_ABA

    meta = service.spreadsheets().get(spreadsheetId=BASE_ID).execute()
    sheet_id = None
    for s in meta["sheets"]:
        if s["properties"]["title"] == aba_usar:
            sheet_id = s["properties"]["sheetId"]
            break

    if sheet_id is None:
        raise Exception(f"Aba '{aba_usar}' não encontrada.")

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


def atualizar_saldo_dados(conta, cliente, pedido, produto, saldo_restante_desejado):
    """
    Ajusta o SALDO_TOTAL de todas as linhas do pedido na planilha DADOS.

    O usuário informa o SALDO RESTANTE que quer (ex: 500t).
    O sistema soma os pesos já carregados e calcula o SALDO_TOTAL real:
        saldo_total = saldo_restante_desejado + total_ja_carregado

    Assim o cálculo saldo_total - soma(pesos) sempre bate com o valor digitado.
    Usa fuzzy match nas palavras do cliente e produto para evitar erros de digitação.
    """
    import re

    def _normalizar(s):
        return re.sub(r"\s+", " ", str(s).strip().upper())

    def _palavras_em_comum(a, b):
        pa = {p for p in _normalizar(a).split() if len(p) > 2}
        pb = {p for p in _normalizar(b).split() if len(p) > 2}
        if not pa or not pb:
            return True
        menor, maior = (pa, pb) if len(pa) <= len(pb) else (pb, pa)
        return sum(1 for p in menor if p in maior) / len(menor) >= 0.6

    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    # Lê colunas A:I (destino, cliente, pedido, produto, saldo_total, data, nota, placa, peso)
    resp = sheet.values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{DADOS_ABA}'!A:I",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    valores = resp.get("values", [])
    linhas_atualizar   = []
    total_ja_carregado = 0.0

    for i, linha in enumerate(valores[1:], 2):
        if len(linha) < 3:
            continue

        ped_plan  = _normalizar(linha[2])
        cli_plan  = _normalizar(linha[1]) if len(linha) > 1 else ""
        prod_plan = _normalizar(linha[3]) if len(linha) > 3 else ""

        if ped_plan != _normalizar(pedido):
            continue
        if not _palavras_em_comum(produto, prod_plan):
            continue
        if not _palavras_em_comum(cliente, cli_plan):
            continue

        linhas_atualizar.append(i)

        # Soma os pesos já carregados (coluna I, índice 8)
        if len(linha) > 8 and str(linha[8]).strip():
            try:
                total_ja_carregado += float(str(linha[8]).replace(",", "."))
            except Exception:
                pass

    if not linhas_atualizar:
        raise Exception(
            f"Pedido {pedido} — {produto} do cliente {cliente} não encontrado na planilha de saldo."
        )

    # saldo_total real = o que o usuário quer de restante + o que já foi carregado
    saldo_total_real = float(saldo_restante_desejado) + total_ja_carregado

    requests = [
        {"range": f"'{DADOS_ABA}'!E{n}", "values": [[saldo_total_real]]}
        for n in linhas_atualizar
    ]

    sheet.values().batchUpdate(
        spreadsheetId=DADOS_ID,
        body={"valueInputOption": "USER_ENTERED", "data": requests},
    ).execute()

    return saldo_total_real, total_ja_carregado


def atualizar_status_base(conta, num_linha, novo_status, aba=None):
    service  = _autenticar(conta)
    sheet    = service.spreadsheets()
    aba_usar = aba or BASE_ABA
    # STATUS é a coluna O (15ª coluna = índice 14)
    rng = f"'{aba_usar}'!O{num_linha}"
    sheet.values().update(
        spreadsheetId=BASE_ID,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": [[novo_status]]},
    ).execute()


def gravar_ordem_dupla(conta, dados_ordem, filial, status="CONFERIDO", aba=None):
    """
    Grava a ordem na aba BASE do mês correto e desconta o peso do saldo
    do pedido na planilha DADOS.
    aba: nome da aba BASE a usar (ex: 'BASE 04/2026'). Se None, usa a mais recente.
    """
    service = _autenticar(conta)

    # Resolve aba: usa a passada, ou detecta a mais recente automaticamente
    if not aba:
        abas = listar_abas_base(conta)
        aba  = _aba_mais_recente(abas)

    # ── 1. BASE (controle de ordens) ──────────────────
    sheet_id_base = _get_base_sheet_id(service, aba=aba)
    ultima_base   = _ultima_linha_base(service, aba=aba)

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

    _inserir_linha_base(service, sheet_id_base, ultima_base, linha_fretes, aba=aba)

    # ── 2. DADOS — desconta peso do saldo do pedido ───
    cliente = pagador
    try:
        peso_num = float(peso.replace(",", "."))
    except Exception:
        peso_num = 0

    if peso_num > 0 and pedido:
        return _descontar_saldo_pedido(
            service, pedido, cliente, peso_num,
            destino  = destino,
            produto  = produto,
            data     = data,
            placa    = placa,
            frete    = frete_emp,
            status   = status,
        )
    return False  # peso zero ou sem pedido — saldo não descontado


def _descontar_saldo_pedido(service, pedido, cliente, peso_num,
                             destino="", produto="", data="", placa="",
                             frete="", status=""):
    """
    Insere uma linha de carregamento na planilha DADOS vinculada ao pedido.
    O saldo_restante é calculado dinamicamente por carregar_blocos_dados
    como: saldo_total - soma(pesos dos carregamentos do pedido).
    Só grava se o pedido já estiver cadastrado na planilha de saldo.
    Retorna True se gravou, False se pedido não encontrado.
    """
    sheet = service.spreadsheets()

    # Busca o registro-mestre do pedido para obter destino/cliente/produto/saldo_total
    resp = sheet.values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{DADOS_ABA}'!A:E",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    valores      = resp.get("values", [])
    dados_mestre = None
    pedido_upper = str(pedido).strip().upper()
    cliente_upper = str(cliente).strip().upper()
    produto_upper = str(produto).strip().upper()

    def _normalizar(s):
        """Remove espaços duplos, strip e upper."""
        import re
        return re.sub(r"\s+", " ", str(s).strip().upper())

    def _palavras_em_comum(a, b):
        """Verifica se as palavras principais de 'a' existem em 'b' ou vice-versa."""
        palavras_a = set(_normalizar(a).split())
        palavras_b = set(_normalizar(b).split())
        # Remove palavras muito curtas (artigos, preposições)
        palavras_a = {p for p in palavras_a if len(p) > 2}
        palavras_b = {p for p in palavras_b if len(p) > 2}
        if not palavras_a or not palavras_b:
            return True  # sem palavras suficientes, não bloqueia
        # Considera match se pelo menos 60% das palavras da menor string estiver na outra
        menor = palavras_a if len(palavras_a) <= len(palavras_b) else palavras_b
        maior = palavras_b if len(palavras_a) <= len(palavras_b) else palavras_a
        matches = sum(1 for p in menor if p in maior)
        return matches / len(menor) >= 0.6

    for linha in valores[1:]:
        if len(linha) < 4:
            continue

        ped_plan  = _normalizar(linha[2])
        cli_plan  = _normalizar(linha[1]) if len(linha) > 1 else ""
        prod_plan = _normalizar(linha[3]) if len(linha) > 3 else ""

        # 1. Pedido deve ser idêntico
        if ped_plan != pedido_upper:
            continue

        # 2. Produto deve ter palavras em comum
        if produto_upper and prod_plan and not _palavras_em_comum(produto_upper, prod_plan):
            continue

        # 3. Cliente deve ter palavras em comum
        if cliente_upper and cli_plan and not _palavras_em_comum(cliente_upper, cli_plan):
            continue

        # Pega saldo_total — ignora linhas com saldo vazio mas continua procurando
        saldo_total = 0
        if len(linha) > 4 and str(linha[4]).strip():
            try:
                saldo_total = float(str(linha[4]).replace(",", "."))
            except Exception:
                saldo_total = 0

        dados_mestre = (
            str(linha[0]).strip() if linha[0] else destino,
            str(linha[1]).strip() if len(linha) > 1 else cliente,
            str(linha[2]).strip(),
            str(linha[3]).strip() if len(linha) > 3 else produto,
            saldo_total,
        )
        break

    if dados_mestre is None:
        return False

    dest_final, cli_final, ped_final, prod_final, saldo_total = dados_mestre

    # Insere linha de carregamento
    sheet_id = _get_dados_sheet_id(service)
    ultima   = _ultima_linha_dados(service)

    valores_novo = [
        dest_final,
        cli_final,
        ped_final,
        prod_final,
        saldo_total,           # repete o saldo_total (referência)
        data,                  # DATA do carregamento
        "",                    # NOTA (em branco — não temos aqui)
        placa,                 # PLACA
        str(peso_num),         # PESO carregado
        frete,                 # FRETE
        status or "CARREGADO", # STATUS
    ]
    _inserir_linha_dados(service, sheet_id, ultima, valores_novo)
    return True


def _get_base_sheet_id(service, aba=None):
    aba_usar = aba or BASE_ABA
    meta = service.spreadsheets().get(spreadsheetId=BASE_ID).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == aba_usar:
            return s["properties"]["sheetId"]
    raise Exception(f"Aba '{aba_usar}' não encontrada na planilha BASE.")


def _ultima_linha_base(service, aba=None):
    aba_usar = aba or BASE_ABA
    resp = service.spreadsheets().values().get(
        spreadsheetId=BASE_ID,
        range=f"'{aba_usar}'!A:A",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    return len(resp.get("values", []))


def _inserir_linha_base(service, sheet_id, linha_idx, valores, aba=None):
    aba_usar = aba or BASE_ABA
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
    rng = f"'{aba_usar}'!A{linha_idx + 1}:{col_fim}{linha_idx + 1}"
    service.spreadsheets().values().update(
        spreadsheetId=BASE_ID,
        range=rng,
        valueInputOption="USER_ENTERED",
        body={"values": [valores]},
    ).execute()