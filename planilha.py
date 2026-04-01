import os
import re
from pathlib import Path

# ── Planilha 1: BASE (ordens por mês) ─────────────────────────────
BASE_ID  = "1xab_LceMGpjIhXKDp1iUp3OUjqk1Z3SOz05vScFqTJg"
BASE_ABA = "BASE 03/2026"   # fallback — sobrescrito dinamicamente

# ── Planilha 2: DADOS (saldo, config, histórico) ───────────────────
DADOS_ID      = "18kpcoEF6dT19s3crJ0forj9n5iAG6HGDRbChnpvUlZI"
PEDIDOS_ABA   = "PEDIDOS"    # cadastro dos pedidos — uma linha por pedido
DADOS_ABA     = "DADOS"      # carregamentos — uma linha por carregamento
CONFIG_ABA    = "CONFIG"
HISTORICO_ABA = "HISTORICO"

TOKENS_DIR       = "gmail_tokens"
CREDENTIALS_FILE = "credentials.json"

SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/spreadsheets",
    "openid",
]



def carregar_usuarios_planilha(conta):
    """
    Lê os usuários da aba CONFIG da planilha BASE.
    Formato esperado na aba CONFIG:
      Coluna A: LOGIN (ex: FELIPE)
      Coluna B: ASSINATURA (ex: Felipe Costa)
    Retorna dict {LOGIN: ASSINATURA}.
    Fallback para arquivo local se aba não existir.
    """
    try:
        service = _autenticar(conta)
        resp = service.spreadsheets().values().get(
            spreadsheetId=DADOS_ID,
            range=f"'{CONFIG_ABA}'!A:B",
            valueRenderOption="FORMATTED_VALUE",
        ).execute()
        valores = resp.get("values", [])
        usuarios = {}
        for linha in valores:
            if len(linha) >= 2 and str(linha[0]).strip():
                login     = str(linha[0]).strip().upper()
                assinatura = str(linha[1]).strip()
                if login and assinatura:
                    usuarios[login] = assinatura
        return usuarios if usuarios else None
    except Exception:
        return None


def gravar_historico_planilha(conta, registro):
    """
    Grava um registro de histórico na aba HISTORICO da planilha BASE.
    Colunas: DATA_HORA | USUARIO | MOTORISTA | PLACA | EMPRESA | ARQUIVO
    Usa append atômico — seguro para múltiplos usuários.
    """
    try:
        service = _autenticar(conta)
        valores = [
            _converter_data_para_sheets(registro.get("data_hora", "")),
            registro.get("usuario",   ""),
            registro.get("motorista", ""),
            registro.get("placa",     ""),
            registro.get("empresa",   ""),
            registro.get("arquivo",   ""),
        ]
        service.spreadsheets().values().append(
            spreadsheetId=DADOS_ID,
            range=f"'{HISTORICO_ABA}'!A1",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": [valores]},
        ).execute()
        return True
    except Exception:
        return False


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




def _normalizar(s):
    """Remove espaços duplos, strip e upper."""
    return re.sub(r"\s+", " ", str(s).strip().upper())


def _palavras_em_comum(a, b, threshold=0.6):
    """True se >= threshold das palavras da string menor existem na maior."""
    pa = {p for p in _normalizar(a).split() if len(p) > 2}
    pb = {p for p in _normalizar(b).split() if len(p) > 2}
    if not pa or not pb:
        return True
    menor, maior = (pa, pb) if len(pa) <= len(pb) else (pb, pa)
    return sum(1 for p in menor if p in maior) / len(menor) >= threshold

def _converter_data_para_sheets(valor):
    """
    Converte data do formato brasileiro DD/MM/YYYY para ISO YYYY-MM-DD
    que o Google Sheets interpreta corretamente independente da localização.
    Também trata formato DD/MM (sem ano) mantendo como string.
    """
    import re
    v = str(valor).strip()
    if not v:
        return v
    # DD/MM/YYYY → YYYY-MM-DD
    m = re.match(r"(\d{2})/(\d{2})/(\d{4})$", v)
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    # Já está em ISO ou outro formato — retorna como está
    return v


def _formatar_data(valor):
    """
    Normaliza datas vindas do Sheets para DD/MM/AAAA.
    Trata:
      - Número serial do Excel (ex: 46111.63)  → converte via datetime
      - '2026-03-02 00:00:00' (datetime str)   → DD/MM/AAAA
      - '2026-03-02' (ISO)                      → DD/MM/AAAA
      - '02/03/2026' (já formatado)             → mantém
      - '28/01' (sem ano)                       → mantém
    """
    import datetime as _dt
    v = str(valor).strip()
    if not v:
        return v

    # Número serial do Excel (float ou int) — ex: 46111.63
    try:
        serial = float(v)
        if 40000 < serial < 60000:   # intervalo razoável para datas 2009-2064
            # Excel serial: dias desde 30/12/1899
            data = _dt.datetime(1899, 12, 30) + _dt.timedelta(days=serial)
            return data.strftime("%d/%m/%Y")
    except (ValueError, TypeError):
        pass

    # Formato datetime completo: 2026-03-02 00:00:00 ou 2026/03/02 00:00:00
    m = re.match(r"(\d{4})[-/](\d{2})[-/](\d{2})(?:[\sT].*)?$", v)
    if m:
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"

    # Já está no formato DD/MM/AAAA ou DD/MM — retorna como está
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
    """
    Lê a aba PEDIDOS (cadastro) e a aba DADOS (carregamentos) separadamente
    e combina para montar os blocos exibidos nos cards do Controle de Pedidos.

    Aba PEDIDOS: DESTINO | CLIENTE | PEDIDO | PRODUTO | SALDO_TOTAL
    Aba DADOS:   PEDIDO  | PRODUTO | CLIENTE | DATA | NOTA | PESO | FRETE | STATUS
    """
    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    # ── 1. Lê cadastro de pedidos ────────────────────────────────
    resp_ped = sheet.values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{PEDIDOS_ABA}'",
        valueRenderOption="FORMATTED_VALUE",
    ).execute()

    pedidos = {}
    ordem   = []

    for linha in resp_ped.get("values", [])[1:]:
        while len(linha) < 5:
            linha.append("")
        destino = str(linha[0]).strip()
        cliente = str(linha[1]).strip()
        pedido  = str(linha[2]).strip()
        produto = str(linha[3]).strip()
        if not any([cliente, pedido]):
            continue
        try:
            saldo_total = float(str(linha[4]).replace(",", "."))
        except Exception:
            saldo_total = 0

        key = f"{_normalizar(cliente)}||{_normalizar(pedido)}||{_normalizar(produto)}"
        if key not in pedidos:
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

    # ── 2. Lê carregamentos e vincula ao pedido correto ──────────
    resp_dados = sheet.values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{DADOS_ABA}'",
        valueRenderOption="FORMATTED_VALUE",
    ).execute()

    for linha in resp_dados.get("values", [])[1:]:
        while len(linha) < 9:
            linha.append("")
        # Aba DADOS: PEDIDO | PRODUTO | CLIENTE | DATA | NOTA | PESO | FRETE | STATUS
        ped_c  = str(linha[0]).strip()
        prod_c = str(linha[1]).strip()
        cli_c  = str(linha[2]).strip()
        data   = _formatar_data(str(linha[3]).strip())
        nota   = _formatar_nota(str(linha[4]).strip())
        placa  = str(linha[5]).strip() if len(linha) > 5 else ""
        peso   = str(linha[6]).strip() if len(linha) > 6 else ""
        frete  = str(linha[7]).strip() if len(linha) > 7 else ""
        status = str(linha[8]).strip() if len(linha) > 8 else ""

        if not any([ped_c, cli_c]):
            continue

        # Encontra o pedido correspondente por fuzzy match
        chave_encontrada = None
        for key in pedidos:
            cli_key, ped_key, prod_key = key.split("||")
            if (_normalizar(ped_c) == ped_key and
                _palavras_em_comum(prod_c, prod_key) and
                _palavras_em_comum(cli_c, cli_key)):
                chave_encontrada = key
                break

        if chave_encontrada and any([data, nota, placa, peso]):
            pedidos[chave_encontrada]["linhas"].append({
                "data": data, "nota": nota, "placa": placa,
                "peso": peso, "frete": frete, "status": status,
            })

    # ── 3. Calcula saldo restante ────────────────────────────────
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




def listar_abas_base(conta):
    """Retorna lista de nomes de abas da planilha BASE, ordenadas."""
    service = _autenticar(conta)
    meta    = service.spreadsheets().get(spreadsheetId=BASE_ID).execute()
    abas    = [s["properties"]["title"] for s in meta["sheets"]]
    return abas


def migrar_base_para_dados(conta, abas_filtro=None, callback_progresso=None):
    """
    Migra carregamentos das abas BASE MM/AAAA para as abas PEDIDOS e DADOS.
    abas_filtro: lista de nomes de abas a migrar. Se None, migra todas.

    Para cada combinação única PEDIDO+PRODUTO+CLIENTE:
      - Cria uma linha na aba PEDIDOS com SALDO_TOTAL = 0 (ajustar manualmente depois)
      - Grava cada carregamento na aba DADOS

    Retorna dict com resumo: {pedidos_criados, carregamentos_migrados, abas_lidas, total_lido}
    """
    service = _autenticar(conta)

    # ── 1. Define quais abas migrar ───────────────────────────────
    if abas_filtro:
        abas_base = abas_filtro
    else:
        abas = listar_abas_base(conta)
        abas_base = [a for a in abas if re.match(r"BASE\s+\d{2}/?\d{4}", a, re.IGNORECASE)]

    if not abas_base:
        raise Exception("Nenhuma aba BASE MM/AAAA encontrada na planilha.")

    # ── 2. Lê carregamentos de todas as abas ──────────────────────
    # BASE colunas: DATA(0) FILIAL(1) PAGADOR(2) AGENCIA(3) MOTORISTA(4)
    #               PLACA(5) FABRICA(6) DESTINO(7) UF(8) PESO(9)
    #               FRETE/E(10) FRETE/M(11) ROTA(12) AGENCIAMENTO(13)
    #               STATUS(14) PEDIDO(15) PRODUTO(16) COLOCADOR(17)

    todos = []  # lista de dicts com cada carregamento

    for aba in abas_base:
        if callback_progresso:
            callback_progresso(f"Lendo {aba}...")

        resp_aba = service.spreadsheets().values().get(
            spreadsheetId=BASE_ID,
            range=f"'{aba}'",
            valueRenderOption="FORMATTED_VALUE",
        ).execute()
        rows = resp_aba.get("values", [])

        for row in rows[1:]:  # pula cabeçalho
            while len(row) < 17:
                row.append("")
            # Ignora linhas de cabeçalho repetido ou sem data
            if str(row[0]).strip().upper() == "DATA":
                continue
            pedido  = str(row[15] or "").strip()
            produto = str(row[16] or "").strip()
            data    = _formatar_data(str(row[0] or "").strip())
            pagador = str(row[2] or "").strip()
            if not pedido or not data or not pagador:
                continue
            todos.append({
                "data":    data,
                "cliente": pagador,
                "destino": str(row[7] or "").strip(),
                "placa":   str(row[5] or "").strip(),
                "pedido":  pedido,
                "produto": produto,
                "peso":    str(row[9] or "").strip(),
                "frete":   str(row[10] or "").strip(),
                "status":  str(row[14] or "").strip(),
            })

    if not todos:
        raise Exception("Nenhum carregamento com PEDIDO preenchido encontrado na BASE.")

    # ── 3. Agrupa pedidos únicos ──────────────────────────────────
    pedidos_unicos = {}
    ordem_pedidos  = []

    for c in todos:
        key = f"{_normalizar(c['pedido'])}||{_normalizar(c['produto'])}||{_normalizar(c['cliente'])}"
        if key not in pedidos_unicos:
            pedidos_unicos[key] = {
                "destino": c["destino"],
                "cliente": c["cliente"],
                "pedido":  c["pedido"],
                "produto": c["produto"],
            }
            ordem_pedidos.append(key)

    # ── 4. Verifica quais pedidos já existem na aba PEDIDOS ───────
    resp_exist = service.spreadsheets().values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{PEDIDOS_ABA}'!A:D",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    ja_existem = set()
    for linha in resp_exist.get("values", [])[1:]:
        if len(linha) >= 3:
            key = f"{_normalizar(linha[2])}||{_normalizar(linha[3] if len(linha)>3 else '')}||{_normalizar(linha[1])}"
            ja_existem.add(key)

    # ── 5. Verifica quais carregamentos já existem na aba DADOS ───
    resp_dados = service.spreadsheets().values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{DADOS_ABA}'!A:F",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    ja_carregados = set()
    for linha in resp_dados.get("values", [])[1:]:
        if len(linha) >= 6:
            # chave: pedido+produto+cliente+data+placa
            chave_c = f"{_normalizar(linha[0])}||{_normalizar(linha[5])}||{_normalizar(linha[4])}"
            ja_carregados.add(chave_c)

    # ── 6. Grava pedidos novos na aba PEDIDOS ─────────────────────
    pedidos_criados = 0
    novos_pedidos   = []

    for key in ordem_pedidos:
        if key in ja_existem:
            continue
        p = pedidos_unicos[key]
        novos_pedidos.append([p["destino"], p["cliente"], p["pedido"], p["produto"], 0])
        pedidos_criados += 1

    if novos_pedidos:
        if callback_progresso:
            callback_progresso(f"Criando {len(novos_pedidos)} pedido(s) na aba PEDIDOS...")
        service.spreadsheets().values().append(
            spreadsheetId=DADOS_ID,
            range=f"'{PEDIDOS_ABA}'!A1",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": novos_pedidos},
        ).execute()

    # ── 7. Grava carregamentos novos na aba DADOS ─────────────────
    carregamentos_migrados = 0
    novos_carregamentos    = []

    for c in todos:
        chave_c = (f"{_normalizar(c['pedido'])}||"
                   f"{_normalizar(c['produto'])}||"
                   f"{_normalizar(c['data'])}||{_normalizar(c['placa'])}")
        if chave_c in ja_carregados:
            continue
        # Aba DADOS: PEDIDO|PRODUTO|CLIENTE|DATA|NOTA|PLACA|PESO|FRETE|STATUS
        novos_carregamentos.append([
            c["pedido"], c["produto"], c["cliente"],
            c["data"], "",
            c["placa"], c["peso"], c["frete"], c["status"],
        ])
        carregamentos_migrados += 1

    if novos_carregamentos:
        if callback_progresso:
            callback_progresso(f"Migrando {len(novos_carregamentos)} carregamento(s) para aba DADOS...")
        # Envia em lotes de 500 para evitar timeout
        lote = 500
        for i in range(0, len(novos_carregamentos), lote):
            service.spreadsheets().values().append(
                spreadsheetId=DADOS_ID,
                range=f"'{DADOS_ABA}'!A1",
                valueInputOption="USER_ENTERED",
                insertDataOption="INSERT_ROWS",
                body={"values": novos_carregamentos[i:i+lote]},
            ).execute()

    return {
        "pedidos_criados":       pedidos_criados,
        "carregamentos_migrados": carregamentos_migrados,
        "total_lido":            len(todos),
        "abas_lidas":            abas_base,
    }


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
        valueRenderOption="FORMATTED_VALUE",
    ).execute()

    valores = resp.get("values", [])
    if not valores:
        return []

    dados = []
    # Primeira linha é sempre o cabeçalho — começa da linha 2
    # Pula linhas que sejam cabeçalho repetido (ex: "DATA" na col 0)
    for i, linha in enumerate(valores[1:], 2):
        while len(linha) < 18:
            linha.append("")
        # Ignora linhas de cabeçalho repetido
        if str(linha[0]).strip().upper() == "DATA":
            continue
        data    = _formatar_data(str(linha[0]).strip())
        pagador = str(linha[2]).strip()
        if not data or not pagador:
            continue
        row    = list(linha[:17]) + [i]
        row[0] = data
        dados.append(row)

    return dados


def carregar_base_com_linhas(conta, aba=None):
    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    aba_usar = aba or BASE_ABA

    resp = sheet.values().get(
        spreadsheetId=BASE_ID,
        range=f"'{aba_usar}'",
        valueRenderOption="FORMATTED_VALUE",
    ).execute()

    valores = resp.get("values", [])
    resultado = []
    for i, linha in enumerate(valores[1:], 2):
        while len(linha) < 18:
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


def _deletar_carregamento_dados(service, pedido, produto, cliente, data, placa):
    """
    Remove o carregamento correspondente da aba DADOS após deletar da BASE.
    Identifica pelo conjunto: PEDIDO + PRODUTO + CLIENTE + DATA + PLACA.
    Remove apenas a primeira ocorrência encontrada.
    """
    if not pedido and not placa:
        return  # sem informação suficiente para identificar

    resp = service.spreadsheets().values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{DADOS_ABA}'!A:F",
        valueRenderOption="FORMATTED_VALUE",
    ).execute()

    valores = resp.get("values", [])
    linha_deletar = None

    ped_n   = _normalizar(pedido)
    prod_n  = _normalizar(produto)
    cli_n   = _normalizar(cliente)
    data_n  = _normalizar(data)
    placa_n = _normalizar(placa)

    for i, linha in enumerate(valores[1:], 2):
        while len(linha) < 6:
            linha.append("")
        # Aba DADOS: PEDIDO|PRODUTO|CLIENTE|DATA|NOTA|PLACA
        if (_normalizar(linha[0]) == ped_n and
            _palavras_em_comum(linha[1], prod_n) and
            _palavras_em_comum(linha[2], cli_n) and
            (_normalizar(linha[3]) == data_n or not data_n) and
            (_normalizar(linha[5]) == placa_n or not placa_n)):
            linha_deletar = i
            break

    if linha_deletar is None:
        return  # carregamento não encontrado — não bloqueia

    # Busca sheet_id da aba DADOS
    meta     = service.spreadsheets().get(spreadsheetId=DADOS_ID).execute()
    sheet_id = None
    for s in meta["sheets"]:
        if s["properties"]["title"] == DADOS_ABA:
            sheet_id = s["properties"]["sheetId"]
            break

    if sheet_id is None:
        return

    service.spreadsheets().batchUpdate(
        spreadsheetId=DADOS_ID,
        body={"requests": [{
            "deleteDimension": {
                "range": {
                    "sheetId":    sheet_id,
                    "dimension":  "ROWS",
                    "startIndex": linha_deletar - 1,
                    "endIndex":   linha_deletar,
                }
            }
        }]}
    ).execute()


def deletar_linha_base(conta, num_linha, aba=None, dados_linha=None):
    """
    Remove a linha da aba BASE e, se o carregamento existir na aba DADOS,
    remove também de lá.
    dados_linha: lista com os valores da linha completa da BASE (opcional).
                 Se fornecido, usa para identificar e remover da aba DADOS.
    """
    service  = _autenticar(conta)
    sheet    = service.spreadsheets()
    aba_usar = aba or BASE_ABA

    # ── 1. Remove da aba DADOS se tiver dados suficientes ──────────
    if dados_linha and len(dados_linha) >= 16:
        # BASE colunas: DATA(0) FILIAL(1) PAGADOR(2) ... PLACA(5) ...
        #               DESTINO(7) ... PEDIDO(15) PRODUTO(16)
        pedido  = str(dados_linha[15] or "").strip() if len(dados_linha) > 15 else ""
        produto = str(dados_linha[16] or "").strip() if len(dados_linha) > 16 else ""
        cliente = str(dados_linha[2]  or "").strip()
        data    = _formatar_data(str(dados_linha[0] or "").strip())
        placa   = str(dados_linha[5]  or "").strip()
        _deletar_carregamento_dados(service, pedido, produto, cliente, data, placa)

    # ── 2. Remove da aba BASE ───────────────────────────────────────
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


# ═══════════════════════════════════════════════════════════════════
# ABA PEDIDOS — cadastro de pedidos (uma linha por pedido)
# Colunas: DESTINO | CLIENTE | PEDIDO | PRODUTO | SALDO_TOTAL
# ═══════════════════════════════════════════════════════════════════

def remover_pedido_dados(conta, cliente, pedido, produto):
    """
    Remove o pedido da aba PEDIDOS e todos os seus carregamentos da aba DADOS.
    Retorna dict com quantas linhas foram removidas de cada aba.
    """
    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    # ── 1. Remove da aba PEDIDOS ──────────────────────────────────
    resp_ped = sheet.values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{PEDIDOS_ABA}'!A:D",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    linhas_pedido = []
    for i, linha in enumerate(resp_ped.get("values", [])[1:], 2):
        if len(linha) < 3:
            continue
        if (_normalizar(linha[2]) == _normalizar(pedido) and
            _palavras_em_comum(linha[3] if len(linha) > 3 else "", produto) and
            _palavras_em_comum(linha[1] if len(linha) > 1 else "", cliente)):
            linhas_pedido.append(i)

    # ── 2. Remove da aba DADOS ────────────────────────────────────
    resp_dados = sheet.values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{DADOS_ABA}'!A:C",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    linhas_dados = []
    for i, linha in enumerate(resp_dados.get("values", [])[1:], 2):
        if len(linha) < 1:
            continue
        if (_normalizar(linha[0]) == _normalizar(pedido) and
            _palavras_em_comum(linha[1] if len(linha) > 1 else "", produto) and
            _palavras_em_comum(linha[2] if len(linha) > 2 else "", cliente)):
            linhas_dados.append(i)

    if not linhas_pedido:
        raise Exception(
            f"Pedido {pedido} — {produto} do cliente {cliente} não encontrado na aba PEDIDOS."
        )

    # Deleta em ordem decrescente para não deslocar índices
    def _deletar_linhas(spreadsheet_id, aba_nome, numeros_linha):
        if not numeros_linha:
            return
        meta     = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheet_id = None
        for s in meta["sheets"]:
            if s["properties"]["title"] == aba_nome:
                sheet_id = s["properties"]["sheetId"]
                break
        if sheet_id is None:
            return
        # Ordena decrescente para deletar de baixo para cima
        requests = []
        for n in sorted(set(numeros_linha), reverse=True):
            requests.append({
                "deleteDimension": {
                    "range": {
                        "sheetId":    sheet_id,
                        "dimension":  "ROWS",
                        "startIndex": n - 1,
                        "endIndex":   n,
                    }
                }
            })
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()

    _deletar_linhas(DADOS_ID, PEDIDOS_ABA, linhas_pedido)
    _deletar_linhas(DADOS_ID, DADOS_ABA,   linhas_dados)

    return {
        "pedidos_removidos":       len(linhas_pedido),
        "carregamentos_removidos": len(linhas_dados),
    }


def criar_pedido_dados(conta, destino, cliente, pedido, produto, saldo_total):
    """Cadastra um novo pedido na aba PEDIDOS. Uma linha por pedido."""
    service = _autenticar(conta)
    valores = [destino, cliente, pedido, produto, saldo_total]
    service.spreadsheets().values().append(
        spreadsheetId=DADOS_ID,
        range=f"'{PEDIDOS_ABA}'!A1",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [valores]},
    ).execute()


def atualizar_saldo_dados(conta, cliente, pedido, produto, saldo_restante_desejado):
    """
    Atualiza o SALDO_TOTAL na aba PEDIDOS.
    O usuário informa o saldo RESTANTE desejado.
    O sistema soma os pesos já carregados na aba DADOS e calcula:
        saldo_total = saldo_restante_desejado + total_já_carregado
    """
    service = _autenticar(conta)
    sheet   = service.spreadsheets()

    # Lê aba PEDIDOS para encontrar a linha do pedido
    resp_ped = sheet.values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{PEDIDOS_ABA}'!A:E",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    linha_pedido = None
    for i, linha in enumerate(resp_ped.get("values", [])[1:], 2):
        if len(linha) < 3:
            continue
        if (_normalizar(linha[2]) == _normalizar(pedido) and
            _palavras_em_comum(produto, linha[3] if len(linha) > 3 else "") and
            _palavras_em_comum(cliente, linha[1] if len(linha) > 1 else "")):
            linha_pedido = i
            break

    if linha_pedido is None:
        raise Exception(
            f"Pedido {pedido} — {produto} do cliente {cliente} não encontrado na aba PEDIDOS."
        )

    # Soma pesos já carregados na aba DADOS
    resp_dados = sheet.values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{DADOS_ABA}'!A:G",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    total_carregado = 0.0
    for linha in resp_dados.get("values", [])[1:]:
        if len(linha) < 3:
            continue
        if (_normalizar(linha[0]) == _normalizar(pedido) and
            _palavras_em_comum(produto, linha[1] if len(linha) > 1 else "") and
            _palavras_em_comum(cliente, linha[2] if len(linha) > 2 else "")):
            try:
                total_carregado += float(str(linha[4] if len(linha) > 4 else 0).replace(",", "."))
            except Exception:
                pass

    saldo_total_real = float(saldo_restante_desejado) + total_carregado

    sheet.values().update(
        spreadsheetId=DADOS_ID,
        range=f"'{PEDIDOS_ABA}'!E{linha_pedido}",
        valueInputOption="USER_ENTERED",
        body={"values": [[saldo_total_real]]},
    ).execute()

    return saldo_total_real, total_carregado


# ═══════════════════════════════════════════════════════════════════
# ABA DADOS — carregamentos (uma linha por carregamento)
# Colunas: PEDIDO | PRODUTO | CLIENTE | DATA | NOTA | PLACA | PESO | FRETE | STATUS
# ═══════════════════════════════════════════════════════════════════

def _inserir_carregamento(service, valores):
    """Append atômico na aba DADOS — seguro para múltiplos usuários.
    Usa RAW para evitar que datas sejam convertidas para número serial."""
    service.spreadsheets().values().append(
        spreadsheetId=DADOS_ID,
        range=f"'{DADOS_ABA}'!A1",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": [valores]},
    ).execute()


# ═══════════════════════════════════════════════════════════════════
# LEITURA — combina PEDIDOS + DADOS para montar os blocos dos cards
# ═══════════════════════════════════════════════════════════════════

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
    return len(resp.get("values", []))


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
    # append é atômico — não precisa buscar última linha

    # Data gravada como texto DD/MM/AAAA — evita interpretação incorreta pelo Sheets
    data         = str(dados_ordem.get("Data Apresentação", "")).strip()
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
    colocador    = str(dados_ordem.get("Colocador", "")).upper()
    filial_upper = filial.upper() if filial else ""

    # BASE colunas: DATA(0) FILIAL(1) PAGADOR(2) AGENCIA(3) MOTORISTA(4)
    #               PLACA(5) FABRICA(6) DESTINO(7) UF(8) PESO(9)
    #               FRETE/E(10) FRETE/M(11) ROTA(12) AGENCIAMENTO(13)
    #               STATUS(14) PEDIDO(15) PRODUTO(16) COLOCADOR(17) COLOCADOR(17)
    linha_fretes = [
        data, filial_upper, pagador, agencia, motorista,
        placa, fabrica, destino, uf, peso,
        frete_emp, frete_mot, rota, agenciamento, status, pedido, produto, colocador
    ]

    _inserir_linha_base(service, None, None, linha_fretes, aba=aba)

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
    Verifica se o pedido existe na aba PEDIDOS e insere uma linha
    de carregamento na aba DADOS.
    Retorna True se gravou, False se pedido não encontrado.

    Aba DADOS: PEDIDO | PRODUTO | CLIENTE | DATA | NOTA | PESO | FRETE | STATUS
    """
    sheet = service.spreadsheets()

    # Verifica se o pedido existe na aba PEDIDOS
    resp = sheet.values().get(
        spreadsheetId=DADOS_ID,
        range=f"'{PEDIDOS_ABA}'!A:E",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()

    pedido_upper  = _normalizar(pedido)
    produto_upper = _normalizar(produto)
    cliente_upper = _normalizar(cliente)

    encontrado = False
    prod_final = produto
    cli_final  = cliente

    for linha in resp.get("values", [])[1:]:
        if len(linha) < 3:
            continue
        ped_plan  = _normalizar(linha[2])
        cli_plan  = _normalizar(linha[1]) if len(linha) > 1 else ""
        prod_plan = _normalizar(linha[3]) if len(linha) > 3 else ""

        if ped_plan != pedido_upper:
            continue
        if produto_upper and prod_plan and not _palavras_em_comum(produto_upper, prod_plan):
            continue
        if cliente_upper and cli_plan and not _palavras_em_comum(cliente_upper, cli_plan):
            continue

        # Usa os valores exatos da planilha para consistência
        prod_final = str(linha[3]).strip() if len(linha) > 3 else produto
        cli_final  = str(linha[1]).strip() if len(linha) > 1 else cliente
        encontrado = True
        break

    if not encontrado:
        return False

    # Insere carregamento na aba DADOS
    # Colunas: PEDIDO | PRODUTO | CLIENTE | DATA | NOTA | PLACA | PESO | FRETE | STATUS
    valores_novo = [
        pedido,
        prod_final,
        cli_final,
        data,
        "",                    # NOTA — preenchida manualmente na planilha
        placa,
        str(peso_num),
        frete,
        status or "CARREGADO",
    ]
    _inserir_carregamento(service, valores_novo)
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
    """Usa values().append — atômico e seguro para múltiplos usuários simultâneos.
    Usa RAW para evitar que datas DD/MM/AAAA sejam convertidas para número serial."""
    aba_usar = aba or BASE_ABA
    service.spreadsheets().values().append(
        spreadsheetId=BASE_ID,
        range=f"'{aba_usar}'!A1",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": [valores]},
    ).execute()