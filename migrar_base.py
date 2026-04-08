"""
Migração BASE 04/2026 → Supabase
Execute: python migrar_base.py
Requer: credentials.json na mesma pasta
"""

import json
import re
import datetime
import urllib.request
import urllib.error
from pathlib import Path

# ── Configuração ───────────────────────────────────────────────────
BASE_ID       = "1xab_LceMGpjIhXKDp1iUp3OUjqk1Z3SOz05vScFqTJg"
ABA           = "BASE 04/2026"
SUPABASE_URL  = "https://xlirwzkmvkzldrssmhxg.supabase.co"
SUPABASE_KEY  = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InhsaXJ3emttdmt6bGRyc3NtaHhnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzU2MDgwMTksImV4cCI6MjA5MTE4NDAxOX0.ofTAEn628a-7JzF3REPj-tBcQJUrlXdfaFSbU5Ysfx4"
CREDENTIALS   = "credentials.json"

# BASE colunas por posição:
# DATA(0) FILIAL(1) PAGADOR(2) AGENCIA(3) MOTORISTA(4) PLACA(5)
# FABRICA(6) DESTINO(7) UF(8) PESO(9) FRETE/E(10) FRETE/M(11)
# ROTA(12) AGENCIAMENTO(13) STATUS(14) PEDIDO(15) PRODUTO(16)
# EMBALAGEM(17) COLOCADOR(18) PAGAMENTO(19)

# ── Google Sheets ──────────────────────────────────────────────────
def _sheets():
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    token_path = Path("token_migracao.json")
    creds = None

    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS, SCOPES)
            creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("sheets", "v4", credentials=creds)

def _ler_aba(svc, aba):
    r = svc.spreadsheets().values().get(
        spreadsheetId=BASE_ID,
        range=f"'{aba}'",
        valueRenderOption="FORMATTED_VALUE"
    ).execute()
    return r.get("values", [])

# ── Parsing ────────────────────────────────────────────────────────
def _parse_data(v):
    s = str(v or "").strip()
    s = re.sub(r"(\d{2}/\d{2}/\d{4})\d+", r"\1", s)
    formatos = [
        ("%d/%m/%Y %H:%M:%S", 19), ("%d/%m/%Y", 10),
        ("%Y/%m/%d %H:%M:%S", 19), ("%Y/%m/%d", 10),
        ("%Y-%m-%d %H:%M:%S", 19), ("%Y-%m-%d", 10),
    ]
    for fmt, tam in formatos:
        try:
            return datetime.datetime.strptime(s[:tam], fmt).strftime("%Y-%m-%d")
        except:
            pass
    return None

def _sfloat(v):
    try:
        s = str(v or "0").strip().replace(" ", "")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", ".")
        return float(s)
    except:
        return 0.0

# ── Supabase ───────────────────────────────────────────────────────
def _inserir(registro):
    body = json.dumps(registro).encode("utf-8")
    req  = urllib.request.Request(
        f"{SUPABASE_URL}/rest/v1/carregamentos",
        data=body,
        headers={
            "apikey":        SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
            "Content-Type":  "application/json",
            "Prefer":        "return=minimal",
        },
        method="POST"
    )
    urllib.request.urlopen(req, timeout=10)

# ── Migração ───────────────────────────────────────────────────────
def migrar():
    print(f"Conectando ao Google Sheets...")
    svc  = _sheets()
    rows = _ler_aba(svc, ABA)
    print(f"Aba '{ABA}' lida: {len(rows)} linhas brutas")

    # Encontrar cabeçalho
    header_idx = None
    for i, row in enumerate(rows):
        if row and str(row[0]).strip().upper() == "DATA":
            header_idx = i
            break

    if header_idx is None:
        print("ERRO: cabeçalho 'DATA' não encontrado na aba.")
        return

    data_rows = rows[header_idx + 1:]
    print(f"Linhas de dados: {len(data_rows)}")

    inseridos  = 0
    ignorados  = 0
    erros      = 0

    for i, row in enumerate(data_rows):
        while len(row) < 20:
            row.append("")

        data = _parse_data(row[0])
        peso = _sfloat(row[9])

        if not data:
            ignorados += 1
            continue

        registro = {
            "data":         data,
            "filial":       str(row[1]).strip().upper(),
            "pagador":      str(row[2]).strip().upper(),
            "motorista":    str(row[4]).strip().upper(),
            "placa":        str(row[5]).strip().upper(),
            "fabrica":      str(row[6]).strip().upper(),
            "destino":      str(row[7]).strip().upper(),
            "uf":           str(row[8]).strip().upper(),
            "peso":         peso,
            "frete_emp":    _sfloat(row[10]),
            "frete_mot":    _sfloat(row[11]),
            "rota":         str(row[12]).strip().upper(),
            "agenciamento": str(row[13]).strip().upper(),
            "status":       str(row[14]).strip().upper(),
            "pedido":       str(row[15]).strip().upper(),
            "produto":      str(row[16]).strip().upper(),
            "embalagem":    str(row[17]).strip().upper(),
            "colocador":    str(row[18]).strip().upper(),
            "pagamento":    str(row[19]).strip().upper(),
            "usuario":      "MIGRACAO",
        }

        # Remove campos vazios
        registro = {k: v for k, v in registro.items()
                    if v not in (None, "", 0, 0.0)}

        try:
            _inserir(registro)
            inseridos += 1
            print(f"  [{i+1}/{len(data_rows)}] OK — {data} | {registro.get('placa','?')} | {registro.get('pedido','?')}")
        except Exception as e:
            erros += 1
            print(f"  [{i+1}/{len(data_rows)}] ERRO — {e}")

    print(f"\n{'='*50}")
    print(f"Migração concluída!")
    print(f"  Inseridos:  {inseridos}")
    print(f"  Ignorados:  {ignorados} (sem data válida)")
    print(f"  Erros:      {erros}")
    print(f"{'='*50}")

if __name__ == "__main__":
    migrar()
