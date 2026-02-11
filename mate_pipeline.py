# PARTE 1A ========================================================================================
# Entrada → resolve pdf_path e resolve:
#   - diario_key (chave estável do diário, usada no fluxo inteiro)
#   - aba (DD/MM/YYYY) e aba_yyyymmdd (YYYYMMDD) quando aplicável
# ================================================================================================

import os
import re
import hashlib
import unicodedata
from datetime import datetime, timedelta, date
from zoneinfo import ZoneInfo
from functools import lru_cache
from pathlib import Path
from urllib.parse import urlparse

from pypdf import PdfReader  # noqa: F401  (usado nas partes seguintes)

URL_BASE = "https://diariolegislativo.almg.gov.br"
TZ_BR = ZoneInfo("America/Sao_Paulo")

# ---- Colab? ----
try:
    from google.colab import files  # type: ignore
    _COLAB = True
except Exception:
    _COLAB = False


# --------------------------------------------------------------------------------
# NÃO-EXPEDIENTE (FERIADOS + RECESSOS)
# --------------------------------------------------------------------------------
def _intervalo_datas(inicio: date, fim: date) -> set[date]:
    out: set[date] = set()
    d = inicio
    while d <= fim:
        out.add(d)
        d += timedelta(days=1)
    return out


NAO_EXPEDIENTE_2025: set[date] = {
    date(2025, 1, 1),
    date(2025, 3, 3), date(2025, 3, 4), date(2025, 3, 5),
    date(2025, 4, 17), date(2025, 4, 18), date(2025, 4, 21),
    date(2025, 5, 1), date(2025, 5, 2),
    date(2025, 6, 19), date(2025, 6, 20),
    date(2025, 8, 15),
    date(2025, 9, 7),
    date(2025, 10, 12), date(2025, 10, 27),
    date(2025, 11, 2), date(2025, 11, 15), date(2025, 11, 20), date(2025, 11, 21),
    date(2025, 12, 8), date(2025, 12, 24), date(2025, 12, 25), date(2025, 12, 26), date(2025, 12, 31),
}

NAO_EXPEDIENTE_2026: set[date] = {
    date(2026, 1, 1), date(2026, 2, 17), date(2026, 6, 4),
    date(2026, 9, 7), date(2026, 10, 12),
    date(2026, 11, 2), date(2026, 11, 15), date(2026, 11, 20),
    date(2026, 12, 25),
}
NAO_EXPEDIENTE_2026 |= {
    date(2026, 2, 18), date(2026, 4, 2), date(2026, 4, 3), date(2026, 6, 5),
}
NAO_EXPEDIENTE_2026 |= _intervalo_datas(date(2026, 12, 7), date(2026, 12, 31))

NAO_EXPEDIENTE_POR_ANO = {
    2025: NAO_EXPEDIENTE_2025,
    2026: NAO_EXPEDIENTE_2026,
}


def proximo_dia_util(yyyymmdd: str) -> str:
    d = datetime.strptime(yyyymmdd, "%Y%m%d").date()
    nao = NAO_EXPEDIENTE_POR_ANO.get(d.year, set())
    while d.weekday() >= 5 or d in nao:
        d += timedelta(days=1)
    return d.strftime("%Y%m%d")


def yyyymmdd_to_ddmmyyyy(yyyymmdd: str) -> str:
    return f"{yyyymmdd[6:8]}/{yyyymmdd[4:6]}/{yyyymmdd[0:4]}"


def normalizar_data(entrada: str) -> str:
    s = (entrada or "").strip().lower()

    if s in ("hoje", "ontem", "anteontem"):
        base = datetime.now(TZ_BR)
        if s == "ontem":
            base -= timedelta(days=1)
        elif s == "anteontem":
            base -= timedelta(days=2)
        return base.strftime("%Y%m%d")

    weekday_map = {
        "segunda": 0, "terça": 1, "terca": 1,
        "quarta": 2, "quinta": 3, "sexta": 4,
        "sábado": 5, "sabado": 5,
    }
    if s in weekday_map:
        today = datetime.now(TZ_BR)
        delta = (today.weekday() - weekday_map[s]) % 7 or 7
        return (today - timedelta(days=delta)).strftime("%Y%m%d")

    digits = "".join(c for c in s if c.isdigit())
    if len(digits) == 4:
        yyyy = datetime.now(TZ_BR).year
        return f"{yyyy:04d}{digits[2:4]}{digits[0:2]}"
    if len(digits) == 6:
        return f"20{digits[4:6]}{digits[2:4]}{digits[0:2]}"
    if len(digits) == 8:
        if digits.startswith(("19", "20")):
            datetime.strptime(digits, "%Y%m%d")
            return digits
        return f"{digits[4:8]}{digits[2:4]}{digits[0:2]}"

    raise ValueError("Data inválida.")


def _sha1_bytes(data: bytes) -> str:
    h = hashlib.sha1()
    h.update(data)
    return h.hexdigest()


def _sha1_file(path: str, max_bytes: int = 8 * 1024 * 1024) -> str:
    h = hashlib.sha1()
    with open(path, "rb") as f:
        remain = max_bytes
        while remain > 0:
            chunk = f.read(min(1024 * 1024, remain))
            if not chunk:
                break
            h.update(chunk)
            remain -= len(chunk)
    return h.hexdigest()


def _cache_dir() -> Path:
    # Colab: /content é garantido; fora do Colab: usa diretório atual.
    base = Path("/content") if _COLAB else Path.cwd()
    d = base / ".cache_diario_legislativo"
    d.mkdir(parents=True, exist_ok=True)
    return d


def baixar_pdf_por_url(url: str) -> str | None:
    """
    Baixa PDF com cache (por URL). Retorna caminho local, ou None se falhar/arquivo não for PDF.
    """
    import urllib.request

    cache_key = _sha1_bytes(url.encode("utf-8"))[:16]
    local = _cache_dir() / f"dl_{cache_key}.pdf"

    if local.exists() and local.stat().st_size > 0:
        try:
            with open(local, "rb") as f:
                if f.read(5) == b"%PDF-":
                    return str(local)
        except Exception:
            pass
        try:
            local.unlink(missing_ok=True)
        except Exception:
            pass

    try:
        req = urllib.request.Request(url, headers={"User-Agent": "mate.ia/1.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
        if not data.startswith(b"%PDF-"):
            return None
        with open(local, "wb") as f:
            f.write(data)
        return str(local)
    except Exception:
        return None


def _extrair_yyyymmdd_da_url(url: str) -> str | None:
    """
    Tenta extrair YYYYMMDD de URLs no padrão .../YYYY/LYYYYMMDD.pdf
    """
    path = urlparse(url).path or ""
    m = re.search(r"/(\d{4})/L(\d{8})\.pdf$", path)
    if not m:
        return None
    yyyy = m.group(1)
    yyyymmdd = m.group(2)
    if yyyymmdd.startswith(yyyy):
        return yyyymmdd
    return None


# ================================================================================================
# BLOCO DE ENTRADA
# ================================================================================================

entrada = input("Data/URL/caminho: ").strip()

pdf_path: str | None = None
aba_yyyymmdd: str | None = None
aba: str | None = None
diario_key: str | None = None  # <- GARANTIDO ao final

ABA_POLICY = "ASK"

if not entrada:
    if not _COLAB:
        raise SystemExit("Entrada vazia.")
    up = files.upload()
    pdf_path = next(iter(up.keys()))
    diario_key = _sha1_file(pdf_path)[:16]

elif entrada.lower().startswith(("http://", "https://")):
    src_yyyymmdd = _extrair_yyyymmdd_da_url(entrada)
    pdf_path = baixar_pdf_por_url(entrada)
    if not pdf_path:
        raise SystemExit("DL inexistente.")
    diario_key = (src_yyyymmdd or _sha1_bytes(entrada.encode("utf-8"))[:16])

    if src_yyyymmdd:
        aba_yyyymmdd = proximo_dia_util(src_yyyymmdd)
        aba = yyyymmdd_to_ddmmyyyy(aba_yyyymmdd)

elif "/" in entrada or "\\" in entrada:
    if not os.path.exists(entrada):
        raise SystemExit("Arquivo não encontrado.")
    pdf_path = entrada
    diario_key = _sha1_file(pdf_path)[:16]

else:
    yyyymmdd = normalizar_data(entrada)
    aba_yyyymmdd = proximo_dia_util(yyyymmdd)
    aba = yyyymmdd_to_ddmmyyyy(aba_yyyymmdd)
    url = f"{URL_BASE}/{yyyymmdd[:4]}/L{yyyymmdd}.pdf"
    pdf_path = baixar_pdf_por_url(url)
    if not pdf_path:
        raise SystemExit("DL inexistente.")
    diario_key = yyyymmdd  # chave estável por data informada

if not pdf_path or not os.path.exists(pdf_path):
    raise FileNotFoundError("PDF não encontrado.")

if not diario_key:
    # fallback duro (não deveria acontecer)
    diario_key = _sha1_file(pdf_path)[:16]

if not aba:
    # Se não veio de DATA/URL padrão, usa a data local atual só para nome de aba (não afeta diario_key).
    aba_yyyymmdd = datetime.now(TZ_BR).strftime("%Y%m%d")
    aba = yyyymmdd_to_ddmmyyyy(aba_yyyymmdd)

print("diario_key:", diario_key)
print("Aba FINAL (Sheets):", aba)
print("Política de aba:", ABA_POLICY)

# ================================================================================================
# DEFS UTILITÁRIAS GLOBAIS (USADAS DEPOIS)
# ================================================================================================

RE_PAG = re.compile(r"\bP[ÁA]GINA\s+(\d{1,4})\b", re.IGNORECASE)

RE_HEADER_LIXO = re.compile(
    r"(DI[ÁA]RIO\s+DO\s+LEGISLATIVO|www\.almg\.gov\.br|"
    r"segunda-feira|terça-feira|quarta-feira|quinta-feira|sexta-feira|"
    r"s[áa]bado|domingo)",
    re.IGNORECASE
)


def limpa_linha(s: str) -> str:
    s = s.replace("\u00a0", " ")
    return re.sub(r"[ \t]+", " ", s).strip()


def primeira_pagina_num(linhas: list[str], fallback: int) -> int:
    for ln in linhas[:220]:
        m = RE_PAG.search(ln)
        if m:
            return int(m.group(1))
    return fallback


@lru_cache(maxsize=20000)
def compact_key(s: str) -> str:
    u = s.upper()
    u = unicodedata.normalize("NFD", u)
    u = "".join(c for c in u if unicodedata.category(c) != "Mn")
    return "".join(c for c in u if c.isalnum())


def _linha_relevante(s: str) -> bool:
    s = limpa_linha(s)
    if not s:
        return False
    if RE_HEADER_LIXO.search(s):
        return False
    if re.fullmatch(r"[-–—_•\.\s]+", s):
        return False
    return True


def is_top_event(line_idx: int, linhas: list[str]) -> bool:
    for prev in linhas[:line_idx]:
        if _linha_relevante(prev):
            return False
    return True


def win_keys(linhas: list[str], i: int, w: int) -> str:
    parts = []
    for k in range(w):
        j = i + k
        if j < len(linhas):
            parts.append(compact_key(linhas[j]))
    return "".join(parts)


def win_any_in(linhas: list[str], i: int, keys: set[str]) -> bool:
    return any(
        win_keys(linhas, i, w) in keys
        for w in (1, 2, 3)
    )


def _checkbox_req(
    sheet_id: int,
    col_idx_0based: int,
    row_1based: int,
    default_checked: bool = False,
):
    """
    Cria checkbox (data validation BOOLEAN) e define o valor padrão (TRUE/FALSE).
    Retorna uma lista de requests para batch_update.
    """
    val = {"boolValue": True} if default_checked else {"boolValue": False}

    dv = {
        "setDataValidation": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row_1based - 1,
                "endRowIndex": row_1based,
                "startColumnIndex": col_idx_0based,
                "endColumnIndex": col_idx_0based + 1,
            },
            "rule": {
                "condition": {"type": "BOOLEAN"},
                "strict": True,
                "showCustomUi": True,
            },
        }
    }

    setv = {
        "updateCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row_1based - 1,
                "endRowIndex": row_1based,
                "startColumnIndex": col_idx_0based,
                "endColumnIndex": col_idx_0based + 1,
            },
            "rows": [
                {
                    "values": [
                        {"userEnteredValue": val}
                    ]
                }
            ],
            "fields": "userEnteredValue",
        }
    }

    return [dv, setv]


def _cf_fontsize_req(sheet_id: int, col0: int, row1: int, font_size: int, formula: str, index: int = 0):
    return {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{
                    "sheetId": sheet_id,
                    "startRowIndex": row1 - 1,
                    "endRowIndex": row1,
                    "startColumnIndex": col0,
                    "endColumnIndex": col0 + 1,
                }],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": formula}],
                    },
                    "format": {"textFormat": {"fontSize": font_size}},
                },
            },
            "index": index,
        }
    }
# PARTE 1B ===============================================================================================================================
# =============================================== 5) GOOGLE SHEETS (layout + dados) =====================================================
# ======================================================================================================================================

import time
import random
import gspread
from google.colab import auth
from google.auth import default

# --------------------------------------------------------------------------------------------------
# AUTH / CLIENT
# --------------------------------------------------------------------------------------------------
auth.authenticate_user()
creds, _ = default(scopes=["https://www.googleapis.com/auth/spreadsheets"])
gc = gspread.authorize(creds)

SHEET_ID: int | None = None

# --------------------------------------------------------------------------------------------------
# UTIL BÁSICOS
# --------------------------------------------------------------------------------------------------
def yyyymmdd_to_ddmmyyyy(yyyymmdd: str) -> str:
    return f"{yyyymmdd[6:8]}/{yyyymmdd[4:6]}/{yyyymmdd[0:4]}"


def rgb_hex_to_api(hex_str: str) -> dict:
    h = hex_str.lstrip("#")
    return {
        "red": int(h[0:2], 16) / 255.0,
        "green": int(h[2:4], 16) / 255.0,
        "blue": int(h[4:6], 16) / 255.0,
    }


def a1_to_grid(a1: str) -> dict:
    a1 = a1.strip()
    if ":" not in a1:
        a1 = f"{a1}:{a1}"
    return gspread.utils.a1_range_to_grid_range(a1)


def _with_backoff(fn, *args, **kwargs):
    for attempt in range(8):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            msg = str(e)
            if any(x in msg for x in ("429", "Quota", "Rate", "503")):
                sleep_s = min(60, (2 ** attempt) + random.random())
                time.sleep(sleep_s)
                continue
            raise


# --------------------------------------------------------------------------------------------------
# REQUEST BUILDERS (mínimo necessário)
# --------------------------------------------------------------------------------------------------
def req_dim_rows(sheet_id: int, start: int, end: int, px: int) -> dict:
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": start, "endIndex": end},
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }


def req_dim_cols(sheet_id: int, start: int, end: int, px: int) -> dict:
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": start, "endIndex": end},
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }


def req_unmerge(sheet_id: int, a1: str) -> dict:
    gr = a1_to_grid(a1)
    return {"unmergeCells": {"range": {"sheetId": sheet_id, **gr}}}


def req_merge(sheet_id: int, a1: str) -> dict:
    gr = a1_to_grid(a1)
    return {"mergeCells": {"range": {"sheetId": sheet_id, **gr}, "mergeType": "MERGE_ALL"}}


def _checkbox_req(sheet_id: int, col0: int, row1: int, default_checked: bool = False) -> list[dict]:
    val = {"boolValue": True} if default_checked else {"boolValue": False}
    return [
        {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row1 - 1,
                    "endRowIndex": row1,
                    "startColumnIndex": col0,
                    "endColumnIndex": col0 + 1,
                },
                "rule": {
                    "condition": {"type": "BOOLEAN"},
                    "strict": True,
                    "showCustomUi": True,
                },
            }
        },
        {
            "updateCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row1 - 1,
                    "endRowIndex": row1,
                    "startColumnIndex": col0,
                    "endColumnIndex": col0 + 1,
                },
                "rows": [{"values": [{"userEnteredValue": val}]}],
                "fields": "userEnteredValue",
            }
        },
    ]


# --------------------------------------------------------------------------------------------------
# CONSTANTES DE LAYOUT (iguais às existentes, sem inventar nada)
# --------------------------------------------------------------------------------------------------
COL_DEFAULT = 60
COL_OVERRIDES = {
    0: 23, 1: 60, 2: 370, 3: 75, 4: 85, 5: 75, 6: 75,
    7: 45, 8: 45, 9: 45, 10: 45, 11: 45, 12: 45, 13: 45,
    14: 45, 15: 60, 16: 75, 17: 70, 18: 70, 19: 60,
    20: 60, 21: 60, 22: 60, 23: 60, 24: 60
}

ROW_HEIGHTS = [
    ("default", 16),
    (0, 4, 14),
    (4, 5, 25),
]

MERGES = [
    "A1:B4", "C1:F4", "G1:G4", "Q1:Y1",
    "A5:B5", "E5:F5", "G5:I5", "T5:Y5",
    "E6:G6", "E8:G8",
]

# --------------------------------------------------------------------------------------------------
# FUNÇÃO PRINCIPAL
# --------------------------------------------------------------------------------------------------
def upsert_tab_diario(
    spreadsheet_url_or_id: str,
    diario_key: str,              # YYYYMMDD (já garantido na PARTE 1A)
    itens: list[tuple[str, str]],
):
    global SHEET_ID

    tab_name = yyyymmdd_to_ddmmyyyy(diario_key)

    sh = (
        gc.open_by_url(spreadsheet_url_or_id)
        if spreadsheet_url_or_id.startswith("http")
        else gc.open_by_key(spreadsheet_url_or_id)
    )

    try:
        ws = sh.worksheet(tab_name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=tab_name, rows=50, cols=25)
        _with_backoff(ws.update_index, 1)

    SHEET_ID = ws.id
    sheet_id = ws.id

    # ----------------------------------------------------------------------------------------------
    # DIMENSÕES (cálculo ÚNICO e determinístico)
    # ----------------------------------------------------------------------------------------------
    base_rows = 9
    footer_rows = 9
    rows_needed = base_rows + len(itens) + footer_rows
    cols_needed = 25

    rows_target = max(ws.row_count, rows_needed + 1, 22)
    cols_target = max(ws.col_count, cols_needed)

    _with_backoff(ws.resize, rows=rows_target, cols=cols_target)

    VIS_LAST_ROW_1BASED = rows_target - 1

    # ----------------------------------------------------------------------------------------------
    # REQUESTS
    # ----------------------------------------------------------------------------------------------
    reqs: list[dict] = []

    # congela cabeçalho
    reqs.append({
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 5}},
            "fields": "gridProperties.frozenRowCount",
        }
    })

    # alturas
    for rh in ROW_HEIGHTS:
        if rh[0] == "default":
            reqs.append(req_dim_rows(sheet_id, 0, rows_target, rh[1]))
        else:
            start, end, px = rh
            reqs.append(req_dim_rows(sheet_id, start, end, px))

    # linha técnica (1px)
    reqs.append(req_dim_rows(sheet_id, rows_target - 1, rows_target, 1))

    # larguras
    reqs.append(req_dim_cols(sheet_id, 0, cols_target, COL_DEFAULT))
    for c, px in COL_OVERRIDES.items():
        reqs.append(req_dim_cols(sheet_id, c, c + 1, px))

    # limpa merges antigos
    reqs.append({
        "unmergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": 0,
                "endRowIndex": rows_target,
                "startColumnIndex": 0,
                "endColumnIndex": cols_target,
            }
        }
    })

    # merges fixos
    for a1 in MERGES:
        reqs.append(req_merge(sheet_id, a1))

    # checkboxes nos itens (coluna I)
    start_items_row = 9
    end_items_row = start_items_row + len(itens) - 1
    if end_items_row >= start_items_row:
        for r in range(start_items_row, end_items_row + 1):
            for rq in _checkbox_req(sheet_id, 8, r, default_checked=False):
                reqs.append(rq)

    # ----------------------------------------------------------------------------------------------
    # EXECUÇÃO (ESTE ERA O BURACO REAL)
    # ----------------------------------------------------------------------------------------------
    _with_backoff(sh.batch_update, {"requests": reqs})

    return ws
# PARTE 2 ===============================================================================================================================
# ========================================================= FOOTER (isolado, determinístico) =============================================
# ======================================================================================================================================

def apply_footer(
    sh,
    sheet_id: int,
    ws,
    extra_end: int,          # 1-based, end EXCLUSIVO dos extras
    LISTA_DROPDOWN_5: list[str],
    LISTA_DROPDOWN_6: list[str],
    LISTA_DROPDOWN_7: list[str],
):
    """
    Aplica o FOOTER completo a partir de extra_end.
    Bloco isolado: só consome sheet_id, ws e listas já existentes.
    """

    reqs: list[dict] = []

    # ----------------------------------------------------------------------------------
    # POSIÇÕES (1-based)
    # ----------------------------------------------------------------------------------
    footer_start = extra_end
    footer_rows  = 9
    footer_end   = footer_start + footer_rows

    r  = footer_start
    r1 = r + 1
    r2 = r + 2
    r3 = r + 3
    r4 = r + 4
    r5 = r + 5
    r6 = r + 6
    r7 = r + 7
    r8 = r + 8

    # ----------------------------------------------------------------------------------
    # VALUES / FÓRMULAS (copiado fielmente, sem alteração semântica)
    # ----------------------------------------------------------------------------------
    def _cell(col0, row1, value, is_formula=False):
        return {
            "updateCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row1 - 1,
                    "endRowIndex": row1,
                    "startColumnIndex": col0,
                    "endColumnIndex": col0 + 1,
                },
                "rows": [{
                    "values": [{
                        "userEnteredValue": (
                            {"formulaValue": value} if is_formula else {"stringValue": value}
                        )
                    }]
                }],
                "fields": "userEnteredValue",
            }
        }

    # links / imagens principais
    reqs.extend([
        _cell(1, r, '=HYPERLINK("http://meet.google.com/api-pefj-mvq";"GDI-GGA")', True),

        _cell(0, r1, '=HYPERLINK("https://mediaserver.almg.gov.br/acervo/511/376/2511376.pdf";IMAGE("https://cdn-icons-png.flaticon.com/512/3079/3079014.png";4;19;19))', True),
        _cell(1, r1, '=HYPERLINK("https://intra.almg.gov.br/export/sites/default/atendimento/docs/lista-telefonica.pdf";IMAGE("https://cdn-icons-png.flaticon.com/512/4783/4783130.png";4;33;33))', True),
        _cell(2, r1, '=HYPERLINK("https://sites.google.com/view/gga-gdi-almg/";IMAGE("https://yt3.ggpht.com/ytc/AKedOLS-fgkzGxYUBgBejVblA1CLhE69pbiZyoH7spcNRQ=s900-c-k-c0x00ffffff-no-rj";4;125;150))', True),

        _cell(4, r1, '=SUM(FILTER(INDIRECT("F"&ROW()+1&":F");INDIRECT("E"&ROW()+1&":E")<>""))', True),
        _cell(5, r1, "TOTAL"),
        _cell(6, r1, "IMPLANTAÇÃO"),
        _cell(8, r1, "CONFERÊNCIA"),
        _cell(12, r1, "PROPOSIÇÕES RELEVANTES"),
    ])

    # ----------------------------------------------------------------------------------
    # MERGES (aplicação direta, sem lógica extra)
    # ----------------------------------------------------------------------------------
    def _merge(r0, r1, c0, c1):
        return {
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": r0 - 1,
                    "endRowIndex": r1,
                    "startColumnIndex": c0,
                    "endColumnIndex": c1,
                },
                "mergeType": "MERGE_ALL",
            }
        }

    reqs.extend([
        _merge(r1, r2, 0, 1),
        _merge(r1, r2, 1, 2),
        _merge(r3, r5, 0, 2),
        _merge(r6, r8, 0, 1),
        _merge(r6, r8, 1, 2),
        _merge(r1, r8, 2, 4),

        _merge(r1, r1, 8, 10),
        _merge(r1, r1, 12, 15),

        _merge(r1, r8, 15, 17),
        _merge(r1, r8, 17, 19),
        _merge(r1, r8, 19, 21),
        _merge(r1, r8, 21, 23),
        _merge(r1, r8, 23, 25),

        _merge(r2, r2, 8, 10),
        _merge(r3, r3, 8, 10),
        _merge(r4, r4, 8, 10),
        _merge(r5, r5, 8, 10),
        _merge(r6, r6, 8, 10),
        _merge(r7, r7, 8, 10),
        _merge(r8, r8, 8, 10),
    ])

    # ----------------------------------------------------------------------------------
    # DATA VALIDATION (listas)
    # ----------------------------------------------------------------------------------
    def _dv(col0, r0, r1, values):
        return {
            "setDataValidation": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": r0 - 1,
                    "endRowIndex": r1,
                    "startColumnIndex": col0,
                    "endColumnIndex": col0 + 1,
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [{"userEnteredValue": v} for v in values],
                    },
                    "strict": False,
                    "showCustomUi": True,
                },
            }
        }

    reqs.extend([
        _dv(4, r1, r6, LISTA_DROPDOWN_5),
        _dv(6, r1, r6, LISTA_DROPDOWN_5),
        _dv(8, r1, r6, LISTA_DROPDOWN_5),
        _dv(12, r1, r8, LISTA_DROPDOWN_6),
        _dv(14, r1, r8, LISTA_DROPDOWN_7),
    ])

    # ----------------------------------------------------------------------------------
    # EXECUÇÃO
    # ----------------------------------------------------------------------------------
    sh.batch_update({"requests": reqs})
# PARTE 3 ============================================================================================================================================================================================
# ============================================================================================= VALUES ===============================================================================================
# ====================================================================================================================================================================================================

# OBS: versão reescrita para ser funcionalmente equivalente ao bloco que você enviou.
# - Mantém as MESMAS fórmulas/valores e ranges.
# - Mantém a mesma estratégia de updates (values_batch_update + batch_update).
# - Evita NameError de `with_backoff` (alias para `_with_backoff`).
# - Mantém a sanitização de requests antes do batch_update.

dd = int(diario_key[6:8])
mm = int(diario_key[4:6])
yyyy = int(diario_key[0:4])
a5_txt = f"{dd}/{mm}"

from datetime import datetime, timedelta

# Alias defensivo: no seu trecho aparece `with_backoff(...)` ao final.
# Para manter equivalência (e evitar NameError), mapeio para `_with_backoff`.
with_backoff = _with_backoff

data = []

def add(a1, values):
    data.append({"range": f"'{tab_name}'!{a1}", "values": values})

def two_business_days_before(d):
    n = 2
    while n > 0:
        d -= timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d

add("A5:B5", [[f"=DATE({yyyy};{mm};{dd})", ""]])
add("A1", [[ '=HYPERLINK("https://www.almg.gov.br/home/index.html";IMAGE("https://sisap.almg.gov.br/banner.png";4;43;110))' ]])
add("C1", [["GERÊNCIA DE GESTÃO ARQUIVÍSTICA"]])
add("Q1", [["DATAS"]])
add("G1", [['=HYPERLINK("https://intra.almg.gov.br/export/sites/default/a-assembleia/calendarios/calendario_2023.pdf";'
    'IMAGE("https://media.istockphoto.com/vectors/flag-map-of-the-brazilian-state-of-minas-gerais-vector-id1248541649?k=20&m=1248541649&s=170667a&w=0&h=V8Ky8c8rddLPjphovytIJXaB6NlMF7dt-ty-2ZJF5Wc="))']])
add("H1", [['=HYPERLINK("https://www.almg.gov.br/atividade_parlamentar/plenario/index.html";''IMAGE("https://www.protestoma.com.br/images/noticia-id_255.jpg";4;27;42))']])
add("H3", [['=HYPERLINK("https://www.almg.gov.br/atividade-parlamentar/comissoes/agenda/";''IMAGE("https://www.ouvidoriageral.mg.gov.br/images/noticias/2019/dezembro/foto_almg.jpg";4;27;42))']])
add("I1", [['=HYPERLINK("https://www.jornalminasgerais.mg.gov.br/";'
    'IMAGE("https://upload.wikimedia.org/wikipedia/commons/thumb/f/f4/Bandeira_de_Minas_Gerais.svg/2560px-Bandeira_de_Minas_Gerais.svg.png";4;35;50))']])
add("I3", [['=HYPERLINK("https://www.almg.gov.br/consulte/arquivo_diario_legislativo/index.html";'
    'IMAGE("https://www.almg.gov.br/favicon.ico";4;25;25))']])
add("J1", [['=HYPERLINK("https://consulta-brs.almg.gov.br/brs/";''IMAGE("https://t4.ftcdn.net/jpg/04/70/40/23/360_F_470402339_5FVE7b1Z2DNI7bATV5a27FGATt6yxcEz.jpg"))']])
add("J3", [['=HYPERLINK("https://silegis.almg.gov.br/silegismg/login/login.jsp";IMAGE("https://silegis.almg.gov.br/silegismg/assets/logotipo.png"))']])
add("K1", [[ '=HYPERLINK("https://webmail.almg.gov.br/";IMAGE("https://images.vexels.com/media/users/3/140138/isolated/lists/88e50689fa3280c748d000aaf0bad480-icone-redondo-de-email-1.png"))' ]])
add("K3", [[ '=HYPERLINK("https://sites.google.com/view/gga-gdi-almg/manuais-e-delibera%C3%A7%C3%B5es#h.no8oprc5oego";IMAGE("http://anthillonline.com/wp-content/uploads/2021/03/mate-logo.jpg";4;65;50))' ]])
add("L1", [[ '=HYPERLINK("https://www.almg.gov.br/atividade-parlamentar/projetos-de-lei/";IMAGE("https://upload.wikimedia.org/wikipedia/commons/thumb/a/a6/Tram-Logo.svg/2048px-Tram-Logo.svg.png";4;23;23))' ]])
add("L3", [[ '=HYPERLINK("https://www.almg.gov.br/consulte/legislacao/index.html";IMAGE("https://cdn-icons-png.flaticon.com/512/3122/3122427.png"))' ]])
add("M1", [[ '=HYPERLINK("https://sei.almg.gov.br/";IMAGE("https://www.gov.br/ebserh/pt-br/media/plataformas/sei/@@images/5a07de59-2af0-45b0-9be9-f0d0438b7a81.png";4;45;50))' ]])
add("M3", [[ '=HYPERLINK("https://stl.almg.gov.br/login.jsp";IMAGE("https://media-exp1.licdn.com/dms/image/C510BAQHc4JZB3kDHoQ/company-logo_200_200/0/1519865605418?e=2147483647&v=beta&t=dE29KDkLy-qxYmZ3TVE95zPf8_PeoMr7YJBQehJbFg8";4;24;28))' ]])
add("N1", [[ '=HYPERLINK("https://docs.google.com/spreadsheets/d/1kJmtsWxoMtBKeMeO0Aex4IrIULRMeyf6yl3UgqatNGs/edit#gid=1276994968";IMAGE("https://cdn-icons-png.flaticon.com/512/3767/3767084.png";4;23;23))' ]])
add("N3", [[ '=HYPERLINK("https://webdrive.almg.gov.br/index.php/login";IMAGE("https://upload.wikimedia.org/wikipedia/en/6/61/WebDrive.png";4;22;22))' ]])
add("O1", [[ '=HYPERLINK("https://www.youtube.com/c/assembleiamg";IMAGE("https://cdn.pixabay.com/photo/2021/02/16/06/00/youtube-logo-6019878_960_720.png";4;20;28))' ]])
add("O3", [[ '=HYPERLINK("https://atom.almg.gov.br/index.php/";IMAGE("https://atom.almg.gov.br/favicon.ico";4;20;20))' ]])
add("P1", [[ '=IMAGE("https://img2.gratispng.com/20180422/slw/kisspng-computer-icons-dice-desktop-wallpaper-clip-art-5adc2023a35a45.9466329215243755876691.jpg")' ]])
add("P2", [["LEGISLATIVO"]])
add("P3", [["ATUAL"]])
add("P4", [["ATA"]])
add("D5", [["#"]])
add("E5", [["IMPLANTAÇÃO"]])
add("J5", [["PROPOSIÇÕES"]])
add("T5", [["EXPRESSÕES DE BUSCA"]])
add("C5", [[ '=HYPERLINK("https://docs.google.com/document/d/1lftfl3SAfJPMdIKYSjATffe-Tvc9qfoLodfGK-f3sLU/edit";"MATE - MATÉRIAS EM TRAMITAÇÃO")' ]])
add("G5", [[ '=HYPERLINK("https://writer.zoho.com/writer/open/fgoh367779094842247dd8313f9c7714f452a";"CONFERÊNCIA")' ]])
add("B6", [['=TEXT(A5;"dd/mm/yyyy")']])
add("C6", [["DIÁRIO DO EXECUTIVO"]])
add("B7", [["-"]])

dl_date = datetime(yyyy, mm, dd).date()
dmenos2_date = two_business_days_before(dl_date)
dmenos2 = f"{dmenos2_date.day}/{dmenos2_date.month}/{dmenos2_date.year}"
add("E8:G8", [[dmenos2]])

# ------------------------------------------------------------------
# A6 (mega fórmula) - preservada integralmente
# ------------------------------------------------------------------
add(
    f"A6:A{footer_start - 1}",
    [[r'''=IFS(

  OR(
  INDIRECT("C"&ROW())="-";
  INDIRECT("C"&ROW())="?")
  ;"-";

  OR(INDIRECT("U"&ROW())<>"TOTAL");
  IFS(
  OR(INDIRECT("C"&ROW())="";INDIRECT("C"&ROW())="IMPLANTAÇÃO DE TEXTOS";INDIRECT("U"&ROW())="IMPLANTAÇÃO");"";

  OR(INDIRECT("C"&ROW())="DIÁRIO DO EXECUTIVO";INDIRECT("C"&ROW())="LEIS";INDIRECT("C"&ROW())="LEI, COM PROPOSIÇÃO ANEXADA";LEFT(INDIRECT("C"&ROW());4)="VETO");
    HYPERLINK(
    "https://www.jornalminasgerais.mg.gov.br/edicao-do-dia?dados=" &
    ENCODEURL("{""dataPublicacaoSelecionada"":""" & TEXT($B$6;"yyyy-mm-dd") & "T03:00:00.000Z""}");
    IMAGE("https://www.almg.gov.br/favicon.ico";4;15;15)
    );

  INDIRECT("C"&ROW())="DIÁRIO DO EXECUTIVO - EDIÇÃO EXTRA";
    HYPERLINK(
    "https://www.jornalminasgerais.mg.gov.br/edicao-do-dia?dados=" &
    ENCODEURL("{""dataPublicacaoSelecionada"":""" & TEXT($B$6;"yyyy-mm-dd") & "T03:00:00.000Z""}");
    IMAGE("https://www.almg.gov.br/favicon.ico";4;15;15)
    );

  INDIRECT("C"&ROW())="DIÁRIO DO LEGISLATIVO";HYPERLINK("https://diariolegislativo.almg.gov.br/"&RIGHT(INDIRECT("B"&ROW());4)&"/L"&RIGHT(INDIRECT("B"&ROW());4)&MID(INDIRECT("B"&ROW());4;2)&LEFT(INDIRECT("B"&ROW());2)&".pdf";IMAGE("https://www.almg.gov.br/favicon.ico";4;15;15));
  INDIRECT("C"&ROW())="DIÁRIO DO LEGISLATIVO - EDIÇÃO EXTRA";HYPERLINK("https://diariolegislativo.almg.gov.br/"&RIGHT(INDIRECT("B"&ROW());4)&"/L"&RIGHT(INDIRECT("B"&ROW());4)&MID(INDIRECT("B"&ROW());4;2)&LEFT(INDIRECT("B"&ROW());2)&"E.pdf";IMAGE("https://www.almg.gov.br/favicon.ico";4;15;15));

  INDIRECT("C"&ROW())="REUNIÕES DE PLENÁRIO";HYPERLINK("https://www.almg.gov.br/atividade-parlamentar/plenario/agenda/?pesquisou=true&q=&tipo=&dataInicio="&TO_TEXT(INDIRECT("B"&ROW()))&"&dataFim="&TO_TEXT(INDIRECT("B"&ROW()));IMAGE("https://www.almg.gov.br/favicon.ico";4;15;15));

  INDIRECT("C"&ROW())="REUNIÕES DE COMISSÕES";HYPERLINK("https://www.almg.gov.br/atividade-parlamentar/comissoes/agenda/?pesquisou=true&q=&tpComissao=&idComissao=&dataInicio="&TO_TEXT(INDIRECT("B"&ROW()))&"&dataFim="&TO_TEXT(INDIRECT("B"&ROW()))&"&pesquisa=todas&ordem=1&tp=30";IMAGE("https://www.almg.gov.br/favicon.ico";4;15;15));

  INDIRECT("C"&ROW())="REQUERIMENTOS DE COMISSÃO";HYPERLINK("https://www.almg.gov.br/atividade-parlamentar/comissoes/agenda/?pesquisou=true&q=&tpComissao=&idComissao=&dataInicio="&TO_TEXT($V$2)&"&dataFim="&TO_TEXT($V$2)&"&pesquisa=todas&ordem=1&tp=30";IMAGE("https://www.almg.gov.br/favicon.ico";4;15;15));
  INDIRECT("C"&ROW())="OFÍCIOS DA SECRETARIA-GERAL DA MESA";HYPERLINK("https://stl.almg.gov.br/";IMAGE("https://www.almg.gov.br/favicon.ico";4;15;15));
  INDIRECT("C"&ROW())="LANÇAMENTOS DE PRECLUSÃO DE PRAZO";HYPERLINK("https://webmail.almg.gov.br/";IMAGE("https://www.almg.gov.br/favicon.ico";4;15;15));
  INDIRECT("C"&ROW())="LANÇAMENTOS DE TRAMITAÇÃO";HYPERLINK("https://www.almg.gov.br/";IMAGE("https://www.almg.gov.br/favicon.ico";4;15;15));
  INDIRECT("C"&ROW())="CADASTRO DE E-MAILS";HYPERLINK("https://webmail.almg.gov.br/";IMAGE("https://www.almg.gov.br/favicon.ico";4;15;15));

  INDIRECT("C"&ROW())="DIÁRIO DO EXECUTIVO";HYPERLINK("https://www.jornalminasgerais.mg.gov.br/?dataJornal="&RIGHT($B$6;4)&"-"&MID($B$6;4;2)&"-"&LEFT($B$6;2)&"";IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;15;15));
  LEFT(INDIRECT("C"&ROW());27)="RECEBIMENTO DE PROPOSIÇÃO: ";HYPERLINK("https://stl.almg.gov.br/html5/?versao=3.1.2#rest-oficios-"&MID($B$6;8;4)&"-"&RIGHT($B$6;4)&"-SGM";IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;15;15));
  INDIRECT("C"&ROW())="DESIGNAÇÃO DE RELATOR";HYPERLINK("https://webmail.almg.gov.br/imp/dynamic.php?page=mailbox#mbox:SU5CT1guREVTSUdOQcOHw4NPIERFIFJFTEFUT1I";IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;15;15));
  INDIRECT("C"&ROW())="CUMPRIMENTO DE DILIGÊNCIA";HYPERLINK("https://webmail.almg.gov.br/imp/dynamic.php?page=mailbox#mbox:SU5CT1guQ1VNUFJJTUVOVE8gREUgRElMSUfDik5DSUE";IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;15;15));
  INDIRECT("C"&ROW())="REUNIÃO ORIGINADA DE RQC";HYPERLINK("https://webmail.almg.gov.br/imp/dynamic.php?page=mailbox#mbox:SU5CT1guUkVVTknDg08gT1JJR0lOQURBIERFIFJRQw";IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;15;15));
  INDIRECT("C"&ROW())="REUNIÃO COM DEBATE DE PROPOSIÇÃO";HYPERLINK("https://webmail.almg.gov.br/imp/dynamic.php?page=mailbox#mbox:SU5CT1guUkVVTknDg08gQ09NIERFQkFURSBERSBQUk9QT1NJw4fDg08";IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;15;15));
  INDIRECT("C"&ROW())="SECRETARIA-GERAL DA MESA";HYPERLINK("https://webmail.almg.gov.br/imp/dynamic.php?page=mailbox#mbox:SU5CT1guU0VDUkVUQVJJQS1HRVJBTCBEQSBNRVNB";IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;15;15));

  OR(
  LEFT(INDIRECT("C"&ROW());9)="ORDINÁRIA";
  LEFT(INDIRECT("C"&ROW());14)="EXTRAORDINÁRIA";
  LEFT(INDIRECT("C"&ROW());8)="ESPECIAL";
  LEFT(INDIRECT("C"&ROW());14)="SOLENE");
  IFS(
  E6="cancelada";
  HYPERLINK("https://www.almg.gov.br/atividade_parlamentar/plenario/interna.html?tipo=pauta&dDet="&LEFT($X$4;2)&"|"&MID($X$4;4;2)&"|"&RIGHT($X$4;4)&"&hDet="&TO_TEXT(INDIRECT("B"&ROW()));
  IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;15;15));
  E6<>"cancelada";
  HYPERLINK("https://www.almg.gov.br/atividade_parlamentar/plenario/interna.html?tipo=res&dia="&LEFT($X$4;2)&"&mes="&MID($X$4;4;2)&"&ano="&RIGHT($X$4;4)&"&hr="&TO_TEXT(INDIRECT("B"&ROW()));
  IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;15;15)));

  OR(LEFT(INDIRECT("C"&ROW());10)="COMISSÃO D";LEFT(INDIRECT("C"&ROW());10)="COMISSÃO E";LEFT(INDIRECT("C"&ROW());6)="GRANDE";LEFT(INDIRECT("C"&ROW());7)="REUNIÃO";RIGHT(INDIRECT("C"&ROW());11)="PERMANENTES";RIGHT(INDIRECT("C"&ROW());8)="CONJUNTA";LEFT(INDIRECT("C"&ROW());4)="CIPE");HYPERLINK("https://www.almg.gov.br/atividade-parlamentar/comissoes/"
  &IFS(RIGHT(INDIRECT("C"&ROW());6)="VISITA";"visita";RIGHT(INDIRECT("C"&ROW());8)<>"VISITA";"reuniao")&"/?idTipo="
  &IFS(
  OR(RIGHT(INDIRECT("C"&ROW());11)="GASTRONOMIA";RIGHT(INDIRECT("C"&ROW());6)="URBANA");"2";
  OR(MID(INDIRECT("C"&ROW());10;14)="EXTRAORDINÁRIA";MID(INDIRECT("C"&ROW());13;5)="ÉTICA";RIGHT(INDIRECT("C"&ROW());8)="ESPECIAL");"5";
  OR(RIGHT(INDIRECT("C"&ROW());14)="EXTRAORDINÁRIA";MID(INDIRECT("C"&ROW());13;8)="PROPOSTA";RIGHT(INDIRECT("C"&ROW());7)="ANIMAIS";RIGHT(INDIRECT("C"&ROW());6)="CÂNCER";RIGHT(INDIRECT("C"&ROW());7)="MARIANA");"2";
  OR(LEFT(INDIRECT("C"&ROW());6)="GRANDE";LEFT(INDIRECT("C"&ROW());7)="REUNIÃO";RIGHT(INDIRECT("C"&ROW());11)="PERMANENTES";RIGHT(INDIRECT("C"&ROW());8)="CONJUNTA");"3";
  RIGHT(INDIRECT("C"&ROW());14)="REFORMA URBANA";"1";
  LEFT(INDIRECT("C"&ROW());4)="CIPE";"7";
  RIGHT(INDIRECT("C"&ROW());14)<>"EXTRAORDINÁRIA";"1")
  &"&idCom="
  &IFS(
  LEFT(INDIRECT("C"&ROW());33)="COMISSÃO DE ADMINISTRAÇÃO PÚBLICA";"1";
  LEFT(INDIRECT("C"&ROW());40)="COMISSÃO DE AGROPECUÁRIA E AGROINDÚSTRIA";"1075";
  LEFT(INDIRECT("C"&ROW());48)="COMISSÃO DE ASSUNTOS MUNICIPAIS E REGIONALIZAÇÃO";"3";
  LEFT(INDIRECT("C"&ROW());34)="COMISSÃO DE CONSTITUIÇÃO E JUSTIÇA";"5";
  LEFT(INDIRECT("C"&ROW());19)="COMISSÃO DE CULTURA";"675";
  LEFT(INDIRECT("C"&ROW());50)="COMISSÃO DE DEFESA DO CONSUMIDOR E DO CONTRIBUINTE";"489";
  LEFT(INDIRECT("C"&ROW());41)="COMISSÃO DE DEFESA DOS DIREITOS DA MULHER";"1132";
  LEFT(INDIRECT("C"&ROW());57)="COMISSÃO DE DEFESA DOS DIREITOS DA PESSOA COM DEFICIÊNCIA";"859";
  LEFT(INDIRECT("C"&ROW());37)="COMISSÃO DE DESENVOLVIMENTO ECONÔMICO";"1077";
  LEFT(INDIRECT("C"&ROW());28)="COMISSÃO DE DIREITOS HUMANOS";"8";
  LEFT(INDIRECT("C"&ROW());42)="COMISSÃO DE EDUCAÇÃO, CIÊNCIA E TECNOLOGIA";"849";
  LEFT(INDIRECT("C"&ROW());38)="COMISSÃO DE ESPORTE, LAZER E JUVENTUDE";"850";
  LEFT(INDIRECT("C"&ROW());50)="COMISSÃO DE FISCALIZAÇÃO FINANCEIRA E ORÇAMENTÁRIA";"10";
  LEFT(INDIRECT("C"&ROW());55)="COMISSÃO DE MEIO AMBIENTE E DESENVOLVIMENTO SUSTENTÁVEL";"799";
  LEFT(INDIRECT("C"&ROW());27)="COMISSÃO DE MINAS E ENERGIA";"800";
  LEFT(INDIRECT("C"&ROW());32)="COMISSÃO DE PARTICIPAÇÃO POPULAR";"585";
  LEFT(INDIRECT("C"&ROW());63)="COMISSÃO DE PREVENÇÃO E COMBATE AO USO DE CRACK E OUTRAS DROGAS";"959";
  LEFT(INDIRECT("C"&ROW());19)="COMISSÃO DE REDAÇÃO";"13";
  LEFT(INDIRECT("C"&ROW());17)="COMISSÃO DE SAÚDE";"14";
  LEFT(INDIRECT("C"&ROW());29)="COMISSÃO DE SEGURANÇA PÚBLICA";"508";
  LEFT(INDIRECT("C"&ROW());60)="COMISSÃO DO TRABALHO, DA PREVIDÊNCIA E DA ASSISTÊNCIA SOCIAL";"1076";
  LEFT(INDIRECT("C"&ROW());52)="COMISSÃO DE TRANSPORTE, COMUNICAÇÃO E OBRAS PÚBLICAS";"12";
  LEFT(INDIRECT("C"&ROW());17)="COMISSÃO DE ÉTICA";"578";
  LEFT(INDIRECT("C"&ROW());71)="COMISSÃO EXTRAORDINÁRIA DAS ENERGIAS RENOVÁVEIS E DOS RECURSOS HÍDRICOS";"1211";
  LEFT(INDIRECT("C"&ROW());62)="COMISSÃO EXTRAORDINÁRIA DE ACOMPANHAMENTO DO ACORDO DE MARIANA";"1232";
  LEFT(INDIRECT("C"&ROW());66)="COMISSÃO EXTRAORDINÁRIA DE DEFESA DA HABITAÇÃO E DA REFORMA URBANA";"1260";
  LEFT(INDIRECT("C"&ROW());62)="COMISSÃO EXTRAORDINÁRIA DE PREVENÇÃO E ENFRENTAMENTO AO CÂNCER";"1258";
  LEFT(INDIRECT("C"&ROW());47)="COMISSÃO EXTRAORDINÁRIA DE PROTEÇÃO AOS ANIMAIS";"1230";
  LEFT(INDIRECT("C"&ROW());41)="COMISSÃO EXTRAORDINÁRIA DAS PRIVATIZAÇÕES";"1212";
  LEFT(INDIRECT("C"&ROW());48)="COMISSÃO EXTRAORDINÁRIA DE TURISMO E GASTRONOMIA";"1261";
  LEFT(INDIRECT("C"&ROW());46)="COMISSÃO EXTRAORDINÁRIA PRÓ-FERROVIAS MINEIRAS";"1217";
  LEFT(INDIRECT("C"&ROW());15)="GRANDE COMISSÃO";"10";
  LEFT(INDIRECT("C"&ROW());53)="COMISSÃO DE PROPOSTA DE EMENDA À CONSTITUIÇÃO 42 2024";"1279";
  LEFT(INDIRECT("C"&ROW());53)="COMISSÃO DE PROPOSTA DE EMENDA À CONSTITUIÇÃO 24 2023";"1280";
  LEFT(INDIRECT("C"&ROW());53)="COMISSÃO DE PROPOSTA DE EMENDA À CONSTITUIÇÃO 58 2025";"1281";
  LEFT(INDIRECT("C"&ROW());45)="COMISSÃO DE MEMBROS DAS COMISSÕES PERMANENTES";"10";
  RIGHT(INDIRECT("C"&ROW());9)="PCD + SPU";959;
  RIGHT(INDIRECT("C"&ROW());9)="CTU + DEC";675;
  LEFT(INDIRECT("C"&ROW());16)="REUNIÃO CONJUNTA";"1";
  LEFT(INDIRECT("C"&ROW());4)="CIPE";"811";
  LEFT(INDIRECT("C"&ROW());24)="COMISSÃO DE VETO 18 2025";"1265";
  LEFT(INDIRECT("C"&ROW());24)="COMISSÃO DE VETO 19 2025";"1264";
  LEFT(INDIRECT("C"&ROW());24)="COMISSÃO DE VETO 20 2025";"1267";
  LEFT(INDIRECT("C"&ROW());24)="COMISSÃO DE VETO 21 2025";"1262";
  LEFT(INDIRECT("C"&ROW());24)="COMISSÃO DE VETO 22 2025";"1266";
  LEFT(INDIRECT("C"&ROW());24)="COMISSÃO DE VETO 23 2025";"1263";
  LEFT(INDIRECT("C"&ROW());24)="COMISSÃO DE VETO 24 2025";"1270"
  )&"&dia="&IFS(MID($A$5;2;1)="/";LEFT($A$5;1);MID($A$5;2;1)<>"/";LEFT($A$5;2))&"&mes="&IFS(MID($A$5;3;1)="/";IFS(MID($A$5;4;1)<>"1";RIGHT($A$5;1);MID($A$5;4;1)="1";IFS(MID($A$5;5;1)="";RIGHT($A$5;1);MID($A$5;5;1)<>"";RIGHT($A$5;2)));MID($A$5;2;1)="/";IFS(MID($A$5;3;1)<>"1";RIGHT($A$5;1);MID($A$5;3;1)="1";IFS(MID($A$5;4;1)="";RIGHT($A$5;1);MID($A$5;4;1)<>"";RIGHT($A$5;2))))&"&ano="&RIGHT($B$6;4)&"&hr="&TO_TEXT(INDIRECT("B"&ROW()))&"&tpCom="&IFS(LEFT(INDIRECT("C"&ROW());45)="COMISSÃO DE MEMBROS DAS COMISSÕES PERMANENTES";"3";LEFT(INDIRECT("C"&ROW());45)<>"COMISSÃO DE MEMBROS DAS COMISSÕES PERMANENTES";"2")&"&aba=js_tabResultado";
  IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;15;15));



  OR(LEFT(INDIRECT("C"&ROW());5)="RQC: ");HYPERLINK("https://www.almg.gov.br/atividade-parlamentar/comissoes/reuniao/?idTipo="
  &IFS(
  OR(RIGHT(INDIRECT("C"&ROW());11)="GASTRONOMIA";RIGHT(INDIRECT("C"&ROW());6)="URBANA");"2";
  OR(RIGHT(INDIRECT("C"&ROW());14)="EXTRAORDINÁRIA";RIGHT(INDIRECT("C"&ROW());25)="EXTRAORDINÁRIA, APROVADOS";RIGHT(INDIRECT("C"&ROW());26)="EXTRAORDINÁRIA - APROVADOS";RIGHT(INDIRECT("C"&ROW());25)="EXTRAORDINÁRIA, RECEBIDOS";RIGHT(INDIRECT("C"&ROW());26)="EXTRAORDINÁRIA - RECEBIDOS";RIGHT(INDIRECT("C"&ROW());37)="EXTRAORDINÁRIA, RECEBIDOS E APROVADOS";RIGHT(INDIRECT("C"&ROW());38)="EXTRAORDINÁRIA - RECEBIDOS E APROVADOS";MID(INDIRECT("C"&ROW());13;8)="PROPOSTA";RIGHT(INDIRECT("C"&ROW());7)="ANIMAIS";RIGHT(INDIRECT("C"&ROW());6)="CÂNCER";RIGHT(INDIRECT("C"&ROW());7)="MARIANA");"2";
  OR(LEFT(INDIRECT("C"&ROW());6)="GRANDE";LEFT(INDIRECT("C"&ROW());7)="REUNIÃO";RIGHT(INDIRECT("C"&ROW());11)="PERMANENTES";RIGHT(INDIRECT("C"&ROW());8)="CONJUNTA";RIGHT(INDIRECT("C"&ROW());19)="CONJUNTA, APROVADOS";RIGHT(INDIRECT("C"&ROW());19)="CONJUNTA, RECEBIDOS");"3";
  OR(MID(INDIRECT("C"&ROW());10;14)="EXTRAORDINÁRIA";RIGHT(INDIRECT("C"&ROW());8)="ESPECIAL");"5";
  LEFT(INDIRECT("C"&ROW());4)="CIPE";"6";
  RIGHT(INDIRECT("C"&ROW());14)<>"EXTRAORDINÁRIA";"1")
  &"&idCom="
  &IFS(
  LEFT(INDIRECT("C"&ROW());26)="RQC: ADMINISTRAÇÃO PÚBLICA";"1";
  LEFT(INDIRECT("C"&ROW());33)="RQC: AGROPECUÁRIA E AGROINDÚSTRIA";"1075";
  LEFT(INDIRECT("C"&ROW());41)="RQC: ASSUNTOS MUNICIPAIS E REGIONALIZAÇÃO";"3";
  LEFT(INDIRECT("C"&ROW());27)="RQC: CONSTITUIÇÃO E JUSTIÇA";"5";
  LEFT(INDIRECT("C"&ROW());12)="RQC: CULTURA";"675";
  LEFT(INDIRECT("C"&ROW());43)="RQC: DEFESA DO CONSUMIDOR E DO CONTRIBUINTE";"489";
  LEFT(INDIRECT("C"&ROW());34)="RQC: DEFESA DOS DIREITOS DA MULHER";"1132";
  LEFT(INDIRECT("C"&ROW());50)="RQC: DEFESA DOS DIREITOS DA PESSOA COM DEFICIÊNCIA";"859";
  LEFT(INDIRECT("C"&ROW());30)="RQC: DESENVOLVIMENTO ECONÔMICO";"1077";
  LEFT(INDIRECT("C"&ROW());21)="RQC: DIREITOS HUMANOS";"8";
  LEFT(INDIRECT("C"&ROW());35)="RQC: EDUCAÇÃO, CIÊNCIA E TECNOLOGIA";"849";
  LEFT(INDIRECT("C"&ROW());31)="RQC: ESPORTE, LAZER E JUVENTUDE";"850";
  LEFT(INDIRECT("C"&ROW());43)="RQC: FISCALIZAÇÃO FINANCEIRA E ORÇAMENTÁRIA";"10";
  LEFT(INDIRECT("C"&ROW());48)="RQC: MEIO AMBIENTE E DESENVOLVIMENTO SUSTENTÁVEL";"799";
  LEFT(INDIRECT("C"&ROW());20)="RQC: MINAS E ENERGIA";"800";
  LEFT(INDIRECT("C"&ROW());25)="RQC: PARTICIPAÇÃO POPULAR";"585";
  LEFT(INDIRECT("C"&ROW());56)="RQC: PREVENÇÃO E COMBATE AO USO DE CRACK E OUTRAS DROGAS";"959";
  LEFT(INDIRECT("C"&ROW());12)="RQC: REDAÇÃO";"13";
  LEFT(INDIRECT("C"&ROW());10)="RQC: SAÚDE";"14";
  LEFT(INDIRECT("C"&ROW());22)="RQC: SEGURANÇA PÚBLICA";"508";
  LEFT(INDIRECT("C"&ROW());53)="RQC: TRABALHO, DA PREVIDÊNCIA E DA ASSISTÊNCIA SOCIAL";"1076";
  LEFT(INDIRECT("C"&ROW());45)="RQC: TRANSPORTE, COMUNICAÇÃO E OBRAS PÚBLICAS";"12";
  LEFT(INDIRECT("C"&ROW());67)="RQC: EXTRAORDINÁRIA DAS ENERGIAS RENOVÁVEIS E DOS RECURSOS HÍDRICOS";"1211";
  LEFT(INDIRECT("C"&ROW());58)="RQC: EXTRAORDINÁRIA DE ACOMPANHAMENTO DO ACORDO DE MARIANA";"1232";
  LEFT(INDIRECT("C"&ROW());62)="RQC: EXTRAORDINÁRIA DE DEFESA DA HABITAÇÃO E DA REFORMA URBANA";"1260";
  LEFT(INDIRECT("C"&ROW());58)="RQC: EXTRAORDINÁRIA DE PREVENÇÃO E ENFRENTAMENTO AO CÂNCER";"1258";
  LEFT(INDIRECT("C"&ROW());43)="RQC: EXTRAORDINÁRIA DE PROTEÇÃO AOS ANIMAIS";"1230";
  LEFT(INDIRECT("C"&ROW());37)="RQC: EXTRAORDINÁRIA DAS PRIVATIZAÇÕES";"1212";
  LEFT(INDIRECT("C"&ROW());44)="RQC: EXTRAORDINÁRIA DE TURISMO E GASTRONOMIA";"1261";
  LEFT(INDIRECT("C"&ROW());42)="RQC: EXTRAORDINÁRIA PRÓ-FERROVIAS MINEIRAS";"1217";
  LEFT(INDIRECT("C"&ROW());38)="RQC: PROPOSTA DE EMENDA À CONSTITUIÇÃO";"1234";
  LEFT(INDIRECT("C"&ROW());38)="RQC: PROPOSTA DE EMENDA À CONSTITUIÇÃO";"1227";
  LEFT(INDIRECT("C"&ROW());38)="RQC: PROPOSTA DE EMENDA À CONSTITUIÇÃO";"1218";
  LEFT(INDIRECT("C"&ROW());45)="RQC: MEMBROS DAS COMISSÕES PERMANENTES";"10";
  LEFT(INDIRECT("C"&ROW());18)="RQC: CIPE RIO DOCE";"811"
  )&"&dia="&IFS(MID($A$5;2;1)="/";LEFT($A$5;1);MID($A$5;2;1)<>"/";LEFT($A$5;2))&"&mes="&IFS(MID($A$5;3;1)="/";IFS(MID($A$5;4;1)<>"1";RIGHT($A$5;1);MID($A$5;4;1)="1";IFS(MID($A$5;5;1)="";RIGHT($A$5;1);MID($A$5;5;1)<>"";RIGHT($A$5;2)));MID($A$5;2;1)="/";IFS(MID($A$5;3;1)<>"1";RIGHT($A$5;1);MID($A$5;3;1)="1";IFS(MID($A$5;4;1)="";RIGHT($A$5;1);MID($A$5;4;1)<>"";RIGHT($A$5;2))))&"&ano="&RIGHT($B$6;4)&"&hr="&TO_TEXT(INDIRECT("B"&ROW()))&"&tpCom="&IFS(LEFT(INDIRECT("C"&ROW());45)="RQC: MEMBROS DAS COMISSÕES PERMANENTES";"3";LEFT(INDIRECT("C"&ROW());45)<>"RQC: MEMBROS DAS COMISSÕES PERMANENTES";"2")&"&aba=js_tabResultado";
  IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;15;15));



  OR(LEFT(INDIRECT("C"&ROW());19)="AUDIÊNCIA PÚBLICA: ");HYPERLINK("https://www.almg.gov.br/atividade-parlamentar/comissoes/reuniao/?idTipo="
  &IFS(
  OR(RIGHT(INDIRECT("C"&ROW());11)="GASTRONOMIA";RIGHT(INDIRECT("C"&ROW());6)="URBANA");"2";
  OR(RIGHT(INDIRECT("C"&ROW());14)="EXTRAORDINÁRIA";RIGHT(INDIRECT("C"&ROW());25)="EXTRAORDINÁRIA, APROVADOS";RIGHT(INDIRECT("C"&ROW());25)="EXTRAORDINÁRIA, RECEBIDOS";MID(INDIRECT("C"&ROW());13;8)="PROPOSTA";RIGHT(INDIRECT("C"&ROW());7)="ANIMAIS";RIGHT(INDIRECT("C"&ROW());6)="CÂNCER";RIGHT(INDIRECT("C"&ROW());7)="MARIANA");"2";
  OR(LEFT(INDIRECT("C"&ROW());6)="GRANDE";LEFT(INDIRECT("C"&ROW());7)="REUNIÃO";RIGHT(INDIRECT("C"&ROW());11)="PERMANENTES";RIGHT(INDIRECT("C"&ROW());8)="CONJUNTA";RIGHT(INDIRECT("C"&ROW());19)="CONJUNTA, APROVADOS";RIGHT(INDIRECT("C"&ROW());19)="CONJUNTA, RECEBIDOS");"3";
  OR(MID(INDIRECT("C"&ROW());10;14)="EXTRAORDINÁRIA";RIGHT(INDIRECT("C"&ROW());8)="ESPECIAL");"5";
  LEFT(INDIRECT("C"&ROW());4)="CIPE";"6";
  RIGHT(INDIRECT("C"&ROW());14)<>"EXTRAORDINÁRIA";"1")
  &"&idCom="
  &IFS(
  MID(INDIRECT("C"&ROW());20;3)="APU";"1";
  MID(INDIRECT("C"&ROW());20;3)="AAG";"1075";
  MID(INDIRECT("C"&ROW());20;3)="AMR";"3";
  MID(INDIRECT("C"&ROW());20;3)="CJU";"5";
  MID(INDIRECT("C"&ROW());20;3)="CTU";"675";
  MID(INDIRECT("C"&ROW());20;3)="DCC";"489";
  MID(INDIRECT("C"&ROW());20;3)="DDM";"1132";
  MID(INDIRECT("C"&ROW());20;3)="DPD";"859";
  MID(INDIRECT("C"&ROW());20;3)="DEC";"1077";
  MID(INDIRECT("C"&ROW());20;3)="DHU";"8";
  MID(INDIRECT("C"&ROW());20;3)="ECT";"849";
  MID(INDIRECT("C"&ROW());20;3)="ELJ";"850";
  MID(INDIRECT("C"&ROW());20;3)="FFO";"10";
  MID(INDIRECT("C"&ROW());20;3)="MAD";"799";
  MID(INDIRECT("C"&ROW());20;3)="MEN";"800";
  MID(INDIRECT("C"&ROW());20;3)="PPO";"585";
  MID(INDIRECT("C"&ROW());20;3)="PCD";"959";
  MID(INDIRECT("C"&ROW());20;3)="RED";"13";
  MID(INDIRECT("C"&ROW());20;3)="SAU";"14";
  MID(INDIRECT("C"&ROW());20;3)="SPU";"508";
  MID(INDIRECT("C"&ROW());20;3)="TPA";"1076";
  MID(INDIRECT("C"&ROW());20;3)="TCO";"12"
  )&"&dia="&MID(INDIRECT("C"&ROW());25;2)
  &"&mes="&MID(INDIRECT("C"&ROW());28;2)
  &"&ano="&MID(INDIRECT("C"&ROW());31;4)
  &"&hr="&TO_TEXT(INDIRECT("B"&ROW()))&"&tpCom="&IFS(LEFT(INDIRECT("C"&ROW());45)="RQC: MEMBROS DAS COMISSÕES PERMANENTES";"3";LEFT(INDIRECT("C"&ROW());45)<>"RQC: MEMBROS DAS COMISSÕES PERMANENTES";"2")&"&aba=js_tabResultado";
  IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;15;15));



  OR(INDIRECT("C"&ROW())<>"REUNIÕES DE PLENÁRIO");
  HYPERLINK("https://www.almg.gov.br/export/sites/default/consulte/arquivo_diario_legislativo/pdfs/"&RIGHT($B$6;4)&"/"&MID($B$6;4;2)&"/L"&RIGHT($B$6;4)&MID($B$6;4;2)&LEFT($B$6;2)&".pdf#page="&IFS(MID(INDIRECT("B"&ROW());3;1)="";IFS(LEFT(INDIRECT("B"&ROW());1)=0;LEFT(INDIRECT("B"&ROW());2);LEFT(INDIRECT("B"&ROW());1)<>0;LEFT(INDIRECT("B"&ROW());2));MID(INDIRECT("B"&ROW());3;1)<>"";LEFT(INDIRECT("B"&ROW());3));

  IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;15;15))
  ))''']]
    * ((footer_start - 1) - 5)
)

# ------------------------------------------------------------------
# P6 (mega fórmula) - preservada integralmente
# ------------------------------------------------------------------
add(
    f"P6:P{footer_start - 1}",
    [[r'''=IFS(

  OR(INDIRECT("E"&ROW())<>"DIOGO");
  IFS(
  OR(INDIRECT("E"&ROW())="-";INDIRECT("E"&ROW())="cancelada";INDIRECT("E"&ROW())="sem quórum");"-";
  INDIRECT("E"&ROW())<>"cancelada";
  IFS(
  OR(INDIRECT("T"&ROW())="-";INDIRECT("H"&ROW())="-");"-";INDIRECT("T"&ROW())="??";
  IMAGE("https://cdn.iconscout.com/icon/premium/png-512-thumb/broken-link-18-610397.png";4;20;20);

  T6<>"??";
  IFS(
  OR(INDIRECT("C"&ROW())="";INDIRECT("C"&ROW())="DIÁRIO DO EXECUTIVO";INDIRECT("C"&ROW())="DIÁRIO DO EXECUTIVO - EDIÇÃO EXTRA";INDIRECT("C"&ROW())="DIÁRIO DO LEGISLATIVO";INDIRECT("C"&ROW())="DIÁRIO DO LEGISLATIVO - EDIÇÃO EXTRA";INDIRECT("C"&ROW())="REUNIÕES DE PLENÁRIO";INDIRECT("C"&ROW())="REUNIÕES DE COMISSÕES";INDIRECT("C"&ROW())="REQUERIMENTOS DE COMISSÕES";INDIRECT("C"&ROW())="LANÇAMENTOS DE TRAMITAÇÃO";INDIRECT("C"&ROW())="CADASTRO DE E-MAILS";INDIRECT("C"&ROW())="OFÍCIOS DA SECRETARIA-GERAL DA MESA";INDIRECT("C"&ROW())="LANÇAMENTOS DE PRECLUSÃO DE PRAZO";INDIRECT("C"&ROW())="IMPLANTAÇÃO DE TEXTOS";INDIRECT("C"&ROW())="REQUERIMENTOS DE COMISSÃO");" ";

  OR(INDIRECT("C"&ROW())<>"REUNIÕES DE PLENÁRIO");IFS($A$683=FALSE;

  HYPERLINK("https://integracao.almg.gov.br/mate-brs/index.html?first=false&search=odp&pagina=1&tp=200&aba=js_tabpesquisaAvancada&txtPalavras="&T6;IMAGE("https://www.almg.gov.br/favicon.ico";4;17;17));
  $A$683=TRUE;HYPERLINK(X6;IMAGE("https://www.almg.gov.br/favicon.ico";4;17;17)))))))''']]
    * ((footer_start - 1) - 5)
)

# ------------------------------------------------------------------
# Q6 (mega fórmula) - preservada integralmente
# ------------------------------------------------------------------
add(
    f"Q6:Q{footer_start - 1}",
    [[r'''=IFS(

  OR(INDIRECT("C"&ROW())="";
  LEFT(INDIRECT("C"&ROW());6)="DIÁRIO";
  LEFT(INDIRECT("C"&ROW());8)="REUNIÕES";
  INDIRECT("C"&ROW())="REQUERIMENTOS DE COMISSÃO";
  INDIRECT("C"&ROW())="LANÇAMENTOS DE TRAMITAÇÃO";
  INDIRECT("C"&ROW())="CADASTRO DE E-MAILS";
  INDIRECT("C"&ROW())="OFÍCIOS DA SECRETARIA-GERAL DA MESA";
  INDIRECT("C"&ROW())="LANÇAMENTOS DE PRECLUSÃO DE PRAZO";
  INDIRECT("C"&ROW())="IMPLANTAÇÃO DE TEXTOS");"";

  OR(
  INDIRECT("C"&ROW())="ALINE";
  INDIRECT("C"&ROW())="ANDRÉ";
  INDIRECT("C"&ROW())="DIOGO";
  INDIRECT("C"&ROW())="KÁTIA";
  INDIRECT("C"&ROW())="LEO";
  INDIRECT("C"&ROW())="WELDER";
  INDIRECT("C"&ROW())="TOTAL";
  ISNUMBER(INDIRECT("C"&ROW()));
  INDIRECT("C"&ROW())="?";INDIRECT("K"&ROW())="-";
  INDIRECT("K"&ROW())="cancelada";
  INDIRECT("K"&ROW())="sem quórum";
  INDIRECT("K"&ROW())="não publicado");"-";

  OR(
  INDIRECT("C"&ROW())="-";
  INDIRECT("C"&ROW())="ERRATAS";
  INDIRECT("C"&ROW())="MANIFESTAÇÕES";
  INDIRECT("C"&ROW())="VOTAÇÕES NOMINAIS";
  RIGHT(INDIRECT("C"&ROW());18)="EMENDAS PUBLICADAS";
  LEFT(INDIRECT("C"&ROW());17)="VOTAÇÕES NOMINAIS");"-";


  LEFT(INDIRECT("C"&ROW());6)<>"DIÁRIO";
  IFS(

  OR(INDIRECT("C"&ROW())="EMENDA À CONSTITUIÇÃO PROMULGADA";INDIRECT("C"&ROW())="EMENDAS À CONSTITUIÇÃO PROMULGADAS");"PL??";
  OR(INDIRECT("C"&ROW())="PROPOSTA DE AÇÃO LEGISLATIVA";INDIRECT("C"&ROW())="PROPOSTAS DE AÇÃO LEGISLATIVA");"PLE1";
  RIGHT(INDIRECT("C"&ROW());48)="PROPOSTAS DE AÇÃO LEGISLATIVA REFERENTES AO PPAG";"PLE1";
  RIGHT(INDIRECT("C"&ROW());24)="VOTAÇÃO DE REQUERIMENTOS";"RQN??/PL??";
  INDIRECT("C"&ROW())="VETO TOTAL A PROPOSIÇÃO DE LEI";"PL80";
  INDIRECT("C"&ROW())="VETO PARCIAL A PROPOSIÇÃO DE LEI";"PL82";
  INDIRECT("C"&ROW())="VETO PARCIAL A PROPOSIÇÃO DE LEI COMPLEMENTAR";"PLC14";
  OR(INDIRECT("C"&ROW())="RESOLUÇÃO";INDIRECT("C"&ROW())="RESOLUÇÕES");"PRE131";
  INDIRECT("C"&ROW())="PROPOSIÇÕES DE LEI";"PL63";
  INDIRECT("C"&ROW())="DECISÃO DA MESA";"PL??";
  INDIRECT("C"&ROW())="DECISÕES DA PRESIDÊNCIA";"PL??";
  OR(INDIRECT("C"&ROW())="DESIGNAÇÃO DE COMISSÕES";INDIRECT("C"&ROW())="TRAMITAÇÃO DE PROPOSIÇÕES: DESIGNAÇÃO DE COMISSÕES");"PL??";
  INDIRECT("C"&ROW())="OFÍCIOS DE PREFEITURA QUE ENCAMINHAM DECRETOS DE CALAMIDADE PÚBLICA";"PL??";
  INDIRECT("C"&ROW())="PROPOSIÇÃO: REQUERIMENTOS - INDICAÇÃO TCE";"PL??";
  INDIRECT("C"&ROW())="SOLENE";"-";
  INDIRECT("C"&ROW())="ESPECIAL";"PL??";
  INDIRECT("C"&ROW())="ORDINÁRIA";"PL??";
  INDIRECT("C"&ROW())="EXTRAORDINÁRIA";"PL??";
  INDIRECT("C"&ROW())="EXTRAORDINÁRIA: PARECERES DE REDAÇÃO FINAL APROVADOS";"PL62";
  OR(INDIRECT("C"&ROW())="ERRATAS";INDIRECT("C"&ROW())="ERRATA");"PL??";
  RIGHT(INDIRECT("C"&ROW());23)="LEITURA DE COMUNICAÇÕES";"-";
  OR(RIGHT(INDIRECT("C"&ROW());11)="PROMULGADAS");"PL112";
  OR(LEFT(INDIRECT("C"&ROW());3)="LEI");"PL81";
  OR(LEFT(INDIRECT("C"&ROW());16)="LEI COMPLEMENTAR");"PL81";
  OR(LEFT(INDIRECT("C"&ROW());27)="LEI, COM PROPOSIÇÃO ANEXADA");"PL81//PL5";
  OR(INDIRECT("C"&ROW())="EMENDAS OU SUBSTITUTIVOS PUBLICADOS";INDIRECT("C"&ROW())="EMENDAS NÃO RECEBIDAS PUBLICADAS");"PL??";
  LEFT(INDIRECT("C"&ROW());8)="COMISSÃO";"RQN??";
  LEFT(INDIRECT("C"&ROW());4)="CIPE";"RQN??";
  LEFT(INDIRECT("C"&ROW());16)="REUNIÃO CONJUNTA";"RQN??";
  OR(LEFT(INDIRECT("C"&ROW());19)="RELATÓRIO DE VISITA";LEFT(INDIRECT("C"&ROW());46)="TRAMITAÇÃO DE PROPOSIÇÕES: RELATÓRIO DE VISITA");"RQC18";
  RIGHT(INDIRECT("C"&ROW());33)="RELATÓRIO DE EVENTO INSTITUCIONAL";"REL1";
  RIGHT(INDIRECT("C"&ROW());24)="REQUERIMENTOS ORDINÁRIOS";"RQO1";
  RIGHT(INDIRECT("C"&ROW());24)="REQUERIMENTOS APROVADOS";"RQN66";
  RIGHT(INDIRECT("C"&ROW());26)="COMUNICAÇÃO DA PRESIDÊNCIA";"RQN26";
  OR(LEFT(INDIRECT("C"&ROW());22)="DECISÃO DA PRESIDÊNCIA";LEFT(INDIRECT("C"&ROW());49)="TRAMITAÇÃO DE PROPOSIÇÕES: DECISÃO DA PRESIDÊNCIA");"PL??";
  RIGHT(INDIRECT("C"&ROW());22)="PALAVRAS DO PRESIDENTE";"PL??";
  LEFT(INDIRECT("C"&ROW());25)="DESPACHO DE REQUERIMENTOS";"RQN83";
  INDIRECT("C"&ROW())="TRAMITAÇÃO DE PROPOSIÇÕES: PARECERES";"PL178";
  INDIRECT("C"&ROW())="PARECERES SOBRE VETO";"PL??";
  INDIRECT("C"&ROW())="PARECERES SOBRE SUBSTITUTIVO";"PL??";
  LEFT(INDIRECT("C"&ROW());17)="AUDIÊNCIA PÚBLICA";"-";
  LEFT(INDIRECT("C"&ROW());23)="AUDIÊNCIA DE CONVIDADOS";"-";

  OR(LEFT(INDIRECT("C"&ROW());36)="MENSAGEM DO GOVERNADOR QUE ENCAMINHA";
  LEFT(INDIRECT("C"&ROW());38)="MENSAGENS DO GOVERNADOR QUE ENCAMINHAM";
  LEFT(INDIRECT("C"&ROW());63)="TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA";
  LEFT(INDIRECT("C"&ROW());65)="TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGENS DO GOVERNADOR QUE ENCAMINHAM";
  LEFT(INDIRECT("C"&ROW());35)="MENSAGEM DO GOVERNADOR QUE COMUNICA";
  LEFT(INDIRECT("C"&ROW());62)="TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE COMUNICA";
  LEFT(INDIRECT("C"&ROW());35)="MENSAGEM DO GOVERNADOR QUE SOLICITA");IFS(
  OR(RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";RIGHT(INDIRECT("C"&ROW());15)="PROJETOS DE LEI";RIGHT(INDIRECT("C"&ROW());36)="PROJETO DE LEI - CRÉDITO SUPLEMENTAR";);"PL??";
  RIGHT(INDIRECT("C"&ROW());42)="EMENDA OU SUBSTITUTIVO COM DESPACHO À MESA";"MSG5 // PL156";
  RIGHT(INDIRECT("C"&ROW());41)="EMENDA OU SUBSTITUTIVO COM DESPACHO À FFO";"MSG7 // PL624";
  RIGHT(INDIRECT("C"&ROW());12)="VETO PARCIAL";"PL??";
  RIGHT(INDIRECT("C"&ROW());10)="VETO TOTAL";"PL??";
  RIGHT(INDIRECT("C"&ROW());29)="REGIME ESPECIAL DE TRIBUTAÇÃO";"MSG21";
  RIGHT(INDIRECT("C"&ROW());9)="INDICAÇÃO";"PL??";
  RIGHT(INDIRECT("C"&ROW());28)="PEDIDO DE REGIME DE URGÊNCIA";"PL??";
  RIGHT(INDIRECT("C"&ROW());18)="CONVÊNIO DO CONFAZ";"MSG11";
  RIGHT(INDIRECT("C"&ROW());16)="CONVÊNIO DO ICMS";"MSG11";
  RIGHT(INDIRECT("C"&ROW());44)="PRESTAÇÃO DE CONTAS DA ADMINISTRAÇÃO PÚBLICA";"PL??";
  RIGHT(INDIRECT("C"&ROW());36)="RELATÓRIO SOBRE A SITUAÇÃO DO ESTADO";"MSG22";
  RIGHT(INDIRECT("C"&ROW());29)="DESARQUIVAMENTO DE PROPOSIÇÃO";"MSG8 // RQN80";
  RIGHT(INDIRECT("C"&ROW());19)="RETIRADA DE PROJETO";"MSG12 // RQN80";
  RIGHT(INDIRECT("C"&ROW());16)="AUSÊNCIA DO PAÍS";"OFI10"
  );

  OR(LEFT(INDIRECT("C"&ROW());20)="OFÍCIO DO GOVERNADOR";
  MID(INDIRECT("C"&ROW());28;20)="OFÍCIO DO GOVERNADOR";
  LEFT(INDIRECT("C"&ROW());25)="OFÍCIO DO VICE-GOVERNADOR");IFS(
  RIGHT(INDIRECT("C"&ROW());28)="COMUNICANDO AUSÊNCIA DO PAÍS";"OFI10";
  RIGHT(INDIRECT("C"&ROW());35)="COMUNICANDO QUE ENCAMINHOU MENSAGEM";"OFI??");

  LEFT(INDIRECT("C"&ROW());27)="REQUERIMENTOS DE COMISSÃO: ";IFS(
  RIGHT(INDIRECT("C"&ROW());21)="RECEBIDOS E APROVADOS";"RQC3";
  RIGHT(INDIRECT("C"&ROW());9)="APROVADOS";"RQC2";
  RIGHT(INDIRECT("C"&ROW());13)="NÃO RECEBIDOS";"RQC130";
  RIGHT(INDIRECT("C"&ROW());9)="RECEBIDOS";"RQC1";
  RIGHT(INDIRECT("C"&ROW());12)="PREJUDICADOS";"RQC6";
  RIGHT(INDIRECT("C"&ROW());10)="REJEITADOS";"RQC17";
  RIGHT(INDIRECT("C"&ROW());9)="RELATÓRIO";"RQC18";
  RIGHT(INDIRECT("C"&ROW());9)="RELATORIA";"RQC19";
  RIGHT(INDIRECT("C"&ROW());10)="REITERADOS";"RQC23";
  RIGHT(INDIRECT("C"&ROW());9)="RETIRADOS";"RQC26";
  RIGHT(INDIRECT("C"&ROW());9)="EMENDADOS";"RQC25";
  RIGHT(INDIRECT("C"&ROW());7)="ADIADOS";"RQC15";
  RIGHT(INDIRECT("C"&ROW());10)<>"RECEBIDOS";"RQC3");

  LEFT(INDIRECT("C"&ROW());3)="RQC";IFS(
  RIGHT(INDIRECT("C"&ROW());21)="RECEBIDOS E APROVADOS";"RQC13";
  RIGHT(INDIRECT("C"&ROW());9)="APROVADOS";"RQC12";
  RIGHT(INDIRECT("C"&ROW());13)="NÃO RECEBIDOS";"RQC130";
  RIGHT(INDIRECT("C"&ROW());9)="RECEBIDOS";"RQC11";
  RIGHT(INDIRECT("C"&ROW());12)="PREJUDICADOS";"RQC6";
  RIGHT(INDIRECT("C"&ROW());9)="RELATÓRIO";"RQC18";
  RIGHT(INDIRECT("C"&ROW());10)="REJEITADOS";"RQC17";
  RIGHT(INDIRECT("C"&ROW());10)="REITERADOS";"RQC23";
  RIGHT(INDIRECT("C"&ROW());9)="RETIRADOS";"RQC26";
  RIGHT(INDIRECT("C"&ROW());9)="RELATORIA";"RQC19";
  RIGHT(INDIRECT("C"&ROW());9)="EMENDADOS";"RQC28 // RQC29";
  RIGHT(INDIRECT("C"&ROW());7)="ADIADOS";"RQC16";
  RIGHT(INDIRECT("C"&ROW());10)<>"RECEBIDOS";"RQC13");

  OR(LEFT(INDIRECT("C"&ROW());10)="OFÍCIOS - ";LEFT(INDIRECT("C"&ROW());27)="CORRESPONDÊNCIA: OFÍCIOS - ");IFS(
  RIGHT(INDIRECT("C"&ROW());15)="PROJETOS DE LEI";"PL330";
  RIGHT(INDIRECT("C"&ROW());13)="REQUERIMENTOS";"RQN30";
  RIGHT(INDIRECT("C"&ROW());5)="VETOS";"PL330";
  RIGHT(INDIRECT("C"&ROW());20)="PRORROGAÇÃO DE PRAZO";"RQN67";
  RIGHT(INDIRECT("C"&ROW());33)="PROPOSTA DE EMENDA À CONSTITUIÇÃO";"PL330");

  OR(LEFT(INDIRECT("C"&ROW());42)="OFÍCIO DO TRIBUNAL DE CONTAS QUE ENCAMINHA";
  LEFT(INDIRECT("C"&ROW());69)="TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO TRIBUNAL DE CONTAS QUE ENCAMINHA");IFS(
  RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";"PL??";
  RIGHT(INDIRECT("C"&ROW());23)="RELATÓRIO DE ATIVIDADES";"PL??";
  RIGHT(INDIRECT("C"&ROW());23)="BALANÇO GERAL DO ESTADO";"PL??";
  RIGHT(INDIRECT("C"&ROW());19)="PRESTAÇÃO DE CONTAS";"PL??");

  OR(LEFT(INDIRECT("C"&ROW());29)="OFÍCIO DO TRIBUNAL DE JUSTIÇA";
  LEFT(INDIRECT("C"&ROW());56)="TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO TRIBUNAL DE JUSTIÇA");IFS(
  OR(RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";RIGHT(INDIRECT("C"&ROW());27)="PROJETO DE LEI COMPLEMENTAR");"OFI4");

  OR(LEFT(INDIRECT("C"&ROW());28)="OFÍCIO DA DEFENSORIA PÚBLICA";
  LEFT(INDIRECT("C"&ROW());55)="TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DA DEFENSORIA PÚBLICA");IFS(
  RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";"OFI4");

  OR(LEFT(INDIRECT("C"&ROW());28)="OFÍCIO DO MINISTÉRIO PÚBLICO";
  LEFT(INDIRECT("C"&ROW());55)="TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO MINISTÉRIO PÚBLICO");IFS(
  RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";"OFI??");

  LEFT(INDIRECT("C"&ROW());39)="OFÍCIO DA PROCURADORIA-GERAL DE JUSTIÇA";IFS(
  RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";"OFI4");

  OR(LEFT(INDIRECT("C"&ROW());42)="APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS";LEFT(INDIRECT("C"&ROW());69)="TRAMITAÇÃO DE PROPOSIÇÕES: APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS");IFS(
  RIGHT(INDIRECT("C"&ROW());19)="COMISSÕES TEMÁTICAS";"RQN27";
  RIGHT(INDIRECT("C"&ROW());15)="COM COMUNICAÇÃO";"RQN26";
  RIGHT(INDIRECT("C"&ROW());15)="SEM COMUNICAÇÃO";"RQN85";
  RIGHT(INDIRECT("C"&ROW());8)="ANEXADOS";"RQN40/RQN47";
  RIGHT(INDIRECT("C"&ROW());19)="CIDADANIA HONORÁRIA";"RQN88";
  RIGHT(INDIRECT("C"&ROW());13)="INDICAÇÃO TCE";"RQN??";
  RIGHT(INDIRECT("C"&ROW());25)="ASSEMBLEIA FISCALIZA MAIS";"RQN??";
  RIGHT(INDIRECT("C"&ROW());18)="FRENTE PARLAMENTAR";"RQN??";
  RIGHT(INDIRECT("C"&ROW());24)="INCLUSÃO EM ORDEM DO DIA";"RQN??";
  RIGHT(INDIRECT("C"&ROW());22)="RETIRADA DE TRAMITAÇÃO";"RQN80";
  RIGHT(INDIRECT("C"&ROW());15)="DESARQUIVAMENTO";"RQN26/RQN85";
  RIGHT(INDIRECT("C"&ROW());15)="DESANEXAÇÃO";"RQN??";
  RIGHT(INDIRECT("C"&ROW());17)="COMISSÃO SEGUINTE";"RQN??";
  RIGHT(INDIRECT("C"&ROW());17)="MAIS UMA COMISSÃO";"RQN??";
  RIGHT(INDIRECT("C"&ROW());7)="RECURSO";"RQN??";
  RIGHT(INDIRECT("C"&ROW());21)="PEDIDO DE INFORMAÇÕES";"RQN??";
  RIGHT(INDIRECT("C"&ROW());22)="PEDIDO DE PROVIDÊNCIAS";"RQN??";
  RIGHT(INDIRECT("C"&ROW());14)="PERDA DE PRAZO";"RQN??";
  RIGHT(INDIRECT("C"&ROW());16)="REUNIÃO ESPECIAL";"RQN80";
  RIGHT(INDIRECT("C"&ROW());38)="MESA DA ASSEMBLEIA, VOTADO EM PLENÁRIO";"RQN14";
  RIGHT(INDIRECT("C"&ROW());57)="MESA DA ASSEMBLEIA, PROVIDÊNCIA INTERNA";"RQN92";
  RIGHT(INDIRECT("C"&ROW());19)="DESPACHO A DEPUTADO";"RQN??";
  RIGHT(INDIRECT("C"&ROW());19)="DESPACHO A SERVIDOR";"RQN??";
  RIGHT(INDIRECT("C"&ROW());13)="SETOR DA CASA";"RQO16";
  RIGHT(INDIRECT("C"&ROW());19)<>"COMISSÕES TEMÁTICAS";"RQN27");

  OR(LEFT(INDIRECT("C"&ROW());62)="APRESENTAÇÃO DE PROPOSIÇÕES: PROPOSTA DE EMENDA À CONSTITUIÇÃO";LEFT(INDIRECT("C"&ROW());60)="TRAMITAÇÃO DE PROPOSIÇÕES: PROPOSTA DE EMENDA À CONSTITUIÇÃO");"PEC5";

  OR(LEFT(INDIRECT("C"&ROW());44)="APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE LEI";
  LEFT(INDIRECT("C"&ROW());42)="TRAMITAÇÃO DE PROPOSIÇÕES: PROJETOS DE LEI");IFS(
  RIGHT(INDIRECT("C"&ROW());19)="COMISSÕES TEMÁTICAS";"PL3";
  RIGHT(INDIRECT("C"&ROW());18)="MESA DA ASSEMBLEIA";"PL282";
  RIGHT(INDIRECT("C"&ROW());8)="ANEXADOS";"PL145/PL204");

  OR(LEFT(INDIRECT("C"&ROW());50)="APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO";
  LEFT(INDIRECT("C"&ROW());48)="TRAMITAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO");IFS(
  RIGHT(INDIRECT("C"&ROW());29)="REGIME ESPECIAL DE TRIBUTAÇÃO";"PRE140";
  RIGHT(INDIRECT("C"&ROW());19)="APROVAÇÃO DE CONTAS";"PRE137";
  RIGHT(INDIRECT("C"&ROW());24)="RATIFICAÇÃO DE CONVÊNIOS";"PRE9";
  RIGHT(INDIRECT("C"&ROW());37)="ESTRUTURA DA SECRETARIA DA ASSEMBLEIA";"PRE134";
  RIGHT(INDIRECT("C"&ROW());19)="COMISSÕES TEMÁTICAS";"PL3";
  RIGHT(INDIRECT("C"&ROW());8)="ANEXADOS";"PL145/PL204";
  RIGHT(INDIRECT("C"&ROW());19)="CIDADANIA HONORÁRIA";"PRE11";
  RIGHT(INDIRECT("C"&ROW());18)="CALAMIDADE PÚBLICA";"PRE12";
  RIGHT(INDIRECT("C"&ROW());21)="LICENÇA AO GOVERNADOR";"PRE13"
  );

  LEFT(INDIRECT("C"&ROW());25)="PROPOSIÇÕES NÃO RECEBIDAS";IFS(
  RIGHT(INDIRECT("C"&ROW());15)="PROJETOS DE LEI";"PL130";
  RIGHT(INDIRECT("C"&ROW());13)="REQUERIMENTOS";"RQN130";
  RIGHT(INDIRECT("C"&ROW());15)<>"PROJETOS DE LEI";"PL130");

  LEFT(INDIRECT("C"&ROW());25)="RECEBIMENTO DE PROPOSIÇÃO";"PL367";
  LEFT(INDIRECT("C"&ROW());23)="DESIGNAÇÃO DE RELATORIA";"PL264";
  LEFT(INDIRECT("C"&ROW());25)="CUMPRIMENTO DE DILIGÊNCIA";"PL373";
  LEFT(INDIRECT("C"&ROW());32)="REUNIÃO COM DEBATE DE PROPOSIÇÃO";"RQC7";
  LEFT(INDIRECT("C"&ROW());33)="REUNIÃO ORIGINADA DE REQUERIMENTO";"RQC5";
  INDIRECT("C"&ROW())="PAUTA COMPLETA DE REUNIÃO COM DEBATE DE PROPOSIÇÃO";"RQC7";
  INDIRECT("C"&ROW())="RESULTADO COMPLETO DE REUNIÃO COM DEBATE DE PROPOSIÇÃO";"RQC8";
  INDIRECT("C"&ROW())="INÍCIO DE APRECIAÇÃO NA PRÓXIMA COMISSÃO";"RQC??";
  INDIRECT("C"&ROW())="CONGRATULAÇÕES ENTREGUES EM REUNIÃO";"RQN18";
  INDIRECT("C"&ROW())="ENTREGA DE DIPLOMA";"RQN18";
  LEFT(INDIRECT("C"&ROW());16)="CONSULTA PÚBLICA";"PL532";

  INDIRECT("C"&ROW())="PROPOSIÇÃO DE LEI ENCAMINHADA PARA SANÇÃO";"PL63";
  INDIRECT("C"&ROW())="REMESSA - PEDIDO DE INFORMAÇÃO";"RQN20";
  INDIRECT("C"&ROW())="REMESSA - REQUERIMENTO APROVADO";"RQN17";
  INDIRECT("C"&ROW())="OFÍCIO - PEDIDO DE INFORMAÇÃO";"RQN20";
  INDIRECT("C"&ROW())="OFÍCIO - PEDIDO DE INFORMAÇÃO";"PL71";
  INDIRECT("C"&ROW())="OFÍCIO - REQUERIMENTO APROVADO";"RQN17";
  INDIRECT("C"&ROW())="OFÍCIO - VOTO DE CONGRATULAÇÕES";"RQN48";
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE APLAUSO";"RQN50";
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE APOIO";"RQN50";
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE REPÚDIO";"RQN50";
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE PROTESTO";"RQN50";
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE PESAR";"RQN49";
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE REPÚDIO";"RQN50";
  INDIRECT("C"&ROW())="OFÍCIO COMUNICANDO MANUTENÇÃO TOTAL DO VETO";"PL93";
  INDIRECT("C"&ROW())="OFÍCIO COMUNICANDO REJEIÇÃO TOTAL DO VETO";"PL92";
  INDIRECT("C"&ROW())="OFÍCIO COMUNICANDO REJEIÇÃO PARCIAL DO VETO";"PL598";
  INDIRECT("C"&ROW())="OFÍCIO COMUNICANDO APROVAÇÃO DA INDICAÇÃO";"IND10";
  INDIRECT("C"&ROW())="OFÍCIO ENCAMINHADO AOS DESTINATÁRIOS POR E-MAIL";"RQN31";

  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: PROJETOS DE LEI";"PL66";
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: REQUERIMENTOS, APROVADOS";"RQN16";
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: REQUERIMENTOS, REJEITADOS";"RQN28";
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: REQUERIMENTOS, RECURSO";"RQN87";
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: INCONSTITUCIONALIDADE";"PL125"

  ))''']]
    * ((footer_start - 1) - 5)
)

# ------------------------------------------------------------------
# R6 / S6 (mega fórmulas) - preservadas integralmente
# ------------------------------------------------------------------
add(f"R6:R{footer_start - 1}", [[r'''=IFS(

  OR(
  INDIRECT("C"&ROW())="";
  LEFT(INDIRECT("C"&ROW());6)="DIÁRIO";
  LEFT(INDIRECT("C"&ROW());8)="REUNIÕES";
  INDIRECT("C"&ROW())="REQUERIMENTOS DE COMISSÃO";
  INDIRECT("C"&ROW())="LANÇAMENTOS DE TRAMITAÇÃO";
  INDIRECT("C"&ROW())="CADASTRO DE E-MAILS";
  INDIRECT("C"&ROW())="OFÍCIOS DA SECRETARIA-GERAL DA MESA";
  INDIRECT("C"&ROW())="LANÇAMENTOS DE PRECLUSÃO DE PRAZO";
  INDIRECT("C"&ROW())="IMPLANTAÇÃO DE TEXTOS");"";

  OR(
  INDIRECT("C"&ROW())="ALINE";
  INDIRECT("C"&ROW())="ANDRÉ";
  INDIRECT("C"&ROW())="DIOGO";
  INDIRECT("C"&ROW())="KÁTIA";
  INDIRECT("C"&ROW())="LEO";
  INDIRECT("C"&ROW())="WELDER";
  INDIRECT("C"&ROW())="TOTAL";
  ISNUMBER(INDIRECT("C"&ROW()));
  INDIRECT("C"&ROW())="?";
  INDIRECT("K"&ROW())="-";
  INDIRECT("K"&ROW())="cancelada";
  INDIRECT("K"&ROW())="sem quórum";
  INDIRECT("K"&ROW())="não publicado");"-";

  OR(
  INDIRECT("C"&ROW())="-";
  INDIRECT("C"&ROW())="VOTAÇÕES NOMINAIS";
  RIGHT(INDIRECT("C"&ROW());18)="EMENDAS PUBLICADAS";
  LEFT(INDIRECT("C"&ROW());17)="VOTAÇÕES NOMINAIS");"-";


  LEFT(INDIRECT("C"&ROW());6)<>"DIÁRIO";
  IFS(

  OR(INDIRECT("C"&ROW())="EMENDA À CONSTITUIÇÃO PROMULGADA";INDIRECT("C"&ROW())="EMENDAS À CONSTITUIÇÃO PROMULGADAS");"3.3.1-D";
  OR(INDIRECT("C"&ROW())="PROPOSTA DE AÇÃO LEGISLATIVA";INDIRECT("C"&ROW())="PROPOSTAS DE AÇÃO LEGISLATIVA");"4.2.6";
  RIGHT(INDIRECT("C"&ROW());48)="PROPOSTAS DE AÇÃO LEGISLATIVA REFERENTES AO PPAG";"4.2.6-B";
  RIGHT(INDIRECT("C"&ROW());24)="VOTAÇÃO DE REQUERIMENTOS";"4.11";
  INDIRECT("C"&ROW())="MANIFESTAÇÕES";"4.14";
  RIGHT(INDIRECT("C"&ROW());24)="REQUERIMENTOS APROVADOS";"4.15";
  OR(INDIRECT("C"&ROW())="ERRATAS";INDIRECT("C"&ROW())="ERRATA");"4.16";
  INDIRECT("C"&ROW())="DECISÕES DA PRESIDÊNCIA";"4.4";
  OR(LEFT(INDIRECT("C"&ROW());22)="DECISÃO DA PRESIDÊNCIA";LEFT(INDIRECT("C"&ROW());49)="TRAMITAÇÃO DE PROPOSIÇÕES: DECISÃO DA PRESIDÊNCIA");"4.4";
  INDIRECT("C"&ROW())="TRAMITAÇÃO DE PROPOSIÇÕES: PARECERES";"6";
  INDIRECT("C"&ROW())="PARECERES SOBRE VETO";"12.7";
  INDIRECT("C"&ROW())="PARECERES SOBRE SUBSTITUTIVO";"-";
  OR(LEFT(INDIRECT("C"&ROW());3)="LEI");"7.1";
  INDIRECT("C"&ROW())="ORDINÁRIA";"10.8";

  INDIRECT("C"&ROW())="VETO TOTAL A PROPOSIÇÃO DE LEI";"7.2";
  LEFT(INDIRECT("C"&ROW());32)="VETO PARCIAL A PROPOSIÇÃO DE LEI";"7.2";
  OR(INDIRECT("C"&ROW())="RESOLUÇÃO";INDIRECT("C"&ROW())="RESOLUÇÕES");"3.3.1-A";
  INDIRECT("C"&ROW())="PROPOSIÇÕES DE LEI";"3.3.2";
  INDIRECT("C"&ROW())="DECISÃO DA MESA";"???";
  OR(INDIRECT("C"&ROW())="DESIGNAÇÃO DE COMISSÕES";INDIRECT("C"&ROW())="TRAMITAÇÃO DE PROPOSIÇÕES: DESIGNAÇÃO DE COMISSÕES");"4.7";
  INDIRECT("C"&ROW())="OFÍCIOS DE PREFEITURA QUE ENCAMINHAM DECRETOS DE CALAMIDADE PÚBLICA";"???";
  INDIRECT("C"&ROW())="PROPOSIÇÃO: REQUERIMENTOS - INDICAÇÃO TCE";"???";
  INDIRECT("C"&ROW())="SOLENE";"-";
  INDIRECT("C"&ROW())="ESPECIAL";"??";
  INDIRECT("C"&ROW())="EXTRAORDINÁRIA";"???";
  INDIRECT("C"&ROW())="EXTRAORDINÁRIA: PARECERES DE REDAÇÃO FINAL APROVADOS";"PL62";
  RIGHT(INDIRECT("C"&ROW());23)="LEITURA DE COMUNICAÇÕES";"-";
  OR(INDIRECT("C"&ROW())="EMENDAS OU SUBSTITUTIVOS PUBLICADOS";INDIRECT("C"&ROW())="EMENDAS NÃO RECEBIDAS PUBLICADAS");"???";
  LEFT(INDIRECT("C"&ROW());16)="REUNIÃO CONJUNTA";"RQN??";
  OR(LEFT(INDIRECT("C"&ROW());19)="RELATÓRIO DE VISITA";LEFT(INDIRECT("C"&ROW());46)="TRAMITAÇÃO DE PROPOSIÇÕES: RELATÓRIO DE VISITA");"5.3";
  RIGHT(INDIRECT("C"&ROW());33)="RELATÓRIO DE EVENTO INSTITUCIONAL";"4.2.9";
  RIGHT(INDIRECT("C"&ROW());24)="REQUERIMENTOS ORDINÁRIOS";"RQO1";
  RIGHT(INDIRECT("C"&ROW());22)="PALAVRAS DO PRESIDENTE";"4.5";
  RIGHT(INDIRECT("C"&ROW());26)="COMUNICAÇÃO DA PRESIDÊNCIA";"4.9";
  LEFT(INDIRECT("C"&ROW());25)="DESPACHO DE REQUERIMENTOS";"4.10";
  LEFT(INDIRECT("C"&ROW());17)="AUDIÊNCIA PÚBLICA";"???";
  LEFT(INDIRECT("C"&ROW());23)="AUDIÊNCIA DE CONVIDADOS";"???";

  OR(LEFT(INDIRECT("C"&ROW());36)="MENSAGEM DO GOVERNADOR QUE ENCAMINHA";
  LEFT(INDIRECT("C"&ROW());38)="MENSAGENS DO GOVERNADOR QUE ENCAMINHAM";
  LEFT(INDIRECT("C"&ROW());63)="TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA";
  LEFT(INDIRECT("C"&ROW());65)="TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGENS DO GOVERNADOR QUE ENCAMINHAM";
  LEFT(INDIRECT("C"&ROW());35)="MENSAGEM DO GOVERNADOR QUE COMUNICA";
  LEFT(INDIRECT("C"&ROW());62)="TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE COMUNICA";
  LEFT(INDIRECT("C"&ROW());35)="MENSAGEM DO GOVERNADOR QUE SOLICITA");IFS(
  OR(RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";RIGHT(INDIRECT("C"&ROW());15)="PROJETOS DE LEI";RIGHT(INDIRECT("C"&ROW());36)="PROJETO DE LEI - CRÉDITO SUPLEMENTAR");"4.2.1-A";
  RIGHT(INDIRECT("C"&ROW());42)="EMENDA OU SUBSTITUTIVO COM DESPACHO À MESA";"4.2.1-B";
  RIGHT(INDIRECT("C"&ROW());41)="EMENDA OU SUBSTITUTIVO COM DESPACHO À FFO";"4.2.1-B";
  RIGHT(INDIRECT("C"&ROW());12)="VETO PARCIAL";"4.2.1-C";
  RIGHT(INDIRECT("C"&ROW());10)="VETO TOTAL";"4.2.1-C";
  RIGHT(INDIRECT("C"&ROW());29)="REGIME ESPECIAL DE TRIBUTAÇÃO";"4.2.1-D";
  RIGHT(INDIRECT("C"&ROW());9)="INDICAÇÃO";"4.2.1-F";
  RIGHT(INDIRECT("C"&ROW());28)="PEDIDO DE REGIME DE URGÊNCIA";"4.2.1-G";
  RIGHT(INDIRECT("C"&ROW());18)="CONVÊNIO DO CONFAZ";"4.2.1-I";
  RIGHT(INDIRECT("C"&ROW());16)="CONVÊNIO DO ICMS";"4.2.1-I";
  RIGHT(INDIRECT("C"&ROW());44)="PRESTAÇÃO DE CONTAS DA ADMINISTRAÇÃO PÚBLICA";"4.2.1-L";
  RIGHT(INDIRECT("C"&ROW());36)="RELATÓRIO SOBRE A SITUAÇÃO DO ESTADO";"4.2.2-L.2";
  RIGHT(INDIRECT("C"&ROW());29)="DESARQUIVAMENTO DE PROPOSIÇÃO";"4.2.1-T";
  RIGHT(INDIRECT("C"&ROW());19)="RETIRADA DE PROJETO";"4.2.1-M";
  RIGHT(INDIRECT("C"&ROW());16)="AUSÊNCIA DO PAÍS";"4.2.2-C.2"
  );

  OR(LEFT(INDIRECT("C"&ROW());20)="OFÍCIO DO GOVERNADOR";
  MID(INDIRECT("C"&ROW());28;20)="OFÍCIO DO GOVERNADOR";
  LEFT(INDIRECT("C"&ROW());25)="OFÍCIO DO VICE-GOVERNADOR");IFS(
  RIGHT(INDIRECT("C"&ROW());28)="COMUNICANDO AUSÊNCIA DO PAÍS";"4.2.2-C.2";
  RIGHT(INDIRECT("C"&ROW());35)="COMUNICANDO QUE ENCAMINHOU MENSAGEM";"4.2.2-C");

  OR(LEFT(INDIRECT("C"&ROW());8)="COMISSÃO";LEFT(INDIRECT("C"&ROW());4)="CIPE");IFS(
  MID(INDIRECT("C"&ROW());13;12)="CONSTITUIÇÃO";"10.2";
  MID(INDIRECT("C"&ROW());13;12)="FISCALIZAÇÃO";"10.3";
  MID(INDIRECT("C"&ROW());13;12)="PARTICIPAÇÃO";"10.4";
  MID(INDIRECT("C"&ROW());13;7)="REDAÇÃO";"10.5";
  MID(INDIRECT("C"&ROW());11;8)="ESPECIAL";"10.6";
  MID(INDIRECT("C"&ROW());13;12)<>"CONSTITUIÇÃO";"10");

  LEFT(INDIRECT("C"&ROW());27)="REQUERIMENTOS DE COMISSÃO: ";IFS(
  RIGHT(INDIRECT("C"&ROW());9)="RECEBIDOS";"5.2.1";
  RIGHT(INDIRECT("C"&ROW());14)="VOTAÇÃO ADIADA";"5.2.2";
  RIGHT(INDIRECT("C"&ROW());9)="APROVADOS";"5.2.3";
  RIGHT(INDIRECT("C"&ROW());9)="EMENDADOS";"5.2.3.1";
  RIGHT(INDIRECT("C"&ROW());7)="ADIADOS";"5.2.2";
  RIGHT(INDIRECT("C"&ROW());9)="RETIRADOS";"5.2.4";
  RIGHT(INDIRECT("C"&ROW());12)="PREJUDICADOS";"5.2.5";
  RIGHT(INDIRECT("C"&ROW());10)="ARQUIVADOS";"5.2.6";
  RIGHT(INDIRECT("C"&ROW());16)="REUNIÃO CONJUNTA";"5.2.7";
  RIGHT(INDIRECT("C"&ROW());11)="RATIFICADOS";"5.2.8";
  RIGHT(INDIRECT("C"&ROW());25)="ASSEMBLEIA FISCALIZA MAIS";"5.2.9";
  RIGHT(INDIRECT("C"&ROW());9)="RELATÓRIO";"5.3";
  RIGHT(INDIRECT("C"&ROW());10)="REITERADOS";"5.5";
  RIGHT(INDIRECT("C"&ROW());10)="REJEITADOS";"5.6";
  RIGHT(INDIRECT("C"&ROW());10)<>"RECEBIDOS";"5.2.3");

  LEFT(INDIRECT("C"&ROW());3)="RQC";IFS(
  RIGHT(INDIRECT("C"&ROW());9)="RECEBIDOS";"5.2.1";
  RIGHT(INDIRECT("C"&ROW());14)="VOTAÇÃO ADIADA";"5.2.2";
  RIGHT(INDIRECT("C"&ROW());9)="APROVADOS";"5.2.3";
  RIGHT(INDIRECT("C"&ROW());9)="EMENDADOS";"5.2.3.1";
  RIGHT(INDIRECT("C"&ROW());7)="ADIADOS";"5.2.2";
  RIGHT(INDIRECT("C"&ROW());9)="RETIRADOS";"5.2.4";
  RIGHT(INDIRECT("C"&ROW());12)="PREJUDICADOS";"5.2.5";
  RIGHT(INDIRECT("C"&ROW());10)="ARQUIVADOS";"5.2.6";
  RIGHT(INDIRECT("C"&ROW());16)="REUNIÃO CONJUNTA";"5.2.7";
  RIGHT(INDIRECT("C"&ROW());11)="RATIFICADOS";"5.2.8";
  RIGHT(INDIRECT("C"&ROW());25)="ASSEMBLEIA FISCALIZA MAIS";"5.2.9";
  RIGHT(INDIRECT("C"&ROW());9)="RELATÓRIO";"5.3";
  RIGHT(INDIRECT("C"&ROW());10)="REITERADOS";"5.5";
  RIGHT(INDIRECT("C"&ROW());10)="REJEITADOS";"5.6";
  RIGHT(INDIRECT("C"&ROW());10)<>"RECEBIDOS";"5.2");

  OR(LEFT(INDIRECT("C"&ROW());10)="OFÍCIOS - ";LEFT(INDIRECT("C"&ROW());27)="CORRESPONDÊNCIA: OFÍCIOS - ");IFS(
  RIGHT(INDIRECT("C"&ROW());15)="PROJETOS DE LEI";"4.1.2";
  RIGHT(INDIRECT("C"&ROW());13)="REQUERIMENTOS";"4.1.1";
  RIGHT(INDIRECT("C"&ROW());5)="VETOS";"4.1.2";
  RIGHT(INDIRECT("C"&ROW());20)="PRORROGAÇÃO DE PRAZO";"4.1.4";
  RIGHT(INDIRECT("C"&ROW());33)="PROPOSTA DE EMENDA À CONSTITUIÇÃO";"4.1.2");

  OR(LEFT(INDIRECT("C"&ROW());42)="OFÍCIO DO TRIBUNAL DE CONTAS QUE ENCAMINHA";
  LEFT(INDIRECT("C"&ROW());69)="TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO TRIBUNAL DE CONTAS QUE ENCAMINHA");IFS(
  RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";"4.2.2-A.1";
  RIGHT(INDIRECT("C"&ROW());23)="RELATÓRIO DE ATIVIDADES";"4.2.2-A.3";
  RIGHT(INDIRECT("C"&ROW());23)="BALANÇO GERAL DO ESTADO";"4.2.2-A.4";
  RIGHT(INDIRECT("C"&ROW());19)="PRESTAÇÃO DE CONTAS";"4.2.2-A.5");

  OR(LEFT(INDIRECT("C"&ROW());29)="OFÍCIO DO TRIBUNAL DE JUSTIÇA";
  LEFT(INDIRECT("C"&ROW());56)="TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO TRIBUNAL DE JUSTIÇA");IFS(
  OR(RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";RIGHT(INDIRECT("C"&ROW());27)="PROJETO DE LEI COMPLEMENTAR");"4.2.2-B.1");

  OR(LEFT(INDIRECT("C"&ROW());28)="OFÍCIO DA DEFENSORIA PÚBLICA";
  LEFT(INDIRECT("C"&ROW());55)="TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DA DEFENSORIA PÚBLICA");IFS(
  RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";"4.2.2-F.1");

  OR(LEFT(INDIRECT("C"&ROW());28)="OFÍCIO DO MINISTÉRIO PÚBLICO";
  LEFT(INDIRECT("C"&ROW());55)="TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO MINISTÉRIO PÚBLICO");IFS(
  RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";"4.2.2-D.1");

  LEFT(INDIRECT("C"&ROW());39)="OFÍCIO DA PROCURADORIA-GERAL DE JUSTIÇA";IFS(
  RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";"4.2.2-D.1");

  OR(LEFT(INDIRECT("C"&ROW());42)="APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS";LEFT(INDIRECT("C"&ROW());69)="TRAMITAÇÃO DE PROPOSIÇÕES: APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS");IFS(
  RIGHT(INDIRECT("C"&ROW());19)="COMISSÕES TEMÁTICAS";"4.2.7-A";
  RIGHT(INDIRECT("C"&ROW());15)="COM COMUNICAÇÃO";"4.2.7-B";
  RIGHT(INDIRECT("C"&ROW());15)="SEM COMUNICAÇÃO";"4.2.7-C";
  RIGHT(INDIRECT("C"&ROW());8)="ANEXADOS";"4.2.7-D";
  RIGHT(INDIRECT("C"&ROW());19)="CIDADANIA HONORÁRIA";"4.2.7-E";
  RIGHT(INDIRECT("C"&ROW());13)="INDICAÇÃO TCE";"4.2.7-F";
  RIGHT(INDIRECT("C"&ROW());25)="ASSEMBLEIA FISCALIZA MAIS";"4.2.7-G";
  RIGHT(INDIRECT("C"&ROW());18)="FRENTE PARLAMENTAR";"4.2.7-H";
  RIGHT(INDIRECT("C"&ROW());24)="INCLUSÃO EM ORDEM DO DIA";"4.2.7-I.1";
  RIGHT(INDIRECT("C"&ROW());22)="RETIRADA DE TRAMITAÇÃO";"4.2.7-I.2";
  RIGHT(INDIRECT("C"&ROW());15)="DESARQUIVAMENTO";"4.2.7-I.3";
  RIGHT(INDIRECT("C"&ROW());15)="DESANEXAÇÃO";"4.2.7-I.4";
  RIGHT(INDIRECT("C"&ROW());17)="COMISSÃO SEGUINTE";"4.2.7-I.5";
  RIGHT(INDIRECT("C"&ROW());17)="MAIS UMA COMISSÃO";"4.2.7-I.6";
  RIGHT(INDIRECT("C"&ROW());7)="RECURSO";"4.2.7-I.7";
  RIGHT(INDIRECT("C"&ROW());21)="PEDIDO DE INFORMAÇÕES";"4.2.7-I";
  RIGHT(INDIRECT("C"&ROW());22)="PEDIDO DE PROVIDÊNCIAS";"4.2.7-I";
  RIGHT(INDIRECT("C"&ROW());14)="PERDA DE PRAZO";"4.2.7-I";
  RIGHT(INDIRECT("C"&ROW());16)="REUNIÃO ESPECIAL";"4.2.7-I";
  RIGHT(INDIRECT("C"&ROW());38)="MESA DA ASSEMBLEIA, VOTADO EM PLENÁRIO";"4.2.7-J";
  RIGHT(INDIRECT("C"&ROW());57)="MESA DA ASSEMBLEIA, ENCAMINHADOS PARA PROVIDÊNCIA INTERNA";"4.2.7-J";
  RIGHT(INDIRECT("C"&ROW());19)="DESPACHO A DEPUTADO";"4.2.7-K.2.1";
  RIGHT(INDIRECT("C"&ROW());19)="DESPACHO A SERVIDOR";"4.2.7-K.2.2";
  RIGHT(INDIRECT("C"&ROW());13)="SETOR DA CASA";"4.2.7-K.2.3";
  RIGHT(INDIRECT("C"&ROW());19)<>"COMISSÕES TEMÁTICAS";"4.2.7-A");

  OR(LEFT(INDIRECT("C"&ROW());62)="APRESENTAÇÃO DE PROPOSIÇÕES: PROPOSTA DE EMENDA À CONSTITUIÇÃO";LEFT(INDIRECT("C"&ROW());60)="TRAMITAÇÃO DE PROPOSIÇÕES: PROPOSTA DE EMENDA À CONSTITUIÇÃO");"4.2.3";

  OR(LEFT(INDIRECT("C"&ROW());44)="APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE LEI";
  LEFT(INDIRECT("C"&ROW());42)="TRAMITAÇÃO DE PROPOSIÇÕES: PROJETOS DE LEI");IFS(
  RIGHT(INDIRECT("C"&ROW());19)="COMISSÕES TEMÁTICAS";"4.2.4-A";
  RIGHT(INDIRECT("C"&ROW());8)="ANEXADOS";"4.2.4-B";
  RIGHT(INDIRECT("C"&ROW());18)="MESA DA ASSEMBLEIA";"4.2.4-D");

  OR(LEFT(INDIRECT("C"&ROW());50)="APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO";
  LEFT(INDIRECT("C"&ROW());48)="TRAMITAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO");IFS(
  RIGHT(INDIRECT("C"&ROW());29)="REGIME ESPECIAL DE TRIBUTAÇÃO";"4.2.5-A";
  RIGHT(INDIRECT("C"&ROW());19)="APROVAÇÃO DE CONTAS";"4.2.5-B";
  RIGHT(INDIRECT("C"&ROW());24)="RATIFICAÇÃO DE CONVÊNIOS";"4.2.5-C";
  RIGHT(INDIRECT("C"&ROW());37)="ESTRUTURA DA SECRETARIA DA ASSEMBLEIA";"4.2.5-D";
  RIGHT(INDIRECT("C"&ROW());19)="COMISSÕES TEMÁTICAS";"4.2.5-E";
  RIGHT(INDIRECT("C"&ROW());8)="ANEXADOS";"4.2.5-F";
  RIGHT(INDIRECT("C"&ROW());19)="CIDADANIA HONORÁRIA";"4.2.5-G";
  RIGHT(INDIRECT("C"&ROW());18)="CALAMIDADE PÚBLICA";"4.2.5-H";
  RIGHT(INDIRECT("C"&ROW());21)="LICENÇA AO GOVERNADOR";"4.2.5-I"
  );

  LEFT(INDIRECT("C"&ROW());25)="PROPOSIÇÕES NÃO RECEBIDAS";IFS(
  RIGHT(INDIRECT("C"&ROW());15)="PROJETOS DE LEI";"4.3";
  RIGHT(INDIRECT("C"&ROW());13)="REQUERIMENTOS";"4.3";
  RIGHT(INDIRECT("C"&ROW());15)<>"PROJETOS DE LEI";"4.3");

  INDIRECT("C"&ROW())="CONGRATULAÇÕES ENTREGUES EM REUNIÃO";"8.4";
  INDIRECT("C"&ROW())="ENTREGA DE DIPLOMA";"8.4";
  LEFT(INDIRECT("C"&ROW());25)="RECEBIMENTO DE PROPOSIÇÃO";"9.1";
  LEFT(INDIRECT("C"&ROW());23)="DESIGNAÇÃO DE RELATORIA";"9.2";
  LEFT(INDIRECT("C"&ROW());25)="CUMPRIMENTO DE DILIGÊNCIA";"9.3";
  LEFT(INDIRECT("C"&ROW());32)="REUNIÃO COM DEBATE DE PROPOSIÇÃO";"9.8";
  LEFT(INDIRECT("C"&ROW());33)="REUNIÃO ORIGINADA DE REQUERIMENTO";"9.7";
  INDIRECT("C"&ROW())="PAUTA COMPLETA DE REUNIÃO COM DEBATE DE PROPOSIÇÃO";"9.8.1";
  INDIRECT("C"&ROW())="RESULTADO COMPLETO DE REUNIÃO COM DEBATE DE PROPOSIÇÃO";"9.8.2";
  LEFT(INDIRECT("C"&ROW());16)="CONSULTA PÚBLICA";"9.9";

  INDIRECT("C"&ROW())="REMESSA - PEDIDO DE INFORMAÇÃO";"8.6";
  INDIRECT("C"&ROW())="REMESSA - REQUERIMENTO APROVADO";"8.5";
  INDIRECT("C"&ROW())="OFÍCIO - PEDIDO DE INFORMAÇÃO";"8.6";
  INDIRECT("C"&ROW())="OFÍCIO - PEDIDO DE INFORMAÇÃO";"8.7";
  INDIRECT("C"&ROW())="OFÍCIO - REQUERIMENTO APROVADO";"8.5";
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE APLAUSO";"8.1";
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE APOIO";"8.1";
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE REPÚDIO";"8.1";
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE PROTESTO";"8.1";
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE PESAR";"8.2";
  INDIRECT("C"&ROW())="OFÍCIO - VOTO DE CONGRATULAÇÕES";"8.3";
  INDIRECT("C"&ROW())="OFÍCIO COMUNICANDO MANUTENÇÃO TOTAL DO VETO";"8.9";
  INDIRECT("C"&ROW())="OFÍCIO COMUNICANDO REJEIÇÃO TOTAL DO VETO";"8.10";
  INDIRECT("C"&ROW())="OFÍCIO COMUNICANDO REJEIÇÃO PARCIAL DO VETO";"8.11";
  INDIRECT("C"&ROW())="OFÍCIO COMUNICANDO APROVAÇÃO DA INDICAÇÃO";"8.11";
  INDIRECT("C"&ROW())="OFÍCIO ENCAMINHADO AOS DESTINATÁRIOS POR E-MAIL";"8.12";

  INDIRECT("C"&ROW())="PROPOSIÇÃO DE LEI ENCAMINHADA PARA SANÇÃO";"8.8";
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: PROJETOS DE LEI";"11.1.1";
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: REQUERIMENTOS, APROVADOS";"11.1.2-A";
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: REQUERIMENTOS, REJEITADOS";"11.1.2-B";
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: REQUERIMENTOS, RECURSO";"11.1.2-C";
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: INCONSTITUCIONALIDADE";"11.2"

  ))''']] * ((footer_start - 1) - 5))

add(f"S6:S{footer_start - 1}", [[r'''=IFS(

  OR(
  INDIRECT("C"&ROW())="";
  LEFT(INDIRECT("C"&ROW());6)="DIÁRIO";
  LEFT(INDIRECT("C"&ROW());8)="REUNIÕES";
  INDIRECT("C"&ROW())="REQUERIMENTOS DE COMISSÃO";
  INDIRECT("C"&ROW())="LANÇAMENTOS DE TRAMITAÇÃO";
  INDIRECT("C"&ROW())="CADASTRO DE E-MAILS";
  INDIRECT("C"&ROW())="OFÍCIOS DA SECRETARIA-GERAL DA MESA";
  INDIRECT("C"&ROW())="LANÇAMENTOS DE PRECLUSÃO DE PRAZO";
  INDIRECT("C"&ROW())="IMPLANTAÇÃO DE TEXTOS");"";

  OR(
  INDIRECT("C"&ROW())="ALINE";
  INDIRECT("C"&ROW())="ANDRÉ";
  INDIRECT("C"&ROW())="DIOGO";
  INDIRECT("C"&ROW())="KÁTIA";
  INDIRECT("C"&ROW())="LEO";
  INDIRECT("C"&ROW())="WELDER";
  INDIRECT("C"&ROW())="TOTAL";
  INDIRECT("C"&ROW())="?";
  ISNUMBER(INDIRECT("C"&ROW()));
  INDIRECT("K"&ROW())="-";
  INDIRECT("K"&ROW())="cancelada";
  INDIRECT("K"&ROW())="sem quórum";
  INDIRECT("K"&ROW())="não publicado");"-";

  OR(
  INDIRECT("C"&ROW())="-";
  INDIRECT("C"&ROW())="VOTAÇÕES NOMINAIS";
  RIGHT(INDIRECT("C"&ROW());18)="EMENDAS PUBLICADAS";
  LEFT(INDIRECT("C"&ROW());17)="VOTAÇÕES NOMINAIS");"-";

  LEFT(INDIRECT("C"&ROW());6)<>"DIÁRIO";IFS(

  INDIRECT("C"&ROW())="IMPLANTAÇÃO DE TEXTOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=27";"PÁG 27");
  OR(INDIRECT("C"&ROW())="EMENDA À CONSTITUIÇÃO PROMULGADA";INDIRECT("C"&ROW())="EMENDAS À CONSTITUIÇÃO PROMULGADAS");HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=47";"PÁG 47");
  OR(INDIRECT("C"&ROW())="PROPOSTA DE AÇÃO LEGISLATIVA";INDIRECT("C"&ROW())="PROPOSTAS DE AÇÃO LEGISLATIVA");HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=194";"PÁG 194");
  RIGHT(INDIRECT("C"&ROW());48)="PROPOSTAS DE AÇÃO LEGISLATIVA REFERENTES AO PPAG";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=195";"PÁG 195");
  INDIRECT("C"&ROW())="ESPECIAL";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=371";"PÁG 371");
  INDIRECT("C"&ROW())="ORDINÁRIA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=386";"PÁG 386");
  INDIRECT("C"&ROW())="EXTRAORDINÁRIA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=386";"PÁG 386");
  INDIRECT("C"&ROW())="SOLENE";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=386";"PÁG 386");
  OR(INDIRECT("C"&ROW())="RESOLUÇÃO";INDIRECT("C"&ROW())="RESOLUÇÕES");HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=43";"PÁG 43");
  INDIRECT("C"&ROW())="PROPOSIÇÕES DE LEI";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=48";"PÁG 48");
  INDIRECT("C"&ROW())="DECISÃO DA MESA";"PÁG ??";
  INDIRECT("C"&ROW())="DECISÕES DA PRESIDÊNCIA";"PÁG ??";
  INDIRECT("C"&ROW())="MANIFESTAÇÕES";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=276";"PÁG 276");
  RIGHT(INDIRECT("C"&ROW());24)="VOTAÇÃO DE REQUERIMENTOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=265";"PÁG 265");
  INDIRECT("C"&ROW())="OFÍCIOS DE PREFEITURA QUE ENCAMINHAM DECRETOS DE CALAMIDADE PÚBLICA";"PÁG ??";
  LEFT(INDIRECT("C"&ROW());16)="REUNIÃO CONJUNTA";"PÁG ??";
  INDIRECT("C"&ROW())="VETO TOTAL A PROPOSIÇÃO DE LEI";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=303";"PÁG 303");
  LEFT(INDIRECT("C"&ROW());32)="VETO PARCIAL A PROPOSIÇÃO DE LEI";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=303";"PÁG 303");
  OR(INDIRECT("C"&ROW())="DESIGNAÇÃO DE COMISSÕES";INDIRECT("C"&ROW())="TRAMITAÇÃO DE PROPOSIÇÕES: DESIGNAÇÃO DE COMISSÕES");HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=251";"PÁG 251");
  INDIRECT("C"&ROW())="PROPOSIÇÃO: REQUERIMENTOS - INDICAÇÃO TCE";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=197";"PÁG 197");
  INDIRECT("C"&ROW())="TRAMITAÇÃO DE PROPOSIÇÕES: PARECERES";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=301";"PÁG 301");
  INDIRECT("C"&ROW())="PARECERES SOBRE VETO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=434";"PÁG 434");
  INDIRECT("C"&ROW())="PARECERES SOBRE SUBSTITUTIVO";"PÁG ??";
  OR(INDIRECT("C"&ROW())="ERRATAS";INDIRECT("C"&ROW())="ERRATA");HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=279";"PÁG 279");
  RIGHT(INDIRECT("C"&ROW());23)="LEITURA DE COMUNICAÇÕES";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=255";"PÁG 255");
  OR(LEFT(INDIRECT("C"&ROW());3)="LEI");HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=315";"PÁG 315");
  OR(INDIRECT("C"&ROW())="EMENDAS OU SUBSTITUTIVOS PUBLICADOS";INDIRECT("C"&ROW())="EMENDAS NÃO RECEBIDAS PUBLICADAS");"PÁG ??";
  OR(LEFT(INDIRECT("C"&ROW());19)="RELATÓRIO DE VISITA";LEFT(INDIRECT("C"&ROW());46)="TRAMITAÇÃO DE PROPOSIÇÕES: RELATÓRIO DE VISITA");HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=290";"PÁG 290");
  RIGHT(INDIRECT("C"&ROW());33)="RELATÓRIO DE EVENTO INSTITUCIONAL";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=223";"PÁG 223");
  RIGHT(INDIRECT("C"&ROW());24)="REQUERIMENTOS APROVADOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=277";"PÁG 277");
  RIGHT(INDIRECT("C"&ROW());26)="COMUNICAÇÃO DA PRESIDÊNCIA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=246";"PÁG 246");
  OR(LEFT(INDIRECT("C"&ROW());22)="DECISÃO DA PRESIDÊNCIA";LEFT(INDIRECT("C"&ROW());49)="TRAMITAÇÃO DE PROPOSIÇÕES: DECISÃO DA PRESIDÊNCIA");HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=243";"PÁG 243");
  RIGHT(INDIRECT("C"&ROW());22)="PALAVRAS DO PRESIDENTE";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=242";"PÁG 242");
  LEFT(INDIRECT("C"&ROW());25)="DESPACHO DE REQUERIMENTOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=262";"PÁG 262");
  LEFT(INDIRECT("C"&ROW());17)="AUDIÊNCIA PÚBLICA";"PÁG ??";
  LEFT(INDIRECT("C"&ROW());23)="AUDIÊNCIA DE CONVIDADOS";"PÁG ??";

  OR(LEFT(INDIRECT("C"&ROW());36)="MENSAGEM DO GOVERNADOR QUE ENCAMINHA";
  LEFT(INDIRECT("C"&ROW());38)="MENSAGENS DO GOVERNADOR QUE ENCAMINHAM";
  LEFT(INDIRECT("C"&ROW());63)="TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA";
  LEFT(INDIRECT("C"&ROW());65)="TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGENS DO GOVERNADOR QUE ENCAMINHAM";
  LEFT(INDIRECT("C"&ROW());35)="MENSAGEM DO GOVERNADOR QUE COMUNICA";
  LEFT(INDIRECT("C"&ROW());62)="TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE COMUNICA";
  LEFT(INDIRECT("C"&ROW());35)="MENSAGEM DO GOVERNADOR QUE SOLICITA");IFS(
  OR(RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";RIGHT(INDIRECT("C"&ROW());15)="PROJETOS DE LEI";RIGHT(INDIRECT("C"&ROW());36)="PROJETO DE LEI - CRÉDITO SUPLEMENTAR");HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=58";"PÁG 58");
  RIGHT(INDIRECT("C"&ROW());42)="EMENDA OU SUBSTITUTIVO COM DESPACHO À MESA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=65";"PÁG 65");
  RIGHT(INDIRECT("C"&ROW());41)="EMENDA OU SUBSTITUTIVO COM DESPACHO À FFO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=65";"PÁG 65");
  RIGHT(INDIRECT("C"&ROW());12)="VETO PARCIAL";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=73";"PÁG 73");
  RIGHT(INDIRECT("C"&ROW());10)="VETO TOTAL";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=73";"PÁG 73");
  RIGHT(INDIRECT("C"&ROW());29)="REGIME ESPECIAL DE TRIBUTAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=78";"PÁG 78");
  RIGHT(INDIRECT("C"&ROW());9)="INDICAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=81";"PÁG 81");
  RIGHT(INDIRECT("C"&ROW());28)="PEDIDO DE REGIME DE URGÊNCIA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=84";"PÁG 84");
  RIGHT(INDIRECT("C"&ROW());18)="CONVÊNIO DO CONFAZ";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=89";"PÁG 89");
  RIGHT(INDIRECT("C"&ROW());16)="CONVÊNIO DO ICMS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=90";"PÁG 90");
  RIGHT(INDIRECT("C"&ROW());44)="PRESTAÇÃO DE CONTAS DA ADMINISTRAÇÃO PÚBLICA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=93";"PÁG 93");
  RIGHT(INDIRECT("C"&ROW());36)="RELATÓRIO SOBRE A SITUAÇÃO DO ESTADO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=97";"PÁG 97");
  RIGHT(INDIRECT("C"&ROW());29)="DESARQUIVAMENTO DE PROPOSIÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=110";"PÁG 110");
  RIGHT(INDIRECT("C"&ROW());19)="RETIRADA DE PROJETO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=110";"PÁG 100");
  RIGHT(INDIRECT("C"&ROW());16)="AUSÊNCIA DO PAÍS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=140";"PÁG 140")
  );

  OR(LEFT(INDIRECT("C"&ROW());20)="OFÍCIO DO GOVERNADOR";
  MID(INDIRECT("C"&ROW());28;20)="OFÍCIO DO GOVERNADOR";
  LEFT(INDIRECT("C"&ROW());25)="OFÍCIO DO VICE-GOVERNADOR");
  IFS(
  RIGHT(INDIRECT("C"&ROW());28)="COMUNICANDO AUSÊNCIA DO PAÍS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=139";"PÁG 139");
  RIGHT(INDIRECT("C"&ROW());35)="COMUNICANDO QUE ENCAMINHOU MENSAGEM";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=138";"PÁG 138"));

  OR(LEFT(INDIRECT("C"&ROW());27)="REQUERIMENTOS DE COMISSÃO: ";LEFT(INDIRECT("C"&ROW());3)="RQC");IFS(
  RIGHT(INDIRECT("C"&ROW());9)="APROVADOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=285";"PÁG 285");
  RIGHT(INDIRECT("C"&ROW());9)="EMENDADOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=291";"PÁG 291");
  RIGHT(INDIRECT("C"&ROW());7)="ADIADOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=291";"PÁG 301");
  RIGHT(INDIRECT("C"&ROW());9)="RECEBIDOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=286";"PÁG 286");
  RIGHT(INDIRECT("C"&ROW());12)="PREJUDICADOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=283";"PÁG 283");
  RIGHT(INDIRECT("C"&ROW());9)="RELATÓRIO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=294";"PÁG 294");
  RIGHT(INDIRECT("C"&ROW());10)="REITERADOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=296";"PÁG 296");
  RIGHT(INDIRECT("C"&ROW());9)<>"APROVADOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=285";"PÁG 285"));

  OR(LEFT(INDIRECT("C"&ROW());10)="OFÍCIOS - ";LEFT(INDIRECT("C"&ROW());27)="CORRESPONDÊNCIA: OFÍCIOS - ");IFS(
  RIGHT(INDIRECT("C"&ROW());15)="PROJETOS DE LEI";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=51";"PÁG 51");
  RIGHT(INDIRECT("C"&ROW());13)="REQUERIMENTOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=50";"PÁG 50");
  RIGHT(INDIRECT("C"&ROW());5)="VETOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=51";"PÁG 51");
  RIGHT(INDIRECT("C"&ROW());20)="PRORROGAÇÃO DE PRAZO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=53";"PÁG 53");
  RIGHT(INDIRECT("C"&ROW());33)="PROPOSTA DE EMENDA À CONSTITUIÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=51";"PÁG 51"));

  OR(LEFT(INDIRECT("C"&ROW());42)="OFÍCIO DO TRIBUNAL DE CONTAS QUE ENCAMINHA";LEFT(INDIRECT("C"&ROW());69)="TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO TRIBUNAL DE CONTAS QUE ENCAMINHA");IFS(
  RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=113";"PÁG 113");
  RIGHT(INDIRECT("C"&ROW());23)="RELATÓRIO DE ATIVIDADES";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=117";"PÁG 117");
  RIGHT(INDIRECT("C"&ROW());23)="BALANÇO GERAL DO ESTADO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=118";"PÁG 118");
  RIGHT(INDIRECT("C"&ROW());19)="PRESTAÇÃO DE CONTAS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=120";"PÁG 120"));

  OR(LEFT(INDIRECT("C"&ROW());29)="OFÍCIO DO TRIBUNAL DE JUSTIÇA";
  LEFT(INDIRECT("C"&ROW());56)="TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO TRIBUNAL DE JUSTIÇA");IFS(
  OR(RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";RIGHT(INDIRECT("C"&ROW());27)="PROJETO DE LEI COMPLEMENTAR");HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=123";"PÁG 123"));

  OR(LEFT(INDIRECT("C"&ROW());28)="OFÍCIO DA DEFENSORIA PÚBLICA";
  LEFT(INDIRECT("C"&ROW());55)="TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DA DEFENSORIA PÚBLICA");IFS(
  RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=148";"PÁG 148"));

  OR(LEFT(INDIRECT("C"&ROW());28)="OFÍCIO DO MINISTÉRIO PÚBLICO";
  LEFT(INDIRECT("C"&ROW());55)="TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO MINISTÉRIO PÚBLICO");IFS(
  RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=144";"PÁG 144"));

  LEFT(INDIRECT("C"&ROW());39)="OFÍCIO DA PROCURADORIA-GERAL DE JUSTIÇA";IFS(
  RIGHT(INDIRECT("C"&ROW());14)="PROJETO DE LEI";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=144";"PÁG 144"));

  OR(LEFT(INDIRECT("C"&ROW());42)="APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS";LEFT(INDIRECT("C"&ROW());69)="TRAMITAÇÃO DE PROPOSIÇÕES: APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS");IFS(
  RIGHT(INDIRECT("C"&ROW());19)="COMISSÕES TEMÁTICAS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=197";"PÁG 197");
  RIGHT(INDIRECT("C"&ROW());15)="COM COMUNICAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=201";"PÁG 201");
  RIGHT(INDIRECT("C"&ROW());15)="SEM COMUNICAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=203";"PÁG 203");
  RIGHT(INDIRECT("C"&ROW());8)="ANEXADOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=204";"PÁG 204");
  RIGHT(INDIRECT("C"&ROW());19)="CIDADANIA HONORÁRIA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=206";"PÁG 206");
  RIGHT(INDIRECT("C"&ROW());13)="INDICAÇÃO TCE";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=207";"PÁG 207");
  RIGHT(INDIRECT("C"&ROW());25)="ASSEMBLEIA FISCALIZA MAIS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=210";"PÁG 210");
  RIGHT(INDIRECT("C"&ROW());18)="FRENTE PARLAMENTAR";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=211";"PÁG 211");
  RIGHT(INDIRECT("C"&ROW());24)="INCLUSÃO EM ORDEM DO DIA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=215";"PÁG 215");
  RIGHT(INDIRECT("C"&ROW());22)="RETIRADA DE TRAMITAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=215";"PÁG 215");
  RIGHT(INDIRECT("C"&ROW());15)="DESARQUIVAMENTO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=216";"PÁG 216");
  RIGHT(INDIRECT("C"&ROW());11)="DESANEXAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=218";"PÁG 218");
  RIGHT(INDIRECT("C"&ROW());17)="COMISSÃO SEGUINTE";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=220";"PÁG 220");
  RIGHT(INDIRECT("C"&ROW());17)="MAIS UMA COMISSÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=221";"PÁG 221");
  RIGHT(INDIRECT("C"&ROW());7)="RECURSO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=222";"PÁG 222");
  RIGHT(INDIRECT("C"&ROW());21)="PEDIDO DE INFORMAÇÕES";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=213";"PÁG 213");
  RIGHT(INDIRECT("C"&ROW());22)="PEDIDO DE PROVIDÊNCIAS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=213";"PÁG 213");
  RIGHT(INDIRECT("C"&ROW());14)="PERDA DE PRAZO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=213";"PÁG 213");
  RIGHT(INDIRECT("C"&ROW());16)="REUNIÃO ESPECIAL";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=213";"PÁG 213");
  RIGHT(INDIRECT("C"&ROW());38)="MESA DA ASSEMBLEIA, VOTADO EM PLENÁRIO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=223";"PÁG 223");
  RIGHT(INDIRECT("C"&ROW());57)="MESA DA ASSEMBLEIA, ENCAMINHADOS PARA PROVIDÊNCIA INTERNA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=223";"PÁG 223");
  RIGHT(INDIRECT("C"&ROW());19)="DESPACHO A DEPUTADO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=226";"PÁG 226");
  RIGHT(INDIRECT("C"&ROW());19)="DESPACHO A SERVIDOR";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=227";"PÁG 227");
  RIGHT(INDIRECT("C"&ROW());13)="SETOR DA CASA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=229";"PÁG 229");
  RIGHT(INDIRECT("C"&ROW());19)<>"COMISSÕES TEMÁTICAS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=197";"PÁG 197"));

  OR(LEFT(INDIRECT("C"&ROW());62)="APRESENTAÇÃO DE PROPOSIÇÕES: PROPOSTA DE EMENDA À CONSTITUIÇÃO";LEFT(INDIRECT("C"&ROW());60)="TRAMITAÇÃO DE PROPOSIÇÕES: PROPOSTA DE EMENDA À CONSTITUIÇÃO");HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=167";"PÁG 167");

  OR(LEFT(INDIRECT("C"&ROW());44)="APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE LEI";
  LEFT(INDIRECT("C"&ROW());42)="TRAMITAÇÃO DE PROPOSIÇÕES: PROJETOS DE LEI");IFS(
  RIGHT(INDIRECT("C"&ROW());19)="COMISSÕES TEMÁTICAS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=168";"PÁG 168");
  RIGHT(INDIRECT("C"&ROW());8)="ANEXADOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=169";"PÁG 169");
  RIGHT(INDIRECT("C"&ROW());18)="MESA DA ASSEMBLEIA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=163";"PÁG 163"));

  OR(LEFT(INDIRECT("C"&ROW());50)="APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO";LEFT(INDIRECT("C"&ROW());48)="TRAMITAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO");IFS(
  RIGHT(INDIRECT("C"&ROW());29)="REGIME ESPECIAL DE TRIBUTAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=170";"PÁG 170");
  RIGHT(INDIRECT("C"&ROW());19)="APROVAÇÃO DE CONTAS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=172";"PÁG 172");
  RIGHT(INDIRECT("C"&ROW());24)="RATIFICAÇÃO DE CONVÊNIOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=173";"PÁG 173");
  RIGHT(INDIRECT("C"&ROW());37)="ESTRUTURA DA SECRETARIA DA ASSEMBLEIA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=175";"PÁG 175");
  RIGHT(INDIRECT("C"&ROW());19)="COMISSÕES TEMÁTICAS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=176";"PÁG 176");
  RIGHT(INDIRECT("C"&ROW());8)="ANEXADOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=178";"PÁG 178");
  RIGHT(INDIRECT("C"&ROW());19)="CIDADANIA HONORÁRIA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=180";"PÁG 180");
  RIGHT(INDIRECT("C"&ROW());18)="CALAMIDADE PÚBLICA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=181";"PÁG 181");
  RIGHT(INDIRECT("C"&ROW());21)="LICENÇA AO GOVERNADOR";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=182";"PÁG 182")
  );


  LEFT(INDIRECT("C"&ROW());25)="PROPOSIÇÕES NÃO RECEBIDAS";IFS(
  RIGHT(INDIRECT("C"&ROW());15)="PROJETOS DE LEI";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=238";"PÁG 238");
  RIGHT(INDIRECT("C"&ROW());13)="REQUERIMENTOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=238";"PÁG 238");
  RIGHT(INDIRECT("C"&ROW());15)<>"PROJETOS DE LEI";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=238";"PÁG 238"));

  LEFT(INDIRECT("C"&ROW());25)="RECEBIMENTO DE PROPOSIÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=339";"PÁG 339");
  LEFT(INDIRECT("C"&ROW());23)="DESIGNAÇÃO DE RELATORIA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=343";"PÁG 343");
  LEFT(INDIRECT("C"&ROW());25)="CUMPRIMENTO DE DILIGÊNCIA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=346";"PÁG 346");
  LEFT(INDIRECT("C"&ROW());33)="REUNIÃO ORIGINADA DE REQUERIMENTO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=350";"PÁG 350");
  LEFT(INDIRECT("C"&ROW());32)="REUNIÃO COM DEBATE DE PROPOSIÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=351";"PÁG 351");
  INDIRECT("C"&ROW())="PAUTA COMPLETA DE REUNIÃO COM DEBATE DE PROPOSIÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=352";"PÁG 352");
  INDIRECT("C"&ROW())="RESULTADO COMPLETO DE REUNIÃO COM DEBATE DE PROPOSIÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=353";"PÁG 353");
  INDIRECT("C"&ROW())="CONGRATULAÇÕES ENTREGUES EM REUNIÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=328";"PÁG 328");
  INDIRECT("C"&ROW())="ENTREGA DE DIPLOMA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=328";"PÁG 328");
  LEFT(INDIRECT("C"&ROW());16)="CONSULTA PÚBLICA";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=356";"PÁG 356");

  OR(LEFT(INDIRECT("C"&ROW());8)="COMISSÃO";LEFT(INDIRECT("C"&ROW());4)="CIPE");IFS(
  MID(INDIRECT("C"&ROW());13;12)="CONSTITUIÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=366";"PÁG 366");
  MID(INDIRECT("C"&ROW());13;12)="FISCALIZAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=371";"PÁG 371");
  MID(INDIRECT("C"&ROW());13;12)="PARTICIPAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=376";"PÁG 376");
  MID(INDIRECT("C"&ROW());13;7)="REDAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=378";"PÁG 378");
  MID(INDIRECT("C"&ROW());11;8)="ESPECIAL";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=380";"PÁG 380");
  MID(INDIRECT("C"&ROW());13;12)<>"CONSTITUIÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=356";"PÁG 356"));

  INDIRECT("C"&ROW())="REMESSA - REQUERIMENTO APROVADO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=329";"PÁG 329");
  INDIRECT("C"&ROW())="REMESSA - PEDIDO DE INFORMAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=330";"PÁG 330");
  INDIRECT("C"&ROW())="OFÍCIO - REQUERIMENTO APROVADO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=329";"PÁG 329");
  INDIRECT("C"&ROW())="OFÍCIO - PEDIDO DE INFORMAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=330";"PÁG 330");
  INDIRECT("C"&ROW())="OFÍCIO - PEDIDO de INFORMAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=331";"PÁG 331");
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE APLAUSO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=326";"PÁG 326");
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE APOIO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=326";"PÁG 326");
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE REPÚDIO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=326";"PÁG 326");
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE PROTESTO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=326";"PÁG 326");
  INDIRECT("C"&ROW())="OFÍCIO - MANIFESTAÇÃO DE PESAR";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=327";"PÁG 327");
  INDIRECT("C"&ROW())="OFÍCIO - VOTO DE CONGRATULAÇÕES";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=328";"PÁG 328");
  INDIRECT("C"&ROW())="OFÍCIO COMUNICANDO MANUTENÇÃO TOTAL DO VETO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=333";"PÁG 333");
  INDIRECT("C"&ROW())="OFÍCIO COMUNICANDO REJEIÇÃO TOTAL DO VETO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=334";"PÁG 334");
  INDIRECT("C"&ROW())="OFÍCIO COMUNICANDO REJEIÇÃO PARCIAL DO VETO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=334";"PÁG 334");
  INDIRECT("C"&ROW())="OFÍCIO COMUNICANDO APROVAÇÃO DA INDICAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=335";"PÁG 335");
  INDIRECT("C"&ROW())="OFÍCIO ENCAMINHADO AOS DESTINATÁRIOS POR E-MAIL";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=335";"PÁG 335");

  INDIRECT("C"&ROW())="PROPOSIÇÃO DE LEI ENCAMINHADA PARA SANÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=332";"PÁG 332");
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: PROJETOS DE LEI";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=403";"PÁG 403");
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: REQUERIMENTOS, APROVADOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=404";"PÁG 404");
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: REQUERIMENTOS, RECURSO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=404";"PÁG 404");
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: REQUERIMENTOS, REJEITADOS";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=404";"PÁG 404");
  INDIRECT("C"&ROW())="PRECLUSÃO DE PRAZO: INCONSTITUCIONALIDADE";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=405";"PÁG 405")

  ))''']] * ((footer_start - 1) - 5))

# ====================================================================================================================================================================================================
# ============================================================================================= DATAS ===============================================================================================
# ====================================================================================================================================================================================================
data += [
    {"range": f"'{tab_name}'!Q2", "values": [["=B6"]]},
    {"range": f"'{tab_name}'!Q3", "values": [["=TODAY()"]]},
    {"range": f"'{tab_name}'!Q4", "values": [['=QUERY(C6:G8;"SELECT E WHERE C MATCHES \'.*DIÁRIO DO LEGISLATIVO.*\'";0)']]},
    {"range": f"'{tab_name}'!S2", "values": [['=TEXT(Q2;"\'dd\' \'mm\' \'yyyy\'")']]},
    {"range": f"'{tab_name}'!S3", "values": [['=TEXT(Q3;"\'d\' \'MM\' yyyy")']]},
    {"range": f"'{tab_name}'!S4", "values": [['=TEXT(Q4;"\'dd\' \'mm\' \'yyyy\'")']]},
    {"range": f"'{tab_name}'!T2", "values": [['=TEXT(Q2;"\'d\' \'m\' \'yyyy\'")']]},
    {"range": f"'{tab_name}'!T3", "values": [['=TEXT(Q3;"\'d\' \'m\' \'yyyy\'")']]},
    {"range": f"'{tab_name}'!T4", "values": [['=TEXT(Q4;"\'d\' \'m\' \'yyyy\'")']]},
    {"range": f"'{tab_name}'!U2", "values": [['=TEXT(Q2;"yyyymmdd")']]},
    {"range": f"'{tab_name}'!U3", "values": [['=TEXT(Q3;"yyyymmdd")']]},
    {"range": f"'{tab_name}'!U4", "values": [['=TEXT(Q4;"yyyymmdd")']]},
    {"range": f"'{tab_name}'!V2", "values": [['=TEXT(Q2;"yyyy-mm-dd")']]},
    {"range": f"'{tab_name}'!V3", "values": [['=TEXT(Q3;"yyyy-mm-dd")']]},
    {"range": f"'{tab_name}'!V4", "values": [['=TEXT(Q4;"yyyy-mm-dd")']]},
    {"range": f"'{tab_name}'!W2", "values": [['=TEXT(Q2;"dd mm yyyy")']]},
    {"range": f"'{tab_name}'!W3", "values": [['=IFERROR(QUERY(C6:G13;"SELECT E WHERE C MATCHES ''.*DIÁRIO DO LEGISLATIVO - EDIÇÃO EXTRA.*''";0);"SEM EXTRA")']]},
    {"range": f"'{tab_name}'!W4", "values": [['=TEXT(QUERY(B6:G33;"SELECT B WHERE C MATCHES \'REQUERIMENTOS DE COMISSÃO\'";0);"\'dd mm yyyy\'")']]},
    {"range": f"'{tab_name}'!X4", "values": [['=IFERROR(TEXT(QUERY(B6:G33;"SELECT B WHERE C MATCHES ''REQUERIMENTOS DE COMISSÃO''";0);"dd/MM/yyyy");"")']]},
    {"range": f"'{tab_name}'!Y2", "values": [["REUNIÃO"]]},
    {"range": f"'{tab_name}'!Y3", "values": [["EXTRA"]]},
    {"range": f"'{tab_name}'!Y4", "values": [["RQC"]]},
]

# EXECUTA O BLOCO PRINCIPAL (values)
body = {"valueInputOption": "USER_ENTERED", "data": data}
_with_backoff(sh.values_batch_update, body)

# ====================================================================================================================================================================================================
# ============================================================================================= TÍTULOS ==============================================================================================
# ====================================================================================================================================================================================================
data2 = [{"range": f"'{tab_name}'!B8:C8", "values": [[tab_name, "DIÁRIO DO LEGISLATIVO"]]}]

ALVOS = (
    "REQUERIMENTOS DE COMISSÃO",
    "LANÇAMENTOS DE TRAMITAÇÃO",
    "CADASTRO DE E-MAILS",
)

extra_rows_c_is_dash = []
for i, row in enumerate(extras):
    if row[0] != "-":
        continue
    if i - 1 < 0:
        continue
    prev_title = extras[i - 1][1] if len(extras[i - 1]) > 1 else ""
    if any(t in str(prev_title) for t in ALVOS):
        extra_rows_c_is_dash.append(start_extra_row + i)

for r in extra_rows_c_is_dash:
    data2.append({"range": f"'{tab_name}'!E{r}:I{r}", "values": [["-","-","-","-","-"]]})

# acha linha do DROPDOWN_2 (para setar D com "-")
dd2_row = next(
    start_extra_row + i
    for i, (_b, c) in enumerate(extras)
    if c == "DROPDOWN_2"
)
data2.append({"range": f"'{tab_name}'!D{dd2_row}", "values": [["-"]]})

# IMPLANTAÇÃO DE TEXTOS (mantém)
impl_row = next(
    start_extra_row + i
    for i, (_b, c) in enumerate(extras)
    if isinstance(c, str) and "IMPLANTAÇÃO DE TEXTOS" in c
)

# linha do título
data2.append({"range": f"'{tab_name}'!E{impl_row}:G{impl_row}", "values": [["TEXTOS", "EMENDAS", "PARECERES"]]})

# linha filha (logo abaixo)
data2.append({"range": f"'{tab_name}'!E{impl_row + 1}:I{impl_row + 1}", "values": [["?", "?", "?", "-", False]]})

# validação boolean em I (coluna 8 -> I, 0-based)
reqs.append({
    "setDataValidation": {
        "range": {
            "sheetId": sheet_id,
            "startRowIndex": impl_row,
            "endRowIndex": impl_row + 1,
            "startColumnIndex": 8,
            "endColumnIndex": 9,
        },
        "rule": {"condition": {"type": "BOOLEAN"}, "strict": True},
    }
})

body2 = {"valueInputOption": "USER_ENTERED", "data": data2}

# fontes vermelhas (mantém)
reqs.append(req_font(sheet_id, f"E{impl_row + 1}", fg_hex="#D32F2F"))
reqs.append(req_font(sheet_id, f"F{impl_row + 1}", fg_hex="#D32F2F"))
reqs.append(req_font(sheet_id, f"G{impl_row + 1}", fg_hex="#D32F2F"))
reqs.append(req_font(sheet_id, f"H{impl_row + 1}", fg_hex="#D32F2F"))
reqs.append(req_font(sheet_id, f"I{impl_row + 1}", fg_hex="#D32F2F"))

_with_backoff(sh.values_batch_update, body2)

# ====================================================================================================================================================================================================
# ======================================================================================= CONTAGEM DINÂMICA ==========================================================================================
# ====================================================================================================================================================================================================
FORMULAS_E = [
    '=LET(total;COUNTIFS(C:C;"ORDINÁRIA")+COUNTIFS(C:C;"EXTRAORDINÁRIA")+COUNTIFS(C:C;"ESPECIAL");total & IF(total=1;" REUNIÃO";" REUNIÕES"))',
    '=LET(total;COUNTIFS(C:C;"COMISSÃO*")+COUNTIFS(C:C;"CIPE*");total & IF(total=1;" REUNIÃO";" REUNIÕES"))',
    '=LET(total;SUMIFS(E:E;C:C;"RQC*");total & IF(total=1;" REQUERIMENTO";" REQUERIMENTOS"))',
    '=LET(total;SUMIFS(E$2:E$39;C$2:C$39;"RECEBIMENTO DE PROPOSIÇÃO")+SUMIFS(E$2:E$39;C$2:C$39;"DESIGNAÇÃO DE RELATORIA")+SUMIFS(E$2:E$39;C$2:C$39;"CUMPRIMENTO DE DILIGÊNCIA")+SUMIFS(E$2:E$39;C$2:C$39;"ENTREGA DE DIPLOMA")+SUMIFS(E$2:E$39;C$2:C$39;"REUNIÃO ORIGINADA DE REQUERIMENTO")+SUMIFS(E$2:E$39;C$2:C$39;"REUNIÃO COM DEBATE DE PROPOSIÇÃO*")+SUMIFS(E$2:E$39;C$2:C$39;"AUDIÊNCIA PÚBLICA*")+SUMIFS(E$2:E$39;C$2:C$39;"REMESSA - *")+SUMIFS(E$2:E$39;C$2:C$39;"OFÍCIO - *")+SUMIFS(E$2:E$39;C$2:C$39;"PROPOSIÇÃO DE LEI ENCAMINHADA PARA SANÇÃO");total & IF(total=1;" LANÇAMENTO";" LANÇAMENTOS"))',
    '=LET(total;SUMIFS(E:E;C:C;"PRECLUSÃO*")+SUMIFS(E:E;C:C;"CONSULTA*");total & IF(total=1;" LANÇAMENTO";" LANÇAMENTOS"))',
]

# linhas onde haverá contagem (as mesmas em que C tem título e você mescla E:G)
# EXCETO dropdown (não pode ter nada em E)
extra_formula_rows = [
    start_extra_row + i
    for i, row in enumerate(extras)
    if (
        (row[1] if len(row) > 1 else "") not in ("-", "", "DROPDOWN_2", "DROPDOWN_4")
        and (row[2] if len(row) > 2 else "") != "DROPDOWN_3"
        and "IMPLANTAÇÃO DE TEXTOS" not in (row[1] if len(row) > 1 else "")
    )
]

data_extra_E = [
    {"range": f"E{r}", "values": [[FORMULAS_E[i]]]}
    for i, r in enumerate(extra_formula_rows[:len(FORMULAS_E)])
]

_with_backoff(ws.batch_update, data_extra_E, value_input_option="USER_ENTERED")

# ====================================================================================================================================================================================================
# ============================================================================================== CALL ================================================================================================
# ====================================================================================================================================================================================================

# --- SANITIZAÇÃO FINAL: remove requests com range inválido/incompleto ---
reqs_ok = []
for i, r in enumerate(reqs):
    rng = None
    for k in ("mergeCells", "updateBorders", "setDataValidation"):
        if k in r and "range" in r[k]:
            rng = r[k]["range"]
            break

    if rng is not None:
        sr = rng.get("startRowIndex")
        er = rng.get("endRowIndex")
        sc = rng.get("startColumnIndex")
        ec = rng.get("endColumnIndex")

        if sr is None or er is None or sc is None or ec is None:
            print(f"[req {i}] range incompleto -> REMOVIDO: {rng}")
            continue

        if er <= sr or ec <= sc:
            print(f"[req {i}] inválido R{sr}:{er} C{sc}:{ec} -> REMOVIDO")
            continue

    reqs_ok.append(r)

reqs = reqs_ok

# Mantém o comportamento do seu trecho: chamadas de batch_update em sequência.
# 1) sempre chama uma vez com requests (pode ser vazio)
_with_backoff(sh.batch_update, body={"requests": reqs})

# 2) repete conforme condição (mesma intenção do seu código)
if reqs:
    _with_backoff(sh.batch_update, body={"requests": reqs})
else:
    _with_backoff(sh.batch_update, body={"requests": []})

# 3) “EXECUTA REQUESTS (LAYOUT)” — preserva o bloco final (com alias)
if reqs:
    with_backoff(sh.batch_update, body={"requests": reqs})

return sh.url, ws.title

SPREADSHEET = "https://docs.google.com/spreadsheets/d/1QUpyjHetLqLcr4LrgQqTnCXPZZfEyPkSQb-ld2RxW1k/edit"

# >>> diario_key PRECISA SER YYYYMMDD (é isso que upsert_tab_diario faz strptime("%Y%m%d"))
# >>> quando a entrada foi DATA, você já tem aba_yyyymmdd (dia útil de trabalho)
diario_key = aba_yyyymmdd if aba_yyyymmdd else yyyymmdd

ret = upsert_tab_diario(
    spreadsheet_url_or_id=SPREADSHEET,
    diario_key=diario_key,
    itens=itens,
)

url, aba = ret if ret is not None else (None, aba)

print("Planilha atualizada:", url)
print("Aba:", aba)
