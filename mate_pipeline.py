# PARTE 1A ========================================================================================
# - intervalo SOBREPOSTO quando configurado: pag_fim = página onde começa o próximo título (sem -1)
#
# Regras de fechamento (pag_fim):
# - Se próximo evento (OUT ou CUT) está em outra página:
#   - Se próximo evento está no TOPO REAL da página: pag_fim = pag_next - 1
#   - Senão:
#       - se evento atual é "sobreposto": pag_fim = pag_next
#       - se não: pag_fim = pag_next - 1
# ================================================================================================

import re
import csv
import os
import hashlib
import urllib.request
import unicodedata
from pathlib import Path
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from functools import lru_cache

from pypdf import PdfReader

# ---- 1) Regex Base ----
RE_PAG = re.compile(r"\bP[ÁA]GINA\s+(\d{1,4})\b", re.IGNORECASE)

URL_BASE = "https://diariolegislativo.almg.gov.br"
CACHE_DIR = "/content/pdfs_cache"
os.makedirs(CACHE_DIR, exist_ok=True)

# ================================================================================================
# ---- 2) Entrada: DATA (DDMMYYYY) -> URL do Diário (ou URL/caminho direto) ----
# ================================================================================================

try:
    from google.colab import files
    _COLAB = True
except Exception:
    _COLAB = False

TZ_BR = ZoneInfo("America/Sao_Paulo")


# --------------------------------------------------------------------------------
# NÃO-EXPEDIENTE (FERIADOS + RECESSOS) — usado para calcular a ABA (dia útil de trabalho)
# --------------------------------------------------------------------------------
from datetime import date

def _intervalo_datas(inicio: date, fim: date) -> set[date]:
    out = set()
    d = inicio
    while d <= fim:
        out.add(d)
        d += timedelta(days=1)
    return out


# --- 2025 (conforme lista oficial "FERIADOS E RECESSOS DE 2025" enviada) ---
NAO_EXPEDIENTE_2025 = {
    date(2025, 1, 1),   # Confraternização Universal

    # Março
    date(2025, 3, 3),   # Recesso
    date(2025, 3, 4),   # Carnaval
    date(2025, 3, 5),   # Recesso – Cinzas

    # Abril
    date(2025, 4, 17),  # Recesso
    date(2025, 4, 18),  # Paixão de Cristo
    date(2025, 4, 21),  # Tiradentes

    # Maio
    date(2025, 5, 1),   # Dia do Trabalho
    date(2025, 5, 2),   # Recesso

    # Junho
    date(2025, 6, 19),  # Corpus Christi
    date(2025, 6, 20),  # Recesso

    # Agosto
    date(2025, 8, 15),  # Assunção de Nossa Senhora

    # Setembro
    date(2025, 9, 7),   # Independência do Brasil

    # Outubro
    date(2025, 10, 12), # Nossa Senhora Aparecida
    date(2025, 10, 30), # Dia do Servidor Público

    # Novembro
    date(2025, 11, 2),  # Finados
    date(2025, 11, 15), # Proclamação da República
    date(2025, 11, 20), # Dia Nacional de Zumbi e da Consciência Negra
    date(2025, 11, 21), # Recesso

    # Dezembro
    date(2025, 12, 8),  # Nossa Senhora da Conceição
    date(2025, 12, 24), # Recesso
    date(2025, 12, 25), # Natal
    date(2025, 12, 26), # Recesso
    date(2025, 12, 31), # Recesso
}


# --- 2026 (conforme calendário enviado; feriados + recessos tratados como não-expediente) ---
NAO_EXPEDIENTE_2026 = set()

# FERIADOS (círculo)
NAO_EXPEDIENTE_2026 |= {
    date(2026, 1, 1),
    date(2026, 2, 17),
    date(2026, 6, 4),
    date(2026, 9, 7),
    date(2026, 10, 12),
    date(2026, 11, 2),
    date(2026, 11, 15),
    date(2026, 11, 20),
    date(2026, 12, 25),
}

# RECESSOS (retângulo)
NAO_EXPEDIENTE_2026 |= {
    date(2026, 2, 18),
    date(2026, 4, 2),
    date(2026, 4, 3),
    date(2026, 6, 5),
}
NAO_EXPEDIENTE_2026 |= _intervalo_datas(date(2026, 12, 7), date(2026, 12, 31))


NAO_EXPEDIENTE_POR_ANO = {
    2025: NAO_EXPEDIENTE_2025,
    2026: NAO_EXPEDIENTE_2026,
}


def proximo_dia_util(yyyymmdd: str) -> str:
    """
    Regra da ABA (trabalho):
    - Se a data do Diário cair em sábado/domingo/feriado/recesso: avança até o próximo dia útil.
    - Se cair em dia útil: retorna a mesma data.
    """
    d = datetime.strptime(yyyymmdd, "%Y%m%d").date()
    nao_expediente = NAO_EXPEDIENTE_POR_ANO.get(d.year, set())

    def eh_util(x: date) -> bool:
        return (x.weekday() < 5) and (x not in nao_expediente)  # Mon=0..Fri=4

    while not eh_util(d):
        d += timedelta(days=1)

    return d.strftime("%Y%m%d")

def normalizar_data(entrada: str) -> str:
    s_raw = "" if entrada is None else str(entrada)
    s = s_raw.strip()
    s_lower = s.lower()

    # --- PALAVRAS-CHAVE ---
    if s_lower in ("hoje", "ontem", "anteontem"):
        base = datetime.now(TZ_BR)
        if s_lower == "ontem":
            base -= timedelta(days=1)
        elif s_lower == "anteontem":
            base -= timedelta(days=2)
        return base.strftime("%Y%m%d")

    # --- DIAS DA SEMANA (última ocorrência passada) ---
    weekday_map = {
        "terça": 1, "terca": 1,
        "quarta": 2,
        "quinta": 3,
        "sexta": 4,
        "sábado": 5, "sabado": 5,
    }

    if s_lower in weekday_map:
        target = weekday_map[s_lower]  # Mon=0 ... Sun=6
        today = datetime.now(TZ_BR)
        days_back = (today.weekday() - target) % 7
        if days_back == 0:
            days_back = 7  # garante "passado"
        return (today - timedelta(days=days_back)).strftime("%Y%m%d")

    digits = "".join(ch for ch in s if ch.isdigit())

    if len(digits) == 4:
        # ddmm -> ano atual
        dd = digits[0:2]
        mm = digits[2:4]
        yyyy = datetime.now(TZ_BR).year
        yyyymmdd = f"{yyyy:04d}{mm}{dd}"

    elif len(digits) == 6:
        # ddmmyy -> assume 20yy
        dd = digits[0:2]
        mm = digits[2:4]
        yy = int(digits[4:6])
        yyyy = 2000 + yy
        yyyymmdd = f"{yyyy:04d}{mm}{dd}"

    elif len(digits) == 8:
        # pode ser yyyymmdd OU ddmmyyyy
        if digits.startswith(("19", "20")):
            try:
                datetime.strptime(digits, "%Y%m%d")
                return digits  # é yyyymmdd válido
            except ValueError:
                pass  # cai para ddmmyyyy

        dd = digits[0:2]
        mm = digits[2:4]
        yyyy = digits[4:8]
        yyyymmdd = f"{yyyy}{mm}{dd}"

    else:
        raise ValueError(
            "Data inválida. Use hoje, ontem, anteontem, "
            "ddmm, ddmmyy, ddmmyyyy, dd/mm/yy, dd/mm/yyyy ou yyyymmdd."
        )

    datetime.strptime(yyyymmdd, "%Y%m%d")
    return yyyymmdd


def montar_url_diario(data_in: str) -> str:
    yyyymmdd = normalizar_data(data_in)
    yyyy = yyyymmdd[:4]
    return f"{URL_BASE}/{yyyy}/L{yyyymmdd}.pdf"


def _parece_pdf(caminho: str) -> bool:
    try:
        with open(caminho, "rb") as f:
            head = f.read(5)
        return head == b"%PDF-"
    except Exception:
        return False


def baixar_pdf_por_url(url: str) -> str | None:
    import requests

    local = "/content/tmp_diario.pdf"

    try:
        r = requests.get(url, timeout=30, allow_redirects=True)
        r.raise_for_status()

        with open(local, "wb") as f:
            f.write(r.content)

        # verifica assinatura PDF
        with open(local, "rb") as f:
            head = f.read(5)

        if head != b"%PDF-":
            print("?? DL não existe para a data informada (conteúdo não é PDF).")
            print("URL:", url)
            print("Head:", head)
            return None

        return local

    except Exception as e:
        print("?? Erro ao baixar o Diário.")
        print("URL:", url)
        print("Erro:", e)
        return None

print("Digite a data do Diário do Legislativo.")
print("EXEMPLOS:")
print("- 19122026 ou 191226 ou 19/12/2026")
print("- hoje, ontem, anteontem")
print("- terça, quarta, quinta, sexta, sábado")
print("- URL completa ou um caminho local.")
print("Se deixar vazio, você poderá fazer upload.\n")

def yyyymmdd_to_ddmmyyyy(yyyymmdd: str) -> str:
    return f"{yyyymmdd[6:8]}/{yyyymmdd[4:6]}/{yyyymmdd[0:4]}"

def main(entrada_override=None, spreadsheet_url_or_id=None):
    # Se veio override, não pede input
    if entrada_override is None:
        print("Digite a data do Diário do Legislativo.")
        print("EXEMPLOS:")
        print("- 19122026 ou 191226 ou 19/12/2026")
        print("- hoje, ontem ou anteontem")
        print("- terça, quarta, quinta, sexta ou sábado")
        print("- URL ou caminho local")
        print("Se deixar vazio, você poderá fazer upload.\n")
        entrada = input("Data/URL/Upload:").strip()
    else:
        entrada = str(entrada_override).strip()

    import re

    # A partir daqui, cole TODO o fluxo atual (o que hoje está global),
    # usando a variável local `entrada` (sem globals()).
    #
    # IMPORTANTE: mantenha suas inicializações exatamente como estão:
    # pdf_path = None, aba_yyyymmdd = None, aba = None, yyyymmdd = None, etc.
    #
    # IMPORTANTE 2: quando chegar na chamada do upsert_tab_diario, use:
    # spreadsheet_url_or_id = spreadsheet_url_or_id or SPREADSHEET
    #
    # E ao final:
    # return url, aba

    pdf_path = None  # sempre inicializa
    aba_yyyymmdd = None  # data da ABA (trabalho), quando a entrada for DATA
    aba = None  # NOME FINAL da aba (DD/MM/YYYY) — deve ser usado no Sheets
    yyyymmdd = None  # fallback seguro p/ diario_key quando entrada não for DATA
    diario = None
        
    if not entrada:
        if not _COLAB:
            raise SystemExit("Entrada vazia fora do Colab. Informe data, URL ou caminho.")
        up = files.upload()
        if not up:
            raise SystemExit("Nenhum arquivo enviado.")
        pdf_path = next(iter(up.keys()))
        print(f"Upload OK: {pdf_path}")

    elif entrada.lower().startswith(("http://", "https://")):
        pdf_path = baixar_pdf_por_url(entrada)
        if not pdf_path:
            raise SystemExit("DL não existe (URL não retornou PDF).")

    elif "/" in entrada or "\\" in entrada or entrada.lower().startswith("/content"):
        pdf_path = entrada
        if not os.path.exists(pdf_path):
            raise SystemExit(f"Arquivo local não encontrado: {pdf_path}")

    else:
        # Entrada é DATA (ou palavra-chave / dia da semana)
        yyyymmdd = normalizar_data(entrada)          # data do Diário (PDF)
        aba_yyyymmdd = proximo_dia_util(yyyymmdd)    # data de Trabalho (ABA)
        diario = yyyymmdd_to_ddmmyyyy(yyyymmdd)      # data de Diário (PLANILHA)

        # --- DIÁRIO - 2 dias úteis ---
        import datetime as dt

        dl_date = dt.datetime.strptime(yyyymmdd, "%Y%m%d").date()

        d = dl_date
        count = 0
        while count < 2:
            d = d - dt.timedelta(days=1)
            # regra simples: segunda–sexta
            if d.weekday() < 5:
                count += 1

        dmenos2 = f"{d.day}/{d.month}/{d.year}"

        yyyy = yyyymmdd[:4]
        url = f"{URL_BASE}/{yyyy}/L{yyyymmdd}.pdf"   # monta URL sem re-normalizar

        print(f"URL: {url}")
        print(f"Diário: {yyyymmdd}")
        print(f"Aba: {aba_yyyymmdd}")

        pdf_path = baixar_pdf_por_url(url)
        if not pdf_path:
            raise SystemExit("DL não existe para a data informada.")

    # ? TRATAMENTO DEFINITIVO DE DL INEXISTENTE
    if pdf_path is None:
        print("? Diário do Legislativo inexistente para a data informada. Execução encerrada.")
        raise SystemExit

    pdf_path = str(pdf_path)
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF não encontrado após processamento: {pdf_path}")

    # --------------------------------------------------------------------------------
    # ABA FINAL (Sheets): usa sempre a ABA de trabalho quando houver DATA; senão cai em HOJE
    # --------------------------------------------------------------------------------
    if aba_yyyymmdd:
        aba = yyyymmdd_to_ddmmyyyy(aba_yyyymmdd)
    else:
        aba = yyyymmdd_to_ddmmyyyy(datetime.now(TZ_BR).strftime("%Y%m%d"))

    print("Diário:", diario)
    print("Planilha:", aba)

    # ================================================================================================
    # ---- 3) Extração e detecção de títulos ----
    # ================================================================================================

    def limpa_linha(s: str) -> str:
        s = s.replace("\u00a0", " ")
        s = re.sub(r"[ \t]+", " ", s).strip()
        return s


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
        u = "".join(ch for ch in u if unicodedata.category(ch) != "Mn")
        return re.sub(r"[^0-9A-Z]", "", u)


    # ---- TOP detection (robusta) ----
    RE_HEADER_LIXO = re.compile(
        r"(DI[ÁA]RIO\s+DO\s+LEGISLATIVO|www\.almg\.gov\.br|"
        r"Segunda-feira|Ter[aç]a-feira|Quarta-feira|Quinta-feira|Sexta-feira|S[aá]bado|Domingo|"
        r"\bP[ÁA]GINA\s+\d+\b)",
        re.IGNORECASE
    )


    def _linha_relevante(s: str) -> bool:
        s = limpa_linha(s)
        if not s:
            return False
        if RE_HEADER_LIXO.search(s):
            return False
        if re.fullmatch(r"[-–—_•\.\s]+", s):
            return False
        return bool(re.search(r"[A-Za-zÀ-ÿ0-9]", s))


    def is_top_event(line_idx: int, linhas: list[str]) -> bool:
        for prev in linhas[:line_idx]:
            if _linha_relevante(prev):
                return False
        return True


    # ---- helper: matching por janela (1–3 linhas) ----
    def win_keys(linhas: list[str], i: int, w: int) -> str:
        parts = []
        for k in range(w):
            j = i + k
            if j < len(linhas):
                parts.append(compact_key(linhas[j]))
        return "".join(parts)


    def win_any_in(linhas: list[str], i: int, keys: set[str]) -> bool:
        k1 = win_keys(linhas, i, 1)
        k2 = win_keys(linhas, i, 2)
        k3 = win_keys(linhas, i, 3)
        return (k1 in keys) or (k2 in keys) or (k3 in keys)


    def _checkbox_req(sheet_id: int, col_idx_0based: int, row_1based: int, default_checked: bool = False):
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
                "rows": [{"values": [{"userEnteredValue": val}]}],
                "fields": "userEnteredValue",
            }
        }

        return [dv, setv]


    def _cf_fontsize_req(sheet_id: int, col0: int, row1: int, font_size: int, formula: str, index: int = 0):
        """
        Conditional formatting: aplica tamanho de fonte quando fórmula custom for TRUE.
        Exemplo: =OR($C9="DIÁRIO DO LEGISLATIVO"; $C9="REUNIÕES DE PLENÁRIO")
        """
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
                        "format": {
                            "textFormat": {"fontSize": font_size}
                        },
                    },
                },
                "index": index,
            }
        }


    # Estruturais / contexto
    C_TRAMITACAO = "TRAMITACAODEPROPOSICOES"
    C_RECEBIMENTO = "RECEBIMENTODEPROPOSICOES"
    C_APRESENTACAO = "APRESENTACAODEPROPOSICOES"

    # CUTs de verdade (não entram no CSV)
    C_ATA = "ATA"
    C_ATAS = "ATAS"
    C_MATERIA_ADM = "MATERIAADMINISTRATIVA"
    C_QUESTAO_ORDEM = "QUESTAODEORDEM"
    CUT_KEYS = {C_ATA, C_ATAS, C_MATERIA_ADM, C_QUESTAO_ORDEM}

    # Contextual CORRESPONDÊNCIA: OFÍCIOS
    C_CORRESP_CAB = "CORRESPONDENCIADESPACHADAPELO1SECRETARIO"
    C_OFICIOS = "OFICIOS"

    # OUTs “simples”
    C_MANIFESTACAO = "MANIFESTACAO"
    C_MANIFESTACOES = "MANIFESTACOES"
    MANIF_KEYS = {C_MANIFESTACAO, C_MANIFESTACOES}

    C_REQ_APROV = "REQUERIMENTOAPROVADO"
    C_REQS_APROV = "REQUERIMENTOSAPROVADOS"
    REQ_APROV_KEYS = {C_REQ_APROV, C_REQS_APROV}

    C_PROPOSICOES_DE_LEI = "PROPOSICOESDELEI"
    C_RESOLUCAO = "RESOLUCAO"
    C_ERRATA = "ERRATA"
    C_ERRATAS = "ERRATAS"
    ERRATA_KEYS = {C_ERRATA, C_ERRATAS}

    C_RECEB_EMENDAS_SUBST = "RECEBIMENTODEEMENDASESUBSTITUTIVO"
    C_RECEB_EMENDAS_SUBSTS = "RECEBIMENTODEEMENDASESUBSTITUTIVOS"
    C_RECEB_EMENDA = "RECEBIMENTODEEMENDA"
    EMENDAS_KEYS = {C_RECEB_EMENDAS_SUBST, C_RECEB_EMENDAS_SUBSTS, C_RECEB_EMENDA}

    # Novos OUTs
    C_LEITURA_COMUNICACOES = "LEITURADECOMUNICACOES"
    C_DESPACHO_REQUERIMENTOS = "DESPACHODEREQUERIMENTOS"
    C_DECISAO_PRESIDENCIA = "DECISAODAPRESIDENCIA"
    C_ACORDO_LIDERES = "ACORDODELIDERES"
    C_COMUNIC_PRESIDENCIA = "COMUNICACAODAPRESIDENCIA"
    C_PROPOSICOES_NAO_RECEBIDAS = "PROPOSICOESNAORECEBIDAS"

    # APRESENTAÇÃO: gatilhos materiais
    C_REQUERIMENTOS = "REQUERIMENTOS"
    C_PROJETO_DE_LEI = "PROJETODELEI"
    C_PROJETOS_DE_LEI = "PROJETOSDELEI"


    def prefix_tramitacao(label: str, in_tramitacao: bool) -> str:
        if in_tramitacao:
            return f"TRAMITAÇÃO DE PROPOSIÇÕES: {label}"
        return label


    def label_apresentacao(tipo_bloco: str, in_tramitacao: bool) -> str:
        if tipo_bloco == "PL":
            base = "APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE LEI"
        else:
            base = "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS"
        return prefix_tramitacao(base, in_tramitacao)


    reader = PdfReader(pdf_path)

    # eventos: (pag, ordem, tipo, label_out, fim_sobreposto, top_flag)
    eventos = []
    ordem = 0

    # estados
    in_tramitacao = False
    sub_tramitacao = None          # None | C_RECEBIMENTO | C_APRESENTACAO
    apresentacao_ativa = False     # True se estamos em APRESENTAÇÃO
    sub_apresentacao = None        # None | "PL" | "REQ"
    viu_corresp_cab = False

    pegou_leis = False
    MAX_PAG_LEIS = 40

    for i, page in enumerate(reader.pages):
        texto = page.extract_text() or ""
        linhas = [limpa_linha(x) for x in texto.splitlines() if limpa_linha(x)]
        pag_num = primeira_pagina_num(linhas, i + 1)

        for li, ln in enumerate(linhas):
            ln_up = ln.upper().strip()
            c = compact_key(ln)
            top_flag = is_top_event(li, linhas)

            # janela compactada (p/ títulos quebrados)
            k1 = win_keys(linhas, li, 1)
            k2 = win_keys(linhas, li, 2)
            k3 = win_keys(linhas, li, 3)

            # ---------------------------
            # CUTs “reais”
            # ---------------------------
            if c in CUT_KEYS:
                ordem += 1
                eventos.append((pag_num, ordem, "CUT", None, False, top_flag))
                # encerra contextos
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            if c.startswith("PARECER"):
                ordem += 1
                eventos.append((pag_num, ordem, "CUT", None, False, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # ---------------------------
            # Estrutural: TRAMITAÇÃO
            # ---------------------------
            if c == C_TRAMITACAO:
                in_tramitacao = True
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                ordem += 1
                eventos.append((pag_num, ordem, "CUT", None, False, top_flag))
                viu_corresp_cab = False
                continue

            # ---------------------------
            # Marcadores RECEBIMENTO/APRESENTAÇÃO dentro de TRAMITAÇÃO
            # ---------------------------
            if in_tramitacao and (c == C_RECEBIMENTO or c == C_APRESENTACAO):
                sub_tramitacao = c
                apresentacao_ativa = (c == C_APRESENTACAO)
                sub_apresentacao = None
                ordem += 1
                eventos.append((pag_num, ordem, "CUT", None, False, top_flag))
                viu_corresp_cab = False
                continue

            # ---------------------------
            # APRESENTAÇÃO fora de TRAMITAÇÃO: só marca contexto
            # ---------------------------
            if (not in_tramitacao) and (c == C_APRESENTACAO):
                apresentacao_ativa = True
                sub_apresentacao = None
                continue

            # se aparecer um “corte natural” fora da lógica, zera apresentação
            if apresentacao_ativa and c in {C_TRAMITACAO, C_ATA, C_ATAS, C_MATERIA_ADM}:
                apresentacao_ativa = False
                sub_apresentacao = None

            # ---------------------------
            # Contexto: CORRESPONDÊNCIA DESPACHADA PELO 1º-SECRETÁRIO
            # ---------------------------
            if c == C_CORRESP_CAB:
                viu_corresp_cab = True
                continue

            # OUT contextual: CORRESPONDÊNCIA: OFÍCIOS
            if viu_corresp_cab and c == C_OFICIOS:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "CORRESPONDÊNCIA: OFÍCIOS", True, top_flag))
                viu_corresp_cab = False
                # encerra contextos gerais
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                continue

            # ---------------------------
            # APRESENTAÇÃO -> subdivisão material (PL vs REQ)
            # ---------------------------
            if apresentacao_ativa:
                # gatilho PL
                if (
                    k1.startswith(C_PROJETO_DE_LEI) or k1.startswith(C_PROJETOS_DE_LEI) or
                    k2.startswith(C_PROJETO_DE_LEI) or k2.startswith(C_PROJETOS_DE_LEI) or
                    k3.startswith(C_PROJETO_DE_LEI) or k3.startswith(C_PROJETOS_DE_LEI)
                ):
                    if sub_apresentacao != "PL":
                        ordem += 1
                        eventos.append((pag_num, ordem, "OUT", label_apresentacao("PL", in_tramitacao), True, top_flag))
                        sub_apresentacao = "PL"
                    continue

                # gatilho REQ
                if (k1.startswith(C_REQUERIMENTOS) or k2.startswith(C_REQUERIMENTOS) or k3.startswith(C_REQUERIMENTOS)):
                    if sub_apresentacao != "REQ":
                        ordem += 1
                        eventos.append((pag_num, ordem, "OUT", label_apresentacao("REQ", in_tramitacao), True, top_flag))
                        sub_apresentacao = "REQ"
                    continue

            # ---------------------------
            # OUTs diretos (fora de APRESENTAÇÃO)
            # ---------------------------

            # OFÍCIOS (comum)
            if c == C_OFICIOS:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "OFÍCIOS", True, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # LEIS PROMULGADAS (linha exatamente LEI/LEIS)
            if (not pegou_leis) and (pag_num <= MAX_PAG_LEIS) and (ln_up == "LEI" or ln_up == "LEIS"):
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "LEIS PROMULGADAS", True, top_flag))
                pegou_leis = True
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # MANIFESTAÇÕES
            if c in MANIF_KEYS:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "MANIFESTAÇÕES", True, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # REQUERIMENTOS APROVADOS
            if c in REQ_APROV_KEYS:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "REQUERIMENTOS APROVADOS", True, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # PROPOSIÇÕES DE LEI
            if c == C_PROPOSICOES_DE_LEI:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "PROPOSIÇÕES DE LEI", True, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # RESOLUÇÃO
            if c == C_RESOLUCAO:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "RESOLUÇÃO", True, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # ERRATAS
            if c in ERRATA_KEYS:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "ERRATAS", True, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # EMENDAS OU SUBSTITUTIVOS PUBLICADOS
            if c in EMENDAS_KEYS:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "EMENDAS OU SUBSTITUTIVOS PUBLICADOS", True, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # ACORDO DE LÍDERES
            if c == C_ACORDO_LIDERES:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "ACORDO DE LÍDERES", True, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # COMUNICAÇÃO DA PRESIDÊNCIA (com prefixo se dentro de TRAMITAÇÃO)
            if c == C_COMUNIC_PRESIDENCIA:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", prefix_tramitacao("COMUNICAÇÃO DA PRESIDÊNCIA", in_tramitacao), True, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # LEITURA DE COMUNICAÇÕES
            if c == C_LEITURA_COMUNICACOES:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "LEITURA DE COMUNICAÇÕES", True, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # DESPACHO DE REQUERIMENTOS
            if c == C_DESPACHO_REQUERIMENTOS:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "DESPACHO DE REQUERIMENTOS", True, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # DECISÃO DA PRESIDÊNCIA
            if c == C_DECISAO_PRESIDENCIA:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "DECISÃO DA PRESIDÊNCIA", True, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue

            # PROPOSIÇÕES NÃO RECEBIDAS
            if c == C_PROPOSICOES_NAO_RECEBIDAS:
                ordem += 1
                eventos.append((pag_num, ordem, "OUT", "PROPOSIÇÕES NÃO RECEBIDAS", True, top_flag))
                in_tramitacao = False
                sub_tramitacao = None
                apresentacao_ativa = False
                sub_apresentacao = None
                viu_corresp_cab = False
                continue


    # ---- ordena eventos ----
    eventos.sort(key=lambda x: (x[0], x[1]))

    # ---- 3.x) pós-processamento: OUTs duplicados na mesma página ----
    KEEP_DUP_OUT = {
        "DECISÃO DA PRESIDÊNCIA",
    }

    _last_idx = {}
    for i, ev in enumerate(eventos):
        pag_ini, ordm, tipo, label_out, fim_sobreposto, top_flag = ev
        if tipo != "OUT":
            continue
        if label_out in KEEP_DUP_OUT:
            continue
        key = (pag_ini, label_out)
        _last_idx[key] = i

    _eventos_filtrados = []
    for i, ev in enumerate(eventos):
        pag_ini, ordm, tipo, label_out, fim_sobreposto, top_flag = ev
        if tipo == "OUT" and label_out not in KEEP_DUP_OUT:
            key = (pag_ini, label_out)
            if _last_idx.get(key, i) != i:
                continue
        _eventos_filtrados.append(ev)

    eventos = _eventos_filtrados

    print("EVENTOS:", len(eventos))

    # ---- 4) intervalos ----
    total_pag_fisica = len(reader.pages)
    itens = []

    for idx, e in enumerate(eventos):
        pag_ini, ordm, tipo, label_out, fim_sobreposto, top_flag = e
        if tipo != "OUT":
            continue

        prox = eventos[idx + 1] if (idx + 1) < len(eventos) else None

        if prox is None:
            pag_fim = total_pag_fisica
        else:
            pag_next, _, tipo_next, _, _, top_next = prox

            if pag_next == pag_ini:
                pag_fim = pag_ini
            else:
                if top_next:
                    pag_fim = pag_next - 1
                else:
                    pag_fim = pag_next if fim_sobreposto else (pag_next - 1)

        if pag_fim < pag_ini:
            pag_fim = pag_ini

        intervalo = f"{pag_ini} - {pag_fim}" if pag_ini != pag_fim else f"{pag_ini}"

        # labels em que repetição na mesma página é "marcador + título real" ? manter só o ÚLTIMO
        DEDUP_ULTIMO_NA_PAG = {"MANIFESTAÇÕES"}

        # labels em que repetição na mesma página é conteúdo distinto ? NÃO deduplica
        NAO_DEDUPLICAR = {"DECISÃO DA PRESIDÊNCIA"}

        if label_out in DEDUP_ULTIMO_NA_PAG and label_out not in NAO_DEDUPLICAR:
            if itens and itens[-1][1] == label_out:
                itens[-1] = (intervalo, label_out)
            else:
                itens.append((intervalo, label_out))
        else:
            itens.append((intervalo, label_out))

    # ---- DEBUG controlado se não achou nada ----
    DEBUG_SEM_OUTS = False        # coloque True apenas quando quiser investigar
    DEBUG_MAX_PAGS = 10           # limite de páginas a varrer
    DEBUG_MAX_LINHAS = 50         # limite de linhas a imprimir

    if not itens and DEBUG_SEM_OUTS:

        achados = []

        for pi, p in enumerate(reader.pages[:DEBUG_MAX_PAGS]):
            t = p.extract_text() or ""

            for raw in t.splitlines():
                ln = limpa_linha(raw)
                if not ln:
                    continue

                if re.search(
                    r"(TRAMITA|APRESENTA|RECEB|REQUER|LEI|MANIFEST|ATA|MATERIA\s+ADMIN|QUESTAO|RESOLU|ERRAT|EMEND|SUBSTIT|ACORDO|PARECER|CORRESP|OFIC|COMUNIC)",
                    ln,
                    re.IGNORECASE,
                ):
                    achados.append(f"p{pi+1}: {ln} || compact={compact_key(ln)}")

                if len(achados) >= DEBUG_MAX_LINHAS:
                    break

            if len(achados) >= DEBUG_MAX_LINHAS:
                break

        print(f"\n=== DEBUG (amostra limitada a {len(achados)} linhas) ===")
        for x in achados:
            print(x)

        print("Nenhum título de interesse encontrado. Prosseguindo com aba sem OUTs.")

    itens = itens or []

    # PARTE 1B ===================================================================================================================================================================================
    # ========================================================================================== 5) GOOGLE SHEETS ========================================================================================
    # ====================================================================================================================================================================================================

    import time, random
    import gspread
    from google.colab import auth
    from google.auth import default

    auth.authenticate_user()
    creds, _ = default(scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds)
    SHEET_ID = None

    def rgb_hex_to_api(hex_str: str):
        h = hex_str.lstrip("#")
        return {
            "red": int(h[0:2], 16) / 255.0,
            "green": int(h[2:4], 16) / 255.0,
            "blue": int(h[4:6], 16) / 255.0,
        }

    def a1_to_grid(a1: str):
        a1 = a1.strip()
        if ":" not in a1:
            a1 = f"{a1}:{a1}"
        return gspread.utils.a1_range_to_grid_range(a1)

    def field_mask_from_fmt(fmt: dict) -> str:
        parts = []
        if "backgroundColor" in fmt:
            parts.append("userEnteredFormat.backgroundColor")
        if "horizontalAlignment" in fmt:
            parts.append("userEnteredFormat.horizontalAlignment")
        if "verticalAlignment" in fmt:
            parts.append("userEnteredFormat.verticalAlignment")
        if "wrapStrategy" in fmt:
            parts.append("userEnteredFormat.wrapStrategy")
        if "textFormat" in fmt:
            parts.append("userEnteredFormat.textFormat")
        if "numberFormat" in fmt:
            parts.append("userEnteredFormat.numberFormat")
        return ",".join(parts) if parts else "userEnteredFormat"


    def req_repeat_cell(sheet_id: int, a1: str, fmt: dict):
        gr = a1_to_grid(a1)
        return {
            "repeatCell": {
                "range": {"sheetId": sheet_id, **gr},
                "cell": {"userEnteredFormat": fmt},
                "fields": field_mask_from_fmt(fmt),
            }
        }


    def req_text(sheet_id: int, a1: str, font_family: str, font_size: int, fg_hex: str):
        gr = a1_to_grid(a1)
        return {
            "repeatCell": {
                "range": {"sheetId": sheet_id, **gr},
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "fontFamily": font_family,
                            "fontSize": int(font_size),
                            "foregroundColor": rgb_hex_to_api(fg_hex),
                        }
                    }
                },
                "fields": "userEnteredFormat.textFormat(fontFamily,fontSize,foregroundColor)",
            }
        }


    def req_font(
        sheet_id: int,
        a1: str,
        font_size: int | None = None,
        fg_hex: str | None = None,
        bold: bool | None = None,
    ):
        gr = a1_to_grid(a1)

        tf = {}
        fields = []

        if font_size is not None:
            tf["fontSize"] = int(font_size)
            fields.append("userEnteredFormat.textFormat.fontSize")

        if fg_hex is not None:
            tf["foregroundColor"] = rgb_hex_to_api(fg_hex)
            fields.append("userEnteredFormat.textFormat.foregroundColor")

        if bold is not None:
            tf["bold"] = bool(bold)
            fields.append("userEnteredFormat.textFormat.bold")

        return {
            "repeatCell": {
                "range": {"sheetId": sheet_id, **gr},
                "cell": {"userEnteredFormat": {"textFormat": tf}},
                "fields": ",".join(fields) if fields else "userEnteredFormat.textFormat",
            }
        }


    def req_merge(sheet_id: int, a1: str):
        gr = a1_to_grid(a1)
        return {"mergeCells": {"range": {"sheetId": sheet_id, **gr}, "mergeType": "MERGE_ALL"}}


    def req_unmerge(sheet_id: int, a1: str):
        gr = a1_to_grid(a1)
        return {"unmergeCells": {"range": {"sheetId": sheet_id, **gr}}}


    def req_dim_rows(sheet_id: int, start: int, end: int, px: int):
        return {
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "ROWS", "startIndex": start, "endIndex": end},
                "properties": {"pixelSize": px},
                "fields": "pixelSize",
            }
        }


    def req_dim_cols(sheet_id: int, start: int, end: int, px: int):
        return {
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": start, "endIndex": end},
                "properties": {"pixelSize": px},
                "fields": "pixelSize",
            }
        }


    def req_tab_color(sheet_id: int, rgb: dict):
        return {"updateSheetProperties": {"properties": {"sheetId": sheet_id, "tabColor": rgb}, "fields": "tabColor"}}


    def req_update_borders(sheet_id: int, a1: str, top=None, bottom=None, left=None, right=None, innerH=None, innerV=None):
        gr = a1_to_grid(a1)

        # --- CORREÇÃO NA ORIGEM: garante GridRange completo ---
        sr = gr.get("startRowIndex")
        er = gr.get("endRowIndex")
        sc = gr.get("startColumnIndex")
        ec = gr.get("endColumnIndex")

        # Se a1_to_grid vier incompleto, completa como 1 linha/1 coluna (caso típico: célula única)
        if sr is not None and er is None:
            gr["endRowIndex"] = sr + 1
        if sc is not None and ec is None:
            gr["endColumnIndex"] = sc + 1

        b = {}
        if top is not None:
            b["top"] = top
        if bottom is not None:
            b["bottom"] = bottom
        if left is not None:
            b["left"] = left
        if right is not None:
            b["right"] = right
        if innerH is not None:
            b["innerHorizontal"] = innerH
        if innerV is not None:
            b["innerVertical"] = innerV

        return {"updateBorders": {"range": {"sheetId": sheet_id, **gr}, **b}}

    def border(style: str, color_rgb: dict):
        return {"style": style, "color": color_rgb}


    def req_merge_row(sheet_id: int, row1: int, col_start0: int, col_end0: int):
        return {
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row1 - 1,
                    "endRowIndex": row1,
                    "startColumnIndex": col_start0,
                    "endColumnIndex": col_end0,
                },
                "mergeType": "MERGE_ALL",
            }
        }


    def _with_backoff(fn, *args, **kwargs):
        for attempt in range(8):
            try:
                return fn(*args, **kwargs)
            except Exception as e:
                msg = str(e)
                if ("429" in msg) or ("Quota exceeded" in msg) or ("Rate Limit" in msg) or ("503" in msg):
                    sleep_s = min(60, (2**attempt) + random.random())
                    print(f"[backoff] tentativa {attempt+1}/8 – esperando {sleep_s:.1f}s por quota...")
                    time.sleep(sleep_s)
                    continue
                raise


    # ====================================================================================================================================================================================================
    # =============================================================================================== CORES ==============================================================================================
    # ====================================================================================================================================================================================================
    DARK_RED_1 = rgb_hex_to_api("#CC0000")
    TAB_RED    = rgb_hex_to_api("#990000")
    BLACK      = rgb_hex_to_api("#000000")
    WHITE      = rgb_hex_to_api("#FFFFFF")
    THIN_BLACK = rgb_hex_to_api("#000000")

    # ====================================================================================================================================================================================================
    # ============================================================================================= LARGURAS =============================================================================================
    # ====================================================================================================================================================================================================
    COL_OVERRIDES = {
        0: 23,  1: 60,  2: 370, 3: 75,  4: 85,  5: 75,  6: 75,
        7: 45,  8: 45,  9: 45,  10: 45, 11: 45, 12: 45, 13: 45,
        14: 45, 15: 60, 16: 75, 17: 70, 18: 70, 19: 60, 20: 60,
        21: 60, 22: 60, 23: 60, 24: 60
    }
    COL_DEFAULT = 60

    # ====================================================================================================================================================================================================
    # ============================================================================================= HEIGHTS ==============================================================================================
    # ====================================================================================================================================================================================================
    ROW_HEIGHTS = [
        ("default", 16),
        (0, 4, 14),   # linhas 1-4
        (4, 5, 25),   # linha 5
    ]

    # ====================================================================================================================================================================================================
    # ============================================================================================ MERGES ================================================================================================
    # ====================================================================================================================================================================================================
    MERGES = [
        "A1:B4", "C1:F4", "G1:G4", "Q1:Y1",
        "A5:B5", "E5:F5", "G5:I5", "T5:Y5",
        "E6:G6", "E8:G8",
        "H1:H2", "H3:H4", "I1:I2", "I3:I4", "J1:J2", "J3:J4", "K1:K2", "K3:K4", "L1:L2", "L3:L4", "M1:M2", "M3:M4", "N1:N2", "N3:N4", "O1:O2", "O3:O4",
        "J5:O5", "J6:O6", "J7:O7", "J8:O8", "J9:O9", "J10:O10",
        "J11:O11", "J12:O12", "J13:O13", "J14:O14", "J15:O15",
        "J16:O16", "J17:O17", "J18:O18", "J19:O19", "J20:O20",
        "J21:O21", "J22:O22",
    ]

    # ====================================================================================================================================================================================================
    # ============================================================================================== STYLES ==============================================================================================
    # ====================================================================================================================================================================================================
    STYLES = [
        # Geral
        ("A1:B", {"h": "CENTER", "v": "MIDDLE", "underline": False}),
        ("B6:I", {"font": "Inconsolata", "size": 8, "bold": True, "underline": False}),
        ("D6:I", {"wrap": "CLIP", "h": "CENTER", "v": "MIDDLE", "font": "Inconsolata", "size": 8, "bold": True, "underline": False}),
        ("P6:S", {"h": "CENTER", "v": "MIDDLE", "font": "Vidaloka", "size": 8, "bold": True, "underline": False, "fg": "BLACK"}),
        ("H6:H", {"font": "Inconsolata", "size": 8, "bold": True, "underline": False}),
        ("I6:I", {"font": "Inconsolata", "size": 6, "bold": True, "underline": False}),
        # Cabeçalho
        ("A5:Y5", {"bg": "BLACK", "h": "CENTER", "v": "MIDDLE", "wrap": "CLIP", "font": "Vidaloka", "size": 10, "bold": True, "fg": "WHITE"}),
        ("A5:B5", {"font": "Vidaloka", "size": 15, "bold": True, "fg": "WHITE", "numfmt": ("DATE", "d/m")}),
        ("C1:F4", {"bg": "DARK_RED_1", "h": "CENTER", "v": "MIDDLE", "wrap": "CLIP", "font": "Oregano", "size": 29, "bold": True, "fg": "WHITE"}),
        ("C5", {"font": "Vidaloka", "size": 15, "bold": True, "underline": False, "fg": "WHITE"}),
        ("D5", {"font": "Vidaloka", "size": 12, "bold": True, "fg": "WHITE"}),
        ("E5:I5", {"font": "Vidaloka", "size": 14, "bold": True, "fg": "WHITE"}),
        ("G5:I5", {"font": "Vidaloka", "size": 14, "bold": True, "underline": False, "fg": "WHITE"}),
        ("J5:O5", {"font": "Vidaloka", "size": 15, "bold": True, "fg": "WHITE"}),
        ("T5:Y5", {"font": "Vidaloka", "size": 15, "bold": True, "fg": "WHITE"}),
        ("P2:Y4", {"wrap": "CLIP", "font": "Special Elite", "size": 6, "bold": True}),
        ("P1:P4", {"h": "RIGHT", "v": "MIDDLE", "wrap": "CLIP", "font": "Special Elite", "size": 6, "bold": True}),
        ("Q1:Y1", {"bg": "TAB_RED", "h": "LEFT", "v": "MIDDLE", "wrap": "CLIP", "font": "Vidaloka", "size": 8, "bold": True, "fg": "WHITE"}),
        ("Y2:Y4", {"font": "Special Elite", "size": 6, "h": "LEFT", "v": "MIDDLE", "wrap": "CLIP", "bold": True}),
        ("G1:O4", {"h": "CENTER", "v": "MIDDLE"}),
        ("Q2:X4", {"h": "CENTER", "v": "MIDDLE"}),
        # Ata
        ("E8:G8", {"h": "CENTER", "v": "MIDDLE", "numfmt": ("DATE", "dd/MM/yyyy")}),
    ]

    # ====================================================================================================================================================================================================
    # ============================================================================================== BORDERS =============================================================================================
    # ====================================================================================================================================================================================================
    rows_needed = 30 + len(itens)
    BORDERS = [
        ("G1:G4", {"right": ("SOLID", "THIN_BLACK")}),
        ("P1:P4", {"left": ("SOLID", "THIN_BLACK")}),
        ("P4:Y4", {"bottom": ("SOLID_MEDIUM", "DARK_RED_1")}),
        ("V2:V4", {"right": ("SOLID_MEDIUM", "DARK_RED_1")}),
        ("G1:O4", {"bottom": ("SOLID_MEDIUM", "DARK_RED_1")}),
        (f"A6:A{rows_needed}", {"right":  ("SOLID", "THIN_BLACK")}),
        (f"H6:H{rows_needed}", {"left":   ("SOLID", "THIN_BLACK")}),
        (f"P6:P{rows_needed}", {"left":   ("SOLID", "THIN_BLACK")}),
        (f"C6:D{rows_needed}", {"right":  ("SOLID_MEDIUM", "BLACK")}),
        (f"S1:S{rows_needed}", {"right":  ("SOLID_MEDIUM", "BLACK")}),
        (f"Y1:Y{rows_needed}", {"right":  ("SOLID_MEDIUM", "BLACK")}),        
    ]

    # ====================================================================================================================================================================================================
    # ============================================================================================= BUILDERS =============================================================================================
    # ====================================================================================================================================================================================================

    _COLOR_MAP = {
        "DARK_RED_1": DARK_RED_1,
        "TAB_RED": TAB_RED,
        "BLACK": BLACK,
        "WHITE": WHITE,
        "THIN_BLACK": THIN_BLACK,
    }


    def _mini_to_user_fmt(mini: dict) -> dict:
        fmt = {}

        if "bg" in mini:
            fmt["backgroundColor"] = _COLOR_MAP[mini["bg"]]
        if "h" in mini:
            fmt["horizontalAlignment"] = mini["h"]
        if "v" in mini:
            fmt["verticalAlignment"] = mini["v"]
        if "wrap" in mini:
            fmt["wrapStrategy"] = mini["wrap"]
        if "numfmt" in mini:
            t, p = mini["numfmt"]  # ex: ("DATE","d/m")
            fmt["numberFormat"] = {"type": t, "pattern": p}

        tf = {}
        if "font" in mini:
            tf["fontFamily"] = mini["font"]
        if "size" in mini:
            tf["fontSize"] = int(mini["size"])
        if "bold" in mini:
            tf["bold"] = bool(mini["bold"])
        if "underline" in mini:
            tf["underline"] = bool(mini["underline"])
        if "fg" in mini:
            tf["foregroundColor"] = _COLOR_MAP[mini["fg"]]
        if tf:
            fmt["textFormat"] = tf

        return fmt


    def _border_from_spec(style_name: str, color_name: str):
        return border(style_name, _COLOR_MAP[color_name])


    def paint_left_of_dropdown(
        sheet_id: int,
        start_row_0: int,
        end_row_0: int,
        dropdown_col_0: int,
        fmt: dict,
    ):
        """
        Pinta a célula à esquerda do dropdown (mesmas linhas).
        Ex.: dropdown em col C (2) -> pinta col B (1).
        start_row_0 / end_row_0: índices 0-based, end exclusivo.
        dropdown_col_0: índice 0-based da coluna do dropdown.
        fmt: dict no padrão userEnteredFormat (igual o dropdown).
        """
        left_col = max(0, dropdown_col_0 - 1)

        return {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": start_row_0,
                    "endRowIndex": end_row_0,
                    "startColumnIndex": left_col,
                    "endColumnIndex": left_col + 1,
                },
                "cell": {"userEnteredFormat": fmt},
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
            }
        }


    # ====================================================================================================================================================================================================
    # ============================================================================================ FUNCTIONS =============================================================================================
    # ====================================================================================================================================================================================================
    from datetime import datetime, timedelta

    def aba_key_from_diario_key(diario_key: str) -> str:
        d = datetime.strptime(diario_key, "%Y%m%d").date()
        if d.weekday() == 5:  # sábado -> segunda
            d += timedelta(days=2)
        return d.strftime("%Y%m%d")

    def upsert_tab_diario(
        spreadsheet_url_or_id: str,
        diario_key: str,                 # YYYYMMDD
        itens: list[tuple[str, str]],
        clear_first: bool = False,
        default_col_width_px: int = COL_DEFAULT,
        col_width_overrides: dict[int, int] | None = None,
    ):
        tab_name = yyyymmdd_to_ddmmyyyy(aba_key_from_diario_key(diario_key))
        sh = gc.open_by_url(spreadsheet_url_or_id) if spreadsheet_url_or_id.startswith("http") else gc.open_by_key(spreadsheet_url_or_id)

        # cria/abre aba
        try:
            ws = sh.worksheet(tab_name)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=tab_name, rows=max(20, 20 + len(itens)), cols=25)
            _with_backoff(ws.update_index, 1)

        sheet_id = ws.id

        # --- GUARDA-CHUVA: garante grid mínimo antes de qualquer merge/unmerge ---
        MIN_ROWS = 1
        MIN_COLS = 25

        if ws.row_count < MIN_ROWS or ws.col_count < MIN_COLS:
            _with_backoff(ws.resize, rows=max(ws.row_count, MIN_ROWS), cols=max(ws.col_count, MIN_COLS))

        # resize da planilha (linhas e colunas) — agora considera EXTRAS também
        extras = [
            ['=TEXT(A5;"dd/mm/yyyy")', '=HYPERLINK("https://www.almg.gov.br/atividade-parlamentar/plenario/agenda/"; "REUNIÕES DE PLENÁRIO")'],
            ["", ""],
            ['=TEXT(A5;"dd/mm/yyyy")', '=HYPERLINK("https://www.almg.gov.br/atividade-parlamentar/comissoes/agenda/"; "REUNIÕES DE COMISSÕES")'],
            ["", ""],
            ['=TEXT(A5;"dd/mm/yyyy")', '=HYPERLINK("https://www.almg.gov.br/atividade-parlamentar/comissoes/agenda/"; "REQUERIMENTOS DE COMISSÃO")'],
            ["-", "-"],
            ['=TEXT(A5;"dd/mm/yyyy")', '=HYPERLINK("https://silegis.almg.gov.br/silegismg/#/processos"; "LANÇAMENTOS DE TRAMITAÇÃO")'],
            ["-", "DROPDOWN_2"],   # <- linha do dropdown 2 (coluna C) + dropdown 3 (coluna D)
            ['=TEXT(A5;"dd/mm/yyyy")', '=HYPERLINK("https://webmail.almg.gov.br/"; "CADASTRO DE E-MAILS")'],
            ["-", "DROPDOWN_4"],   # <- linha do dropdown 4 (coluna C)
            ['=TEXT(A5;"dd/mm/yyyy")', '=HYPERLINK("https://consulta-brs.almg.gov.br/brs/"; "IMPLANTAÇÃO DE TEXTOS")'],
            ["", '=SUM(INDIRECT("B"&ROW());INDIRECT("E"&ROW());INDIRECT("F"&ROW());INDIRECT("G"&ROW()))'],   # <- linha da implantação de textos
        ]

        # o que realmente vai aparecer na planilha (troca DROPDOWN_x por "-")
        extras_out = [[b, ("-" if str(c).startswith("DROPDOWN_") else c)] for b, c in extras]

        itens_len = len(itens) if itens else 0
        start_extra_row = 9 + itens_len + (1 if itens_len == 0 else 0)

        footer_rows = 9  # RODAPÉ: quantidade de linhas reservadas
        rows_needed = 9 + itens_len + len(extras) + footer_rows - 1
        cols_needed = 25

        MIN_ROWS = 1
        MIN_COLS = 25

        rows_target = max(ws.row_count, rows_needed + 1, MIN_ROWS)
        cols_target = max(ws.col_count, cols_needed, MIN_COLS)

        _with_backoff(ws.resize, rows=rows_target, cols=cols_target)

        VIS_LAST_ROW_1BASED = rows_target - 1  # última linha "visível" (a última é técnica 1px)

        # linha técnica (1px) — NÃO usa reqs aqui (reqs ainda não existe neste ponto)
        _with_backoff(sh.batch_update, {
            "requests": [{
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": rows_target - 1,
                        "endIndex": rows_target
                    },
                    "properties": {"pixelSize": 1},
                    "fields": "pixelSize"
                }
            }]
        })

        # ====================================================================================================================================================================================================
        # ============================================================================================ REQUESTS ==============================================================================================
        # ====================================================================================================================================================================================================

        reqs = []

        # cor da aba
        reqs.append(req_tab_color(sheet_id, DARK_RED_1))

        # congela linhas 1–5
        reqs.append({
            "updateSheetProperties": {
                "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 5}},
                "fields": "gridProperties.frozenRowCount"
            }
        })

        # alturas
        for rh in ROW_HEIGHTS:
            if rh[0] == "default":
                reqs.append(req_dim_rows(sheet_id, 0, ws.row_count, rh[1]))
            else:
                start, end, px = rh
                reqs.append(req_dim_rows(sheet_id, start, end, px))

        # linha técnica (1px) — mantém a linha extra invisível (depois das alturas)
        reqs.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": rows_target - 1,
                    "endIndex": rows_target
                },
                "properties": {"pixelSize": 1},
                "fields": "pixelSize"
            }
        })

        # larguras
        reqs.append(req_dim_cols(sheet_id, 0, 25, default_col_width_px))
        ow = col_width_overrides or COL_OVERRIDES
        for col_idx, px in ow.items():
            reqs.append(req_dim_cols(sheet_id, col_idx, col_idx + 1, px))

        reqs.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 19,   # T
                    "endIndex": 25      # Y (exclusivo)
                },
                "properties": {
                    "hiddenByUser": True
                },
                "fields": "hiddenByUser"
            }
        })

        # --- UNMERGE GERAL: zera qualquer mesclagem antiga antes de aplicar os merges desta execução ---
        reqs.append({
            "unmergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": rows_target,
                    "startColumnIndex": 0,
                    "endColumnIndex": 25
                }
            }
        })

        # merges fixos (MERGES) — sempre (layout base não depende de itens)
        for r in MERGES:
            reqs.append(req_unmerge(sheet_id, r))
            reqs.append(req_merge(sheet_id, r))

        # merges dinâmicos dos EXTRAS (nas linhas onde C != "-"), EXCETO "IMPLANTAÇÃO DE TEXTOS
        MERGE_TITLES = (
            "REUNIÕES DE PLENÁRIO",
            "REUNIÕES DE COMISSÕES",
            "REQUERIMENTOS DE COMISSÃO",
            "LANÇAMENTOS DE TRAMITAÇÃO",
            "CADASTRO DE E-MAILS",
        )

        extra_merge_rows = [
            start_extra_row + i
            for i, row in enumerate(extras)
            if (
                (row[1] if len(row) > 1 else "") not in ("-", "", "DROPDOWN_2", "DROPDOWN_4")
                and (row[2] if len(row) > 2 else "") != "DROPDOWN_3"
                and any(t in str(row[1]).upper() for t in MERGE_TITLES)
            )
        ]

        for r in extra_merge_rows:
            reqs.append(req_unmerge(sheet_id, f"C{r}:D{r}"))
            reqs.append(req_merge(sheet_id, f"C{r}:D{r}"))
            reqs.append(req_unmerge(sheet_id, f"E{r}:G{r}"))
            reqs.append(req_merge(sheet_id, f"E{r}:G{r}"))

        # -----------------------------
        # Linhas de DROPDOWN (não podem ter DV BOOLEAN em H)
        # - C: "DROPDOWN_2" ou "DROPDOWN_4"
        # - D: "DROPDOWN_3"
        # -----------------------------
        dropdown_rows = [
            start_extra_row + i
            for i, row in enumerate(extras)
            if (
                (row[1] if len(row) > 1 else "") in ("DROPDOWN_2", "DROPDOWN_4")
                or (row[2] if len(row) > 2 else "") == "DROPDOWN_3"
            )
        ]

        # -----------------------------
        # Checkbox em H somente nas linhas de título (C != "-"/vazio),
        # NÃO dropdown e NÃO DIÁRIO
        # -----------------------------
        extra_checkbox_rows = [
            start_extra_row + i
            for i, row in enumerate(extras)
            if (row[1] if len(row) > 1 else "") not in ("-", "")
            and (start_extra_row + i) not in dropdown_rows
        ]

        # range total dos EXTRAS (row 1-based -> grid 0-based endIndex exclusivo)
        extra_start = start_extra_row
        extra_end   = start_extra_row + len(extras) -1

        # ====================================================================================================================================================================================================
        # ============================================================================================ DROPDOWNS =============================================================================================
        # ====================================================================================================================================================================================================

        LISTA_DROPDOWNS_1_BG = "#fff2cc"   # cor leve p/ destacar que foi selecionado
        LISTA_DROPDOWNS_1_FG = "#7f6000"

        LISTA_DROPDOWN_1 = [
            "-", "?",
            "LEIS",
            "LEI, COM PROPOSIÇÃO ANEXADA",
            "LEIS, SEM PROPOSIÇÃO DE LEI PUBLICADA",
            "EMENDAS À CONSTITUIÇÃO PROMULGADAS",
            "LEIS PROMULGADAS",
            "PROPOSIÇÕES DE LEI",
            "RESOLUÇÃO",
            "PROPOSTAS DE AÇÃO LEGISLATIVA",
            "OFÍCIOS - PROJETOS DE LEI",
            "OFÍCIOS - REQUERIMENTOS",
            "OFÍCIOS - VETOS",
            "OFÍCIOS - PRORROGAÇÃO DE PRAZO",
            "OFÍCIO DA DEFENSORIA PÚBLICA QUE ENCAMINHA PROJETO DE LEI",
            "OFÍCIO DO TRIBUNAL DE CONTAS QUE ENCAMINHA PRESTAÇÃO DE CONTAS",
            "OFÍCIO DO TRIBUNAL DE CONTAS QUE ENCAMINHA PARECER PRÉVIO SOBRE BALANÇO GERAL DO ESTADO",
            "OFÍCIO DO TRIBUNAL DE CONTAS QUE ENCAMINHA RELATÓRIO DE ATIVIDADES",
            "OFÍCIO DO TRIBUNAL DE JUSTIÇA QUE ENCAMINHA PROJETO DE LEI",
            "OFÍCIO DO TRIBUNAL DE JUSTIÇA QUE ENCAMINHA PROJETO DE LEI COMPLEMENTAR",
            "OFÍCIO DO VICE-GOVERNADOR COMUNICANDO AUSÊNCIA DO PAÍS",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA PROJETO DE LEI COMPLEMENTAR",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA PROJETO DE LEI",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA PROJETO DE LEI - COMISSÕES TEMÁTICAS",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA PROJETO DE LEI - CRÉDITO SUPLEMENTAR",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA EMENDA OU SUBSTITUTIVO COM DESPACHO À FFO",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA EMENDA OU SUBSTITUTIVO COM DESPACHO À MESA",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA ABERTURA DE CRÉDITO SUPLEMENTAR",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA PRESTAÇÃO DE CONTAS DA ADMINISTRAÇÃO PÚBLICA",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA PEDIDO DE REGIME DE URGÊNCIA",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA CONVÊNIO DO ICMS",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA CONVÊNIO DO CONFAZ",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA REGIME ESPECIAL DE TRIBUTAÇÃO",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA RELATÓRIO TRIMESTRAL",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA INDICAÇÃO",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA VETO TOTAL",
            "MENSAGEM DO GOVERNADOR QUE ENCAMINHA VETO PARCIAL",
            "MENSAGEM DO GOVERNADOR QUE SOLICITA DESARQUIVAMENTO DE PROPOSIÇÃO",
            "MENSAGEM DO GOVERNADOR QUE SOLICITA RETIRADA DE PROJETO",
            "MENSAGEM DO GOVERNADOR QUE COMUNICA AUSÊNCIA DO PAÍS",
            "PROPOSIÇÃO: REQUERIMENTOS - INDICAÇÃO TCE",
            "APRESENTAÇÃO DE PROPOSIÇÕES: PROPOSTA DE EMENDA À CONSTITUIÇÃO",
            "APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE LEI COMPLEMENTAR - COMISSÕES TEMÁTICAS",
            "APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE LEI COMPLEMENTAR - ANEXADOS",
            "APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE LEI - COMISSÕES TEMÁTICAS",
            "APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE LEI - ANEXADOS",
            "APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO - COMISSÕES TEMÁTICAS",
            "APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO - CALAMIDADE PÚBLICA",
            "APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO - CIDADANIA HONORÁRIA",
            "APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO - ESTRUTURA DA SECRETARIA DA ASSEMBLEIA",
            "APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO - RATIFICAÇÃO DE CONVÊNIOS",
            "APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO - REGIME ESPECIAL DE TRIBUTAÇÃO",
            "APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO - APROVAÇÃO DE CONTAS",
            "APRESENTAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO - ANEXADOS",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - COMISSÕES TEMÁTICAS",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - MESA DA ASSEMBLEIA",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - MESA DA ASSEMBLEIA, ENCAMINHADOS PARA PROVIDÊNCIA INTERNA",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - MESA DA ASSEMBLEIA, VOTADO EM PLENÁRIO",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - SEM DESPACHO, SEM COMUNICAÇÃO",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - SEM DESPACHO, COM COMUNICAÇÃO",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - SEM DESPACHO, DESANEXAÇÃO",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - SEM DESPACHO, DESARQUIVAMENTO",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - SEM DESPACHO, RETIRADA DE TRAMITAÇÃO",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - SEM DESPACHO, REUNIÃO ESPECIAL",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - SEM DESPACHO, REDISTRIBUIÇÃO",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - ANEXADOS",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - CIDADANIA HONORÁRIA",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - INDICAÇÃO TCE",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - RQO COM DESPACHO À MAS",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - RQN COM DESPACHO A SERVIDOR",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - RQO COM DESPACHO A SETOR DA CASA",
            "APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS ORDINÁRIOS",
            "COMUNICAÇÃO DA PRESIDÊNCIA",
            "LEITURA DE COMUNICAÇÕES",
            "LEITURA DE COMUNICAÇÕES - CIENTE ANEXAR",
            "PALAVRAS DO PRESIDENTE",
            "DECISÃO DA PRESIDÊNCIA",
            "DECISÃO DA PRESIDÊNCIA, ANEXAÇÃO",
            "DECISÃO DA PRESIDÊNCIA, DESANEXAÇÃO",
            "DECISÃO DA PRESIDÊNCIA, REDISTRIBUIÇÃO",
            "DECISÃO DA MESA",
            "DESIGNAÇÃO DE COMISSÕES",
            "DESPACHO DE REQUERIMENTOS",
            "DESPACHO DE REQUERIMENTOS, COMISSÃO SEGUINTE",
            "DESPACHO DE REQUERIMENTOS, DESANEXAÇÃO",
            "DESPACHO DE REQUERIMENTOS, DESARQUIVAMENTO",
            "DESPACHO DE REQUERIMENTOS, RETIRADA DE TRAMITAÇÃO",
            "EMENDAS OU SUBSTITUTIVOS PUBLICADOS",
            "EMENDAS NÃO RECEBIDAS PUBLICADAS",
            "MANIFESTAÇÕES",
            "PROPOSIÇÕES NÃO RECEBIDAS",
            "REQUERIMENTOS APROVADOS",
            "TRAMITAÇÃO DE PROPOSIÇÕES: PARECERES",
            "TRAMITAÇÃO DE PROPOSIÇÕES: DESIGNAÇÃO DE COMISSÕES",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA PROJETO DE LEI COMPLEMENTAR",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA PROJETO DE LEI",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA PROJETO DE LEI - COMISSÕES TEMÁTICAS",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA PROJETO DE LEI - CRÉDITO SUPLEMENTAR",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA EMENDA OU SUBSTITUTIVO COM DESPACHO À MESA",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA EMENDA OU SUBSTITUTIVO COM DESPACHO À FFO",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA PEDIDO DE URGÊNCIA",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA CONVÊNIO DO ICMS",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA CONVÊNIO DO CONFAZ",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA REGIME ESPECIAL DE TRIBUTAÇÃO",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA RELATÓRIO TRIMESTRAL",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA INDICAÇÃO",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA VETO TOTAL",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MENSAGEM DO GOVERNADOR QUE ENCAMINHA VETO PARCIAL",
            "TRAMITAÇÃO DE PROPOSIÇÕES: MSG DO GOVERNADOR QUE REQUER RETIRADA DE REGIME DE URGÊNCIA",
            "TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO TRIBUNAL DE CONTAS QUE ENCAMINHA PRESTAÇÃO DE CONTAS",
            "TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO TRIBUNAL DE CONTAS QUE ENCAMINHA BALANÇO GERAL",
            "TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO TRIBUNAL DE CONTAS QUE ENCAMINHA PROJETO DE LEI",
            "TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO TRIBUNAL DE CONTAS QUE ENCAMINHA RELATÓRIO DE ATIVIDADES",
            "TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO TRIBUNAL DE JUSTIÇA QUE ENCAMINHA PROPOSTA DE EMENDA OU SUBSTITUTIVO",
            "TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO TRIBUNAL DE JUSTIÇA QUE ENCAMINHA PROJETO DE LEI",
            "TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DO MINISTÉRIO PÚBLICO QUE ENCAMINHA PROJETO DE LEI",
            "TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DA PROCURADORIA-GERAL DE JUSTIÇA QUE ENCAMINHA PROPOSTA DE EMENDA OU SUBSTITUTIVO",
            "TRAMITAÇÃO DE PROPOSIÇÕES: OFÍCIO DE PREFEITURA QUE ENCAMINHA DECRETOS DE CALAMIDADE PÚBLICA",
            "TRAMITAÇÃO DE PROPOSIÇÕES: PROJETOS DE LEI - COMISSÕES TEMÁTICAS",
            "TRAMITAÇÃO DE PROPOSIÇÕES: PROJETOS DE LEI - ANEXADOS",
            "TRAMITAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO - CALAMIDADE PÚBLICA",
            "TRAMITAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO - ESTRUTURA DA SECRETARIA DA ASSEMBLEIA",
            "TRAMITAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO - LICENÇA AO GOVERNADOR",
            "TRAMITAÇÃO DE PROPOSIÇÕES: PROJETOS DE RESOLUÇÃO - CIDADANIA HONORÁRIA",
            "TRAMITAÇÃO DE PROPOSIÇÕES: DECISÃO DA PRESIDÊNCIA",
            "TRAMITAÇÃO DE PROPOSIÇÕES: PALAVRAS DO PRESIDENTE",
            "TRAMITAÇÃO DE PROPOSIÇÕES: APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - COMISSÕES TEMÁTICAS",
            "TRAMITAÇÃO DE PROPOSIÇÕES: APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - MESA DA ASSEMBLEIA",
            "TRAMITAÇÃO DE PROPOSIÇÕES: APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - ANEXADOS",
            "TRAMITAÇÃO DE PROPOSIÇÕES: APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - CIDADANIA HONORÁRIA",
            "TRAMITAÇÃO DE PROPOSIÇÕES: APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - DESANEXAÇÃO",
            "TRAMITAÇÃO DE PROPOSIÇÕES: APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - SEM DESPACHO, COM COMUNICAÇÃO",
            "TRAMITAÇÃO DE PROPOSIÇÕES: APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - INDICAÇÃO TCE",
            "TRAMITAÇÃO DE PROPOSIÇÕES: APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - RQO COM DESPACHO À MAS",
            "TRAMITAÇÃO DE PROPOSIÇÕES: APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - RQO COM DESPACHO A SERVIDOR",
            "TRAMITAÇÃO DE PROPOSIÇÕES: APRESENTAÇÃO DE PROPOSIÇÕES: REQUERIMENTOS - RQO COM DESPACHO A SETOR DA CASA",
            "TRAMITAÇÃO DE PROPOSIÇÕES: COMUNICAÇÃO DA PRESIDÊNCIA",
            "TRAMITAÇÃO DE PROPOSIÇÕES: DECISÃO DA PRESIDÊNCIA, ACORDO DE LÍDERES",
            "TRAMITAÇÃO DE PROPOSIÇÕES: DESPACHO DE REQUERIMENTOS",
            "TRAMITAÇÃO DE PROPOSIÇÕES: RELATÓRIO DE VISITA",
            "TRAMITAÇÃO DE PROPOSIÇÕES: PROPOSTAS DE AÇÃO LEGISLATIVA REFERENTES AO PPAG",
            "TRAMITAÇÃO DE PROPOSIÇÕES: DESPACHO DE REQUERIMENTOS, RETIRADA DE TRAMITAÇÃO",
            "CORRESPONDÊNCIA: OFÍCIOS - PROJETOS DE LEI",
            "CORRESPONDÊNCIA: OFÍCIOS - REQUERIMENTOS",
            "CORRESPONDÊNCIA: OFÍCIOS - PRORROGAÇÃO DE PRAZO",
            "CORRESPONDÊNCIA: OFÍCIOS - PROPOSTA DE EMENDA À CONSTITUIÇÃO",
            "VOTAÇÕES NOMINAIS - PROJETOS DE LEI",
            "VOTAÇÕES NOMINAIS - PROJETOS DE RESOLUÇÃO",
            "VOTAÇÕES NOMINAIS - REDAÇÃO FINAL",
            "VOTAÇÕES NOMINAIS - REQUERIMENTOS",
            "VOTAÇÃO DE REQUERIMENTOS",
            "ERRATAS",
            "PARECERES SOBRE VETO",
            "VETO PARCIAL A PROPOSIÇÃO DE LEI",
            "VETO PARCIAL A PROPOSIÇÃO DE LEI, COM ANEXADOS",
            "VETO TOTAL A PROPOSIÇÃO DE LEI",
            "VETO TOTAL A PROPOSIÇÃO DE LEI, COM ANEXADOS",
            "VETO PARCIAL A PROPOSIÇÃO DE LEI COMPLEMENTAR",
            "VETO PARCIAL A PROPOSIÇÃO DE LEI COMPLEMENTAR, COM ANEXADOS",
            "VETO TOTAL A PROPOSIÇÃO DE LEI COMPLEMENTAR",
            "VETO TOTAL A PROPOSIÇÃO DE LEI COMPLEMENTAR, COM ANEXADOS",
        ]

        LISTA_DROPDOWN_2 = [
            "-",
            "DESIGNAÇÃO DE RELATORIA",
            "RECEBIMENTO DE PROPOSIÇÃO",
            "CUMPRIMENTO DE DILIGÊNCIA",
            "CONSULTA PÚBLICA",
            "ENTREGA DE DIPLOMA",
            "REUNIÃO ORIGINADA DE REQUERIMENTO",
            "REUNIÃO COM DEBATE DE PROPOSIÇÃO AGENDADA",
            "REUNIÃO COM DEBATE DE PROPOSIÇÃO REALIZADA",
            "REUNIÃO COM DEBATE DE PROPOSIÇÃO CANCELADA",
            "REMESSA - PEDIDO DE INFORMAÇÃO",
            "REMESSA - REQUERIMENTO APROVADO",
            "OFÍCIO - REQUERIMENTO APROVADO",
            "OFÍCIO - PEDIDO DE INFORMAÇÃO",
            "OFÍCIO - MANIFESTAÇÃO DE APOIO",
            "OFÍCIO - VOTO DE CONGRATULAÇÕES",
            "OFÍCIO - MANIFESTAÇÃO DE PESAR",
            "OFÍCIO - MANIFESTAÇÃO DE REPÚDIO",
            "PROPOSIÇÃO DE LEI ENCAMINHADA PARA SANÇÃO",
            "AUDIÊNCIA PÚBLICA",
        ]

        LISTA_DROPDOWN_3 = [
            "-",
            "CCJ",
            "APU", "AAG", "AMR",
            "CDM", "CHR", "CTA",
            "DCC", "DPD", "DEC", "DHU",
            "ECT", "ELJ",
            "FFO",
            "MAD", "MEN",
            "PPO", "PCD",
            "SAU",
            "SPU",
            "TPA",
            "TCO",
            "ESP",
            "CTG",
            "CIP",
            "RED",
            "SGM",
        ]

        LISTA_DROPDOWN_4 = [
            "-",
            "PRECLUSÃO DE PRAZO: PROJETOS DE LEI",
            "PRECLUSÃO DE PRAZO: REQUERIMENTOS, APROVADOS",
            "PRECLUSÃO DE PRAZO: REQUERIMENTOS, REJEITADOS",
            "PRECLUSÃO DE PRAZO: REQUERIMENTOS, RECURSO",
            "PRECLUSÃO DE PRAZO: INCONSTITUCIONALIDADE",
        ]

        LISTA_DROPDOWN_5 = [
            "-",
            "?",
            "ALINE",
            "ANDRÉ",
            "DIOGO",
            "KÁTIA",
            "LEO",
        ]

        LISTA_DROPDOWN_6 = [
            "PL",
            "PLC",
            "PEC",
            "PRE",
            "RQN",
            "MSG",
            "OFI",
            "IND",
            "VET",
            "REL",
        ]

        LISTA_DROPDOWN_7 = [
            "2026",
            "2025",
            "2024",
            "2023",
            "2022",
            "2021",
            "2020",
            "2019",
            "2018",
            "2017",
            "2016",
            "2015",
        ]

        def _dv_req(col0: int, row1: int, values_list: list[str], strict: bool = True):
            return {
                "setDataValidation": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row1 - 1,
                        "endRowIndex": row1,
                        "startColumnIndex": col0,
                        "endColumnIndex": col0 + 1,
                    },
                    "rule": {
                        "condition": {
                            "type": "ONE_OF_LIST",
                            "values": [{"userEnteredValue": v} for v in values_list],
                        },
                        "strict": strict,
                        "showCustomUi": True,
                    },
                }
            }

        def _hex_to_rgb01(hex_color: str) -> dict:
            h = hex_color.lstrip("#")
            return {
                "red": int(h[0:2], 16) / 255.0,
                "green": int(h[2:4], 16) / 255.0,
                "blue": int(h[4:6], 16) / 255.0,
            }

        def _cf_req(col0: int, row1: int, bg_hex: str, fg_hex: str, index: int = 0):
            """Conditional formatting: pinta a célula quando NÃO estiver vazia (vale p/ qualquer opção do dropdown)."""
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
                            "condition": {"type": "NOT_BLANK"},
                            "format": {
                                "backgroundColor": _hex_to_rgb01(bg_hex),
                                "textFormat": {"foregroundColor": _hex_to_rgb01(fg_hex)},
                            },
                        },
                    },
                    "index": index,
                }
            }

        def _cf_left_of_c_req(row1: int, bg_hex: str, fg_hex: str, index: int = 0):
            """
            Pinta a célula B{row1} (à esquerda da coluna C) com as mesmas cores do dropdown da coluna C.
            Condição: C{row1} não vazia.
            """
            return {
                "addConditionalFormatRule": {
                    "rule": {
                        "ranges": [{
                            "sheetId": sheet_id,
                            "startRowIndex": row1 - 1,
                            "endRowIndex": row1,
                            "startColumnIndex": 1,   # B
                            "endColumnIndex": 2
                        }],
                        "booleanRule": {
                            "condition": {
                                "type": "CUSTOM_FORMULA",
                                "values": [{"userEnteredValue": f'=$C{row1}<>""'}]
                            },
                            "format": {
                                "backgroundColor": _hex_to_rgb01(bg_hex),
                                "textFormat": {"foregroundColor": _hex_to_rgb01(fg_hex)},
                            },
                        },
                    },
                    "index": index,
                }
            }

        def _set_value_req(col0: int, row1: int, value: str):
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
                            "userEnteredValue": {"stringValue": value}
                        }]
                    }],
                    "fields": "userEnteredValue",
                }
            }

        # ---------------------------
        # DROPDOWN 1 (BLOCO PRINCIPAL)
        # ---------------------------
        # Aplica na coluna C, nas linhas de itens "normais", antes dos extras.
        # Usa start_items_row se existir; caso não exista no seu código, ele precisa ser o início real dos itens.
        start_items_row = locals().get("start_items_row", 9)
        end_items_row = start_extra_row - 1

        if end_items_row >= start_items_row:
            for r in range(start_items_row, end_items_row + 1):
                # garante C:D mesclado
                reqs.append(req_unmerge(sheet_id, f"C{r}:D{r}"))
                reqs.append(req_merge(sheet_id, f"C{r}:D{r}"))

                # dropdown 1 na coluna C
                reqs.append(_dv_req(2, r, LISTA_DROPDOWN_1, strict=False))
                reqs.append(_cf_req(2, r, bg_hex=LISTA_DROPDOWNS_1_BG, fg_hex=LISTA_DROPDOWNS_1_FG, index=0))
                reqs.append(_cf_left_of_c_req(r, bg_hex=LISTA_DROPDOWNS_1_BG, fg_hex=LISTA_DROPDOWNS_1_FG, index=0))

        # ----------------------------------
        # DROPDOWN 5 (F e G) – default "?"
        # + COLUNA H: só "?" (sem dropdown) — só ITENS do DL (OUTs)
        # + COLUNA I: checkbox — só ITENS do DL (OUTs)
        # ----------------------------------

        start_items_row = locals().get("start_items_row", 9)
        start_extra_row = locals().get("start_extra_row", 9 + (len(itens) if itens else 0))
        end_items_row   = start_extra_row - 1

        if end_items_row >= start_items_row:
            for row1 in range(start_items_row, end_items_row + 1):

                # coluna F (dropdown + ?)
                reqs.append(_dv_req(5, row1, LISTA_DROPDOWN_5, strict=False))
                reqs.append(_set_value_req(5, row1, "?"))
                reqs.append(_cf_req(5, row1, bg_hex="#ffffff", fg_hex="#cc0000", index=0))

                # coluna G (dropdown + ?)
                reqs.append(_dv_req(6, row1, LISTA_DROPDOWN_5, strict=False))
                reqs.append(_set_value_req(6, row1, "?"))
                reqs.append(_cf_req(6, row1, bg_hex="#ffffff", fg_hex="#cc0000", index=0))

                # coluna H (SÓ ? — SEM dropdown)
                reqs.append(_set_value_req(7, row1, "?"))  # 7 = H
                reqs.append(_cf_req(7, row1, bg_hex="#ffffff", fg_hex="#cc0000", index=0))

                # coluna I (checkbox) — só itens
                for req in _checkbox_req(sheet_id, 8, row1, default_checked=False):  # 8 = I
                    reqs.append(req)

        # ---------------------------
        # DROPDOWNS NOS EXTRAS
        # + COLUNA H checkbox (fonte 6 vermelha) nas linhas válidas
        # ---------------------------
        for i, (_b, c) in enumerate(extras):
            r = start_extra_row + i

            if c == "DROPDOWN_2":
                reqs.append(_dv_req(2, r, LISTA_DROPDOWN_2))
                reqs.append(_cf_req(2, r, bg_hex="#e6cff2", fg_hex="#5a3286", index=0))
                reqs.append(_cf_left_of_c_req(r, bg_hex="#e6cff2", fg_hex="#5a3286", index=0))

                reqs.append(_dv_req(3, r, LISTA_DROPDOWN_3))
                reqs.append(_cf_req(3, r, bg_hex="#e6cff2", fg_hex="#5a3286", index=0))

            elif c == "DROPDOWN_4":
                reqs.append(_dv_req(2, r, LISTA_DROPDOWN_4))
                reqs.append(_cf_req(2, r, bg_hex="#c6dbe1", fg_hex="#215a6c", index=0))
                reqs.append(_cf_left_of_c_req(r, bg_hex="#c6dbe1", fg_hex="#215a6c", index=0))

                reqs.append(req_unmerge(sheet_id, f"C{r}:D{r}"))
                reqs.append(req_merge(sheet_id, f"C{r}:D{r}"))

            else:
                if c not in ("-", ""):
                    reqs.append(req_unmerge(sheet_id, f"C{r}:D{r}"))
                    reqs.append(req_merge(sheet_id, f"C{r}:D{r}"))

            # checkbox dinâmico na coluna H só nas linhas "válidas"
            CHECKBOX_TITLES = (
                "DIÁRIO DO EXECUTIVO",
                "DIÁRIO DO LEGISLATIVO",
                "REUNIÕES DE PLENÁRIO",
                "REUNIÕES DE COMISSÕES",
                "REQUERIMENTOS DE COMISSÃO",
                "LANÇAMENTOS DE TRAMITAÇÃO",
                "CADASTRO DE E-MAILS",
                "IMPLANTAÇÃO DE TEXTOS",
            )
            if c and any(t in c.upper() for t in CHECKBOX_TITLES):
                for req in _checkbox_req(sheet_id, 7, r, default_checked=False):  # 7 = H
                    reqs.append(req)

        # styles
        for a1, mini in STYLES:
            reqs.append(req_repeat_cell(sheet_id, a1, _mini_to_user_fmt(mini)))

        # CHECKBOX FIXO NA BARRA DO TÍTULO (H6 e H8)
        for rr in (6, 8):
            for req in _checkbox_req(sheet_id, 7, rr, default_checked=False):  # 7 = H
                reqs.append(req)

        # OVERRIDE: checkbox H6/H8 com o mesmo tamanho dos outros (fonte 6)
        for r in (6, 8):
            reqs.append(req_repeat_cell(sheet_id, f"H{r}:H{r}", {
                "textFormat": {
                    "fontFamily": "Inconsolata",
                    "fontSize": 6,
                    "foregroundColor": rgb_hex_to_api("#cc0000"),
                }
            }))

        # ---------------------------------------------------------------------------------------
        # OVERRIDES (imediatamente após STYLES) — pra não ser sobrescrito
        # - H (itens/OUTs): Inconsolata 8 vermelho
        # - I (itens/OUTs): Inconsolata 6 vermelho (checkbox menor)
        # - H (extras com checkbox): Inconsolata 6 vermelho
        # ---------------------------------------------------------------------------------------

        # garante ranges (sempre recalcule aqui, não dependa de variável antiga)
        start_items_row = locals().get("start_items_row", 9)
        start_extra_row = locals().get("start_extra_row", 9 + (len(itens) if itens else 0))
        end_items_row   = start_extra_row - 1
        extra_end_row   = start_extra_row + (len(extras) if extras else 0) - 1

        # ITENS (OUTs): H=8, I=6
        if end_items_row >= start_items_row:
            reqs.append(req_text(sheet_id, f"H{start_items_row}:H{end_items_row}", "Inconsolata", 8, "#cc0000"))
            reqs.append(req_text(sheet_id, f"I{start_items_row}:I{end_items_row}", "Inconsolata", 6, "#cc0000"))

        # EXTRAS: H=6 (só pra manter os checkboxes dos extras pequenos e vermelhos)
        if extra_end_row >= start_extra_row:
            reqs.append(req_text(sheet_id, f"H{start_extra_row}:H{extra_end_row}", "Inconsolata", 6, "#cc0000"))

        # borders
        for a1, spec in BORDERS:
            # corta qualquer range que vá até rows_needed e evita encostar na linha técnica
            a1_fix = a1.replace(f"{rows_needed}", f"{VIS_LAST_ROW_1BASED}") if "rows_needed" in locals() else a1
            kwargs = {}
            for side, (style_name, color_name) in spec.items():
                kwargs[side] = _border_from_spec(style_name, color_name)
            reqs.append(req_update_borders(sheet_id, a1_fix, **kwargs))

    # PARTE 2 =====================================================================================================================================================================================
    # ============================================================================================ FOOTER ==============================================================================================
    # ====================================================================================================================================================================================================

        # Se NÃO houver OUTs, ancora o footer após o template fixo (até a linha 21 no layout atual)
        # (isso evita depender de extra_end gerado por OUT)
        TEMPLATE_END_ROW = 21  # 1-based: ajuste só se você mudar o template fixo

        if not extras_out:
            extra_end = TEMPLATE_END_ROW  # força base
        # se tiver OUTs, extra_end já veio correto do bloco de extras

        # EXTRA: extra_end é "fim" (1-based) -> footer começa na PRÓXIMA linha
        footer_start = extra_end + 1

        footer_rows = 9
        footer_end  = footer_start + footer_rows - 1

        r  = footer_start
        r1 = footer_start + 1
        r2 = footer_start + 2
        r3 = footer_start + 3
        r4 = footer_start + 4
        r5 = footer_start + 5
        r6 = footer_start + 6
        r7 = footer_start + 7
        r8 = footer_start + 8

        # -------------------------------------------------------------------------------------------------------------------------------------------------
        # -------------------------------------------------------------------- VALUES ---------------------------------------------------------------------
        # -------------------------------------------------------------------------------------------------------------------------------------------------
        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r - 1,  "endRowIndex": r,  "startColumnIndex": 1,  "endColumnIndex": 2},  # B
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=HYPERLINK("http://meet.google.com/api-pefj-mvq";"GDI-GGA")'}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 0,  "endColumnIndex": 1},  # A
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=HYPERLINK("https://mediaserver.almg.gov.br/acervo/511/376/2511376.pdf";IMAGE("https://cdn-icons-png.flaticon.com/512/3079/3079014.png";4;19;19))'}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 1,  "endColumnIndex": 2},  # B
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=HYPERLINK("https://intra.almg.gov.br/export/sites/default/atendimento/docs/lista-telefonica.pdf";IMAGE("https://cdn-icons-png.flaticon.com/512/4783/4783130.png";4;33;33))'}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 2,  "endColumnIndex": 3},  # C
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=HYPERLINK("https://sites.google.com/view/gga-gdi-almg/";IMAGE("https://yt3.ggpht.com/ytc/AKedOLS-fgkzGxYUBgBejVblA1CLhE69pbiZyoH7spcNRQ=s900-c-k-c0x00ffffff-no-rj";4;112;125))'}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 4,  "endColumnIndex": 5},  # E
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=SUM(FILTER(INDIRECT("F"&ROW()+1&":F");INDIRECT("E"&ROW()+1&":E")<>""))'}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 5,  "endColumnIndex": 6},  # F
                "rows": [{"values": [{"userEnteredValue": {"stringValue": "TOTAL"}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": r2 - 1, "endRowIndex": r6, "startColumnIndex": 5, "endColumnIndex": 6},  # F (hyperlink)
                "cell": {"userEnteredValue": {"formulaValue": '=SUMIFS(INDEX(H:H;ROW()):INDEX(H:H;ROW()+6);INDEX(G:G;ROW()):INDEX(G:G;ROW()+6);INDEX(E:E;ROW()))+SUMIFS(INDEX(K:K;ROW()):INDEX(K:K;ROW()+6);INDEX(I:I;ROW()):INDEX(I:I;ROW()+6);INDEX(E:E;ROW()))'}},
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 6,  "endColumnIndex": 7},  # G
                "rows": [{"values": [{"userEnteredValue": {"stringValue": "IMPLANTAÇÃO"}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 7,  "endColumnIndex": 8},  # H (ícone)
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=IMAGE("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRYV-RpYwK3orapycj_CXJGevAVSORX9_E2jUYZLgID8L3bLwfSRXMX7ksvRTsEEoRBeNE&usqp=CAU";4;17;17)'}}]}],
                "fields": "userEnteredValue"}})
        reqs.append({"repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": r2 - 1, "endRowIndex": r6, "startColumnIndex": 7, "endColumnIndex": 8},  # L (hyperlink)
                "cell": {"userEnteredValue": {"formulaValue": '=SUMIFS($E$1:INDEX($E:$E;ROW()-1);$F$1:INDEX($F:$F;ROW()-1);INDEX($G:$G;ROW()))'}},
                "fields": "userEnteredValue"}})
        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 8,  "endColumnIndex": 9},  # I
                "rows": [{"values": [{"userEnteredValue": {"stringValue": "CONFERÊNCIA"}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 10, "endColumnIndex": 11},  # K (ícone)
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=IMAGE("https://w7.pngwing.com/pngs/894/494/png-transparent-black-male-symbol-art-avatar-education-professor-user-profile-faculty-boss-face-heroes-service-thumbnail.png";4;17;17)'}}]}],

                "fields": "userEnteredValue"}})
        reqs.append({"repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": r2 - 1, "endRowIndex": r6, "startColumnIndex": 10, "endColumnIndex": 11},  # K (hyperlink)
                "cell": {"userEnteredValue": {"formulaValue": '=SUMIFS($H$6:INDEX($H:$H;ROW()-1);$G$6:INDEX($G:$G;ROW()-1);INDEX($I:$I;ROW()))'}},
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 11, "endColumnIndex": 12},  # L (ícone)
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=IMAGE("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRyxXB7iHrkoP3waMJDQVtKeDlVpA7sno_XMNVpY20s5rmcQyJh")'}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": r2 - 1, "endRowIndex": r8, "startColumnIndex": 11, "endColumnIndex": 12},  # L (hyperlink)
                "cell": {"userEnteredValue": {"formulaValue": '=HYPERLINK("https://www.almg.gov.br/atividade_parlamentar/tramitacao_projetos/interna.html?a="&INDIRECT("O"&ROW())&"&n="&INDIRECT("N"&ROW())&"&t="&INDIRECT("M"&ROW())&"&aba=js_tabTramitacao";IMAGE("https://seeklogo.com/images/B/bandeira-minas-gerais-logo-AD7B6F3604-seeklogo.com.png";4;14;14))'}},
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 12, "endColumnIndex": 13},  # M
                "rows": [{"values": [{"userEnteredValue": {"stringValue": "PROPOSIÇÕES RELEVANTES"}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 15, "endColumnIndex": 16},  # P
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=HYPERLINK("https://dspace.almg.gov.br/server/api/core/bitstreams/7cd591b0-1a2c-41cc-9341-78919e827df1/content";IMAGE("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcS5Y_YMt-RiPPZVOKRYkqMru870B2Pa2Kbsg-Ck-1KphTkHW3XM0Vtlb7MgVZRRxsfFtJY&usqp=CAU";4;130;140))'}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 17, "endColumnIndex": 18},  # R
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=HYPERLINK("https://www.cbhdoce.org.br/wp-content/uploads/2016/01/ConstituicaoEstadual.pdf";IMAGE("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcT5c5T7Fvx5UqHvPvb1EWmn6zxEyl9XZua3dQ&s";4;130;140))'}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 19, "endColumnIndex": 20},  # T
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=HYPERLINK("https://www.planalto.gov.br/ccivil_03/constituicao/ConstituicaoCompilado.htm";IMAGE("https://www2.camara.leg.br/atividade-legislativa/legislacao/Constituicoes_Brasileiras/constituicao-cidada/regulamentacao/imagens/copy_of_1.jpg/@@images/8beeb113-f656-495f-9c90-c81fb62a2ebb.jpeg";4;130;125))'}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 21, "endColumnIndex": 22},  # V
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=HYPERLINK("https://dspace.almg.gov.br/server/api/core/bitstreams/7cd591b0-1a2c-41cc-9341-78919e827df1/content";IMAGE("https://www.aracruz.es.leg.br/imagens/PORTLETREGIMENTOINTERNO.png/image_preview";4;130;125))'}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 23, "endColumnIndex": 24},  # X
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=HYPERLINK("https://www.cbhdoce.org.br/wp-content/uploads/2016/01/ConstituicaoEstadual.pdf";IMAGE("https://upload.wikimedia.org/wikipedia/commons/d/d2/Bras%C3%A3o_de_Minas_Gerais.svg"))'}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r3 - 1, "endRowIndex": r3, "startColumnIndex": 0,  "endColumnIndex": 1},  # A
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=HYPERLINK("https://intra.almg.gov.br/acontece/noticias/";IMAGE("https://intra.almg.gov.br/.content/imagens/logo-intra.svg";4;20;75))'}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r6 - 1, "endRowIndex": r6, "startColumnIndex": 0,  "endColumnIndex": 1},  # A
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=HYPERLINK("https://calendar.google.com/calendar/u/0?cid=a3RyajJsZmRwdGpxYTdrczBqNXVhbXBldmdAZ3JvdXAuY2FsZW5kYXIuZ29vZ2xlLmNvbQ";IMAGE("https://cdn-icons-png.flaticon.com/512/217/217837.png";4;18;18))'}}]}],
                "fields": "userEnteredValue"}})

        reqs.append({"updateCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": r6 - 1, "endRowIndex": r6, "startColumnIndex": 1,  "endColumnIndex": 2},  # B
                "rows": [{"values": [{"userEnteredValue": {"formulaValue": '=HYPERLINK("https://ead.almg.gov.br/moodle/";IMAGE("https://ead.almg.gov.br/moodle/pluginfile.php/2/course/section/288/servidor_ALMG.png?time=1657626782411"))'}}]}],
                "fields": "userEnteredValue"}})

        # ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        # ----------------------------------------------------------------------------------------------- MERGES -----------------------------------------------------------------------------------------------
        # ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r2, "startColumnIndex": 0, "endColumnIndex": 1}, "mergeType": "MERGE_ALL"}})  # CALENDAR A
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r2, "startColumnIndex": 1, "endColumnIndex": 2}, "mergeType": "MERGE_ALL"}})  # PHONE B
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r3 - 1, "endRowIndex": r5, "startColumnIndex": 0, "endColumnIndex": 2}, "mergeType": "MERGE_ALL"}})  # INTRA A:B
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r6 - 1, "endRowIndex": r8, "startColumnIndex": 0, "endColumnIndex": 1}, "mergeType": "MERGE_ALL"}})  # AGENDA A
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r6 - 1, "endRowIndex": r8, "startColumnIndex": 1, "endColumnIndex": 2}, "mergeType": "MERGE_ALL"}})  # GGA B
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r8, "startColumnIndex": 2, "endColumnIndex": 4}, "mergeType": "MERGE_ALL"}})  # ALMG C:D
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 8,  "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}})  # CONFERÊNCIA (título)
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 12, "endColumnIndex": 15}, "mergeType": "MERGE_ALL"}})  # PROPOSIÇÕES (título)
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r8, "startColumnIndex": 15, "endColumnIndex": 17}, "mergeType": "MERGE_ALL"}})  # REGIMENTO
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r8, "startColumnIndex": 17, "endColumnIndex": 19}, "mergeType": "MERGE_ALL"}})  # CONST. ESTADUAL
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r8, "startColumnIndex": 19, "endColumnIndex": 21}, "mergeType": "MERGE_ALL"}})  # CONST. FEDERAL
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r8, "startColumnIndex": 21, "endColumnIndex": 23}, "mergeType": "MERGE_ALL"}})  # REGIMENTO (img)
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r8, "startColumnIndex": 23, "endColumnIndex": 25}, "mergeType": "MERGE_ALL"}})  # CONST. ESTADUAL (img)
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r2 - 1, "endRowIndex": r2, "startColumnIndex": 8, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}})  # ALINE
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r3 - 1, "endRowIndex": r3, "startColumnIndex": 8, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}})  # ANDRÉ
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r4 - 1, "endRowIndex": r4, "startColumnIndex": 8, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}})  # DIOGO
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r5 - 1, "endRowIndex": r5, "startColumnIndex": 8, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}})  # KÁTIA
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r6 - 1, "endRowIndex": r6, "startColumnIndex": 8, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}})  # LEO
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r7 - 1, "endRowIndex": r7, "startColumnIndex": 8, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}})  # VINÍCIUS
        reqs.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r8 - 1, "endRowIndex": r8, "startColumnIndex": 8, "endColumnIndex": 10}, "mergeType": "MERGE_ALL"}})  # WELDER

        for r in range(6, extra_end + 1):   # COLUNAS J:O
            reqs.append(req_merge(sheet_id, f"J{r}:O{r}"))

        # ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        # ----------------------------------------------------------------------------------------------- STYLES -----------------------------------------------------------------------------------------------
        # ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": extra_end, "endRowIndex": extra_end, "startColumnIndex": 2, "endColumnIndex": 3},
            "cell": {"userEnteredFormat": {"horizontalAlignment": "LEFT", "verticalAlignment": "MIDDLE"}},
            "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment)"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r - 1, "endRowIndex": r8, "startColumnIndex": 2, "endColumnIndex": 3},
            "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE", "backgroundColor": {"red": 0.953, "green": 0.953, "blue": 0.953}}},
            "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment,backgroundColor)"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r5 - 1, "startColumnIndex": 0, "endColumnIndex": 2},
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.9764706, "green": 0.7960784, "blue": 0.6117647}}},
            "fields": "userEnteredFormat(backgroundColor)"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r3 - 1, "endRowIndex": r6 - 1, "startColumnIndex": 0, "endColumnIndex": 2},
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.988, "green": 0.820, "blue": 0.800}}},
            "fields": "userEnteredFormat(backgroundColor)"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 4, "endColumnIndex": 6},
            "cell": {"userEnteredFormat": {
                "backgroundColor": {"red": 0.6, "green": 0.0, "blue": 0.0},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "textFormat": {"fontFamily": "Vidaloka", "fontSize": 8, "bold": True, "foregroundColor": {"red": 0.85, "green": 0.67, "blue": 0.10}}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 6, "endColumnIndex": 11},
            "cell": {"userEnteredFormat": {
                "backgroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "textFormat": {"fontFamily": "Vidaloka", "fontSize": 7, "bold": True, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 11, "endColumnIndex": 15},
            "cell": {"userEnteredFormat": {
                "backgroundColor": {"red": 0.9019608, "green": 0.5686275, "blue": 0.21960788},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "textFormat": {"fontFamily": "Vidaloka", "fontSize": 7, "bold": True, "foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r1, "endRowIndex": r8, "startColumnIndex": 11, "endColumnIndex": 15},
            "cell": {"userEnteredFormat": {
                "backgroundColor": {"red": 1.0, "green": 0.949, "blue": 0.8},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
                "textFormat": {"fontFamily": "Special Elite", "fontSize": 8, "bold": True, "foregroundColor": {"red": 0.6, "green": 0.0, "blue": 0.0}}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r2 - 1, "endRowIndex": r8, "startColumnIndex": 4, "endColumnIndex": 5},
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.741, "green": 0.741, "blue": 0.741}, "horizontalAlignment": "RIGHT", "verticalAlignment": "MIDDLE",
                    "textFormat": {"fontFamily": "Boogaloo", "fontSize": 8, "bold": False}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r2 - 1, "endRowIndex": r8, "startColumnIndex": 5, "endColumnIndex": 6},
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.741, "green": 0.741, "blue": 0.741}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                    "textFormat": {"fontFamily": "Boogaloo", "fontSize": 8, "bold": False}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r2 - 1, "endRowIndex": r8, "startColumnIndex": 6, "endColumnIndex": 7},
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.741, "green": 0.741, "blue": 0.741}, "horizontalAlignment": "RIGHT", "verticalAlignment": "MIDDLE",
                    "textFormat": {"fontFamily": "Boogaloo", "fontSize": 8, "bold": False}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r2 - 1, "endRowIndex": r8, "startColumnIndex": 7, "endColumnIndex": 8},
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.741, "green": 0.741, "blue": 0.741}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                    "textFormat": {"fontFamily": "Boogaloo", "fontSize": 8, "bold": False}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r2 - 1, "endRowIndex": r8, "startColumnIndex": 8, "endColumnIndex": 9},
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.741, "green": 0.741, "blue": 0.741}, "horizontalAlignment": "RIGHT", "verticalAlignment": "MIDDLE",
                    "textFormat": {"fontFamily": "Boogaloo", "fontSize": 8, "bold": False}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}})
        reqs.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r2 - 1, "endRowIndex": r8, "startColumnIndex": 10, "endColumnIndex": 11},
            "cell": {"userEnteredFormat": {"backgroundColor": {"red": 0.741, "green": 0.741, "blue": 0.741}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                    "textFormat": {"fontFamily": "Boogaloo", "fontSize": 8, "bold": False}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)"}})

        # ----------------------------------------------------------------------------------------------------------
        # -------------------------------------------------- VALUES ------------------------------------------------
        # ----------------------------------------------------------------------------------------------------------
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": 4,
                "endColumnIndex": 5},"rows": [{"values": [{"userEnteredValue": {"stringValue": "ALINE"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": 4,
                "endColumnIndex": 5},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r2,
                "endRowIndex": r3,
                "startColumnIndex": 4,
                "endColumnIndex": 5},"rows": [{"values": [{"userEnteredValue": {"stringValue": "ANDRÉ"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r2,
                "endRowIndex": r3,
                "startColumnIndex": 4,
                "endColumnIndex": 5},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r3,
                "endRowIndex": r4,
                "startColumnIndex": 4,
                "endColumnIndex": 5},"rows": [{"values": [{"userEnteredValue": {"stringValue": "DIOGO"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r3,
                "endRowIndex": r4,
                "startColumnIndex": 4,
                "endColumnIndex": 5},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r4,
                "endRowIndex": r5,
                "startColumnIndex": 4,
                "endColumnIndex": 5},"rows": [{"values": [{"userEnteredValue": {"stringValue": "KÁTIA"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r4,
                "endRowIndex": r5,
                "startColumnIndex": 4,
                "endColumnIndex": 5},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r5,
                "endRowIndex": r6,
                "startColumnIndex": 4,
                "endColumnIndex": 5},"rows": [{"values": [{"userEnteredValue": {"stringValue": "LEO"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r5,
                "endRowIndex": r6,
                "startColumnIndex": 4,
                "endColumnIndex": 5},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])

        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": 6,
                "endColumnIndex": 7},"rows": [{"values": [{"userEnteredValue": {"stringValue": "ALINE"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": 6,
                "endColumnIndex": 7},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r2,
                "endRowIndex": r3,
                "startColumnIndex": 6,
                "endColumnIndex": 7},"rows": [{"values": [{"userEnteredValue": {"stringValue": "ANDRÉ"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r2,
                "endRowIndex": r3,
                "startColumnIndex": 6,
                "endColumnIndex": 7},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r3,
                "endRowIndex": r4,
                "startColumnIndex": 6,
                "endColumnIndex": 7},"rows": [{"values": [{"userEnteredValue": {"stringValue": "DIOGO"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r3,
                "endRowIndex": r4,
                "startColumnIndex": 6,
                "endColumnIndex": 7},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r4,
                "endRowIndex": r5,
                "startColumnIndex": 6,
                "endColumnIndex": 7},"rows": [{"values": [{"userEnteredValue": {"stringValue": "KÁTIA"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r4,
                "endRowIndex": r5,
                "startColumnIndex": 6,
                "endColumnIndex": 7},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r5,
                "endRowIndex": r6,
                "startColumnIndex": 6,
                "endColumnIndex": 7},"rows": [{"values": [{"userEnteredValue": {"stringValue": "LEO"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r5,
                "endRowIndex": r6,
                "startColumnIndex": 6,
                "endColumnIndex": 7},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])

        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": 6,
                "endColumnIndex": 7},"rows": [{"values": [{"userEnteredValue": {"stringValue": "ALINE"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": 6,
                "endColumnIndex": 7},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r2,
                "endRowIndex": r3,
                "startColumnIndex": 6,
                "endColumnIndex": 7},"rows": [{"values": [{"userEnteredValue": {"stringValue": "ANDRÉ"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r2,
                "endRowIndex": r3,
                "startColumnIndex": 6,
                "endColumnIndex": 7},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r3,
                "endRowIndex": r4,
                "startColumnIndex": 6,
                "endColumnIndex": 7},"rows": [{"values": [{"userEnteredValue": {"stringValue": "DIOGO"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r3,
                "endRowIndex": r4,
                "startColumnIndex": 6,
                "endColumnIndex": 7},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r4,
                "endRowIndex": r5,
                "startColumnIndex": 6,
                "endColumnIndex": 7},"rows": [{"values": [{"userEnteredValue": {"stringValue": "KÁTIA"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r4,
                "endRowIndex": r5,
                "startColumnIndex": 6,
                "endColumnIndex": 7},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r5,
                "endRowIndex": r6,
                "startColumnIndex": 6,
                "endColumnIndex": 7},"rows": [{"values": [{"userEnteredValue": {"stringValue": "LEO"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r5,
                "endRowIndex": r6,
                "startColumnIndex": 6,
                "endColumnIndex": 7},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])

        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": 8,
                "endColumnIndex": 9},"rows": [{"values": [{"userEnteredValue": {"stringValue": "ALINE"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r2,
                "startColumnIndex": 8,
                "endColumnIndex": 9},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r2,
                "endRowIndex": r3,
                "startColumnIndex": 8,
                "endColumnIndex": 9},"rows": [{"values": [{"userEnteredValue": {"stringValue": "ANDRÉ"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r2,
                "endRowIndex": r3,
                "startColumnIndex": 8,
                "endColumnIndex": 9},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r3,
                "endRowIndex": r4,
                "startColumnIndex": 8,
                "endColumnIndex": 9},"rows": [{"values": [{"userEnteredValue": {"stringValue": "DIOGO"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r3,
                "endRowIndex": r4,
                "startColumnIndex": 8,
                "endColumnIndex": 9},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r4,
                "endRowIndex": r5,
                "startColumnIndex": 8,
                "endColumnIndex": 9},"rows": [{"values": [{"userEnteredValue": {"stringValue": "KÁTIA"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r4,
                "endRowIndex": r5,
                "startColumnIndex": 8,
                "endColumnIndex": 9},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"updateCells": {"range": {"sheetId": sheet_id,
                "startRowIndex": r5,
                "endRowIndex": r6,
                "startColumnIndex": 8,
                "endColumnIndex": 9},"rows": [{"values": [{"userEnteredValue": {"stringValue": "LEO"}}]}],"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r5,
                "endRowIndex": r6,
                "startColumnIndex": 8,
                "endColumnIndex": 9},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_5]},"showCustomUi": True,"strict": False}}}])
        reqs.extend([
        {"repeatCell": {"range": {"sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r8,
                "startColumnIndex": 12,
                "endColumnIndex": 13},
            "cell": {"userEnteredValue": {"stringValue": "PL"}},"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r8,
                "startColumnIndex": 12,
                "endColumnIndex": 13},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_6]},"showCustomUi": True,"strict": True}}}])
        reqs.extend([
        {"repeatCell": {"range": {"sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r8,
                "startColumnIndex": 14,
                "endColumnIndex": 15},
            "cell": {"userEnteredValue": {"stringValue": "2026"}},"fields": "userEnteredValue"}},
        {"setDataValidation": {"range": {"sheetId": sheet_id,
                "startRowIndex": r1,
                "endRowIndex": r8,
                "startColumnIndex": 14,
                "endColumnIndex": 15},
            "rule": {"condition": {"type": "ONE_OF_LIST","values": [{"userEnteredValue": v} for v in LISTA_DROPDOWN_7]},"showCustomUi": True,"strict": True}}}])

        # ----------------------------------------------------------------------------------------------------------
        # -------------------------------------------------- NOTES --------------------------------------------------
        # ----------------------------------------------------------------------------------------------------------
        reqs.append({"updateCells": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r1, "startColumnIndex": 1, "endColumnIndex": 2},
            "rows": [{"values": [{"note": "7776 ar-condicionado\n7870 gerência de saúde (- Marcos/Alberto)\n7710 informática\n7468 frequência (Milena)\n7786 polícia legislativa\n7885 plenário (Elton)"}]}],
            "fields": "note"}})

        reqs.append({"updateCells": {"range": {"sheetId": sheet_id, "startRowIndex": 7, "endRowIndex": 8, "startColumnIndex": 3, "endColumnIndex": 4},
            "rows": [{"values": [{"note": "CAPÍTULO 2-6, pág 19.\n\n"
                                        "Prioridades de lançamento (o que devo lançar primeiro?)\n\n"
                                        "1) PLs, PLCs, PREs e PECs novos\n"
                                        "2) RQNs novos\n"
                                        "3) Proposições não recebidas\n"
                                        "4) RQNs aprovados\n"
                                        "5) Todos os lançamentos que envolvam a implantação de um boletim novo em proposição anteriormente já implantada\n"
                                        "6) Lançamentos que consistam não em acrescentar boletins, mas apenas em adicionar uma frase a um boletim já implantado (ex: \"Publicado no DL em...\")"}]}],
            "fields": "note"}})

        # ----------------------------------------------------------------------------------------------------------
        # -------------------------------------------------- BORDERS ------------------------------------------------
        # ----------------------------------------------------------------------------------------------------------
        reqs.append({"updateBorders": {"range": {"sheetId": sheet_id, "startRowIndex": 5, "endRowIndex": footer_start - 1, "startColumnIndex": 8, "endColumnIndex": 9},
            "right": {"style": "SOLID_MEDIUM", "color": {"red": 0.0, "green": 0.0, "blue": 0.0}}}})

        reqs.append({"updateBorders": {"range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r8, "startColumnIndex": 0, "endColumnIndex": 1},
            "right": {"style": "DOTTED", "color": {"red": 0.8, "green": 0.0, "blue": 0.0}}}})

        reqs.append({"updateBorders": {"range": {"sheetId": sheet_id, "startRowIndex": r7, "endRowIndex": r7 + 1, "startColumnIndex": 0, "endColumnIndex": 2},
            "top": {"style": "DOTTED", "color": {"red": 0.8, "green": 0.0, "blue": 0.0}}}})

        reqs.append({"updateBorders": {"range": {"sheetId": sheet_id, "startRowIndex": r - 1, "endRowIndex": r8, "startColumnIndex": 1, "endColumnIndex": 2},
            "right": {"style": "SOLID", "color": {"red": 0.0, "green": 0.0, "blue": 0.0}},
            "bottom": {"style": "SOLID_MEDIUM", "color": {"red": 0.0, "green": 0.0, "blue": 0.0}}}})

        reqs.append({"updateBorders": {"range": {"sheetId": sheet_id, "startRowIndex": r - 1, "endRowIndex": r8, "startColumnIndex": 4, "endColumnIndex": 5},
            "left": {"style": "SOLID", "color": {"red": 0.0, "green": 0.0, "blue": 0.0}}}})

        reqs.append({"updateBorders": {"range": {"sheetId": sheet_id, "startRowIndex": r, "endRowIndex": r8, "startColumnIndex": 17, "endColumnIndex": 18},
            "left": {"style": "SOLID_MEDIUM", "color": {"red": 0.8, "green": 0.0, "blue": 0.0}}}})

        reqs.append({"updateBorders": {"range": {"sheetId": sheet_id, "startRowIndex": r8, "endRowIndex": r8 + 1, "startColumnIndex": 0, "endColumnIndex": 25},
            "top": {"style": "SOLID_MEDIUM", "color": {"red": 0.0, "green": 0.0, "blue": 0.0}}}})

        reqs.append({"updateBorders": {"range": {"sheetId": sheet_id, "startRowIndex": r7, "endRowIndex": r8, "startColumnIndex": 0, "endColumnIndex": 25},
            "bottom": {"style": "SOLID_MEDIUM", "color": {"red": 0.0, "green": 0.0, "blue": 0.0}}}})

        # ====================================================================================================================================================================================================
        # ========================================================================================== CONDICIONAIS ============================================================================================
        # ====================================================================================================================================================================================================

        reqs.append({"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": sheet_id, "startRowIndex": 5, "endRowIndex": ws.row_count, "startColumnIndex": 0, "endColumnIndex": 25}],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=$H6=TRUE'}]},
                "format": {"backgroundColor": {"red": 0.2627450980392157, "green": 0.2627450980392157, "blue": 0.2627450980392157},"textFormat": {"foregroundColor": {"red": 0.6, "green": 0.6, "blue": 0.6}, "bold": True}}}},
            "index": 0}})

        reqs.append({"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": sheet_id, "startRowIndex": 5, "endRowIndex": ws.row_count, "startColumnIndex": 0, "endColumnIndex": 25}],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=OR($I6=TRUE;REGEXMATCH(TO_TEXT($I6);"-"))'}]},
                "format": {"backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8},"textFormat": {"foregroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}}}}},
            "index": 1}})

        reqs.append({"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": sheet_id, "startRowIndex": 5, "endRowIndex": ws.row_count, "startColumnIndex": 0, "endColumnIndex": 25}],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=REGEXMATCH($C6;"^DIÁRIO")'}]},
                "format": {"backgroundColor": {"red": 102/255, "green": 0.0, "blue": 0.0}, "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True}}}},
            "index": 2}})

        reqs.append({"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": sheet_id, "startRowIndex": 5, "endRowIndex": ws.row_count, "startColumnIndex": 0, "endColumnIndex": 25}],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=REGEXMATCH($C6;"^REUNIÕES")'}]},
                "format": {"backgroundColor": {"red": 39/255, "green": 78/255, "blue": 19/255}, "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True}}}},
            "index": 3}})

        reqs.append({"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": sheet_id, "startRowIndex": 5, "endRowIndex": ws.row_count, "startColumnIndex": 0, "endColumnIndex": 25}],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=REGEXMATCH($C6;"^REQUERIMENTOS DE COMISSÃO")'}]},
                "format": {"backgroundColor": {"red": 255/255, "green": 153/255, "blue": 0/255}, "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True}}}},
            "index": 4}})

        reqs.append({"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": sheet_id, "startRowIndex": 5, "endRowIndex": ws.row_count, "startColumnIndex": 0, "endColumnIndex": 25}],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=REGEXMATCH($C6;"^LANÇAMENTOS DE TRAMITAÇÃO")'}]},
                "format": {"backgroundColor": {"red": 32/255, "green": 18/255, "blue": 77/255}, "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True}}}},
            "index": 5}})

        reqs.append({"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": sheet_id, "startRowIndex": 5, "endRowIndex": ws.row_count, "startColumnIndex": 0, "endColumnIndex": 25}],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=REGEXMATCH($C6;"^CADASTRO DE E-MAILS")'}]},
                "format": {"backgroundColor": {"red": 7/255, "green": 55/255, "blue": 99/255}, "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True}}}},
            "index": 6}})

        reqs.append({"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": sheet_id, "startRowIndex": 5, "endRowIndex": ws.row_count, "startColumnIndex": 0, "endColumnIndex": 25}],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=REGEXMATCH($C6;"^IMPLANTAÇÃO DE TEXTOS")'}]},
                "format": {"backgroundColor": {"red": 127/255, "green": 96/255, "blue": 0/255}, "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True}}}},
            "index": 7}})

        reqs.append({"addConditionalFormatRule": {"rule": {"ranges": [{"sheetId": sheet_id, "startRowIndex": 5, "endRowIndex": ws.row_count, "startColumnIndex": 0, "endColumnIndex": 25}],
            "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": '=AND(OR(REGEXMATCH($B6;"^GDI-GGA");REGEXMATCH($C6;"^MATE")))'}]},
                "format": {"backgroundColor": {"red": 0.0, "green": 0.0, "blue": 0.0}, "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True}}}},
            "index": 8}})

        # ====================================================================================================================================================================================================
        # ============================================================================================ CHECKBOX ==============================================================================================
        # ====================================================================================================================================================================================================
        # (se você já está usando _checkbox_req no PARTE 1B, NÃO duplica aqui — mantém só o "for r in (6,8)" antigo fora do footer)

        # PARTE 3 ============================================================================================================================================================================================
        # ============================================================================================= VALUES ===============================================================================================
        # ====================================================================================================================================================================================================

        dd = int(diario_key[6:8])
        mm = int(diario_key[4:6])
        yyyy = int(diario_key[0:4])
        a5_txt = f"{dd}/{mm}"

        from datetime import datetime, timedelta

        data = []

        def add(a1, values):
            data.append({"range": f"'{tab_name}'!{a1}", "values": values})

        add("A5:B5", [[f"=DATE({yyyy};{mm};{dd})", ""]])
        add("A1", [[ '=HYPERLINK("https://www.almg.gov.br/home/index.html";IMAGE("https://sisap.almg.gov.br/banner.png";4;43;110))' ]])
        add("C1", [["GERÊNCIA DE GESTÃO ARQUIVÍSTICA"]])
        add("Q1", [["DATAS"]])
        add("G1", [['=HYPERLINK("https://intra.almg.gov.br/acontece/noticias/index.html?lq=&reloaded=&q=&di=&df=&tema=direcionamento-estrategico/";'
            'IMAGE("https://media.istockphoto.com/vectors/flag-map-of-the-brazilian-state-of-minas-gerais-vector-id1248541649?k=20&m=1248541649&s=170667a&w=0&h=V8Ky8c8rddLPjphovytIJXaB6NlMF7dt-ty-2ZJF5Wc="))']])
        add("H1", [['=HYPERLINK("https://www.almg.gov.br/atividade_parlamentar/plenario/index.html";''IMAGE("https://www.protestoma.com.br/images/noticia-id_255.jpg";4;27;42))']])
        add("H3", [['=HYPERLINK("https://www.almg.gov.br/atividade-parlamentar/comissoes/agenda/";''IMAGE("https://www.ouvidoriageral.mg.gov.br/images/noticias/2019/dezembro/foto_almg.jpg";4;27;42))']])
        add("I1", [['=HYPERLINK("https://www.jornalminasgerais.mg.gov.br/";'
            'IMAGE("https://upload.wikimedia.org/wikipedia/commons/thumb/f/f4/Bandeira_de_Minas_Gerais.svg/2560px-Bandeira_de_Minas_Gerais.svg.png";4;35;50))']])
        add("I3", [['=HYPERLINK("https://www.almg.gov.br/consulte/arquivo_diario_legislativo/index.html";'
            'IMAGE("https://www.almg.gov.br/favicon.ico";4;25;25))']])
        add("J1", [['=HYPERLINK("https://consulta-brs.almg.gov.br/brs/";''IMAGE("https://t4.ftcdn.net/jpg/04/70/40/23/360_F_470402339_5FVE7b1Z2DNI7bATV5a27FGATt6yxcEz.jpg"))']])
        add("J3", [['=HYPERLINK("https://silegis.almg.gov.br/silegismg/#/processos";IMAGE("https://silegis.almg.gov.br/silegismg/assets/logotipo.png"))']])
        add("K1", [[ '=HYPERLINK("https://webmail.almg.gov.br/";IMAGE("https://images.vexels.com/media/users/3/140138/isolated/lists/88e50689fa3280c748d000aaf0bad480-icone-redondo-de-email-1.png"))' ]])
        add("K3", [[ '=HYPERLINK("https://integracao.almg.gov.br/mate-brs/index.html";IMAGE("http://anthillonline.com/wp-content/uploads/2021/03/mate-logo.jpg";4;65;50))' ]])
        add("L1", [[ '=HYPERLINK("https://www.almg.gov.br/atividade-parlamentar/projetos-de-lei/";IMAGE("https://upload.wikimedia.org/wikipedia/commons/thumb/a/a6/Tram-Logo.svg/2048px-Tram-Logo.svg.png";4;23;23))' ]])
        add("L3", [[ '=HYPERLINK("https://www.almg.gov.br/consulte/legislacao/index.html";IMAGE("https://cdn-icons-png.flaticon.com/512/3122/3122427.png"))' ]])
        add("M1", [[ '=HYPERLINK("https://sei.almg.gov.br/";IMAGE("https://www.gov.br/ebserh/pt-br/media/plataformas/sei/@@images/5a07de59-2af0-45b0-9be9-f0d0438b7a81.png";4;45;50))' ]])
        add("M3", [[ '=HYPERLINK("https://stl.almg.gov.br/login.jsp";IMAGE("https://media-exp1.licdn.com/dms/image/C510BAQHc4JZB3kDHoQ/company-logo_200_200/0/1519865605418?e=2147483647&v=beta&t=dE29KDkLy-qxYmZ3TVE95zPf8_PeoMr7YJBQehJbFg8";4;24;28))' ]])
        add("N1", [[ '=HYPERLINK("https://docs.google.com/spreadsheets/d/1kJmtsWxoMtBKeMeO0Aex4IrIULRMeyf6yl3UgqatNGs/edit#gid=1276994968";IMAGE("https://cdn-icons-png.flaticon.com/512/3767/3767084.png";4;23;23))' ]])
        add("N3", [[ '=HYPERLINK("https://webdrive.almg.gov.br/index.php/login";IMAGE("https://upload.wikimedia.org/wikipedia/en/6/61/WebDrive.png";4;22;22))' ]])
        add("O1", [[ '=HYPERLINK("https://www.youtube.com/c/assembleiamg";IMAGE("https://upload.wikimedia.org/wikipedia/commons/thumb/0/09/YouTube_full-color_icon_%282017%29.svg/960px-YouTube_full-color_icon_%282017%29.svg.png";4;20;25))' ]])
        add("O3", [[ '=HYPERLINK("https://atom.almg.gov.br/index.php/";IMAGE("https://atom.almg.gov.br/favicon.ico";4;22;22))' ]])
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
        add("B6", [[f'=TEXT(DATEVALUE("{diario}");"dd/mm/yyyy")']])
        add("C6", [["DIÁRIO DO EXECUTIVO"]])
        add("B7", [["-"]])
        add("E8:G8", [[dmenos2]])

        add(f"A6:A{footer_start - 1}", [['''=IFS(

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
    ))''']] * ((footer_start - 1) - 5))

        add(f"P6:P{footer_start - 1}", [['''=IFS(

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
    $A$683=TRUE;HYPERLINK(X6;IMAGE("https://www.almg.gov.br/favicon.ico";4;17;17)))))))''']] * ((footer_start - 1) - 5))

        add(f"Q6:Q{footer_start - 1}", [['''=IFS(

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

    ))''']] * ((footer_start - 1) - 5))

        add(f"R6:R{footer_start - 1}", [['''=IFS(

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

        add(f"S6:S{footer_start - 1}", [['''=IFS(

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
    INDIRECT("C"&ROW())="OFÍCIO - PEDIDO DE INFORMAÇÃO";HYPERLINK("http://welder.eci.ufmg.br/wp-content/uploads/2024/03/MANUAL-MATE-2024.pdf#page=331";"PÁG 331");
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
            {"range": f"'{tab_name}'!W3", "values": [['=IFERROR(QUERY(C6:G13;"SELECT E WHERE C MATCHES \'.*DIÁRIO DO LEGISLATIVO - EDIÇÃO EXTRA.*\'";0);"SEM EXTRA")']]},
            {"range": f"'{tab_name}'!W4", "values": [['=IFERROR(TEXT(QUERY(B6:G33;"SELECT B WHERE C MATCHES \'REQUERIMENTOS DE COMISSÃO\'";0);"\'dd mm yyyy\'");"")']]},
            {"range": f"'{tab_name}'!X3", "values": [['=IFERROR(TEXT(QUERY(B6:G33;"SELECT B WHERE C MATCHES \'REQUERIMENTOS DE COMISSÃO\'";0);"\'d m yyyy\'");"")']]},
            {"range": f"'{tab_name}'!X4", "values": [['=IFERROR(TEXT(QUERY(B6:G33;"SELECT B WHERE C MATCHES \'REQUERIMENTOS DE COMISSÃO\'";0);"dd/MM/yyyy");"")']]},
            {"range": f"'{tab_name}'!Y2", "values": [["REUNIÃO"]]},
            {"range": f"'{tab_name}'!Y3", "values": [["EXTRA"]]},
            {"range": f"'{tab_name}'!Y4", "values": [["RQC"]]},
        ]

        # EXECUTA O BLOCO PRINCIPAL
        body = {"valueInputOption": "USER_ENTERED", "data": data}
        _with_backoff(sh.values_batch_update, body)

    # ====================================================================================================================================================================================================
    # ============================================================================================= TÍTULOS ==============================================================================================
    # ====================================================================================================================================================================================================
        data2 = []
        data2.append({"range": f"'{tab_name}'!B8:C8", "values": [[diario, "DIÁRIO DO LEGISLATIVO"]]})

        if itens:
            data2.append({"range": f"'{tab_name}'!B9:C{9 + len(itens) - 1}", "values": [[a, b] for a, b in itens]})

        data2.append({
            "range": f"'{tab_name}'!B{start_extra_row}:C{start_extra_row + len(extras_out) - 1}",
            "values": extras_out
        })

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
            data2.append({
            "range": f"'{tab_name}'!E{r}:I{r}",
            "values": [["-","-","-","-","-"]]})

        # acha linha do DROPDOWN_2 (para setar D com "-")
        dd2_row = next(
            (start_extra_row + i for i, (_b, c) in enumerate(extras) if c == "DROPDOWN_2"),
            None
        )

        if dd2_row is not None:
            data2.append({
                "range": f"'{tab_name}'!D{dd2_row}",
                "values": [["-"]]
            })

        # IMPLANTAÇÃO DE TEXTOS (mantém)
        impl_row = next(
            (start_extra_row + i for i, (_b, c) in enumerate(extras)
            if isinstance(c, str) and "IMPLANTAÇÃO DE TEXTOS" in c),
            None
        )

        if impl_row is not None:
            data2.append({"range": f"'{tab_name}'!E{impl_row}", "values": [["..."]]})

            # linha filha (logo abaixo)
            data2.append({
                "range": f"'{tab_name}'!E{impl_row + 1}:I{impl_row + 1}",
                "values": [["?","?","?","-",False]]
            })

            # linha do título
            data2.append({
                "range": f"'{tab_name}'!E{impl_row}:G{impl_row}",
                "values": [["TEXTOS", "EMENDAS", "PARECERES"]]
            })

            reqs.append({
                "setDataValidation": {"range": {
                    "sheetId": sheet_id,
                    "startRowIndex": impl_row,
                    "endRowIndex": impl_row + 1,
                    "startColumnIndex": 8,  # coluna I (0-based)
                    "endColumnIndex": 9
                },
                "rule": {"condition": {"type": "BOOLEAN"},
                        "strict": True}}
            })

            reqs.append(req_font(sheet_id, f"C{impl_row + 1}:I{impl_row + 1}", fg_hex="#CC0000"))

        body2 = {"valueInputOption": "USER_ENTERED", "data": data2}
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
            and "IMPLANTAÇÃO DE TEXTOS" not in (row[1] if len(row) > 1 else ""))]

        data_extra_E = [
            {"range": f"E{r}", "values": [[FORMULAS_E[i]]]}
            for i, r in enumerate(extra_formula_rows[:len(FORMULAS_E)])]

    # ====================================================================================================================================================================================================
    # ============================================================================================== CALL ================================================================================================
    # ====================================================================================================================================================================================================

        _with_backoff(ws.batch_update, data_extra_E, value_input_option="USER_ENTERED")

        # --- SANITIZAÇÃO FINAL: remove mergeCells com intervalo vazio ---
        reqs_ok = []
        for i, r in enumerate(reqs):
            # tenta extrair um range padrão, se existir
            rng = None
            for k in ("mergeCells", "updateBorders", "setDataValidation"):
                if k in r and "range" in r[k]:
                    rng = r[k]["range"]
                    break

            if rng is not None:
                sr = rng.get("startRowIndex"); er = rng.get("endRowIndex")
                sc = rng.get("startColumnIndex"); ec = rng.get("endColumnIndex")

                if sr is None or er is None or sc is None or ec is None:
                    print(f"[req {i}] range incompleto -> REMOVIDO: {rng}")
                    continue

                if er <= sr or ec <= sc:
                    print(f"[req {i}] inválido R{sr}:{er} C{sc}:{ec} -> REMOVIDO")
                    continue

            reqs_ok.append(r)

        reqs = reqs_ok

        # --- AJUSTE DE GRID: garante que a aba tem linhas/colunas suficientes para os ranges dos requests ---
        max_er = 0
        max_ec = 0
        for r in reqs:
            rng = None
            for k in ("mergeCells", "updateBorders", "setDataValidation", "updateCells", "repeatCell", "addConditionalFormatRule"):
                if k in r:
                    if "range" in r[k]:
                        rng = r[k]["range"]
                        break
                    # conditional format usa "ranges": [ {range}, ... ]
                    if k == "addConditionalFormatRule":
                        rr = r[k].get("rule", {}).get("ranges", [])
                        if rr:
                            rng = rr[0]
                            break

            if rng is None:
                continue

            er = rng.get("endRowIndex")
            ec = rng.get("endColumnIndex")
            if isinstance(er, int) and er > max_er:
                max_er = er
            if isinstance(ec, int) and ec > max_ec:
                max_ec = ec

        # endRowIndex/endColumnIndex são EXCLUSIVOS (0-based)
        need_rows = max_er
        need_cols = max(ws  .col_count, max_ec)

        if need_rows > ws.row_count  or need_cols > ws.col_count:
            ws.resize(rows=need_rows, cols=need_cols)

        _with_backoff(sh.batch_update, body={"requests": reqs})

        return sh.url, ws.title


    SPREADSHEET = "https://docs.google.com/spreadsheets/d/1QUpyjHetLqLcr4LrgQqTnCXPZZfEyPkSQb-ld2RxW1k/edit"

    # >>> diario_key PRECISA SER YYYYMMDD (é isso que upsert_tab_diario faz strptime("%Y%m%d"))
    # >>> quando a entrada foi DATA, você já tem aba_yyyymmdd (dia útil de trabalho)
    if not aba_yyyymmdd and entrada and "L20" in entrada:
        import re
        m = re.search(r"L(\d{8})\.pdf", entrada)
        if m:
            yyyymmdd = m.group(1)
            aba_yyyymmdd = proximo_dia_util(yyyymmdd)
    diario_key = aba_yyyymmdd if aba_yyyymmdd else datetime.now(TZ_BR).strftime("%Y%m%d")

    url, aba = upsert_tab_diario(
        spreadsheet_url_or_id=(spreadsheet_url_or_id or SPREADSHEET),
        diario_key=diario_key,
        itens=itens,
        clear_first=False,
        default_col_width_px=COL_DEFAULT,
        col_width_overrides=COL_OVERRIDES
    )

    print("Planilha atualizada:", url)
    print("Aba:", aba)
    
    return url, aba

if __name__ == "__main__":
    main()
