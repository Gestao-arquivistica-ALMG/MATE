# playwright_fetch_jmg.py
#
# Download do "Diário do Executivo" (Jornal Minas Gerais) SEM Playwright.
# Estratégia:
#  1) Usa a mesma data/parametro "dados=" (JSON URL-encoded) que você usa no Google Sheets.
#  2) Consulta a API observada no DevTools:
#       /api/v1/Jornal/ObterEdicaoPorDataPublicacao?dataPublicacao=YYYY-MM-DD
#  3) Tenta extrair um link/rota de PDF do JSON (busca recursiva por strings que pareçam URL/rota).
#  4) Se não achar, tenta uma lista de endpoints de download usando "dados=".
#  5) Valida se é PDF (Content-Type e/ou assinatura %PDF-).
#
# Observação:
#  - O portal pode mudar os endpoints. Este módulo foi feito para “descobrir” o PDF com heurísticas
#    e te devolver logs claros do que tentou.

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Callable, Optional
from urllib.parse import quote, urljoin

import requests


JMG_BASE = "https://www.jornalminasgerais.mg.gov.br"


def _sanitize_filename(name: str) -> str:
    name = (name or "").strip() or "arquivo.pdf"
    name = re.sub(r"[^\w\-. ()áàâãéèêíïóôõöúçÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇ]", "_", name)
    name = re.sub(r"\s+", " ", name)
    if not name.lower().endswith(".pdf"):
        name += ".pdf"
    return name[:180] if len(name) > 180 else name


def _make_dados_param(data_publicacao_yyyy_mm_dd: str) -> str:
    """
    Replica sua fórmula do GS:
      {"dataPublicacaoSelecionada":"YYYY-MM-DDT03:00:00.000Z"}
    URL-encoded via ENCODEURL (aqui: quote).
    """
    payload = {"dataPublicacaoSelecionada": f"{data_publicacao_yyyy_mm_dd}T03:00:00.000Z"}
    # json.dumps com separators compactos deixa mais parecido com o que o front costuma gerar
    s = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return quote(s, safe="")


def _is_pdf_response(resp: requests.Response) -> bool:
    ctype = (resp.headers.get("Content-Type") or "").lower()
    if "application/pdf" in ctype:
        return True
    # fallback: assinatura do PDF
    try:
        head = resp.content[:5]
        return head == b"%PDF-"
    except Exception:
        return False


def _deep_collect_strings(obj: Any, out: list[str]) -> None:
    """
    Coleta recursivamente strings dentro de dict/list.
    """
    if obj is None:
        return
    if isinstance(obj, str):
        s = obj.strip()
        if s:
            out.append(s)
        return
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(k, str) and k.strip():
                out.append(k.strip())
            _deep_collect_strings(v, out)
        return
    if isinstance(obj, (list, tuple)):
        for it in obj:
            _deep_collect_strings(it, out)
        return


def _pick_pdf_like_urls(strings: list[str]) -> list[str]:
    """
    Heurísticas para extrair URLs/rotas candidatas a PDF.
    """
    cand: list[str] = []

    # URLs absolutas ou relativas que contenham 'pdf' (não precisa terminar com .pdf)
    for s in strings:
        sl = s.lower()
        if "pdf" in sl:
            # url absoluta
            if s.startswith("http://") or s.startswith("https://"):
                cand.append(s)
            # rota relativa
            elif s.startswith("/"):
                cand.append(urljoin(JMG_BASE, s))

    # remove duplicadas mantendo ordem
    seen = set()
    out = []
    for u in cand:
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


def _http_get(url: str, timeout_s: float) -> requests.Response:
    # user-agent simples ajuda em alguns WAFs
    headers = {"User-Agent": "Mozilla/5.0 (compatible; MATE/1.0; +https://www.almg.gov.br/)"}
    return requests.get(url, headers=headers, timeout=timeout_s)


def download_diario_executivo(
    *,
    data_publicacao_yyyy_mm_dd: str,
    out_dir: str = "downloads",
    headless: bool = True,  # mantido só para compatibilidade com a UI (não usado)
    timeout_ms: int = 60_000,
    log: Optional[Callable[[str], None]] = None,
) -> Path:
    """
    Baixa o PDF da edição do dia (Diário do Executivo / Jornal Minas Gerais) para a data informada.
    Retorna o caminho do arquivo salvo.
    """
    def _log(msg: str) -> None:
        if log:
            try:
                log(msg)
            except Exception:
                pass

    timeout_s = max(5.0, float(timeout_ms) / 1000.0)

    out_path = Path(out_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    dados = _make_dados_param(data_publicacao_yyyy_mm_dd)

    # 1) API observada no DevTools
    api_url = f"{JMG_BASE}/api/v1/Jornal/ObterEdicaoPorDataPublicacao?dataPublicacao={data_publicacao_yyyy_mm_dd}"

    _log("[1/4] Consultando API (ObterEdicaoPorDataPublicacao)...")
    resp = _http_get(api_url, timeout_s=timeout_s)
    resp.raise_for_status()

    try:
        payload = resp.json()
    except Exception as e:
        raise RuntimeError("API não retornou JSON válido.") from e

    # 2) tentar achar URLs/rotas com 'pdf' dentro do JSON
    _log("[2/4] Procurando link/rota do PDF no JSON...")
    strings: list[str] = []
    _deep_collect_strings(payload, strings)
    pdf_candidates = _pick_pdf_like_urls(strings)

    tried: list[str] = []

    # tenta candidatos encontrados no JSON
    for u in pdf_candidates:
        tried.append(u)
        try:
            _log(f"[3/4] Tentando candidato do JSON: {u}")
            r = _http_get(u, timeout_s=timeout_s)
            if r.status_code == 200 and _is_pdf_response(r):
                fname = _sanitize_filename(f"Diario_do_Executivo_{data_publicacao_yyyy_mm_dd}.pdf")
                final_file = out_path / fname
                final_file.write_bytes(r.content)
                _log(f"[4/4] OK (PDF via JSON): {final_file.name}")
                return final_file
        except Exception:
            continue

    # 3) fallback: tentar endpoints usando "dados=" (igual seu link do GS)
    _log("[3/4] Fallback: tentando endpoints com parametro dados= ...")

    # Página do viewer (às vezes já entrega PDF direto; na maioria não, mas deixamos)
    fallback_urls = [
        f"{JMG_BASE}/edicao-do-dia?dados={dados}",

        # Endpoints de download comuns (variam por implementação)
        f"{JMG_BASE}/api/v1/Jornal/DownloadEdicao?dados={dados}",
        f"{JMG_BASE}/api/v1/Jornal/BaixarEdicao?dados={dados}",
        f"{JMG_BASE}/api/v1/Jornal/DownloadEdicaoPorDataPublicacao?dados={dados}",
        f"{JMG_BASE}/api/v1/Jornal/DownloadPdf?dados={dados}",
        f"{JMG_BASE}/api/v1/Jornal/BaixarPdf?dados={dados}",
        f"{JMG_BASE}/api/v1/Jornal/ObterPdf?dados={dados}",
    ]

    for u in fallback_urls:
        tried.append(u)
        try:
            _log(f"[3/4] Tentando: {u}")
            r = _http_get(u, timeout_s=timeout_s)

            # se a página HTML voltar, não é o PDF; seguimos
            if r.status_code != 200:
                continue

            if _is_pdf_response(r):
                fname = _sanitize_filename(f"Diario_do_Executivo_{data_publicacao_yyyy_mm_dd}.pdf")
                final_file = out_path / fname
                final_file.write_bytes(r.content)
                _log(f"[4/4] OK (PDF via endpoint): {final_file.name}")
                return final_file
        except Exception:
            continue

    # 4) se não achou, devolve erro com diagnóstico
    #    (isso é importante para você me mandar e eu cravar o endpoint real)
    tried_txt = "\n".join(tried[-25:])  # limita
    raise RuntimeError(
        "Não consegui obter o PDF por URL direta.\n"
        "Provavelmente o PDF é servido por outro endpoint (ex.: token/hash/id) que vem no JSON.\n\n"
        "Últimas URLs tentadas:\n"
        f"{tried_txt}\n\n"
        "Sugestão: no DevTools (Network), filtre por 'pdf' ou 'application/pdf' e copie o Request URL."
    )