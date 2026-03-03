# playwright_fetch_jmg.py
# Download do Diário do Executivo (Jornal Minas Gerais) via API (sem Playwright)
#
# Fonte:
#   GET /api/v1/Jornal/ObterEdicaoPorDataPublicacao?dataPublicacao=YYYY-MM-DD
# Retorno:
#   dados.arquivoCadernoPrincipal.arquivo  (base64, normalmente CMS/PKCS#7 contendo PDF embutido)
#
# Estratégia:
#   - baixa JSON
#   - base64 decode
#   - extrai bytes do PDF procurando por "%PDF-" ... "%%EOF"
#   - salva em downloads/

from __future__ import annotations

import base64
import re
from pathlib import Path
from typing import Callable, Optional

import requests


JMG_BASE = "https://www.jornalminasgerais.mg.gov.br"


def _sanitize_filename(name: str) -> str:
    name = (name or "").strip() or "arquivo.pdf"
    name = re.sub(r"[^\w\-. ()áàâãéèêíïóôõöúçÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇ]", "_", name)
    name = re.sub(r"\s+", " ", name)
    if not name.lower().endswith(".pdf"):
        name += ".pdf"
    return name[:180] if len(name) > 180 else name


def _http_get(url: str, timeout_s: float) -> requests.Response:
    headers = {"User-Agent": "Mozilla/5.0 (compatible; MATE/1.0)"}
    return requests.get(url, headers=headers, timeout=timeout_s)


def _extract_pdf_from_container(blob: bytes) -> bytes:
    """
    O campo 'arquivo' frequentemente vem como CMS/PKCS#7 (DER) e contém o PDF embutido.
    Estratégia robusta: procurar assinatura do PDF (%PDF-) e recortar até o último %%EOF.
    """
    pdf_magic = b"%PDF-"
    eof_magic = b"%%EOF"

    i = blob.find(pdf_magic)
    if i < 0:
        # às vezes vem zip/gzip; mas como você já viu virar data:application/pdf no front,
        # o mais provável é estar embutido. Se não achou, falha com diagnóstico.
        head = blob[:32]
        raise RuntimeError(
            "Não encontrei '%PDF-' dentro do blob decodificado. "
            "O conteúdo pode ter mudado (ou estar em outro campo), "
            f"ou precisar de outra etapa. Head bytes={head!r}"
        )

    # pega do %PDF- até o último %%EOF
    j = blob.rfind(eof_magic)
    if j < 0:
        # sem EOF: ainda assim salva do %PDF- até o fim
        return blob[i:]

    return blob[i : j + len(eof_magic)]


def download_diario_executivo(
    *,
    data_publicacao_yyyy_mm_dd: str,
    out_dir: str = "downloads",
    headless: bool = True,  # compat (não usado)
    timeout_ms: int = 60_000,
    log: Optional[Callable[[str], None]] = None,
) -> Path:
    """
    Baixa o PDF do Diário do Executivo da data informada, via API do Jornal Minas Gerais.
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

    api_url = (
        f"{JMG_BASE}/api/v1/Jornal/ObterEdicaoPorDataPublicacao"
        f"?dataPublicacao={data_publicacao_yyyy_mm_dd}"
    )

    _log("[1/4] Consultando API do Jornal...")
    resp = _http_get(api_url, timeout_s=timeout_s)
    resp.raise_for_status()

    data = resp.json()

    _log("[2/4] Lendo campo arquivoCadernoPrincipal.arquivo ...")
    try:
        b64 = data["dados"]["arquivoCadernoPrincipal"]["arquivo"]
    except Exception as e:
        raise RuntimeError(
            "Estrutura inesperada no JSON: não achei dados.arquivoCadernoPrincipal.arquivo"
        ) from e

    if not isinstance(b64, str) or not b64.strip():
        raise RuntimeError("Campo 'arquivo' veio vazio.")

    _log("[3/4] Decodificando e extraindo PDF...")
    try:
        raw = base64.b64decode(b64, validate=False)
    except Exception as e:
        raise RuntimeError("Falha ao decodificar base64 do campo 'arquivo'.") from e

    pdf_bytes = _extract_pdf_from_container(raw)

    filename = _sanitize_filename(f"Diario_do_Executivo_{data_publicacao_yyyy_mm_dd}.pdf")
    final_file = out_path / filename
    final_file.write_bytes(pdf_bytes)

    _log(f"[4/4] OK: {final_file.name}")
    return final_file