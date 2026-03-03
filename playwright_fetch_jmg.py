# playwright_fetch_jmg.py
# Versão SEM Playwright — download direto via API

import requests
from pathlib import Path
from typing import Optional
import re


def _sanitize_filename(name: str) -> str:
    name = name.strip()
    name = re.sub(r"[^\w\-. ()áàâãéèêíïóôõöúçÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇ]", "_", name)
    name = re.sub(r"\s+", " ", name)
    return name[:180] if len(name) > 180 else name


def download_diario_executivo(
    *,
    data_publicacao_yyyy_mm_dd: str,
    out_dir: str = "downloads",
    headless: bool = True,          # mantido só para compatibilidade
    timeout_ms: int = 60_000,
    log: Optional[callable] = None,
) -> Path:

    def _log(msg: str):
        if log:
            log(msg)

    out_path = Path(out_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    _log("[1/4] Consultando API do Jornal...")

    api_url = (
        "https://www.jornalminasgerais.mg.gov.br"
        "/api/v1/Jornal/ObterEdicaoPorDataPublicacao"
        f"?dataPublicacao={data_publicacao_yyyy_mm_dd}"
    )

    resp = requests.get(api_url, timeout=timeout_ms / 1000)
    resp.raise_for_status()

    data = resp.json()

    if not data or "dados" not in data:
        raise RuntimeError("API retornou resposta inesperada.")

    _log("[2/4] Extraindo identificador da edição...")

    # A API retorna estrutura com metadados da edição
    # Precisamos descobrir onde está o ID/URL do PDF
    # Geralmente vem em data['dados']['idEdicao'] ou similar

    edicao = data.get("dados")

    # Tentativas comuns
    pdf_url = None

    # Caso 1: já venha link direto
    if isinstance(edicao, dict):
        for k in edicao.keys():
            if "pdf" in k.lower():
                pdf_url = edicao[k]
                break

    if not pdf_url:
        # fallback padrão do portal
        pdf_url = (
            "https://www.jornalminasgerais.mg.gov.br"
            f"/api/v1/Jornal/DownloadEdicao?dataPublicacao={data_publicacao_yyyy_mm_dd}"
        )

    _log("[3/4] Baixando PDF...")

    pdf_resp = requests.get(pdf_url, timeout=timeout_ms / 1000)
    pdf_resp.raise_for_status()

    filename = f"Diario_do_Executivo_{data_publicacao_yyyy_mm_dd}.pdf"
    filename = _sanitize_filename(filename)

    final_file = out_path / filename
    final_file.write_bytes(pdf_resp.content)

    _log(f"[4/4] OK: {final_file.name}")

    return final_file