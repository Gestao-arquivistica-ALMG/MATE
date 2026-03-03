import re
from typing import Optional
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

import sys
import subprocess
import os
from pathlib import Path

# força Playwright a usar diretório gravável
pw_path = Path("/tmp/playwright")
pw_path.mkdir(parents=True, exist_ok=True)
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(pw_path)

def _ensure_chromium_installed() -> None:
    """
    No Streamlit Cloud, o postBuild pode falhar/ser ignorado por cache.
    Este fallback instala o Chromium em runtime (uma vez) quando necessário.
    """
    marker = Path(".cache") / "playwright_chromium_ok"
    marker.parent.mkdir(parents=True, exist_ok=True)

    if marker.exists():
        return

    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=True,
            capture_output=True,
            text=True,
            env=os.environ.copy(),
        )
    except subprocess.CalledProcessError as e:
        raise RuntimeError(
            f"playwright install chromium falhou:\nSTDOUT:\n{e.stdout}\n\nSTDERR:\n{e.stderr}"
        ) from e

    marker.write_text("ok", encoding="utf-8")

def _sanitize_filename(name: str) -> str:
    name = name.strip()
    name = re.sub(r"[^\w\-. ()áàâãéèêíïóôõöúçÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇ]", "_", name)
    name = re.sub(r"\s+", " ", name)
    return name[:180] if len(name) > 180 else name


def download_diario_executivo(
    *,
    data_publicacao_yyyy_mm_dd: str,
    out_dir: str = "downloads",
    headless: bool = True,
    timeout_ms: int = 60_000,
) -> Path:
    """
    Abre o Jornal Minas Gerais na data escolhida e baixa o PDF do Diário do Executivo via botão do viewer.
    Retorna o caminho do arquivo salvo.
    """
    out_path = Path(out_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    # Página do "Edição do dia" (o site controla a data via UI; vamos navegar e operar no viewer).
    base_url = "https://www.jornalminasgerais.mg.gov.br/edicao-do-dia"

    with sync_playwright() as p:
        try:
            _ensure_chromium_installed()
            browser = p.chromium.launch(
                headless=headless,
                args=["--no-sandbox", "--disable-dev-shm-usage"],
            )
        except Exception:
            # tenta instalar de novo (caso o marker exista mas a pasta sumiu por cache)
            if Path(".playwright_chromium_ok").exists():
                Path(".playwright_chromium_ok").unlink()
            _ensure_chromium_installed()
            browser = p.chromium.launch(
                headless=headless,
                args=["--no-sandbox", "--disable-dev-shm-usage"],
            )
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        page.goto(base_url, wait_until="domcontentloaded", timeout=timeout_ms)

        # IMPORTANTE:
        # A seleção de data no site é via componente; o caminho mais robusto é:
        # 1) abrir o seletor,
        # 2) digitar ou navegar até a data,
        # 3) confirmar.
        #
        # Como o HTML pode variar, aqui deixo duas estratégias:
        # - tentativa 1: usar o parâmetro no endpoint da API para forçar carregamento (nem sempre funciona)
        # - tentativa 2: operar no UI por seletores genéricos (pode precisar ajuste fino).

        # ---- Tentativa 1: Se o site aceitar querystring (às vezes não) ----
        # page.goto(f"{base_url}?data={data_publicacao_yyyy_mm_dd}", wait_until="networkidle", timeout=timeout_ms)

        # ---- Tentativa 2: Interação na UI (genérica) ----
        # Clique na "Edição do dia" / calendário (área à esquerda)
        # Ajuste os seletores caso mude no site.
        try:
            # abre dropdown do calendário (seta ao lado da data)
            page.locator("div:has-text('Edição do dia')").first.wait_for(timeout=timeout_ms)
        except PWTimeout:
            pass

        # DICA: se você já abre a edição pela data correta manualmente, esse bloco pode ser removido.
        # Aqui vamos tentar setar a data pelo campo/controle se existir.
        # Se falhar, a página já estará na data padrão e você pode navegar manualmente.

        # Espera o viewer do PDF aparecer (barra de ferramentas do PDF.js)
        page.locator("input[aria-label='Page']").first.wait_for(timeout=timeout_ms)

        # Agora: clicar no botão de download do PDF.js
        # No PDF.js o botão costuma ter id #download ou title "Download" / "Baixar"
        download_button = page.locator("#download").first
        if download_button.count() == 0:
            download_button = page.locator("button[title*='Download'], button[title*='Baixar'], a[title*='Download'], a[title*='Baixar']").first

        if download_button.count() == 0:
            raise RuntimeError("Não encontrei o botão de download no viewer (PDF.js).")

        try:
            with page.expect_download(timeout=timeout_ms) as dl_info:
                download_button.click()
            download = dl_info.value
        except PWTimeout:
            # Em alguns casos, o site abre o diálogo de download por blob/data: sem evento de download.
            # Ainda assim, frequentemente o Playwright captura. Se não capturar, precisamos interceptar a chamada XHR/base64.
            raise RuntimeError("Cliquei no download, mas não capturou evento de download. Pode ser fluxo via data: URL.")

        suggested = download.suggested_filename or f"Diario_do_Executivo_{data_publicacao_yyyy_mm_dd}.pdf"
        suggested = _sanitize_filename(suggested)
        final_file = out_path / suggested

        download.save_as(final_file.as_posix())

        context.close()
        browser.close()

        return final_file