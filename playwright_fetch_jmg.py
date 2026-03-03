# playwright_fetch_jmg.py
#
# Downloader robusto do "Diário do Executivo" (Jornal Minas Gerais) via Playwright,
# preparado para Streamlit Cloud:
# - instala Chromium em runtime em diretório gravável (/tmp)
# - evita reinstalar na mesma sessão (marker em /tmp)
# - flags cloud-friendly no Chromium
# - tenta fechar banner de cookies
# - tenta capturar download normal (expect_download)
# - fallback: captura href data:/blob: gerado pelo PDF.js (hook em <a>.click)
# - instrumentação de etapas via callback log(msg) + print(flush=True)

import os
import re
import sys
import time
import subprocess
from pathlib import Path
from typing import Optional, Callable

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout


# ======================================================================================
# Playwright: diretório gravável para browsers no Streamlit Cloud
# ======================================================================================
PW_DIR = Path("/tmp/playwright")
PW_DIR.mkdir(parents=True, exist_ok=True)
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(PW_DIR)

MARKER = PW_DIR / "chromium_ok"


def _ensure_chromium_installed(timeout_s: int = 600) -> None:
    """
    Garante que o Chromium do Playwright está instalado em diretório gravável.
    No Streamlit Cloud, postBuild pode não rodar / cache pode variar.
    """
    if MARKER.exists():
        return

    print("[playwright] Installing chromium (runtime fallback)...", flush=True)
    try:
        r = subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=True,
            capture_output=True,
            text=True,
            env=os.environ.copy(),
            timeout=timeout_s,
        )
        # log curto (evita spam)
        if r.stdout:
            print("[playwright] install stdout:", r.stdout[:2000], flush=True)
        if r.stderr:
            print("[playwright] install stderr:", r.stderr[:2000], flush=True)
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(f"playwright install chromium TIMEOUT ({timeout_s}s)") from e
    except subprocess.CalledProcessError as e:
        raise RuntimeError(
            f"playwright install chromium falhou:\nSTDOUT:\n{e.stdout}\n\nSTDERR:\n{e.stderr}"
        ) from e

    MARKER.write_text("ok", encoding="utf-8")
    print("[playwright] chromium ok", flush=True)


def _sanitize_filename(name: str) -> str:
    name = name.strip()
    name = re.sub(r"[^\w\-. ()áàâãéèêíïóôõöúçÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇ]", "_", name)
    name = re.sub(r"\s+", " ", name)
    return name[:180] if len(name) > 180 else name


def _install_download_hooks(page) -> None:
    """
    Instala hooks no DOM para capturar o download gerado pelo PDF.js mesmo quando
    não há evento "download" (ex.: data:application/pdf ou blob:).
    """
    page.add_init_script(
        """
        (() => {
          try {
            window.__pw_last_pdf = null;

            const origClick = HTMLAnchorElement.prototype.click;
            HTMLAnchorElement.prototype.click = function(...args) {
              try {
                const href = (this && this.href) ? String(this.href) : "";
                const dl = (this && this.download) ? String(this.download) : "";
                if (href.startsWith("data:application/pdf") || href.startsWith("blob:")) {
                  window.__pw_last_pdf = { href, download: dl, ts: Date.now() };
                }
              } catch (e) {}
              return origClick.apply(this, args);
            };
          } catch (e) {}
        })();
        """
    )


def _try_close_cookies(page) -> None:
    # O portal costuma ter "Ok, entendi"
    for txt in ("Ok, entendi", "OK, entendi", "Aceitar", "Aceito", "Concordo"):
        try:
            page.locator(f"text={txt}").first.click(timeout=1500)
            print("[playwright] cookie banner closed", flush=True)
            return
        except Exception:
            pass


def _goto_best_effort(page, data_publicacao_yyyy_mm_dd: str, timeout_ms: int) -> None:
    """
    O site pode aceitar (ou não) querystring; tentamos algumas.
    Se nenhuma funcionar, fica no /edicao-do-dia padrão.
    """
    base = "https://www.jornalminasgerais.mg.gov.br/edicao-do-dia"
    candidates = [
        base,
        f"{base}?dataJornal={data_publicacao_yyyy_mm_dd}",
        f"{base}?dataPublicacao={data_publicacao_yyyy_mm_dd}",
        f"https://www.jornalminasgerais.mg.gov.br/?dataJornal={data_publicacao_yyyy_mm_dd}",
    ]

    last_err: Optional[Exception] = None
    for url in candidates:
        try:
            print("[playwright] goto:", url, flush=True)
            page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
            _try_close_cookies(page)
            return
        except Exception as e:
            last_err = e
            continue

    if last_err is not None:
        print("[playwright] goto candidates failed, fallback to base:", repr(last_err), flush=True)

    page.goto(base, wait_until="domcontentloaded", timeout=timeout_ms)
    _try_close_cookies(page)


def _wait_viewer_ready(page, timeout_ms: int) -> None:
    """
    Espera o viewer do PDF.js aparecer.
    """
    print("[playwright] waiting for pdf viewer toolbar...", flush=True)

    # Em alguns layouts, pode ser iframe; tentamos página principal primeiro
    try:
        page.locator("input[aria-label='Page']").first.wait_for(timeout=timeout_ms)
        print("[playwright] viewer ready (page)", flush=True)
        return
    except PWTimeout:
        pass

    # fallback: tenta achar em iframes
    try:
        iframe = page.frame_locator("iframe").first
        iframe.locator("input[aria-label='Page']").first.wait_for(timeout=timeout_ms)
        print("[playwright] viewer ready (iframe)", flush=True)
        return
    except Exception as e:
        raise RuntimeError("Não encontrei a barra do viewer do PDF.js (Page input) — possível mudança de layout.") from e


def _find_download_button(page):
    """
    Localiza o botão de download do PDF.js.
    Tenta no documento principal e em iframe.
    """
    # 1) principal
    btn = page.locator("#download").first
    if btn.count() > 0:
        return btn

    btn = page.locator(
        "button[title*='Download'], button[title*='Baixar'], a[title*='Download'], a[title*='Baixar']"
    ).first
    if btn.count() > 0:
        return btn

    # 2) iframe
    iframe = page.frame_locator("iframe").first
    btn = iframe.locator("#download").first
    if btn.count() > 0:
        return btn

    btn = iframe.locator(
        "button[title*='Download'], button[title*='Baixar'], a[title*='Download'], a[title*='Baixar']"
    ).first
    if btn.count() > 0:
        return btn

    return None


def _get_last_pdf_data_url(page, timeout_ms: int) -> Optional[dict]:
    """
    Busca window.__pw_last_pdf (setado pelo hook).
    """
    deadline = time.time() + (timeout_ms / 1000.0)
    while time.time() < deadline:
        try:
            val = page.evaluate("window.__pw_last_pdf")
            if val and isinstance(val, dict) and val.get("href"):
                return val
        except Exception:
            pass
        time.sleep(0.2)
    return None


def _dataurl_to_pdf_bytes(data_url: str) -> bytes:
    """
    data:application/pdf;base64,....
    """
    if not data_url.startswith("data:"):
        raise ValueError("Não é data URL")

    try:
        header, b64 = data_url.split(",", 1)
    except ValueError:
        raise ValueError("data URL inválida")

    if ";base64" not in header:
        raise ValueError("data URL não-base64")

    import base64

    return base64.b64decode(b64.encode("utf-8"))


def _bloburl_to_dataurl_in_page(page, blob_url: str, timeout_ms: int) -> str:
    """
    Converte blob: URL para data URL dentro do contexto do browser.
    """
    return page.evaluate(
        """
        async ({ url, timeoutMs }) => {
          const ctrl = new AbortController();
          const t = setTimeout(() => ctrl.abort(), timeoutMs);
          try {
            const resp = await fetch(url, { signal: ctrl.signal });
            const blob = await resp.blob();
            const dataUrl = await new Promise((resolve, reject) => {
              const reader = new FileReader();
              reader.onload = () => resolve(reader.result);
              reader.onerror = reject;
              reader.readAsDataURL(blob);
            });
            return dataUrl;
          } finally {
            clearTimeout(t);
          }
        }
        """,
        {"url": blob_url, "timeoutMs": timeout_ms},
    )


def download_diario_executivo(
    *,
    data_publicacao_yyyy_mm_dd: str,
    out_dir: str = "downloads",
    headless: bool = True,
    timeout_ms: int = 90_000,
    log: Optional[Callable[[str], None]] = None,
) -> Path:
    """
    Abre o Jornal Minas Gerais e baixa o PDF do Diário do Executivo via botão do viewer.
    Retorna o caminho do arquivo salvo.

    Instrumentação:
      - log(msg): callback opcional (ex.: para exibir etapas no Streamlit)
      - também faz print(flush=True) para logs do Cloud
    """

    def _log(msg: str) -> None:
        if log:
            try:
                log(msg)
            except Exception:
                pass
        print(msg, flush=True)

    out_path = Path(out_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        _log("[1/8] ensure chromium")
        _ensure_chromium_installed()

        _log("[2/8] launching browser")
        browser = p.chromium.launch(
            headless=headless,
            args=["--no-sandbox", "--disable-dev-shm-usage"],
        )

        _log("[3/8] new context/page")
        context = browser.new_context(accept_downloads=True)
        context.set_default_timeout(timeout_ms)
        context.set_default_navigation_timeout(timeout_ms)
        page = context.new_page()

        _log("[4/8] install hooks + goto")
        _install_download_hooks(page)
        _goto_best_effort(page, data_publicacao_yyyy_mm_dd, timeout_ms)

        _log("[5/8] wait viewer")
        _wait_viewer_ready(page, timeout_ms)

        _log("[6/8] find download button")
        download_button = _find_download_button(page)
        if download_button is None:
            context.close()
            browser.close()
            raise RuntimeError("Não encontrei o botão de download no viewer (PDF.js).")

        _log("[7/8] click download (expect_download)")
        try:
            with page.expect_download(timeout=timeout_ms) as dl_info:
                download_button.click()
            download = dl_info.value

            suggested = download.suggested_filename or f"Diario_do_Executivo_{data_publicacao_yyyy_mm_dd}.pdf"
            suggested = _sanitize_filename(suggested)
            final_file = out_path / suggested

            download.save_as(final_file.as_posix())
            _log(f"[OK] download event: {final_file.name}")

            context.close()
            browser.close()
            return final_file

        except PWTimeout:
            _log("[7/8] sem evento de download; tentando fallback data/blob...")

        _log("[8/8] fallback data/blob (hook)")
        # clica de novo para garantir que o hook capture href
        try:
            download_button.click()
        except Exception:
            pass

        last = _get_last_pdf_data_url(page, timeout_ms)
        if not last:
            context.close()
            browser.close()
            raise RuntimeError(
                "Cliquei no download, mas não capturou evento nem data/blob. "
                "Pode ter mudado o fluxo do site."
            )

        href = str(last.get("href", ""))
        dlname = str(last.get("download", "")).strip()

        suggested = dlname or f"Diario_do_Executivo_{data_publicacao_yyyy_mm_dd}.pdf"
        suggested = _sanitize_filename(suggested)
        final_file = out_path / suggested

        if href.startswith("data:application/pdf"):
            pdf_bytes = _dataurl_to_pdf_bytes(href)
            final_file.write_bytes(pdf_bytes)
            _log(f"[OK] salvo de data:url: {final_file.name}")

            context.close()
            browser.close()
            return final_file

        if href.startswith("blob:"):
            data_url = _bloburl_to_dataurl_in_page(page, href, timeout_ms)
            if not isinstance(data_url, str) or not data_url.startswith("data:application/pdf"):
                context.close()
                browser.close()
                raise RuntimeError("Falha ao converter blob: em data URL PDF.")

            pdf_bytes = _dataurl_to_pdf_bytes(data_url)
            final_file.write_bytes(pdf_bytes)
            _log(f"[OK] salvo de blob:url: {final_file.name}")

            context.close()
            browser.close()
            return final_file

        context.close()
        browser.close()
        raise RuntimeError(f"Fallback capturou href inesperado: {href[:200]}")