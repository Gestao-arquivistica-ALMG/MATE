import re
import streamlit as st
import threading
import time
import base64
import streamlit.components.v1 as components
from datetime import datetime
from mate_pipeline import main, normalizar_data
from playwright import fetch_diario_executivo_pdf_bytes

# ================= CONFIG =================
st.set_page_config(
    page_title="MATE - Matérias em Tramitação",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ================= ESTILO =================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700&display=swap');

/* Fundo externo */
html, body, [data-testid="stAppViewContainer"]{
  background:#b30000;
}

/* Container geral branco (fecha tudo) */
.block-container{
  padding-top: 0rem !important;
  max-width:600px;
  margin:90px auto 30px auto;
  padding:22px 22px 26px 22px;
  background:#fff;
  border-radius:18px;
  box-shadow:0 10px 30px rgba(0,0,0,0.25);
}

/* Título e subtítulo */
.title{
  font-family:Montserrat;
  font-size:52px;
  font-weight:700;
  text-align:center;
  color:#111;
  margin:0;
}
.subtitle{
  font-family:Montserrat;
  font-size:14px;
  font-weight:700;
  text-align:center;
  color:#444;
  margin:6px 0 10px 0;
}

/* Card interno (o próprio form) */
div[data-testid="stForm"]{
  background:#fff !important;
  padding:30px !important;
  border-radius:18px !important;
  box-shadow:0 10px 30px rgba(0,0,0,0.25) !important;
  max-width:560px !important;
  margin:0 auto !important;
}

/* Input com fundo claro */
div[data-testid="stTextInput"] > div{
  background:#f3f4f6 !important;
}

/* Conteúdo do form inteiro em uma coluna central (label+input+lista+botões) */
div[data-testid="stForm"] > div{
  max-width:320px;
  margin:0 auto;
}

/* Conteúdo legível: alinhamento à esquerda */
div[data-testid="stForm"] label,
div[data-testid="stForm"] .stCaption,
div[data-testid="stForm"] ul,
div[data-testid="stForm"] li{
  text-align:left !important;
}

div[data-testid="stForm"] div[data-testid="stHorizontalBlock"]:last-of-type{
  max-width:320px !important;
  margin:12px 0 0 40px !important;
  flex-wrap:nowrap !important;
  gap:8px !important;
}

/* 🚀 Gerar (PRIMARY) */
  div[data-testid="stForm"] div[data-testid="stHorizontalBlock"]:last-of-type
  button[kind="primary"]{
    min-width:100px !important;
    padding:10px 14px !important;
  }
}

/* 🧹 Limpar (SECONDARY) */
div[data-testid="stForm"]
div[data-testid="stHorizontalBlock"]:last-of-type
button[kind="secondary"]{
  min-width:40px !important;
  padding:8px 0 !important;
  justify-content: flex-start !important;
}

.small-gap{ margin-top:10px; }

/* ===== BOTÃO MENU OVERLAY (ÚNICO E LIMPO) ===== */

/* botão invisível exatamente em cima do ☰ no cabeçalho */
button#menu_btn{
  position: fixed !important;
  top: 118px !important;                     /* ajuste fino */
  left: 20px !important; /* ajuste fino: 560/2 = 280 */
  width: 45px !important;
  height: 45px !important;
  padding: 0 !important;
  background: transparent !important;
  border: 0 !important;
  box-shadow: none !important;
  color: transparent !important;
  z-index: 10000 !important;
}

/* overlay invisível que fecha ao clicar fora */
button#close_menu_btn{
  position: fixed !important;
  inset: 0 !important;
  padding: 0 !important;
  background: rgba(0,0,0,0.25) !important;
  border: 0 !important;
  box-shadow: none !important;
  color: transparent !important;
  z-index: 9998 !important;
}
button#close_menu_btn *{ display:none !important; }

/* drawer por cima do overlay */
#almg_menu_drawer{
  position: fixed;
  left: 0;
  top: 0;
  bottom: 0;
  width: 260px;
  background: white;
  padding: 20px;
  z-index: 9999;
  box-shadow: 3px 0 12px rgba(0,0,0,0.2);
}

#almg_menu_drawer_hidden{
position: fixed;
left: 0;
top: 0;
bottom: 0;
width:260px;
background:white;
padding:20px;
z-index:9999;
box-shadow:3px 0 12px rgba(0,0,0,0.2);
transform:translateX(-100%);
transition:transform 0.25s ease;
}

/* overlay invisível que fecha o menu ao clicar fora */
button#close_menu_btn{
  position: fixed !important;
  inset: 0 !important;
  width: 100vw !important;
  height: 100vh !important;
  padding: 0 !important;

  background: rgba(0,0,0,0.25) !important;
  border: 0 !important;
  box-shadow: none !important;

  color: transparent !important; /* some o texto */
  z-index: 9998 !important;
}

</style>
""", unsafe_allow_html=True)

# ================= HEADER ALMG =================
if "menu_open" not in st.session_state:
    st.session_state.menu_open = False

# "cabeçalho" feito com layout Streamlit (sem HTML clicável)
c1, c2 = st.columns([2.2, 7.8], gap="small")

with c1:
    # LOGO vira o botão do menu
    if st.button(
        " ",
        key="btn_menu_toggle",
        use_container_width=True,
    ):
        st.session_state.menu_open = not st.session_state.menu_open
        st.rerun()

    # desenha o logo por cima do botão
    st.markdown(
        """
        <div style="
            margin-top:-70px;
            display:flex;
            align-items:left;
            justify-content:flex-start;
            pointer-events:none;
        ">
          <img src="https://www.almg.gov.br/system/modules/br.gov.almg.portal/resources/img/logo/logo.svg"
               style="height:45px;width:100px;">
        </div>
        """,
        unsafe_allow_html=True,
    )

with c2:
    st.markdown(
        """
        <div style="display:flex; align-items:center; justify-content:flex-end; height:45px; gap:12px; font-size:24px;">
          <a href="https://silegis.almg.gov.br/silegismg/login/login.jsp#/processos" target="_blank" style="text-decoration:none;">🔍</a>
          <a href="https://intra.almg.gov.br/" target="_blank" style="text-decoration:none;">👤</a>
        </div>
        """,
        unsafe_allow_html=True,
    )

# style do "container" do header (1 caixa branca como antes)
st.markdown("""
<style>
/* aplica no row acima (Streamlit) */
div[data-testid="stHorizontalBlock"]{
  max-width:100%;
  margin:0 auto 0 auto;
  background:white;
  padding:2px 18px;
  border-radius:12px;
}
button[kind="secondary"][data-testid="baseButton-secondary"]{
  color:#cc0000 !important;
  font-size:26px !important;
  width:200px !important;
  height:100px !important;
  padding:0 !important;
}
</style>
""", unsafe_allow_html=True)

# ================= MENU (OVERLAY NO CORPO) =================
if st.session_state.get("menu_open", False):

    # 2) drawer (menu) por cima do overlay
    st.markdown("""
    <div id="almg_menu_drawer" style="
      display: block !important;
      z-index: 2147483647 !important;
      position: fixed;
      left: 0;
      top: 0;
      bottom: 0;
      width: 260px;
      background: white;
      padding: 20px;
      z-index: 9999;
      box-shadow: 3px 0 12px rgba(0,0,0,0.2);
    ">
    
      <div style="display:flex; align-items:center; justify-content:space-between;">
        <div style="font-family:Montserrat; font-weight:700; font-size:16px;">MENU</div>
      </div>

      <div style="margin-top:14px; display:flex; flex-direction:column; gap:10px; font-family:Montserrat;">
        <a href="https://www.almg.gov.br/" target="_blank" style="text-decoration:none; color:#111;">ALMG</a>
        <a href="https://www.jornalminasgerais.mg.gov.br/?dataJornal=" target="_blank" style="text-decoration:none; color:#111;">Diário do Executivo</a>
        <a href="https://www.almg.gov.br/transparencia/diario-do-legislativo/index.html" target="_blank" style="text-decoration:none; color:#111;">Diário do Legislativo</a>
        <a href="https://www.almg.gov.br/atividade-parlamentar/plenario/agenda/" target="_blank" style="text-decoration:none; color:#111;">Reuniões de Plenário</a>
        <a href="https://www.almg.gov.br/atividade-parlamentar/comissoes/agenda/" target="_blank" style="text-decoration:none; color:#111;">Reuniões de Comissões</a>
        <a href="https://silegis.almg.gov.br/silegismg/login/login.jsp#/processos" target="_blank" style="text-decoration:none; color:#111;">Silegis</a>
        <a href="https://webmail.almg.gov.br/" target="_blank" style="text-decoration:none; color:#111;">Webmail</a>
      </div>

      <div style="margin-top:18px;">
        <button onclick="window.location.reload()" style="
          display:none;
        "></button>
      </div>
    </div>
    """, unsafe_allow_html=True)
            
# ================= HEADER =================
st.markdown(
    '<div class="title" style="font-size:24px; font-weight:2000; font-height:100;">GERÊNCIA DE GESTÃO ARQUIVÍSTICA</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle" style="font-size:16px; font-weight:1000;">MATE - MATÉRIAS EM TRAMITAÇÃO</div>',
    unsafe_allow_html=True
)

if "menu_open" not in st.session_state:
    st.session_state.menu_open = False

# ================= CARD =================

with st.form("form_mate", clear_on_submit=False):
    entrada = st.text_input(
        "Informe uma data de publicação válida",
        placeholder="Ex.: 24/02/2026 ou dia...",
    )

    st.caption(
        "- 24022026 ou 240226 ou 24/02/2026\n"
        "- hoje, ontem, anteontem\n"
        "- terça, quarta, quinta, sexta, sábado"
    )

    st.markdown('<div class="small-gap"></div>', unsafe_allow_html=True)

    col1, col2 = st.columns([2,1], gap=None)
    with col1:
        rodar = st.form_submit_button("📝 Gerar Planilha", type="primary")
    with col2:
        limpar = st.form_submit_button("🧹")

# ================= EXECUÇÃO =================
if limpar:
    st.session_state.clear()
    st.rerun()

if rodar:
    entrada_clean = (entrada or "").strip()

    if not entrada_clean:
        st.warning("Informe uma data, palavra ou URL.")
        st.stop()

    # --- VALIDAÇÃO AQUI ---
    yyyymmdd_check = normalizar_data(entrada_clean)
    dt_check = datetime.strptime(yyyymmdd_check, "%Y%m%d").date()
    data_pub_exec = dt_check.strftime("%Y-%m-%d")
    diario_exec_page = (
        f"https://www.jornalminasgerais.mg.gov.br/edicao-do-dia?"
        f"dados=%7B%22dataPublicacaoSelecionada%22:%22{data_pub_exec}T03:00:00.000Z%22%7D"
    )

    if dt_check.weekday() in (6, 0):  # domingo ou segunda
        st.error("Não há Diário do Legislativo para a data informada. Informe uma data válida.")
        st.stop()

    # busca o Diário do Executivo
    try:
        pdf_bytes_exec, filename_exec = fetch_diario_executivo_pdf_bytes(
            data_publicacao_yyyy_mm_dd=data_pub_exec,
            timeout_ms=90_000,
        )
        st.session_state["exec_pdf_bytes"] = pdf_bytes_exec
        st.session_state["exec_filename"] = filename_exec
    except Exception as e:
        st.session_state.pop("exec_pdf_bytes", None)
        st.session_state.pop("exec_filename", None)
        st.warning(f"Falha ao obter Diário do Executivo: {e}")
    
    # só agora começa a execução visual
    try:
        progress_bar = st.progress(0)
        status_text = st.empty()

        progress_bar.progress(5)
        status_text.write("Inicializando… 5%")

        result = {"url": None, "aba": None, "gid": None, "diario_url": None}
        err = {"exc": None}
        done = threading.Event()

        def run_main():
            try:
                r = main(
                    entrada_override=entrada_clean,
                    spreadsheet_url_or_id=st.secrets["SPREADSHEET_URL_OR_ID"],
                    auth_mode="service_account",
                    sa_info=st.secrets["gcp_service_account"],
                )
                result["url"] = r.get("url")
                result["aba"] = r.get("aba")
                result["gid"] = r.get("gid")
                result["diario_url"] = r.get("diario_url")
            except Exception as e:
                err["exc"] = e
            finally:
                done.set()

        # só inicia thread/progresso se passou na validação
        threading.Thread(target=run_main, daemon=True).start()

        pct_fake = 5

        while not done.is_set():
            if pct_fake < 99:
                pct_fake += 0.2  # sobe devagar até 99
            else:
                pct_fake = 99  # fica fixo em 99

            progress_bar.progress(int(pct_fake))

            spinner = ("⠋","⠙","⠹","⠸","⠼","⠴","⠦","⠧","⠇","⠏")
            frame = int(time.time() * 10) % len(spinner)
            status_text.write(f"{spinner[frame]} Processando Diários de Minas Gerais… {int(pct_fake)}%")

            time.sleep(0.1)

        if err["exc"] is not None:
            raise err["exc"]

        progress_bar.progress(100)
        status_text.write("Concluído 100%")

        if not result["url"] or result["gid"] is None:
            st.warning("Processo concluído, mas não foi possível montar o link da planilha.")
            st.write("Retorno:", result)
            st.stop()

        url_base = result["url"]
        gid = result["gid"]

        if "/edit" not in url_base:
            url_base = url_base.rstrip("/") + "/edit"

        url_com_aba = f"{url_base}#gid={gid}"

        st.error(f"Aba da Planilha: {result['aba']}")
        st.success(f"Diário do Legislativo: {result['diario_url']}")
        st.success(f"Diário do Executivo: {diario_exec_page}")

        # --- botões lado a lado: Planilha + Diário ---
        diario_url = (result.get("diario_url") or "").strip()

        c_btn1, c_btn2, c_btn3 = st.columns(3, gap="small")

        btn_style = """
            display:block;
            text-align:center;
            padding:10px;
            border-radius:8px;
            background-color:#e9e9e9;
            text-decoration:none;
            font-weight:500;
            color:black;
        """

        with c_btn1:
            st.markdown(
                f"""
                <a href="{url_com_aba}" target="_blank" rel="noopener noreferrer" style="{btn_style}">
                    Abrir Planilha
                </a>
                """,
                unsafe_allow_html=True
            )

        with c_btn2:
            if diario_url:
                html_btn2 = f"""
                <a href="{diario_url}" target="_blank" rel="noopener noreferrer" style="{btn_style}">
                    Diário do Legislativo
                </a>
                """
            else:
                html_btn2 = f"""
                <button style="{btn_style} opacity:0.55; cursor:not-allowed;">
                    Diário do Legislativo
                </button>
                """

            st.markdown(html_btn2, unsafe_allow_html=True)

        with c_btn3:
            pdf_bytes_exec = st.session_state.get("exec_pdf_bytes")
            filename_exec = st.session_state.get("exec_filename")

            if pdf_bytes_exec:
                b64_exec = base64.b64encode(pdf_bytes_exec).decode("ascii")
                safe_name_exec = filename_exec.replace("'", "").replace('"', "")

                components.html(
                    f"""
                    <button id="openPdfBtnTopExec" style="{btn_style}">
                    Diário do Executivo
                    </button>

                    <script>
                    (function() {{
                      const b64 = "{b64_exec}";
                      const fileName = "{safe_name_exec}";

                      function b64ToUint8Array(base64) {{
                        const binary = atob(base64);
                        const len = binary.length;
                        const bytes = new Uint8Array(len);
                        for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
                        return bytes;
                      }}

                      document.getElementById("openPdfBtnTopExec").addEventListener("click", () => {{
                        const bytes = b64ToUint8Array(b64);
                        const blob = new Blob([bytes], {{ type: "application/pdf" }});
                        const url = URL.createObjectURL(blob);

                        const w = window.open(url, "_blank");
                        if (!w) {{
                          alert("Popup bloqueado. Permita popups para este site e tente novamente.");
                          return;
                        }}
                        try {{ w.document.title = fileName; }} catch(e) {{}}
                      }});
                    }})();
                    </script>
                    """,
                    height=70,
                )
            else:
                st.markdown(
                    """
                    <div style="
                        display:inline-block;
                        width:100%;
                        padding:10px 12px;
                        background:#f0f0f0;
                        border:1px solid #d0d0d0;
                        border-radius:8px;
                        color:#999;
                        text-align:center;">
                        Diário do Executivo
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
    except Exception as e:
        st.error("Erro ao processar.")
        st.exception(e)

# ======================================================================================
# EXTRA: Jornal Minas Gerais (Diário do Executivo) — abrir PDF em nova aba (sem Playwright)
# ======================================================================================
import base64
import streamlit.components.v1 as components
from playwright import fetch_diario_executivo_pdf_bytes

st.divider()
st.subheader("Diário do Executivo")

data_pub = st.text_input("Data de publicação (YYYY-MM-DD)", value="2026-03-03", key="jmg_data_pub")

if st.button("Gerar link do PDF (nova aba)", key="jmg_btn_open"):
    try:
        status = st.empty()

        def ui_log(msg: str) -> None:
            status.write(msg)

        with st.spinner("Obtendo PDF pela API..."):
            pdf_bytes, filename = fetch_diario_executivo_pdf_bytes(
                data_publicacao_yyyy_mm_dd=data_pub,
                timeout_ms=90_000,
                log=ui_log,
            )

        st.success(f"OK: {filename}")

        # 1) Botão download (sempre funciona)
        st.download_button(
            "Baixar PDF",
            data=pdf_bytes,
            file_name=filename,
            mime="application/pdf",
            key="jmg_dl_btn",
        )

        # 2) Abrir em nova aba via Blob URL (evita about:blank#blocked)
        b64 = base64.b64encode(pdf_bytes).decode("ascii")
        safe_name = filename.replace("'", "").replace('"', "")

        components.html(
            f"""
            <button id="openPdfBtn" style="
                display:inline-block;padding:8px 12px;background:#e9e9e9;border-radius:6px;
                border:0;cursor:pointer;margin-top:8px;">
                Abrir PDF em nova aba
            </button>

            <script>
            (function() {{
              const b64 = "{b64}";
              const fileName = "{safe_name}";

              function b64ToUint8Array(base64) {{
                const binary = atob(base64);
                const len = binary.length;
                const bytes = new Uint8Array(len);
                for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
                return bytes;
              }}

              document.getElementById("openPdfBtn").addEventListener("click", () => {{
                const bytes = b64ToUint8Array(b64);
                const blob = new Blob([bytes], {{ type: "application/pdf" }});
                const url = URL.createObjectURL(blob);

                const w = window.open(url, "_blank");
                if (!w) {{
                  alert("Popup bloqueado. Permita popups para este site e tente novamente.");
                  return;
                }}
                try {{ w.document.title = fileName; }} catch(e) {{}}
              }});
            }})();
            </script>
            """,
            height=90,
        )

    except Exception as e:
        st.error(f"Falhou: {e}")