import streamlit as st
import threading
import time
import base64
import requests
import streamlit.components.v1 as components
from datetime import datetime
from mate_pipeline import main, normalizar_data, proximo_dia_util, yyyymmdd_to_ddmmyyyy
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

</style>
""", unsafe_allow_html=True)

# ================= HEADER ALMG =================
if "menu_open" not in st.session_state:
    st.session_state.menu_open = False

# "cabeçalho" feito com layout Streamlit (sem HTML clicável)
c1, c2, c3 = st.columns([0.7, 7, 5], gap="small")

with c1:
    if st.button("☰", key="btn_menu_toggle", use_container_width=True):
        st.session_state.menu_open = not st.session_state.menu_open
        st.rerun()

with c2:
    st.markdown(
        """
        <div style="
            display:flex;
            align-items:center;
            justify-content:flex-start;
            height:56px;
        ">
        <img src="https://www.almg.gov.br/system/modules/br.gov.almg.portal/resources/img/logo/logo.svg"
            style="height:40px; display:block; transform:translateY(-12px);">
        </div>
        """,
        unsafe_allow_html=True,
    )

with c3:
    st.markdown(
        """
        <div style="
            display:flex;
            align-items:right;
            justify-content:flex-end;
            height:56px;
            gap:8px;
        ">
          <a href="https://silegis.almg.gov.br/silegismg/login/login.jsp#/processos" target="_blank"
             style="
                text-decoration:none;
                width:25px;
                height:25px;
                display:flex;
                align-items:flex-start;
                justify-content:center;
                border-radius:8px;
                font-size:22px;
             "><img src="https://silegis.almg.gov.br/silegismg/favicon.ico"></a>

          <a href="https://intra.almg.gov.br/" target="_blank"
             style="
                text-decoration:none;
                width:100px;
                height:100px;
                display:flex;
                align-items:center;
                justify-content:center;
                border-radius:8px;
                font-size:22px;
             "><img src="https://intra.almg.gov.br/.content/imagens/logo-intra.svg"
                    style="height:24px; width:auto; display:block;"></a>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("""
<style>
div[data-testid="stHorizontalBlock"]{
  background:white;
  padding:10px 18px;
  border-radius:14px;
  align-items:center !important;
}

div[data-testid="stHorizontalBlock"] > div{
  display:flex;
  align-items:center;
}

/* botão menu */
button[kind="secondary"]{
  height:26px !important;
  min-height:26px !important;
  line-height:1 !important;
  border-radius:8px !important;
  border:1px solid #d0d0d0 !important;
  background:#f7f7f7 !important;
  color:#666666 !important;
  font-size:20px !important;
  font-weight:500 !important;
  padding:0 !important;
  box-shadow:none !important;
}

button[kind="secondary"]:hover{
  border:1px solid #cc0000 !important;
  color:#cc0000 !important;
  background:#fff5f5 !important;
}

/* links dos ícones */
a:hover{
  background:#f5f5f5;
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
    """, unsafe_allow_html=True)
            
# ================= HEADER =================
st.markdown(
    '<div class="title" style="font-size:24px; font-weight:2000; font-height:100;">GERÊNCIA-GERAL DE ARQUIVO</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle" style="font-size:16px; font-weight:1000;">MATE - MATÉRIAS EM TRAMITAÇÃO</div>',
    unsafe_allow_html=True
)

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

    col1, col2 = st.columns([9,8], gap=None)
    with col1:
        rodar = st.form_submit_button("📝 Gerar Planilha", type="primary", help="Gerar nova aba no Google Sheets")
    with col2:
        limpar = st.form_submit_button("🧹", key="limpar", help="Limpar campos e reiniciar processamento")

# ================= EXECUÇÃO =================
if limpar:
    st.session_state.clear()
    st.rerun()

if rodar:
    progress_bar = st.progress(0)
    status_text = st.empty()

    status_text.markdown(
        """
        <div style="
            font-family:'Montserrat',sans-serif;
            font-size:13px;
            color:#31333F;
        ">
            Inicializando<span class="loading-dots"></span>
        </div>

        <style>
        .loading-dots::after {
            content: "";
            display: inline-block;
            width: 1.2em;
            text-align: left;
            animation: loadingDots 1s steps(4, end) infinite;
        }

        @keyframes loadingDots {
            0%   { content: ""; }
            25%  { content: "."; }
            50%  { content: ".."; }
            75%  { content: "..."; }
            100% { content: ""; }
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    entrada_clean = (entrada or "").strip()
    if not entrada_clean:
        st.warning("Informe uma data, palavra ou URL.")
        st.stop()

    # --- VALIDAÇÃO AQUI ---
    yyyymmdd_check = normalizar_data(entrada_clean)
    dt_check = datetime.strptime(yyyymmdd_check, "%Y%m%d").date()
    data_pub_exe = dt_check.strftime("%Y-%m-%d")
    data_pub_leg = dt_check.strftime("%d/%m/%Y")
    aba_yyyymmdd_check = proximo_dia_util(yyyymmdd_check)
    data_reuniao = yyyymmdd_to_ddmmyyyy(aba_yyyymmdd_check)

    diario_exe_page = (
        f"https://www.jornalminasgerais.mg.gov.br/edicao-do-dia?"
        f"dados=%7B%22dataPublicacaoSelecionada%22:%22{data_pub_exe}T03:00:00.000Z%22%7D"
    )

    diario_leg_page = f"https://diariolegislativo.almg.gov.br/{yyyymmdd_check[:4]}/L{yyyymmdd_check}.pdf"

    reuniao_plenario = (
        f"https://www.almg.gov.br/atividade-parlamentar/plenario/agenda/"
        f"?pesquisou=true&q=&tipo=&dataInicio={data_reuniao}&dataFim={data_reuniao}"
    )

    reuniao_comissoes = (
        f"https://www.almg.gov.br/atividade-parlamentar/comissoes/agenda/"
        f"?pesquisou=true&q=&tpComissao=&idComissao=&dataInicio={data_reuniao}"
        f"&dataFim={data_reuniao}&pesquisa=todas&ordem=1&tp=30"
    )

    if dt_check.weekday() in (6, 0):  # domingo ou segunda
        st.error("Não há Diário do Legislativo para a data informada. Informe uma data válida.")
        st.stop()

    # busca o Diário do Executivo
    try:
        pdf_bytes_exec, filename_exec = fetch_diario_executivo_pdf_bytes(
            data_publicacao_yyyy_mm_dd=data_pub_exe,
            timeout_ms=90_000,
        )
        st.session_state["exec_pdf_bytes"] = pdf_bytes_exec
        st.session_state["exec_filename"] = filename_exec
    except Exception as e:
        st.session_state.pop("exec_pdf_bytes", None)
        st.session_state.pop("exec_filename", None)
        st.warning(f"Falha ao obter Diário do Executivo: {e}")

    # busca o Diário do Legislativo
    try:
        resp_leg = requests.get(diario_leg_page, timeout=30)
        resp_leg.raise_for_status()
        st.session_state["leg_pdf_bytes"] = resp_leg.content
        st.session_state["leg_filename"] = f"L{yyyymmdd_check}.pdf"
    except Exception as e:
        st.session_state.pop("leg_pdf_bytes", None)
        st.session_state.pop("leg_filename", None)
        st.warning(f"Falha ao obter Diário do Legislativo: {e}")

    status_text.empty()

    # EXIBE ANTES DO PROCESSAMENTO
    open_icon = "https://cdn-icons-png.flaticon.com/512/4949/4949024.png"
    pdf_icon = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTcAfrx37j4CuKi12fymiEDw35IRZj62mplHM0PeeGo2w&s"

    c_links, c_pdfs = st.columns([1.2, 1.8], gap=None)

    with c_links:
        components.html(
            f"""
            <link rel="preconnect" href="https://fonts.googleapis.com">
            <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
            <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600&display=swap" rel="stylesheet">
            <style>
            .link-line a {{
                font-family:'Montserrat',sans-serif;
                font-size:13px;
                color:#31333F;
                text-decoration:none;
                font-weight:600;
            }}

            .link-line img {{
                height:16px;
                vertical-align:middle;
            }}
            </style>

            <div class="link-line">

                <div style="margin:0 0 8px 0;">
                    <a href="javascript:void(0)" id="openExecTop" style="margin-right:6px;">
                        <img src="{open_icon}">
                    </a>
                    <a href="{diario_exe_page}" target="_blank" rel="noopener noreferrer">
                        Diário do Executivo
                    </a>
                </div>

                <div style="margin:0 0 8px 0;">
                    <a href="{diario_leg_page}" target="_blank" rel="noopener noreferrer" style="margin-right:6px;">
                        <img src="{open_icon}">
                    </a>
                    <a href="https://www.almg.gov.br/transparencia/diario-do-legislativo/index.html" target="_blank" rel="noopener noreferrer">
                        Diário do Legislativo
                    </a>
                </div>

                <div style="margin:0 0 8px 0;">
                    <a href="{reuniao_plenario}" target="_blank" rel="noopener noreferrer" style="margin-right:6px;">
                        <img src="{open_icon}">
                    </a>
                    <a href="https://www.almg.gov.br/atividade-parlamentar/plenario/agenda/" target="_blank" rel="noopener noreferrer">
                        Reuniões de Plenário
                    </a>
                </div>

                <div style="margin:0 0 8px 0;">
                    <a href="{reuniao_comissoes}" target="_blank" rel="noopener noreferrer" style="margin-right:6px;">
                        <img src="{open_icon}">
                    </a>
                    <a href="https://www.almg.gov.br/atividade-parlamentar/comissoes/agenda/" target="_blank" rel="noopener noreferrer">
                        Reuniões de Comissões
                    </a>
                </div>

                <script>
                (function() {{
                const b64Exec = "{base64.b64encode(st.session_state.get('exec_pdf_bytes', b'')).decode('ascii')}";
                const fileNameExec = "{st.session_state.get('exec_filename', 'diario-executivo.pdf').replace("'", "").replace('"', "")}";

                const btnExecTop = document.getElementById("openExecTop");

                function b64ToUint8Array(base64) {{
                    const binary = atob(base64);
                    const len = binary.length;
                    const bytes = new Uint8Array(len);
                    for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
                    return bytes;
                }}

                if (btnExecTop && b64Exec) {{
                    btnExecTop.addEventListener("click", () => {{
                    const bytes = b64ToUint8Array(b64Exec);
                    const blob = new Blob([bytes], {{ type: "application/pdf" }});
                    const url = URL.createObjectURL(blob);

                    const w = window.open(url, "_blank");
                    if (!w) {{
                        alert("Popup bloqueado. Permita popups para este site e tente novamente.");
                        return;
                    }}
                    try {{ w.document.title = fileNameExec; }} catch(e) {{}}
                    }});
                }}
                }})();
                </script>

            </div>
            """,
            height=120,
        )

    with c_pdfs:
        components.html(
            f"""
            <div style="margin:0 0 8px 0; display:flex; justify-content:flex-start;">
                <a href="javascript:void(0)" id="downloadExecPdf" style="text-decoration:none;">
                    <img src="{pdf_icon}" style="height:16px; vertical-align:middle; position:relative; top:-3px;">
                </a>
            </div>

            <div style="margin:0 0 8px 0; display:flex; justify-content:flex-start;">
                <a href="javascript:void(0)" id="downloadLegPdf" style="text-decoration:none;">
                    <img src="{pdf_icon}" style="height:16px; vertical-align:middle; position:relative; top:-3px;">
                </a>
            </div>

            <div style="margin:0 0 8px 0; height:16px;"></div>
            <div style="margin:0 0 8px 0; height:16px;"></div>

            <script>
            (function() {{

            const b64Exec = "{base64.b64encode(st.session_state.get('exec_pdf_bytes', b'')).decode('ascii')}";
            const fileNameExec = "{st.session_state.get('exec_filename', 'diario-executivo.pdf').replace('"','').replace("'",'')}";

            const btnExec = document.getElementById("downloadExecPdf");

            function baixarExecPdf() {{
                const binary = atob(b64Exec);
                const len = binary.length;
                const bytes = new Uint8Array(len);

                for (let i = 0; i < len; i++) {{
                bytes[i] = binary.charCodeAt(i);
                }}

                const blob = new Blob([bytes], {{ type: "application/pdf" }});
                const url = URL.createObjectURL(blob);

                const a = document.createElement("a");
                a.href = url;
                a.download = fileNameExec;
                document.body.appendChild(a);
                a.click();
                a.remove();
            }}

            if (btnExec) {{
                btnExec.addEventListener("click", baixarExecPdf);
            }}

            const b64Leg = "{base64.b64encode(st.session_state.get('leg_pdf_bytes', b'')).decode('ascii')}";
            const fileNameLeg = "{st.session_state.get('leg_filename', 'diario-legislativo.pdf').replace('"','').replace("'",'')}";

            const btnLeg = document.getElementById("downloadLegPdf");
            if (btnLeg) {{
                btnLeg.addEventListener("click", () => {{
                const binary = atob(b64Leg);
                const len = binary.length;
                const bytes = new Uint8Array(len);

                for (let i = 0; i < len; i++) {{
                    bytes[i] = binary.charCodeAt(i);
                }}

                const blob = new Blob([bytes], {{ type: "application/pdf" }});
                const url = URL.createObjectURL(blob);

                const a = document.createElement("a");
                a.href = url;
                a.download = fileNameLeg;
                document.body.appendChild(a);
                a.click();
                a.remove();
                }});
            }}

            }})();
            </script>
            """,
            height=110
        )
    # só agora começa a execução visual
    try:
        progress_bar = st.progress(0)
        status_text = st.empty()

        result = {"url": None, "aba": None, "gid": None, "diario_url": None}
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
            status_text.markdown(
                f"""
                <div style="font-family:'Montserrat',sans-serif; font-size:12px; color:#31333F;">
                    {spinner[frame]} Processando publicações oficiais de Minas Gerais… {int(pct_fake)}%
                </div>
                """,
                unsafe_allow_html=True
            )

            time.sleep(0.1)

        progress_bar.progress(100)

        url_base = result["url"]
        gid = result["gid"]

        if "/edit" not in url_base:
            url_base = url_base.rstrip("/") + "/edit"

        url_com_aba = f"{url_base}#gid={gid}"

        status_text.markdown(
            f"""
        <div style="display:flex;justify-content:space-between;align-items:center;
        font-family:'Montserrat',sans-serif;font-size:13px;color:#31333F;width:100%;">

        <span>Concluído 100%</span>

        <a href="{url_com_aba}" target="_blank" rel="noopener noreferrer"
        style="text-decoration:none;color:#31333F;font-weight:400;
        position:relative; top:1px;">
        Abrir planilha
        </a>

        <span>{result['aba']}</span>

        </div>
        """,
            unsafe_allow_html=True
        )

        if not result["url"] or result["gid"] is None:
            st.warning("Processo concluído, mas não foi possível montar o link da planilha.")
            st.write("Retorno:", result)
            st.stop()

        #st.write("")

    except Exception as e:
        st.error("Erro ao processar.")
        st.exception(e)