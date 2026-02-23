import time
from concurrent.futures import ThreadPoolExecutor

import streamlit as st
from mate_pipeline import main


# ================= EXECUTOR (1 worker, persistente entre reruns) =================
@st.cache_resource
def get_executor():
    return ThreadPoolExecutor(max_workers=1)


def _run_pipeline(entrada: str):
    return main(
        entrada_override=entrada,
        spreadsheet_url_or_id=st.secrets["SPREADSHEET_URL_OR_ID"],
        auth_mode="service_account",
        sa_info=st.secrets["gcp_service_account"],
    )


# ================= CONFIG =================
st.set_page_config(
    page_title="MATE",
    page_icon="🧠",
    layout="wide",
)

EXEC = get_executor()

# ================= ESTILO =================
st.markdown(
    """
<style>
/* Página compacta */
.block-container {
    padding-top: 1rem !important;
    padding-bottom: 2rem;
    max-width: 600px;
    margin: auto;
}

/* Título principal */
.title {
    font-size: 34px;
    font-weight: 700;
    text-align: center;
    margin-bottom: 0.2rem;
}

/* Subtítulo */
.subtitle {
    text-align: center;
    color: #6b7280;
    margin-top: 0;
    margin-bottom: 1rem;
}

/* Card */
.card {
    background: #ffffff;
    padding: 1.4rem;
    border: 1px solid rgba(0,0,0,0.08);
    border-radius: 12px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.04);
}

/* Deixa o input mais “web” */
div[data-baseweb="input"] > div {
    background: #f3f4f6;
}

/* Reduz um pouco o espaço padrão dos elementos */
.small-gap {
    margin-top: 0.6rem;
}
</style>
""",
    unsafe_allow_html=True,
)

# ================= HEADER =================
st.markdown('<div class="title">GERÊNCIA DE GESTÃO ARQUIVÍSTICA</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">MATE - MATÉRIAS EM TRAMITAÇÃO</div>', unsafe_allow_html=True)

# ================= STATE (antes do card) =================
st.session_state.setdefault("job", None)         # future
st.session_state.setdefault("job_result", None)  # (url, aba)
st.session_state.setdefault("job_error", None)   # exception

job = st.session_state.get("job")
running = bool(job and not job.done())

# ================= CARD =================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    # ENTER só funciona como submit dentro de form
    with st.form("form_mate", clear_on_submit=False):
        entrada = st.text_input(
            """Informe uma data do Diário do Legislativo

- 19122026 ou 191226 ou 19/12/2026
- hoje, ontem, anteontem
- terça, quarta, quinta, sexta, sábado
""",
            placeholder="Ex: 19/02/2026 ou https://...",
        )

        st.markdown('<div class="small-gap"></div>', unsafe_allow_html=True)

        col1, col2 = st.columns(2, gap="small")
        with col1:
            rodar = st.form_submit_button("🚀 Gerar", use_container_width=True, type="primary", disabled=running)
        with col2:
            limpar = st.form_submit_button("🧹 Limpar", use_container_width=True, disabled=running)

    st.markdown("</div>", unsafe_allow_html=True)

# ================= AÇÕES =================
if limpar:
    st.session_state.clear()
    st.rerun()

if rodar:
    entrada_clean = (entrada or "").strip()
    if not entrada_clean:
        st.warning("Informe uma data, palavra ou URL.")
        st.stop()

    # dispara 1 job por vez
    job = st.session_state.get("job")
    if job is None or job.done():
        st.session_state["job_result"] = None
        st.session_state["job_error"] = None
        st.session_state["job"] = EXEC.submit(_run_pipeline, entrada_clean)

# ================= STATUS / RESULTADO =================
job = st.session_state.get("job")

# Enquanto roda: mostra status e faz rerun leve (sem loop agressivo)
if job and not job.done():
    st.info("Processando Diário...")
    time.sleep(0.8)  # dá tempo da UI respirar e evita rerun frenético
    st.experimental_rerun()

# Captura resultado uma vez
if job and job.done() and st.session_state["job_result"] is None and st.session_state["job_error"] is None:
    try:
        st.session_state["job_result"] = job.result()
    except Exception as e:
        st.session_state["job_error"] = e

# Mostra resultado / erro
if st.session_state["job_result"]:
    url, aba = st.session_state["job_result"]
    st.success("Concluído.")
    st.write("Aba:", aba)
    st.link_button("Abrir planilha", url, use_container_width=True)

if st.session_state["job_error"] is not None:
    st.error("Erro ao processar.")
    st.exception(st.session_state["job_error"])