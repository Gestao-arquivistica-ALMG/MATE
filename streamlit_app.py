import streamlit as st
from mate_pipeline import main
import time
from concurrent.futures import ThreadPoolExecutor

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
EXEC = ThreadPoolExecutor(max_workers=1)

# ================= ESTILO =================
st.markdown("""
<style>

/* Remove o padding superior padrão do Streamlit */
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
    margin-bottom: 0.2rem;  /* reduz espaço abaixo */
}

/* Subtítulo */
.subtitle {
    text-align: center;
    color: #6b7280;
    margin-top: 0;
    margin-bottom: 1rem;  /* reduz espaço antes do input */
}

</style>
""", unsafe_allow_html=True)

# ================= HEADER =================
st.markdown('<div class="title">GERÊNCIA DE GESTÃO ARQUIVÍSTICA</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">MATE - MATÉRIAS EM TRAMITAÇÃO</div>', unsafe_allow_html=True)

# ================= CARD =================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    with st.form("form_mate", clear_on_submit=False):

        entrada = st.text_input(
            """Informe uma data do Diário do Legislativo

    - 19122026 ou 191226 ou 19/12/2026
    - hoje, ontem, anteontem
    - terça, quarta, quinta, sexta, sábado

    """,
            placeholder="Ex: 19/02/2026 ou https://...",
        )

        st.write("")

        running = bool(st.session_state.get("job") and not st.session_state["job"].done())

        col1, col2 = st.columns(2, gap="small")

        with col1:
            rodar = st.form_submit_button("🚀 Gerar", use_container_width=True, type="primary", disabled=running)

        with col2:
            limpar = st.form_submit_button("🧹 Limpar", use_container_width=True, disabled=running)

    st.markdown('</div>', unsafe_allow_html=True)

# ================= LÓGICA =================
# ================= EXECUÇÃO (NÃO BLOQUEANTE) =================

# inicializa estados
st.session_state.setdefault("job", None)         # future
st.session_state.setdefault("job_input", "")     # entrada da execução
st.session_state.setdefault("job_result", None)  # (url, aba)
st.session_state.setdefault("job_error", None)   # exception

# LIMPAR
if limpar:
    st.session_state.clear()
    st.rerun()

# DISPARAR
if rodar:
    entrada_clean = (entrada or "").strip()
    if not entrada_clean:
        st.warning("Informe uma data, palavra ou URL.")
        st.stop()

    # evita duplo disparo
    if st.session_state["job"] is None or st.session_state["job"].done():
        st.session_state["job_input"] = entrada_clean
        st.session_state["job_result"] = None
        st.session_state["job_error"] = None
        st.session_state["job"] = EXEC.submit(_run_pipeline, entrada_clean)

# UI DE STATUS (fora do card, como no seu print)
job = st.session_state.get("job")

if job and not job.done():
    st.info("Processando Diário...")
    time.sleep(0.3)  # permite redesenhar e evita loop agressivo
    st.rerun()

# capturar resultado uma vez
if job and job.done() and st.session_state["job_result"] is None and st.session_state["job_error"] is None:
    try:
        url, aba = job.result()
        st.session_state["job_result"] = (url, aba)
    except Exception as e:
        st.session_state["job_error"] = e

# RESULTADO / ERRO
if st.session_state["job_result"]:
    url, aba = st.session_state["job_result"]
    st.success("Concluído.")
    st.write("Aba:", aba)
    st.link_button("Abrir planilha", url, use_container_width=True)

if st.session_state["job_error"] is not None:
    st.error("Erro ao processar.")
    st.exception(st.session_state["job_error"])
    
    with st.status("Processando Diário do Legislativo...", expanded=True) as status:
        try:
            url, aba = main(
                entrada_override=entrada.strip(),
                spreadsheet_url_or_id=st.secrets["SPREADSHEET_URL_OR_ID"],
                auth_mode="service_account",
                sa_info=st.secrets["gcp_service_account"],
            )

            status.update(label="Concluído com sucesso ✅", state="complete")

            st.success("Planilha gerada com sucesso.")
            st.info(f"Aba criada: **{aba}**")

            st.link_button("📊 Abrir planilha", url, use_container_width=True)

        except Exception as e:
            status.update(label="Erro no processamento ❌", state="error")
            st.error("Erro ao processar o Diário.")
            st.exception(e)