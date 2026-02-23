import re
import streamlit as st
from mate_pipeline import main

# ================= CONFIG =================
st.set_page_config(
    page_title="MATE",
    page_icon="🧠",
    layout="wide",
)

# ================= ESTILO =================
st.markdown(
    """
<style>
.block-container {
    padding-top: 1rem !important;
    padding-bottom: 2rem;
    max-width: 600px;
    margin: auto;
}

.title {
    font-size: 34px;
    font-weight: 700;
    text-align: center;
    margin-bottom: 0.2rem;
}

.subtitle {
    text-align: center;
    color: #6b7280;
    margin-top: 0;
    margin-bottom: 1rem;
}

.card {
    background: #ffffff;
    padding: 1.4rem;
    border: 1px solid rgba(0,0,0,0.10);
    border-radius: 12px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.04);
}

div[data-baseweb="input"] > div {
    background: #f3f4f6;
}

.small-gap { margin-top: 0.6rem; }
</style>
""",
    unsafe_allow_html=True,
)

# ================= HEADER =================
st.markdown('<div class="title">GERÊNCIA DE GESTÃO ARQUIVÍSTICA</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">MATE - MATÉRIAS EM TRAMITAÇÃO</div>', unsafe_allow_html=True)

# ================= CARD =================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    with st.form("form_mate", clear_on_submit=False):
        entrada = st.text_input(
            "Informe uma data do Diário do Legislativo",
            placeholder="Ex: 21/02/2026 ou https://...",
        )

        st.caption(
            "- 19122026 ou 191226 ou 19/12/2026\n"
            "- hoje, ontem, anteontem\n"
            "- terça, quarta, quinta, sexta, sábado"
        )

        st.markdown('<div class="small-gap"></div>', unsafe_allow_html=True)

        col1, col2 = st.columns(2, gap="small")
        with col1:
            rodar = st.form_submit_button("🚀 Gerar", use_container_width=True, type="primary")
        with col2:
            st.write("")

    # Limpar fora do form (ENTER = Gerar garantido)
    col1, col2 = st.columns(2, gap="small")
    with col1:
        st.write("")
    with col2:
        limpar = st.button("🧹 Limpar", use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

# ================= EXECUÇÃO =================
if limpar:
    st.session_state.clear()
    st.rerun()

if rodar:
    entrada_clean = (entrada or "").strip()

    if not entrada_clean.strip():
        st.warning("Informe uma data, palavra ou URL.")
        st.stop()

    st.info(f"DEBUG: entrada enviada = {entrada_clean!r}")

    try:
        with st.status("Processando Diário...", expanded=True) as status:
            status.write("Checkpoint 1: antes do main()")

            url, aba = main(
                entrada_override=entrada_clean,
                spreadsheet_url_or_id=st.secrets["SPREADSHEET_URL_OR_ID"],
                auth_mode="service_account",
                sa_info=st.secrets["gcp_service_account"],
            )

            status.write("Checkpoint 2: depois do main()")
            status.update(label="Concluído ✅", state="complete", expanded=False)

        st.success("Concluído.")
        st.write("Aba:", aba)
        st.link_button("Abrir planilha", url, use_container_width=True)

    except Exception as e:
        st.error("Erro ao processar.")
        st.exception(e)