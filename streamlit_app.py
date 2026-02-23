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
    border: 1px solid rgba(0,0,0,0.10);
    border-radius: 12px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.04);
}

/* Input mais “web” */
div[data-baseweb="input"] > div {
    background: #f3f4f6;
}

/* Pequena folga */
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
            placeholder="Ex: 19/02/2026 ou https://...",
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
            limpar = st.form_submit_button("🧹 Limpar", use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

# ================= EXECUÇÃO (SÍNCRONA / ESTÁVEL) =================
if limpar:
    st.session_state.clear()
    st.rerun()

if rodar:
    entrada_clean = (entrada or "").strip()
    if not entrada_clean:
        st.warning("Informe uma data, palavra ou URL.")
        st.stop()

    try:
        with st.spinner("Processando Diário..."):
            url, aba = main(
                entrada_override=entrada_clean,
                spreadsheet_url_or_id=st.secrets["SPREADSHEET_URL_OR_ID"],
                auth_mode="service_account",
                sa_info=st.secrets["gcp_service_account"],
            )

        st.success("Concluído.")
        st.write("Aba:", aba)
        st.link_button("Abrir planilha", url, use_container_width=True)

    except Exception as e:
        st.error("Erro ao processar.")
        st.exception(e)