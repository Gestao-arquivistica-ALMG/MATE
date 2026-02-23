import streamlit as st
from mate_pipeline import main

# ================= CONFIG =================
st.set_page_config(
    page_title="MATE",
    page_icon="🧠",
    layout="wide",
)

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

        col1, col2 = st.columns(2, gap="small")

        with col1:
            rodar = st.form_submit_button("🚀 Gerar", use_container_width=True, type="primary")

        with col2:
            limpar = st.form_submit_button("🧹 Limpar", use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ================= LÓGICA =================
if limpar:
    st.session_state.clear()
    st.rerun()

if rodar:

    if not entrada.strip():
        st.warning("⚠️ Informe uma data, palavra ou URL.")
        st.stop()

    with st.status("Processando Diário...", expanded=True) as status:
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