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
.main {
    background-color: #f5f7fa;
}

/* Limita largura da página */
.block-container {
    max-width: 720px;
    padding-top: 3rem;
    padding-bottom: 2rem;
    margin: auto;
}

/* Título */
.title {
    font-size: 36px;
    font-weight: 700;
    text-align: center;
    margin-bottom: 5px;
}

/* Subtítulo */
.subtitle {
    text-align: center;
    color: #6b7280;
    margin-bottom: 30px;
}

/* Card */
.card {
    background: white;
    padding: 2rem;
    border-radius: 14px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.05);
}
</style>
""", unsafe_allow_html=True)

# ================= HEADER =================
st.markdown('<div class="title">MATE.IA</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Automação do Diário do Legislativo</div>', unsafe_allow_html=True)

# ================= CARD =================
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)

    entrada = st.text_input(
        "Data / Palavra / URL do Diário"
        "EXEMPLOS:"
        "- 19122026 ou 191226 ou 19/12/2026"
        "- hoje, ontem, anteontem"
        "- terça, quarta, quinta, sexta, sábado"
        "- URL completa ou um caminho local.",
        print("Digite a data do Diário do Legislativo.")

        placeholder="Ex: 12/02/2026 ou https://...",
    )

    col1, col2 = st.columns(2)

    with col1:
        rodar = st.button("🚀 Gerar planilha", use_container_width=True, type="primary")

    with col2:
        limpar = st.button("🧹 Limpar", use_container_width=True)

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