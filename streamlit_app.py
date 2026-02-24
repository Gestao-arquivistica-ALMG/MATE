import re
import streamlit as st
from mate_pipeline import main

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

html, body, [data-testid="stAppViewContainer"]{
background:#b30000;
}

.block-container{
max-width:650px;
margin:auto;
padding-top:60px;
padding-bottom:60px;
}

.title{
font-family: Montserrat;
color:white;
}

.subtitle{
text-align:center;
color:white;
font-family:Inter;
margin-bottom:20px;
}

.card{
background:white !important;
padding:20px;
border-radius:18px;
box-shadow:0 10px 30px rgba(0,0,0,0.25);
}

div[data-testid="stForm"]{
background:white !important;
}

div[data-testid="stTextInput"] > div{
background:white !important;
}

.stButton>button{
font-family:Inter;
font-weight:700;
border-radius:14px;
padding:14px;
}

.small-gap{margin-top:10px;}

</style>
""",unsafe_allow_html=True)

# ================= HEADER ALMG =================

st.markdown("""
<div style='
display:flex;
align-items:center;
justify-content:space-between;
margin-bottom:20px;
background:white;
'>

<div style='font-size:28px;color:#cc0000'>☰</div>

<div>
<img src="https://www.almg.gov.br/system/modules/br.gov.almg.portal/resources/img/logo/logo.svg"
style="height:36px;">
</div>

<div style='font-size:22px;color:#cc0000'>
🔍 👤
</div>

</div>
""",unsafe_allow_html=True)

# ================= HEADER =================
st.markdown(
    '<div class="title" style="font-size:32px;">GERÊNCIA DE GESTÃO ARQUIVÍSTICA</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle" style="font-size:9px;">MATE - MATÉRIAS EM TRAMITAÇÃO</div>',
    unsafe_allow_html=True
)

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
        with st.status("Processando Diário do Legislativo...", expanded=True) as status:
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