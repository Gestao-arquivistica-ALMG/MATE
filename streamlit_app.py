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

/* Fundo externo */
html, body, [data-testid="stAppViewContainer"]{
  background:#b30000;
}

/* Container geral branco (fecha tudo) */
.block-container{
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
  padding:20px !important;
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
  max-width:340px;
  margin:0 auto;
}

/* Conteúdo legível: alinhamento à esquerda */
div[data-testid="stForm"] label,
div[data-testid="stForm"] .stCaption,
div[data-testid="stForm"] ul,
div[data-testid="stForm"] li{
  text-align:left !important;
}

/* Linha dos botões: só a ÚLTIMA linha de colunas dentro do form */
div[data-testid="stForm"] div[data-testid="stHorizontalBlock"]:last-of-type{
  max-width:260px !important;      /* largura do grupo [Gerar][🧹] */
  margin:12px auto 0 auto !important; /* centraliza o grupo */
}

/* Botões: tamanhos fixos (não espicha) */
div[data-testid="stFormSubmitButton"] button{
  font-family:Montserrat !important;
  font-weight:700 !important;
  border-radius:14px !important;
  padding:12px 14px !important;
  min-width:140px !important;      /* Gerar */
}

/* Segundo botão (🧹) menor */
div[data-testid="stForm"] div[data-testid="stHorizontalBlock"]:last-of-type
div[data-testid="stFormSubmitButton"]:nth-of-type(2) button{
  min-width:60px !important;
  padding:12px 0 !important;
}

.small-gap{ margin-top:10px; }

</style>
""", unsafe_allow_html=True)

# ================= HEADER ALMG =================

st.markdown("""
<div style='
display:flex;
align-items:center;
justify-content:space-between;
margin:0 auto 20px auto;
background:white;
max-width:560px;
padding:10px 18px;
border-radius:12px;
'>

<div style='font-size:35px;color:#cc0000'>☰</div>

<div>
<img src="https://www.almg.gov.br/system/modules/br.gov.almg.portal/resources/img/logo/logo.svg"
style="height:45px;">
</div>

<div style='font-size:30px;color:#cc0000'>
🔍 👤
</div>

</div>
""",unsafe_allow_html=True)

# ================= HEADER =================
st.markdown(
    '<div class="title" style="font-size:24px; font-weight:1000;">GERÊNCIA DE GESTÃO ARQUIVÍSTICA</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle" style="font-size:14px; font-weight:1000;">MATE - MATÉRIAS EM TRAMITAÇÃO</div>',
    unsafe_allow_html=True
)

# ================= CARD =================

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

    col1, col2 = st.columns([3,1], gap="small")
    with col1:
        rodar = st.form_submit_button("🚀 Gerar", type="primary")
    with col2:
        limpar = st.form_submit_button("🧹")

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