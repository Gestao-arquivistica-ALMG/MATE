import re
import streamlit as st
import threading
import time
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

/* Linha dos botões: força ficar em uma linha no mobile */
div[data-testid="stForm"] div[data-testid="stHorizontalBlock"]:last-of-type{
  max-width:300px !important;          /* ajuste manual aqui */
  margin:12px auto 0 0 !important;     /* ajuste manual aqui */
  flex-wrap:nowrap !important;         /* NÃO quebra */
  gap:10px !important;                 /* ajuste manual aqui */
}

/* 🚀 Gerar (PRIMARY) */
div[data-testid="stForm"] div[data-testid="stHorizontalBlock"]:last-of-type
button[kind="primary"]{
  min-width:240px !important;   /* ajuste aqui */
  padding:12px 18px !important;
}

div[data-testid="stForm"] div[data-testid="stHorizontalBlock"]:last-of-type{
  max-width:320px !important;
  margin:12px 0 0 40px !important;
  flex-wrap:nowrap !important;
  gap:8px !important;
}

@media (max-width: 520px){
  /* reduz o grupo no celular */
  div[data-testid="stForm"] div[data-testid="stHorizontalBlock"]:last-of-type{
    max-width:240px !important;
  }

  /* reduz o Gerar no celular */
  div[data-testid="stForm"] div[data-testid="stHorizontalBlock"]:last-of-type
  button[kind="primary"]{
    min-width:100px !important;
    padding:10px 14px !important;
  }
}

/* 🧹 Limpar (SECONDARY) */
div[data-testid="stForm"] div[data-testid="stHorizontalBlock"]:last-of-type
button[kind="secondary"]{
  min-width:40px !important;
  padding:8px 0 !important;
}

.small-gap{ margin-top:10px; }

/* --- Overlay do botão menu (SEM :has) --- */

/* tira a cara de botão */
div[data-testid="stButton"] > button{
  background:transparent !important;
  border:none !important;
  box-shadow:none !important;
}

/* o menu_btn vira um "overlay" fixo em cima do header */
div[data-testid="stButton"] > button#menu_btn{
  position: absolute !important;
  top:48px !important;   /* sobe/desce */
  left:calc(50% - 300px + 22px) !important; /* esquerda/direita */
  width: 45px !important;
  height: 45px !important;
  padding: 0 !important;
  font-size: 0 !important;
  z-index: 9999 !important;
}

</style>
""", unsafe_allow_html=True)

# ================= HEADER ALMG =================

if "menu_open" not in st.session_state:
    st.session_state.menu_open = False

# Wrapper com o mesmo visual do seu header (branco, largura 560, padding, radius)
st.markdown("""
<div id="almg_header" style='
margin:0 auto 20px auto;
background:white;
max-width:560px;
padding:10px 18px;
border-radius:12px;
'>
</div>
""", unsafe_allow_html=True)

# Agora o conteúdo REAL do header em colunas (fica no lugar certo sempre)
col_left, col_mid, col_right = st.columns([1.2, 6, 1.6])

with col_left:
    # ☰ vira botão de verdade (clicável)
    if st.button("☰", key="menu_btn"):
        st.session_state.menu_open = not st.session_state.menu_open

with col_mid:
    st.markdown("""
    <div style="display:flex; justify-content:center; align-items:center; height:48px;">
      <a href="https://www.almg.gov.br/" target="_blank">
        <img src="https://www.almg.gov.br/system/modules/br.gov.almg.portal/resources/img/logo/logo.svg"
             style="height:45px;">
      </a>
    </div>
    """, unsafe_allow_html=True)

with col_right:
    st.markdown("""
    <div style="display:flex; justify-content:flex-end; gap:10px; align-items:center; height:48px; font-size:30px;">
      <a href="https://silegis.almg.gov.br/silegismg/login/login.jsp#/processos"
         target="_blank" style="text-decoration:none;color:#cc0000;">🔍</a>
      <a href="https://intra.almg.gov.br/"
         target="_blank" style="text-decoration:none;color:#cc0000;">👤</a>
    </div>
    """, unsafe_allow_html=True)
    
# ================= HEADER =================
st.markdown(
    '<div class="title" style="font-size:24px; font-weight:1000; font-height:100;">GERÊNCIA DE GESTÃO ARQUIVÍSTICA</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle" style="font-size:16px; font-weight:1000;">MATE - MATÉRIAS EM TRAMITAÇÃO</div>',
    unsafe_allow_html=True
)

# ================= CARD =================

with st.form("form_mate", clear_on_submit=False):
    entrada = st.text_input(
        "Informe uma data do Diário do Legislativo",
        placeholder="Ex: 24/02/2026 ou https://...",
    )

    st.caption(
        "- 24022026 ou 240226 ou 24/02/2026\n"
        "- hoje, ontem, anteontem\n"
        "- terça, quarta, quinta, sexta, sábado"
    )

    st.markdown('<div class="small-gap"></div>', unsafe_allow_html=True)

    col1, col2 = st.columns([3,1], gap="small")
    with col1:
        rodar = st.form_submit_button("🚀 Gerar Planilha", type="primary")
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

    st.info(f"DATA = {entrada_clean!r}")

    try:
        progress_bar = st.progress(0)
        status_text = st.empty()

        progress_bar.progress(5)
        status_text.write("Inicializando… 5%")

        result = {"url": None, "aba": None}
        err = {"exc": None}
        done = threading.Event()

        def run_main():
            try:
                url, aba = main(
                    entrada_override=entrada_clean,
                    spreadsheet_url_or_id=st.secrets["SPREADSHEET_URL_OR_ID"],
                    auth_mode="service_account",
                    sa_info=st.secrets["gcp_service_account"],
                )
                result["url"] = url
                result["aba"] = aba
            except Exception as e:
                err["exc"] = e
            finally:
                done.set()

        t = threading.Thread(target=run_main, daemon=True)
        t.start()

        # Animação de progresso: sobe lentamente até 90% e fica “vivo”
        pct = 8
        while not done.is_set():
            pct = min(90, pct + 1)
            progress_bar.progress(pct)
            status_text.write(f"Processando Diário do Legislativo… {pct}%")
            time.sleep(0.10)

            # quando chega em 90, mantém “batendo” sem ficar parado
            if pct >= 90 and not done.is_set():
                for dots in (".", "..", "..."):
                    status_text.write(f"Processando Diário do Legislativo{dots} {pct}%")
                    time.sleep(0.35)
                    if done.is_set():
                        break

        # terminou: se houve erro na thread, explode aqui no principal
        if err["exc"] is not None:
            raise err["exc"]

        progress_bar.progress(100)
        status_text.write("Concluído 100%")

        st.success("")
        st.write("Aba:", result["aba"])
        st.link_button("Abrir planilha", result["url"], use_container_width=True)

    except Exception as e:
        st.error("Erro ao processar.")
        st.exception(e)