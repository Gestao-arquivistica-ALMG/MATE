import streamlit as st
from mate_pipeline import main

st.set_page_config(page_title="MATE", layout="centered")
st.title("MATE — Assembleia Inteligente")

entrada = st.text_input("Data (DDMMYYYY), palavra (hoje/ontem/sabado), URL ou caminho")

col1, col2 = st.columns([1, 1])
with col1:
    rodar = st.button("Gerar planilha", type="primary")
with col2:
    limpar = st.button("Limpar")

if limpar:
    st.session_state.clear()
    st.rerun()

if rodar:
    if not entrada.strip():
        st.error("Informe uma data/palavra/URL/caminho.")
        st.stop()

    with st.spinner("Processando..."):
        url, aba = main(
            entrada_override=entrada.strip(),
            spreadsheet_url_or_id=st.secrets["SPREADSHEET_URL_OR_ID"],
            auth_mode="service_account",
            sa_info=st.secrets["gcp_service_account"],
        )

    st.success("Concluído.")
    st.write("Aba:", aba)
    st.link_button("Abrir planilha", url)
