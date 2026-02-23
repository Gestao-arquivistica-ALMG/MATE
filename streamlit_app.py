import streamlit as st
from mate_pipeline import main

st.set_page_config(page_title="MATE", layout="centered")
st.title("MATE.IA")

entrada = st.text_input("Data (DDMMYYYY), palavra (hoje/ontem/sabado), URL ou caminho")
    print("Digite a data do Diário do Legislativo.")
    print("EXEMPLOS:")
    print("- 19122026 ou 191226 ou 19/12/2026")
    print("- hoje, ontem ou anteontem")
    print("- terça, quarta, quinta, sexta ou sábado")
    print("- URL ou caminho local")
    print("Se deixar vazio, você poderá fazer upload.\n")

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
