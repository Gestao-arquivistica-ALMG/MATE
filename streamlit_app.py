import streamlit as st
from mate_pipeline import main
from datetime import datetime

st.set_page_config(page_title="MATE", layout="centered")
st.title("MATE.IA")

entrada = st.text_input("Digite a data do Diário do Legislativo.", key="entrada")

col1, col2 = st.columns([1, 1])
with col1:
    rodar = st.button("Gerar planilha", type="primary", key="rodar")
with col2:
    limpar = st.button("Limpar", key="limpar")

if limpar:
    st.session_state.clear()
    st.rerun()

if rodar:
    print("=== CLICK RODAR ===", datetime.now().isoformat())
    print("entrada:", repr(entrada))

    if not entrada.strip():
        st.error("Informe uma data/palavra/URL/caminho.")
        st.stop()

    # valida secrets ANTES do main (pra não travar e você achar que foi o main)
    print("tem SPREADSHEET_URL_OR_ID?", "SPREADSHEET_URL_OR_ID" in st.secrets)
    print("tem gcp_service_account?", "gcp_service_account" in st.secrets)

    st.info("Entrou no handler. Chamando main()...")
    print("ANTES DO MAIN", datetime.now().isoformat())

    try:
        with st.spinner("Processando..."):
            url, aba = main(
                entrada_override=entrada.strip(),
                spreadsheet_url_or_id=st.secrets["SPREADSHEET_URL_OR_ID"],
                auth_mode="service_account",
                sa_info=st.secrets["gcp_service_account"],
            )
    except Exception as e:
        print("ERRO NO MAIN:", repr(e))
        st.exception(e)
        raise

    print("DEPOIS DO MAIN", datetime.now().isoformat())

    st.success("Concluído.")
    st.write("Aba:", aba)
    st.link_button("Abrir planilha", url)