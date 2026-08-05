import streamlit as st

st.set_page_config(page_title="LinkedIn Report App", layout="wide")

st.title("LinkedIn Report App")
st.write("L'app è online.")

st.header("Caricamento report")
uploaded_file = st.file_uploader("Carica il PDF del report", type=["pdf"])

if uploaded_file is not None:
    st.success(f"File caricato: {uploaded_file.name}")
    st.info("Qui poi puoi aggiungere la lettura e l'analisi del report.")
