import streamlit as st
import pandas as pd
import os
import base64

# ==========================================
# 1. CONFIGURAÇÕES DA PÁGINA
# ==========================================
st.set_page_config(page_title="Relatório de Expurgo", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 1rem !important; padding-bottom: 2rem !important; }
    html, body, [class*="css"] { font-size: 12px !important; }
    .eh-dark { background-color: #047857; color: white; font-weight: 700; padding: 10px; border-radius: 4px; text-align: center; text-transform: uppercase;}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. LOGO NO TOPO
# ==========================================
st.markdown("<br>", unsafe_allow_html=True) 
if os.path.exists("LOGO_NIP.png"):
    with open("LOGO_NIP.png", "rb") as image_file:
        b64_logo = base64.b64encode(image_file.read()).decode()
    
    st.markdown(f'''
        <div style="text-align: center; margin-bottom: 20px;">
            <img src="data:image/png;base64,{b64_logo}" style="max-width: 150px; width: 100%; height: auto; pointer-events: none;">
        </div>
    ''', unsafe_allow_html=True)

# ==========================================
# 3. CABEÇALHO DA PÁGINA
# ==========================================
st.markdown('<div class="eh-dark">📊 Relatório de Expurgo</div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

st.info("Página de Relatório de Expurgo criada com sucesso e logomarca inserida. Aguardando a lógica de filtragem dos dados.")

# Aqui entrará a lógica para ler a aba NOTAS e DADOS e gerar o relatório
# ...
