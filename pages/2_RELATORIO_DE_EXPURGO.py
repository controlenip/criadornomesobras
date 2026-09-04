import streamlit as st
import pandas as pd
import unicodedata
import os
import base64
import re

# ==========================================
# 1. CONFIGURAÇÕES DA PÁGINA E CSS
# ==========================================
st.set_page_config(page_title="Relatório de Expurgo", page_icon="📊", layout="centered")

st.markdown("""
<style>
    .block-container { padding-top: 1.5rem !important; padding-bottom: 2rem !important; max-width: 850px;}
    
    @media print {
        body * { visibility: hidden; }
        #area-impressao, #area-impressao * { visibility: visible !important; }
        #area-impressao { position: absolute; left: 0; top: 0; width: 100%; margin: 0; padding: 0;}
        header, [data-testid="stSidebar"], .painel-topo { display: none !important; }
        .st-emotion-cache-1z1uz4k { display: none !important; } 
    }
</style>
""", unsafe_allow_html=True)

# Logo da Equatorial em Base64 (Para não precisar de arquivo externo)
LOGO_EQTL_B64 = "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"

# ==========================================
# FUNÇÕES DE DADOS
# ==========================================
def formatar_data(data_raw):
    try:
        if pd.isna(data_raw) or str(data_raw).lower() in ['nan', 'none', '']: return ""
        return pd.to_datetime(data_raw).strftime('%d/%m/%Y')
    except:
        return str(data_raw)[:10]

@st.cache_data(show_spinner=False)
def carregar_dados(file_path, mtime):
    xls = pd.ExcelFile(file_path)
    try: df_notas = pd.read_excel(xls, sheet_name='NOTAS')
    except: 
        try: df_notas = pd.read_excel(xls, sheet_name='NotasSisgb')
        except: df_notas = pd.DataFrame()
        
    if not df_notas.empty:
        df_notas.columns = df_notas.columns.str.strip().str.upper()
        if 'PROTOCOLO' in df_notas.columns:
            df_notas['PROTOCOLO'] = df_notas['PROTOCOLO'].astype(str).replace(r'\.0$', '', regex=True).str.strip()
            
    return df_notas

arquivo_bd = "BASE_LEVANTAMENTO_ATUALIZADA.xlsx"
df_notas = pd.DataFrame()

if os.path.exists(arquivo_bd):
    mtime = os.path.getmtime(arquivo_bd) 
    df_notas = carregar_dados(arquivo_bd, mtime)
else:
    st.warning(f"⚠️ Planilha '{arquivo_bd}' não encontrada.")

# ==========================================
# PAINEL MINIMALISTA DE ENTRADA (TOPO)
# ==========================================
st.markdown('<div class="painel-topo">', unsafe_allow_html=True)
c1, c2 = st.columns([1.5, 1.5])
prot_input = c1.text_input("Nº da Nota / Protocolo:")
foto_upload = c2.file_uploader("Foto da Evidência:", type=['png', 'jpg', 'jpeg'])
st.markdown('</div>', unsafe_allow_html=True)

# Lógica de Cruzamento de Dados (Fiel às regras)
nota_val = prot_input.strip().upper() if prot_input else ""
regional, data_ab, cc, endereco, parceiro = "", "", "", "", ""

if nota_val and not df_notas.empty:
    res = df_notas[df_notas['PROTOCOLO'] == nota_val]
    if not res.empty:
        r = res.iloc[0]
        regional = str(r.get('REGIONAL', '')).upper().replace('"', '')
        if regional.lower() == 'nan': regional = ""
            
        data_ab = formatar_data(r.get('DATA ABERTURA', r.get('DATA DA SOLICITAÇÃO', '')))
        
        cc = str(r.get('CONTA CONTRATO', '')).replace('.0', '').replace('"', '')
        if cc.lower() == 'nan': cc = ""
            
        parceiro = str(r.get('NOME', '')).replace('"', '').upper()
        if parceiro.lower() == 'nan': parceiro = ""
            
        endereco_bruto = str(r.get('ENDEREÇO', '')).replace('"', '')
        cidade = str(r.get('MUNICIPIO', '')).replace('"', '')
        if endereco_bruto.lower() != 'nan':
            endereco = f"{endereco_bruto} - {cidade}".strip(" -").upper()

# Preparar Imagem
foto_base64 = ""
if foto_upload is not None:
    foto_base64 = base64.b64encode(foto_upload.read()).decode()
    img_html = f'<img src="data:image/png;base64,{foto_base64}" style="max-height: 280px; width: 100%; object-fit: contain;">'
else:
    img_html = '<span style="color: #ccc; font-style: italic; font-size: 11px;">Nenhuma imagem anexada</span>'

# Botão de imprimir discreto
st.markdown("""
<div class="painel-topo" style="text-align: right; margin-bottom: 5px;">
    <button onclick="window.print()" style="background-color: #f1f5f9; color: #0f172a; border: 1px solid #cbd5e1; padding: 4px 12px; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 11px;">🖨️ IMPRIMIR / PDF</button>
</div>
""", unsafe_allow_html=True)

# ==========================================
# HTML MINIFICADO PARA ALINHAMENTO MATEMÁTICO PERFEITO
# ==========================================
html_form = f"""
<style>
.form-container {{ width: 100%; max-width: 800px; margin: 0 auto; font-family: Arial, sans-serif; font-size: 11px; color: black; background: white; }}
.header-box {{ background-color: #003399; color: white; display: flex; align-items: center; height: 50px; padding: 0 15px; border: 1.5px solid black; margin-bottom: 12px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
.header-logo {{ border-right: 1.5px solid white; padding-right: 15px; margin-right: 15px; height: 35px; display: flex; align-items: center; justify-content: center; }}
.header-title {{ font-size: 16px; margin: 0; text-align: center; flex-grow: 1; margin-left: -80px; font-weight: bold; }}

.section-title {{ background-color: #1b365d; color: white; font-weight: bold; padding: 4px 8px; border: 1.5px solid black; margin-bottom: 4px; font-size: 11px; -webkit-print-color-adjust: exact; print-color-adjust: exact; line-height: 1; }}

.t-table {{ width: 100%; border-collapse: collapse; margin-bottom: 6px; table-layout: fixed; }}
.t-table td {{ border: 1.5px solid black; vertical-align: middle; height: 26px; padding: 0; margin: 0; box-sizing: border-box; }}
.no-border {{ border: none !important; }}

.lbl {{ background-color: white; font-weight: bold; text-align: center; color: black; font-size: 11px; }}
.val {{ background-color: white; }}

.fi {{ width: 100%; height: 100%; border: none; outline: none; background: transparent; font-family: Arial, sans-serif; font-size: 11px; font-weight: bold; color: black; text-align: center; text-transform: uppercase; padding: 0; margin: 0; display: block; box-sizing: border-box; }}
.bg-blue {{ background-color: #cbe0f5 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
</style>

<div id="area-impressao" class="form-container">

    <!-- CABEÇALHO AZUL COM A LOGO EMBUTIDA -->
    <div class="header-box">
        <div class="header-logo"><img src="data:image/png;base64,{LOGO_EQTL_B64}" style="height: 30px;"></div>
        <div class="header-title">Formulário de Não Atendimento Expansão</div>
    </div>

    <!-- DISTRIBUIDORA / REGIONAL -->
    <table class="t-table" style="margin-bottom: 12px;">
      <colgroup><col style="width: 14%;"><col style="width: 14%;"><col style="width: 20%;"><col style="width: 14%;"><col style="width: 14%;"><col style="width: 6%;"><col style="width: 9%;"><col style="width: 9%;"></colgroup>
      <tr>
        <td class="lbl">Distribuidora:</td>
        <td class="val"><input type="text" class="fi" value="EQTL MA"></td>
        <td class="no-border"></td>
        <td class="lbl">Regional:</td>
        <td class="val"><input type="text" class="fi" value="{regional}"></td>
        <td class="no-border"></td>
        <td class="lbl" style="font-size: 8px; line-height: 1.1;">Data da<br>solicitação:</td>
        <td class="val"><input type="text" class="fi" value="{data_ab}"></td>
      </tr>
    </table>

    <!-- DADOS DO CLIENTE -->
    <div class="section-title">Dados do Cliente:</div>
    <table class="t-table">
      <colgroup><col style="width: 18%;"><col style="width: 30%;"><col style="width: 4%;"><col style="width: 18%;"><col style="width: 30%;"></colgroup>
      <tr>
        <td class="lbl">Nº da nota:</td>
        <td class="val"><input type="text" class="fi" value="{nota_val}"></td>
        <td class="no-border"></td>
        <td class="lbl">Conta Contrato:</td>
        <td class="val"><input type="text" class="fi" value="{cc}"></td>
      </tr>
    </table>
    <table class="t-table">
      <colgroup><col style="width: 18%;"><col style="width: 82%;"></colgroup>
      <tr>
        <td class="lbl">Parceiro de Negócios:</td>
        <td class="val"><input type="text" class="fi" value="{parceiro}"></td>
      </tr>
    </table>
    <table class="t-table" style="margin-bottom: 12px;">
      <colgroup><col style="width: 18%;"><col style="width: 82%;"></colgroup>
      <tr>
        <td class="lbl">Endereço:</td>
        <td class="val"><input type="text" class="fi" value="{endereco}"></td>
      </tr>
    </table>

    <!-- DADOS DA VISITA -->
    <div class="section-title">Dados da Visita:</div>
    <table class="t-table">
      <colgroup><col style="width: 18%;"><col style="width: 30%;"><col style="width: 4%;"><col style="width: 18%;"><col style="width: 30%;"></colgroup>
      <tr>
        <td class="lbl">Data:</td>
        <td class="val"><input type="text" class="fi" value=""></td>
        <td class="no-border"></td>
        <td class="lbl">Latitude:</td>
        <td class="val"><input type="text" class="fi" value=""></td>
      </tr>
    </table>
    <table class="t-table">
      <colgroup><col style="width: 18%;"><col style="width: 30%;"><col style="width: 4%;"><col style="width: 18%;"><col style="width: 30%;"></colgroup>
      <tr>
        <td class="lbl">Horário:</td>
        <td class="val"><input type="text" class="fi" value=""></td>
        <td class="no-border"></td>
        <td class="lbl">Longitude:</td>
        <td class="val"><input type="text" class="fi" value=""></td>
      </tr>
    </table>
    <table class="t-table" style="margin-bottom: 12px;">
      <colgroup><col style="width: 18%;"><col style="width: 82%;"></colgroup>
      <tr>
        <td class="lbl" style="font-size: 10px; line-height: 1.1;">Identificação da<br>equipe:</td>
        <td class="val"><input type="text" class="fi" value="EQP NIP"></td>
      </tr>
    </table>

    <!-- MOTIVO DO EXPURGO -->
    <div class="section-title">Motivo do expurgo:</div>
    <table class="t-table">
      <colgroup><col style="width: 18%;"><col style="width: 82%;"></colgroup>
      <tr>
        <td class="lbl">Justificativa:</td>
        <td class="val"><input type="text" class="fi" value=""></td>
      </tr>
    </table>
    <table class="t-table">
      <colgroup><col style="width: 18%;"><col style="width: 82%;"></colgroup>
      <tr>
        <td class="lbl">Descrição do Expurgo:</td>
        <td class="val"><input type="text" class="fi" value=""></td>
      </tr>
    </table>
    <table class="t-table" style="margin-bottom: 12px;">
      <colgroup><col style="width: 18%;"><col style="width: 82%;"></colgroup>
      <tr>
        <td class="lbl bg-blue" style="font-size: 10px; line-height: 1.1;">Tratativa no Sistema<br>Comercial:</td>
        <td class="val bg-blue"><input type="text" class="fi bg-blue" value=""></td>
      </tr>
    </table>

    <!-- EVIDÊNCIAS -->
    <div class="section-title">Evidências:</div>
    <table class="t-table">
      <colgroup><col style="width: 25%;"><col style="width: 23%;"><col style="width: 4%;"><col style="width: 25%;"><col style="width: 23%;"></colgroup>
      <tr>
        <td class="lbl" style="font-size: 10px; line-height: 1.2;">Número do medidor<br>do cliente atendido:</td>
        <td class="val"><input type="text" class="fi" value=""></td>
        <td class="no-border"></td>
        <td class="lbl" style="font-size: 10px; line-height: 1.2;">Número da nota do<br>atendimento em campo:</td>
        <td class="val"><input type="text" class="fi" value="{nota_val}"></td>
      </tr>
    </table>
    <table class="t-table" style="margin-bottom: 8px;">
      <colgroup><col style="width: 25%;"><col style="width: 23%;"><col style="width: 4%;"><col style="width: 25%;"><col style="width: 23%;"></colgroup>
      <tr>
        <td class="lbl" style="font-size: 10px; line-height: 1.2;">Número do medidor<br>do vizinho:</td>
        <td class="val"><input type="text" class="fi" value=""></td>
        <td class="no-border"></td>
        <td class="lbl" style="font-size: 10px; line-height: 1.2;">Número da estrutura<br>mais próxima:</td>
        <td class="val"><input type="text" class="fi" value=""></td>
      </tr>
    </table>

    <!-- FOTO -->
    <div style="border: 1.5px solid black; width: 100%; height: 320px; display: flex; align-items: center; justify-content: center; background-color: white; overflow: hidden; margin-top: 5px;">
        {img_html}
    </div>

</div>
"""

# Destrói TODOS os espaços invisíveis antes de injetar no Streamlit
html_form = "".join([line.strip() for line in html_form.split('\n')])
st.markdown(html_form, unsafe_allow_html=True)
