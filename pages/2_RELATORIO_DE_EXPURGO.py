import streamlit as st
import pandas as pd
import unicodedata
import os
import base64
import re

# ==========================================
# 1. CONFIGURAÇÕES DA PÁGINA
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
c1, c2, c3 = st.columns([1.5, 1.5, 0.8])
prot_input = c1.text_input("Nº da Nota / Protocolo:")
foto_upload = c2.file_uploader("Foto da Evidência:", type=['png', 'jpg', 'jpeg'])
st.markdown('</div>', unsafe_allow_html=True)

# Lógica de Cruzamento de Dados
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
    img_html = f'<img src="data:image/png;base64,{foto_base64}" style="max-height: 250px; width: auto; object-fit: contain;">'
else:
    img_html = '<span style="color: #ccc; font-style: italic;">Nenhuma imagem anexada</span>'

# Botão de imprimir discreto
st.markdown("""
<div class="painel-topo" style="text-align: right; margin-bottom: 5px;">
    <button onclick="window.print()" style="background-color: #f1f5f9; color: #0f172a; border: 1px solid #cbd5e1; padding: 4px 12px; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 11px;">🖨️ IMPRIMIR / PDF</button>
</div>
""", unsafe_allow_html=True)

# ==========================================
# CONSTRUÇÃO DO HTML (FORMULÁRIO IDÊNTICO)
# ==========================================
html_form = f"""
<style>
.form-container {{ width: 100%; max-width: 800px; margin: 0 auto; font-family: Arial, sans-serif; font-size: 12px; color: black; background: white; }}
.header-box {{ background-color: #1b365d; color: white; display: flex; align-items: center; height: 50px; padding: 0 15px; border: 1.5px solid black; margin-bottom: 12px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
.header-logo {{ font-weight: bold; font-size: 13px; border-right: 1.5px solid white; padding-right: 15px; margin-right: 15px; line-height: 1; }}
.header-title {{ font-size: 18px; margin: 0; text-align: center; flex-grow: 1; margin-left: -80px; font-weight: bold; }}
.section-title {{ background-color: #1b365d; color: white; font-weight: bold; padding: 4px 8px; border: 1.5px solid black; margin-bottom: 5px; font-size: 12px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
.t-table {{ width: 100%; border-collapse: collapse; margin-bottom: 5px; }}
.t-table td {{ border: 1.5px solid black; padding: 4px 5px; vertical-align: middle; height: 26px; }}
.lbl {{ font-weight: bold; text-align: center; font-size: 11px; background-color: white; }}
.val {{ padding: 0 !important; }}
.fi {{ width: 100%; height: 100%; border: none; outline: none; background: transparent; font-family: Arial, sans-serif; font-size: 11px; font-weight: bold; color: black; padding: 4px; box-sizing: border-box; text-transform: uppercase; text-align: center; }}
.fa {{ width: 100%; border: none; outline: none; background: transparent; font-family: Arial, sans-serif; font-size: 11px; font-weight: bold; color: black; padding: 4px; box-sizing: border-box; resize: none; overflow: hidden; min-height: 26px; text-transform: uppercase; text-align: center; }}
.bg-blue {{ background-color: #cbe0f5 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
.no-border {{ border: none !important; }}
</style>

<div id="area-impressao" class="form-container">

    <!-- CABEÇALHO AZUL -->
    <div class="header-box">
        <div class="header-logo">GRUPO<br><span style="font-size: 16px;">equatorial</span><br><span style="font-size: 7px; font-weight: normal; letter-spacing: 1px;">ENERGIA</span></div>
        <div class="header-title">Formulário de Não Atendimento Expansão</div>
    </div>

    <!-- DISTRIBUIDORA / REGIONAL -->
    <table class="t-table" style="margin-bottom: 15px;">
      <tr>
        <td class="lbl" style="width: 15%;">Distribuidora:</td>
        <td class="val" style="width: 15%;"><input type="text" class="fi" value="EQTL MA"></td>
        <td class="no-border" style="width: 18%;"></td>
        <td class="lbl" style="width: 15%;">Regional:</td>
        <td class="val" style="width: 15%;"><input type="text" class="fi" value="{regional}"></td>
        <td class="no-border" style="width: 2%;"></td>
        <td class="lbl" style="width: 10%; font-size: 9px; line-height: 1.1; padding: 2px;">Data da<br>solicitação:</td>
        <td class="val" style="width: 10%;"><input type="text" class="fi" value="{data_ab}"></td>
      </tr>
    </table>

    <!-- DADOS DO CLIENTE -->
    <div class="section-title">Dados do Cliente:</div>
    <table class="t-table">
      <tr>
        <td class="lbl" style="width: 15%;">Nº da nota:</td>
        <td class="val" style="width: 35%;"><input type="text" class="fi" value="{nota_val}"></td>
        <td class="no-border" style="width: 10%;"></td>
        <td class="lbl" style="width: 15%;">Conta Contrato:</td>
        <td class="val" style="width: 25%;"><input type="text" class="fi" value="{cc}"></td>
      </tr>
    </table>
    <table class="t-table">
      <tr>
        <td class="lbl" style="width: 20%;">Parceiro de Negócios:</td>
        <td class="val" style="width: 80%;"><input type="text" class="fi" value="{parceiro}"></td>
      </tr>
    </table>
    <table class="t-table" style="margin-bottom: 15px;">
      <tr>
        <td class="lbl" style="width: 15%;">Endereço:</td>
        <td class="val" style="width: 85%;"><input type="text" class="fi" value="{endereco}"></td>
      </tr>
    </table>

    <!-- DADOS DA VISITA -->
    <div class="section-title">Dados da Visita:</div>
    <table class="t-table">
      <tr>
        <td class="lbl" style="width: 15%;">Data:</td>
        <td class="val" style="width: 35%;"><input type="text" class="fi" value=""></td>
        <td class="no-border" style="width: 10%;"></td>
        <td class="lbl" style="width: 15%;">Latitude:</td>
        <td class="val" style="width: 25%;"><input type="text" class="fi" value=""></td>
      </tr>
    </table>
    <table class="t-table">
      <tr>
        <td class="lbl" style="width: 15%;">Horário:</td>
        <td class="val" style="width: 35%;"><input type="text" class="fi" value=""></td>
        <td class="no-border" style="width: 10%;"></td>
        <td class="lbl" style="width: 15%;">Longitude:</td>
        <td class="val" style="width: 25%;"><input type="text" class="fi" value=""></td>
      </tr>
    </table>
    <table class="t-table" style="margin-bottom: 15px;">
      <tr>
        <td class="lbl" style="width: 25%;">Identificação da equipe:</td>
        <td class="val" style="width: 75%;"><input type="text" class="fi" value="EQP NIP"></td>
      </tr>
    </table>

    <!-- MOTIVO DO EXPURGO -->
    <div class="section-title">Motivo do expurgo:</div>
    <table class="t-table">
      <tr>
        <td class="lbl" style="width: 20%;">Justificativa:</td>
        <td class="val" style="width: 80%;"><input type="text" class="fi" value=""></td>
      </tr>
    </table>
    <table class="t-table">
      <tr>
        <td class="lbl" style="width: 20%;">Descrição do Expurgo:</td>
        <td class="val" style="width: 80%; padding: 0;"><textarea class="fa"></textarea></td>
      </tr>
    </table>
    <table class="t-table" style="margin-bottom: 15px;">
      <tr>
        <td class="lbl bg-blue" style="width: 20%;">Tratativa no Sistema<br>Comercial:</td>
        <td class="val bg-blue" style="width: 80%;"><input type="text" class="fi bg-blue" value=""></td>
      </tr>
    </table>

    <!-- EVIDÊNCIAS -->
    <div class="section-title">Evidências:</div>
    <table class="t-table" style="margin-bottom: 10px;">
      <tr>
        <td class="lbl" style="width: 23%; text-align: left; font-size: 10px; padding-left: 8px;">Número do medidor<br>do cliente atendido:</td>
        <td class="val" style="width: 25%;"><input type="text" class="fi" value=""></td>
        <td class="no-border" style="width: 4%;"></td>
        <td class="lbl" style="width: 23%; text-align: left; font-size: 10px; padding-left: 8px;">Número da nota do<br>atendimento em campo:</td>
        <td class="val" style="width: 25%;"><input type="text" class="fi" value="{nota_val}"></td>
      </tr>
      <tr>
        <td class="lbl" style="width: 23%; text-align: left; font-size: 10px; border-top: none; padding-left: 8px;">Número do medidor<br>do vizinho:</td>
        <td class="val" style="width: 25%; border-top: none;"><input type="text" class="fi" value=""></td>
        <td class="no-border" style="width: 4%;"></td>
        <td class="lbl" style="width: 23%; text-align: left; font-size: 10px; border-top: none; padding-left: 8px;">Número da estrutura<br>mais próxima:</td>
        <td class="val" style="width: 25%; border-top: none;"><input type="text" class="fi" value=""></td>
      </tr>
    </table>

    <!-- FOTO -->
    <div style="border: 1.5px solid black; width: 100%; height: 260px; display: flex; align-items: center; justify-content: center; background-color: white; overflow: hidden; margin-top: 5px;">
        {img_html}
    </div>

</div>
"""

html_form = re.sub(r'^[ \t]+', '', html_form, flags=re.MULTILINE)
st.markdown(html_form, unsafe_allow_html=True)
