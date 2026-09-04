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
c1, c2 = st.columns([1.5, 1.5])
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
    img_html = f'<img src="data:image/png;base64,{foto_base64}" style="max-height: 250px; width: 100%; object-fit: contain;">'
else:
    img_html = '<span style="color: #ccc; font-style: italic;">Nenhuma imagem anexada</span>'

# Botão de imprimir discreto
st.markdown("""
<div class="painel-topo" style="text-align: right; margin-bottom: 5px;">
    <button onclick="window.print()" style="background-color: #f1f5f9; color: #0f172a; border: 1px solid #cbd5e1; padding: 4px 12px; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 11px;">🖨️ IMPRIMIR / PDF</button>
</div>
""", unsafe_allow_html=True)

# ==========================================
# CONSTRUÇÃO DO HTML (FORMULÁRIO ESTRITO)
# ==========================================
html_form = f"""
<style>
.form-container {{ width: 100%; max-width: 800px; margin: 0 auto; font-family: Arial, sans-serif; font-size: 11px; color: black; background: white; }}
.header-box {{ background-color: #1b365d; color: white; display: flex; align-items: center; height: 50px; padding: 0 15px; border: 1.5px solid black; margin-bottom: 10px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
.header-logo {{ font-weight: bold; font-size: 13px; border-right: 1.5px solid white; padding-right: 15px; margin-right: 15px; line-height: 1; }}
.header-title {{ font-size: 16px; margin: 0; text-align: center; flex-grow: 1; margin-left: -80px; font-weight: bold; }}

.section-title {{ background-color: #1b365d; color: white; font-weight: bold; padding: 4px 8px; border: 1.5px solid black; margin-bottom: 4px; font-size: 11px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}

.t-table {{ width: 100%; border-collapse: collapse; margin-bottom: 8px; table-layout: fixed; }}
.t-table td {{ border: 1.5px solid black; vertical-align: middle; height: 26px; overflow: hidden; padding: 0; margin: 0; box-sizing: border-box; }}

.celula-label {{ font-weight: bold; text-align: center; background-color: white; padding: 0 4px !important; }}
.celula-label-esq {{ font-weight: bold; text-align: left; background-color: white; padding: 0 4px !important; font-size: 9px; line-height: 1.1; }}

.celula-valor {{ text-align: center; }}
.celula-valor-esq {{ text-align: left; }}

.fi {{ width: 100%; height: 100%; border: none; outline: none; background: transparent; font-family: Arial, sans-serif; font-size: 11px; font-weight: bold; color: black; text-align: center; text-transform: uppercase; padding: 0; margin: 0; display: block; }}
.fi-esq {{ width: 100%; height: 100%; border: none; outline: none; background: transparent; font-family: Arial, sans-serif; font-size: 11px; font-weight: bold; color: black; text-align: left; text-transform: uppercase; padding: 0 4px; margin: 0; display: block; }}
.fa {{ width: 100%; border: none; outline: none; background: transparent; font-family: Arial, sans-serif; font-size: 11px; font-weight: bold; color: black; resize: none; min-height: 26px; text-transform: uppercase; text-align: center; padding: 4px 0; margin: 0; display: block; }}

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
    <table class="t-table">
      <tr>
        <td class="celula-label" style="width: 15%;">Distribuidora:</td>
        <td class="celula-valor" style="width: 15%;"><input type="text" class="fi" value="EQTL MA"></td>
        <td class="no-border" style="width: 20%;"></td>
        <td class="celula-label" style="width: 15%;">Regional:</td>
        <td class="celula-valor" style="width: 15%;"><input type="text" class="fi" value="{regional}"></td>
        <td class="no-border" style="width: 5%;"></td>
        <td class="celula-label" style="width: 10%; font-size: 9px; line-height: 1.1;">Data da<br>solicitação:</td>
        <td class="celula-valor" style="width: 10%;"><input type="text" class="fi" value="{data_ab}"></td>
      </tr>
    </table>

    <!-- DADOS DO CLIENTE -->
    <div class="section-title">Dados do Cliente:</div>
    <table class="t-table" style="margin-bottom: 4px;">
      <tr>
        <td class="celula-label" style="width: 15%;">Nº da nota:</td>
        <td class="celula-valor" style="width: 40%;"><input type="text" class="fi" value="{nota_val}"></td>
        <td class="no-border" style="width: 10%;"></td>
        <td class="celula-label" style="width: 15%;">Conta Contrato:</td>
        <td class="celula-valor" style="width: 20%;"><input type="text" class="fi" value="{cc}"></td>
      </tr>
    </table>
    <table class="t-table" style="margin-bottom: 4px;">
      <tr>
        <td class="celula-label-esq" style="width: 20%; font-size: 11px;">Parceiro de Negócios:</td>
        <td class="celula-valor" style="width: 80%;"><input type="text" class="fi" value="{parceiro}"></td>
      </tr>
    </table>
    <table class="t-table">
      <tr>
        <td class="celula-label" style="width: 15%;">Endereço:</td>
        <td class="celula-valor" style="width: 85%;"><input type="text" class="fi" value="{endereco}"></td>
      </tr>
    </table>

    <!-- DADOS DA VISITA -->
    <div class="section-title">Dados da Visita:</div>
    <table class="t-table" style="margin-bottom: 4px;">
      <tr>
        <td class="celula-label" style="width: 15%;">Data:</td>
        <td class="celula-valor" style="width: 40%;"><input type="text" class="fi" value=""></td>
        <td class="no-border" style="width: 10%;"></td>
        <td class="celula-label" style="width: 15%;">Latitude:</td>
        <td class="celula-valor" style="width: 20%;"><input type="text" class="fi" value=""></td>
      </tr>
    </table>
    <table class="t-table" style="margin-bottom: 4px;">
      <tr>
        <td class="celula-label" style="width: 15%;">Horário:</td>
        <td class="celula-valor" style="width: 40%;"><input type="text" class="fi" value=""></td>
        <td class="no-border" style="width: 10%;"></td>
        <td class="celula-label" style="width: 15%;">Longitude:</td>
        <td class="celula-valor" style="width: 20%;"><input type="text" class="fi" value=""></td>
      </tr>
    </table>
    <table class="t-table">
      <tr>
        <td class="celula-label-esq" style="width: 20%; font-size: 11px;">Identificação da<br>equipe:</td>
        <td class="celula-valor" style="width: 80%;"><input type="text" class="fi" value="EQP NIP"></td>
      </tr>
    </table>

    <!-- MOTIVO DO EXPURGO -->
    <div class="section-title">Motivo do expurgo:</div>
    <table class="t-table" style="margin-bottom: 4px;">
      <tr>
        <td class="celula-label-esq" style="width: 20%; font-size: 11px;">Justificativa:</td>
        <td class="celula-valor" style="width: 80%;"><input type="text" class="fi" value=""></td>
      </tr>
    </table>
    <table class="t-table" style="margin-bottom: 4px;">
      <tr>
        <td class="celula-label-esq" style="width: 20%; font-size: 11px;">Descrição do Expurgo:</td>
        <td class="celula-valor" style="width: 80%;"><textarea class="fa"></textarea></td>
      </tr>
    </table>
    <table class="t-table">
      <tr>
        <td class="celula-label-esq bg-blue" style="width: 20%; font-size: 11px;">Tratativa no Sistema<br>Comercial:</td>
        <td class="celula-valor bg-blue" style="width: 80%;"><input type="text" class="fi bg-blue" value=""></td>
      </tr>
    </table>

    <!-- EVIDÊNCIAS -->
    <div class="section-title">Evidências:</div>
    <table class="t-table">
      <tr>
        <td class="celula-label-esq" style="width: 25%;">Número do medidor<br>do cliente atendido:</td>
        <td class="celula-valor" style="width: 23%;"><input type="text" class="fi" value=""></td>
        <td class="no-border" style="width: 4%;"></td>
        <td class="celula-label-esq" style="width: 25%;">Número da nota do<br>atendimento em campo:</td>
        <td class="celula-valor" style="width: 23%;"><input type="text" class="fi" value="{nota_val}"></td>
      </tr>
      <tr>
        <td class="celula-label-esq" style="width: 25%; border-top: none;">Número do medidor<br>do vizinho:</td>
        <td class="celula-valor" style="width: 23%; border-top: none;"><input type="text" class="fi" value=""></td>
        <td class="no-border" style="width: 4%;"></td>
        <td class="celula-label-esq" style="width: 25%; border-top: none;">Número da estrutura<br>mais próxima:</td>
        <td class="celula-valor" style="width: 23%; border-top: none;"><input type="text" class="fi" value=""></td>
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
