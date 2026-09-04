import streamlit as st
import pandas as pd
import unicodedata
import os
import base64
import datetime

# ==========================================
# 1. CONFIGURAÇÕES DA PÁGINA E CSS
# ==========================================
st.set_page_config(page_title="Relatório de Expurgo", page_icon="📊", layout="centered")

st.markdown("""
<style>
    .block-container { padding-top: 1rem !important; padding-bottom: 2rem !important; max-width: 900px;}
    
    /* Configurações perfeitas para esconder painéis e focar só no formulário na Impressão */
    @media print {
        body * { visibility: hidden; }
        #area-impressao, #area-impressao * { visibility: visible !important; }
        #area-impressao { position: absolute; left: 0; top: 0; width: 100%; margin: 0; padding: 0;}
        header, [data-testid="stSidebar"], .painel-controle, .print-btn { display: none !important; }
        .st-emotion-cache-1z1uz4k { display: none !important; } 
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# FUNÇÕES DE LEITURA E DADOS
# ==========================================
def formatar_data(data_raw):
    try:
        if pd.isna(data_raw) or str(data_raw).lower() == 'nan' or str(data_raw).strip() == "": return ""
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
            
    try: df_dados = pd.read_excel(xls, sheet_name='DADOS', header=1)
    except:
        try: df_dados = pd.read_excel(xls, sheet_name='Dados', header=1)
        except: df_dados = pd.DataFrame()
        
    if not df_dados.empty:
        df_dados.columns = df_dados.columns.str.strip().str.upper()
    
    return df_notas, df_dados

arquivo_bd = "BASE_LEVANTAMENTO_ATUALIZADA.xlsx"
df_notas, df_dados = pd.DataFrame(), pd.DataFrame()

if os.path.exists(arquivo_bd):
    mtime = os.path.getmtime(arquivo_bd) 
    df_notas, df_dados = carregar_dados(arquivo_bd, mtime)
else:
    st.warning(f"⚠️ Planilha '{arquivo_bd}' não encontrada.")

# ==========================================
# PAINEL DE CONTROLE (BUSCA E FOTO - TOPO)
# ==========================================
st.markdown('<div class="painel-controle">', unsafe_allow_html=True)
st.markdown('<div style="background-color: #047857; color: white; font-weight: bold; padding: 10px; border-radius: 4px; text-align: center; margin-bottom: 15px;">⚙️ 1º PASSO: BUSCAR DADOS E ANEXAR FOTO (Não sai na impressão)</div>', unsafe_allow_html=True)

c1, c2 = st.columns([1, 1])
prot_input = c1.text_input("🔍 Digite a Nota/Protocolo para puxar os dados:", placeholder="Ex: 1113804258")
foto_upload = c2.file_uploader("📸 Anexe a Foto da Evidência AQUI:", type=['png', 'jpg', 'jpeg'])

st.info("⚠️ **2º PASSO:** Role a tela para baixo, clique nas linhas em branco do formulário e digite as informações (Justificativa, Horário, Medidor, etc) **diretamente nele**.")
st.markdown('</div>', unsafe_allow_html=True)

# Lógica de Cruzamento
regional, data_ab, cc, endereco, lat, lon, parceiro = "", "", "", "", "", "", ""
if prot_input and not df_notas.empty:
    res = df_notas[df_notas['PROTOCOLO'] == prot_input.strip()]
    if not res.empty:
        r = res.iloc[0]
        regional = str(r.get('REGIONAL', '')).upper().replace('"', '')
        data_ab = formatar_data(r.get('DATA ABERTURA', r.get('DATA DA SOLICITAÇÃO', '')))
        cc = str(r.get('CONTA CONTRATO', '')).replace('.0', '').replace('"', '')
        if cc.lower() == 'nan': cc = ""
        
        endereco_bruto = str(r.get('ENDEREÇO', '')).replace('"', '')
        cidade = str(r.get('MUNICIPIO', '')).replace('"', '')
        endereco = f"{endereco_bruto} - {cidade}".strip(" -") if endereco_bruto.lower() != 'nan' else ""
        
        lat = str(r.get('LATITUDE', '')).replace('"', '')
        if lat.lower() == 'nan': lat = ""
        lon = str(r.get('LONGITUDE', '')).replace('"', '')
        if lon.lower() == 'nan': lon = ""
        
        pi_auto = str(r.get('TIPO LIGAÇÃO', r.get('TIPO NOTA', ''))).strip().upper()
        if pi_auto and not df_dados.empty and 'PI' in df_dados.columns:
            d_pi = df_dados[df_dados['PI'] == pi_auto]
            if not d_pi.empty:
                col_idx = 0
                if "SUL" in regional: col_idx = 14
                elif "CENTRO" in regional: col_idx = 19
                elif "LESTE" in regional: col_idx = 24
                elif "NORTE" in regional: col_idx = 4
                elif "NOROESTE" in regional: col_idx = 9
                
                if col_idx > 0:
                    empresa_busca = str(d_pi.iloc[0, col_idx + 2]).replace('"', '')
                    parceiro = empresa_busca if empresa_busca.lower() != 'nan' else ""
        st.success("✅ Dados encontrados com sucesso!")
    else:
        st.error("❌ Nota não encontrada.")

# Tratamento da Imagem
foto_base64 = ""
if foto_upload is not None:
    foto_base64 = base64.b64encode(foto_upload.read()).decode()
    img_html = f'<img src="data:image/png;base64,{foto_base64}" style="max-height: 400px; max-width: 100%; object-fit: contain;">'
else:
    img_html = '<span style="color: #ccc; font-style: italic;">Nenhuma imagem anexada</span>'

hoje = datetime.date.today().strftime('%d/%m/%Y')

# ==========================================
# RENDERIZAÇÃO DO FORMULÁRIO INTERATIVO E PDF
# ==========================================

st.markdown("""
<div class="print-btn" style="text-align: right; margin-bottom: 10px;">
    <button onclick="window.print()" style="background-color: #0ea5e9; color: white; border: none; padding: 10px 20px; border-radius: 4px; font-weight: bold; cursor: pointer; font-size: 14px;">🖨️ 3º PASSO: IMPRIMIR / SALVAR PDF</button>
</div>
""", unsafe_allow_html=True)

# Construção do HTML do Formulário sem espaços laterais
html_form = f"""
<style>
.tg {{border-collapse:collapse;border-spacing:0; width: 100%; margin-bottom: 15px;}}
.tg td {{border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:12px; overflow:hidden;padding:4px 5px;word-break:normal;}}
.lbl {{font-weight:bold; text-align:center; width: 15%; background-color: white;}}
.val {{width: 35%;}}
.fi {{width: 100%; border: none; outline: none; background: transparent; font-family: inherit; font-size: 12px; font-weight: bold; color: #000; text-transform: uppercase;}}
.fa {{width: 100%; border: none; outline: none; background: transparent; font-family: inherit; font-size: 12px; font-weight: bold; color: #000; resize: none; overflow: hidden; min-height: 35px; text-transform: uppercase;}}
.secao {{background-color: #1b365d; color: white; font-weight: bold; font-size: 12px; padding: 4px 8px; margin-bottom: 4px; border: 1px solid black; -webkit-print-color-adjust: exact; print-color-adjust: exact;}}
.no-border {{border: none !important;}}
</style>

<div id="area-impressao" style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; color: black; background-color: white; padding: 20px; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">

<!-- CABEÇALHO AZUL COM LOGO SIMULADA -->
<div style="background-color: #1b365d; width: 100%; height: 60px; display: flex; align-items: center; margin-bottom: 15px; border: 1px solid black; padding: 0 15px; box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact;">
<div style="color: white; font-weight: bold; font-size: 14px; border-right: 2px solid white; padding-right: 15px; margin-right: 15px; line-height: 1;">GRUPO<br><span style="font-size: 18px;">equatorial</span><br><span style="font-size: 8px; font-weight: normal; letter-spacing: 2px;">ENERGIA</span></div>
<h2 style="color: white; margin: 0; font-size: 18px; flex-grow: 1; text-align: center; margin-left: -80px;">Formulário de Não Atendimento Expansão</h2>
</div>

<!-- DISTRIBUIDORA / REGIONAL -->
<table class="tg" style="margin-bottom: 12px;">
<tr>
<td class="lbl" style="width: 15%;">Distribuidora:</td>
<td class="val" style="width: 15%;"><input type="text" class="fi" value="EQTL MA" style="text-align: center;"></td>
<td class="no-border" style="width: 20%;"></td>
<td class="lbl" style="width: 15%;">Regional:</td>
<td class="val" style="width: 15%;"><input type="text" class="fi" value="{regional}" style="text-align: center;"></td>
<td class="no-border" style="width: 5%;"></td>
<td class="lbl" style="width: 10%; font-size: 10px;">Data da<br>abertura:</td>
<td class="val" style="width: 10%;"><input type="text" class="fi" value="{data_ab}" style="text-align: center;"></td>
</tr>
</table>

<!-- DADOS DO CLIENTE -->
<div class="secao">Dados do Cliente:</div>
<table class="tg" style="margin-bottom: 4px;">
<tr>
<td class="lbl" style="width: 15%;">Nº da nota:</td>
<td class="val" style="width: 35%;"><input type="text" class="fi" value="{prot_input.upper()}"></td>
<td class="no-border" style="width: 10%;"></td>
<td class="lbl" style="width: 15%;">Conta Contrato:</td>
<td class="val" style="width: 25%;"><input type="text" class="fi" value="{cc}"></td>
</tr>
</table>
<table class="tg" style="margin-bottom: 4px;">
<tr>
<td class="lbl" style="width: 20%;">Parceiro de Negócios:</td>
<td class="val" style="width: 80%;"><input type="text" class="fi" value="{parceiro}"></td>
</tr>
</table>
<table class="tg" style="margin-bottom: 12px;">
<tr>
<td class="lbl" style="width: 15%;">Endereço:</td>
<td class="val" style="width: 85%;"><input type="text" class="fi" value="{endereco}"></td>
</tr>
</table>

<!-- DADOS DA VISITA -->
<div class="secao">Dados da Visita:</div>
<table class="tg" style="margin-bottom: 4px;">
<tr>
<td class="lbl" style="width: 15%;">Data:</td>
<td class="val" style="width: 35%;"><input type="text" class="fi" value="{hoje}"></td>
<td class="no-border" style="width: 10%;"></td>
<td class="lbl" style="width: 15%;">Latitude:</td>
<td class="val" style="width: 25%;"><input type="text" class="fi" value="{lat}"></td>
</tr>
</table>
<table class="tg" style="margin-bottom: 4px;">
<tr>
<td class="lbl" style="width: 15%;">Horário:</td>
<td class="val" style="width: 35%;"><input type="text" class="fi" value="" placeholder="14:30"></td>
<td class="no-border" style="width: 10%;"></td>
<td class="lbl" style="width: 15%;">Longitude:</td>
<td class="val" style="width: 25%;"><input type="text" class="fi" value="{lon}"></td>
</tr>
</table>
<table class="tg" style="margin-bottom: 12px;">
<tr>
<td class="lbl" style="width: 25%;">Identificação da equipe:</td>
<td class="val" style="width: 75%;"><input type="text" class="fi" value="EQP NIP" style="text-align: center;"></td>
</tr>
</table>

<!-- MOTIVO DO EXPURGO -->
<div class="secao">Motivo do expurgo:</div>
<table class="tg" style="margin-bottom: 4px;">
<tr>
<td class="lbl" style="width: 15%;">Justificativa:</td>
<td class="val" style="width: 85%;"><input type="text" class="fi" value="" placeholder="CLIQUE AQUI E DIGITE..."></td>
</tr>
</table>
<table class="tg" style="margin-bottom: 4px;">
<tr>
<td class="lbl" style="width: 20%;">Descrição do Expurgo:</td>
<td class="val" style="width: 80%; padding: 0;"><textarea class="fa" placeholder="CLIQUE AQUI E DIGITE..."></textarea></td>
</tr>
</table>
<table class="tg" style="margin-bottom: 12px;">
<tr>
<td class="lbl" style="width: 25%; background-color: #cbe0f5; -webkit-print-color-adjust: exact; print-color-adjust: exact;">Tratativa no Sistema Comercial:</td>
<td class="val" style="width: 75%; background-color: #cbe0f5; -webkit-print-color-adjust: exact; print-color-adjust: exact;"><input type="text" class="fi" value=""></td>
</tr>
</table>

<!-- EVIDÊNCIAS -->
<div class="secao">Evidências:</div>
<table class="tg" style="margin-bottom: 12px;">
<tr>
<td class="lbl" style="width: 22%; text-align: left; font-size: 11px;">Número do medidor<br>do cliente atendido:</td>
<td class="val" style="width: 23%;"><input type="text" class="fi" value="" style="text-align: center;"></td>
<td class="no-border" style="width: 10%;"></td>
<td class="lbl" style="width: 22%; text-align: left; font-size: 11px;">Número da nota do<br>atendimento em campo:</td>
<td class="val" style="width: 23%;"><input type="text" class="fi" value="0" style="text-align: center;"></td>
</tr>
<tr>
<td class="lbl" style="width: 22%; text-align: left; font-size: 11px; border-top: none;">Número do medidor<br>do vizinho:</td>
<td class="val" style="width: 23%; border-top: none;"><input type="text" class="fi" value="" style="text-align: center;"></td>
<td class="no-border" style="width: 10%;"></td>
<td class="lbl" style="width: 22%; text-align: left; font-size: 11px; border-top: none;">Número da estrutura<br>mais próxima:</td>
<td class="val" style="width: 23%; border-top: none;"><input type="text" class="fi" value="" style="text-align: center;"></td>
</tr>
</table>

<!-- FOTO -->
<div style="border: 1px solid black; width: 100%; height: 400px; display: flex; align-items: center; justify-content: center; background-color: #fafafa; overflow: hidden;">
{img_html}
</div>

</div>
"""

st.markdown(html_form, unsafe_allow_html=True)
