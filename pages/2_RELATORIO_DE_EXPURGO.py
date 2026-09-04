import streamlit as st
import pandas as pd
import unicodedata
import os
import base64
import re
import datetime

# ==========================================
# 1. CONFIGURAÇÕES DA PÁGINA
# ==========================================
st.set_page_config(page_title="Relatório de Expurgo", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 1rem !important; padding-bottom: 2rem !important; }
    html, body, [class*="css"] { font-size: 12px !important; }
    
    /* Esconde os menus laterais e o resto do site na hora de imprimir o PDF */
    @media print {
        body * { visibility: hidden; }
        #area-impressao, #area-impressao * { visibility: visible !important; }
        #area-impressao { position: absolute; left: 0; top: 0; width: 100%; margin: 0; padding: 0;}
        .stButton, header, [data-testid="stSidebar"] { display: none !important; }
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# FUNÇÕES BLINDADAS E LEITURA DE DADOS
# ==========================================
def remover_acentos(texto):
    if pd.isna(texto) or texto == "": return ""
    texto = str(texto).upper().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

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
# ESTRUTURA DA INTERFACE (2 COLUNAS)
# ==========================================
col_input, col_preview = st.columns([1.2, 2.8])

with col_input:
    st.markdown('<div style="background-color: #047857; color: white; font-weight: bold; padding: 8px; border-radius: 4px; text-align: center; margin-bottom: 15px;">⚙️ CONFIGURAR EXPURGO</div>', unsafe_allow_html=True)
    
    prot_input = st.text_input("Nº da Nota / Protocolo:", placeholder="Ex: 1113804258")
    
    # Variáveis de autocompletar
    regional, data_ab, cc, endereco, lat, lon, parceiro = "", "", "", "", "", "", ""
    
    if prot_input and not df_notas.empty:
        res = df_notas[df_notas['PROTOCOLO'] == prot_input.strip()]
        if not res.empty:
            r = res.iloc[0]
            regional = str(r.get('REGIONAL', '')).upper()
            data_ab = formatar_data(r.get('DATA ABERTURA', r.get('DATA DA SOLICITAÇÃO', '')))
            cc = str(r.get('CONTA CONTRATO', '')).replace('.0', '')
            if cc.lower() == 'nan': cc = ""
            
            endereco_bruto = str(r.get('ENDEREÇO', ''))
            cidade = str(r.get('MUNICIPIO', ''))
            endereco = f"{endereco_bruto} - {cidade}".strip(" -") if endereco_bruto.lower() != 'nan' else ""
            
            lat = str(r.get('LATITUDE', ''))
            if lat.lower() == 'nan': lat = ""
            lon = str(r.get('LONGITUDE', ''))
            if lon.lower() == 'nan': lon = ""
            
            # Cruzamento para buscar Parceiro de Negócios (Empresa)
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
                        empresa_busca = str(d_pi.iloc[0, col_idx + 2])
                        parceiro = empresa_busca if empresa_busca.lower() != 'nan' else ""
            st.success("✅ Dados encontrados!")
        else:
            st.error("❌ Nota não encontrada.")

    # Campos Manuais
    st.markdown("---")
    st.markdown("**Dados da Visita & Motivo**")
    data_visita = st.date_input("Data da Visita:", value=datetime.date.today(), format="DD/MM/YYYY")
    hora_visita = st.time_input("Horário da Visita:")
    justificativa = st.text_input("Justificativa:")
    descricao = st.text_area("Descrição do Expurgo:")
    tratativa = st.text_input("Tratativa no Sistema Comercial:")
    
    st.markdown("---")
    st.markdown("**Evidências Numéricas**")
    c1, c2 = st.columns(2)
    medidor_cli = c1.text_input("Medidor Cliente Atendido:")
    medidor_viz = c1.text_input("Medidor Vizinho:")
    nota_campo = c2.text_input("Nota Atend. em Campo:", value="0")
    estrutura = c2.text_input("Estrutura mais próxima:")
    
    st.markdown("---")
    st.markdown("**Foto da Evidência**")
    foto_upload = st.file_uploader("Anexar Imagem", type=['png', 'jpg', 'jpeg'])

# ==========================================
# PREPARAÇÃO DA IMAGEM E HTML
# ==========================================
foto_base64 = ""
if foto_upload is not None:
    foto_base64 = base64.b64encode(foto_upload.read()).decode()
    img_html = f'<img src="data:image/png;base64,{foto_base64}" style="max-height: 380px; max-width: 100%; object-fit: contain;">'
else:
    img_html = '<span style="color: #ccc; font-style: italic;">Nenhuma imagem anexada</span>'

# ==========================================
# RENDERIZAÇÃO DO FORMULÁRIO (DIREITA)
# ==========================================
with col_preview:
    
    st.markdown("""
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
            <div style="font-size: 16px; font-weight: bold; color: #1b365d;">👁️ Visualização do Relatório</div>
            <button onclick="window.print()" style="background-color: #0ea5e9; color: white; border: none; padding: 6px 12px; border-radius: 4px; font-weight: bold; cursor: pointer;">🖨️ IMPRIMIR / SALVAR PDF</button>
        </div>
    """, unsafe_allow_html=True)
    
    # IMPORTANTE: O HTML abaixo foi estruturado sem recuos e espaços à esquerda
    # para impedir que o Streamlit transforme a tabela em um "bloco de código".
    html_form = f"""
<div id="area-impressao" style="font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; color: black; background-color: white; padding: 20px; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">

<!-- CABEÇALHO AZUL COM LOGO SIMULADA -->
<div style="background-color: #1b365d; width: 100%; height: 60px; display: flex; align-items: center; margin-bottom: 15px; border: 1px solid black; padding: 0 15px; box-sizing: border-box;">
<div style="color: white; font-weight: bold; font-size: 14px; border-right: 2px solid white; padding-right: 15px; margin-right: 15px; line-height: 1;">GRUPO<br><span style="font-size: 18px;">equatorial</span><br><span style="font-size: 8px; font-weight: normal; letter-spacing: 2px;">ENERGIA</span></div>
<h2 style="color: white; margin: 0; font-size: 18px; flex-grow: 1; text-align: center; margin-left: -80px;">Formulário de Não Atendimento Expansão</h2>
</div>

<!-- DISTRIBUIDORA / REGIONAL -->
<div style="margin-bottom: 12px;">
<table style="width: 100%; border-collapse: collapse;">
<tr>
<td style="border: 1px solid black; width: 15%; font-weight: bold; text-align: center; font-size: 12px; padding: 4px;">Distribuidora:</td>
<td style="border: 1px solid black; width: 15%; text-align: center; font-size: 12px; font-weight: bold;">EQTL MA</td>
<td style="border: none; width: 20%;"></td>
<td style="border: 1px solid black; width: 15%; font-weight: bold; text-align: center; font-size: 12px;">Regional:</td>
<td style="border: 1px solid black; width: 15%; text-align: center; font-size: 12px;">{regional}</td>
<td style="border: none; width: 5%;"></td>
<td style="border: 1px solid black; width: 10%; font-weight: bold; text-align: center; font-size: 10px; line-height: 1;">Data da<br>abertura:</td>
<td style="border: 1px solid black; width: 10%; text-align: center; font-size: 12px;">{data_ab}</td>
</tr>
</table>
</div>

<!-- DADOS DO CLIENTE -->
<div style="background-color: #1b365d; color: white; padding: 4px 8px; font-weight: bold; font-size: 12px; margin-bottom: 4px; border: 1px solid black;">Dados do Cliente:</div>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 4px;">
<tr>
<td style="border: 1px solid black; width: 15%; font-weight: bold; text-align: center; font-size: 12px; padding: 4px;">Nº da nota:</td>
<td style="border: 1px solid black; width: 35%; font-size: 12px; padding-left: 8px; font-weight: bold;">{prot_input.upper()}</td>
<td style="border: none; width: 10%;"></td>
<td style="border: 1px solid black; width: 15%; font-weight: bold; text-align: center; font-size: 12px;">Conta Contrato:</td>
<td style="border: 1px solid black; width: 25%; font-size: 12px; padding-left: 8px; font-weight: bold;">{cc}</td>
</tr>
</table>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 4px;">
<tr>
<td style="border: 1px solid black; width: 20%; font-weight: bold; text-align: center; font-size: 12px; padding: 4px;">Parceiro de Negócios:</td>
<td style="border: 1px solid black; width: 80%; font-size: 12px; padding-left: 8px;">{parceiro}</td>
</tr>
</table>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 12px;">
<tr>
<td style="border: 1px solid black; width: 15%; font-weight: bold; text-align: center; font-size: 12px; padding: 4px;">Endereço:</td>
<td style="border: 1px solid black; width: 85%; font-size: 12px; padding-left: 8px;">{endereco}</td>
</tr>
</table>

<!-- DADOS DA VISITA -->
<div style="background-color: #1b365d; color: white; padding: 4px 8px; font-weight: bold; font-size: 12px; margin-bottom: 4px; border: 1px solid black;">Dados da Visita:</div>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 4px;">
<tr>
<td style="border: 1px solid black; width: 15%; font-weight: bold; text-align: center; font-size: 12px; padding: 4px;">Data:</td>
<td style="border: 1px solid black; width: 35%; font-size: 12px; padding-left: 8px;">{data_visita.strftime('%d/%m/%Y')}</td>
<td style="border: none; width: 10%;"></td>
<td style="border: 1px solid black; width: 15%; font-weight: bold; text-align: center; font-size: 12px;">Latitude:</td>
<td style="border: 1px solid black; width: 25%; font-size: 12px; padding-left: 8px;">{lat}</td>
</tr>
</table>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 4px;">
<tr>
<td style="border: 1px solid black; width: 15%; font-weight: bold; text-align: center; font-size: 12px; padding: 4px;">Horário:</td>
<td style="border: 1px solid black; width: 35%; font-size: 12px; padding-left: 8px;">{hora_visita.strftime('%H:%M') if hora_visita else ''}</td>
<td style="border: none; width: 10%;"></td>
<td style="border: 1px solid black; width: 15%; font-weight: bold; text-align: center; font-size: 12px;">Longitude:</td>
<td style="border: 1px solid black; width: 25%; font-size: 12px; padding-left: 8px;">{lon}</td>
</tr>
</table>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 12px;">
<tr>
<td style="border: 1px solid black; width: 25%; font-weight: bold; text-align: center; font-size: 12px; padding: 4px;">Identificação da equipe:</td>
<td style="border: 1px solid black; width: 75%; font-weight: bold; font-size: 12px; text-align: center;">EQP NIP</td>
</tr>
</table>

<!-- MOTIVO DO EXPURGO -->
<div style="background-color: #1b365d; color: white; padding: 4px 8px; font-weight: bold; font-size: 12px; margin-bottom: 4px; border: 1px solid black;">Motivo do expurgo:</div>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 4px;">
<tr>
<td style="border: 1px solid black; width: 15%; font-weight: bold; text-align: center; font-size: 12px; padding: 4px;">Justificativa:</td>
<td style="border: 1px solid black; width: 85%; font-size: 12px; padding-left: 8px;">{justificativa.upper()}</td>
</tr>
</table>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 4px;">
<tr>
<td style="border: 1px solid black; width: 20%; font-weight: bold; text-align: center; font-size: 12px; padding: 4px;">Descrição do Expurgo:</td>
<td style="border: 1px solid black; width: 80%; font-size: 12px; padding-left: 8px; min-height: 25px;">{descricao.upper()}</td>
</tr>
</table>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 12px;">
<tr>
<td style="border: 1px solid black; width: 25%; font-weight: bold; text-align: center; font-size: 12px; padding: 4px; background-color: #cbe0f5;">Tratativa no Sistema Comercial:</td>
<td style="border: 1px solid black; width: 75%; font-size: 12px; padding-left: 8px; background-color: #cbe0f5;">{tratativa.upper()}</td>
</tr>
</table>

<!-- EVIDÊNCIAS -->
<div style="background-color: #1b365d; color: white; padding: 4px 8px; font-weight: bold; font-size: 12px; margin-bottom: 4px; border: 1px solid black;">Evidências:</div>
<table style="width: 100%; border-collapse: collapse; margin-bottom: 4px;">
<tr>
<td style="border: 1px solid black; width: 22%; font-weight: bold; text-align: left; font-size: 11px; padding: 4px;">Número do medidor<br>do cliente atendido:</td>
<td style="border: 1px solid black; width: 23%; font-size: 12px; text-align: center;">{medidor_cli}</td>
<td style="border: none; width: 10%;"></td>
<td style="border: 1px solid black; width: 22%; font-weight: bold; text-align: left; font-size: 11px; padding: 4px;">Número da nota do<br>atendimento em campo:</td>
<td style="border: 1px solid black; width: 23%; font-size: 12px; text-align: center;">{nota_campo}</td>
</tr>
<tr>
<td style="border: 1px solid black; border-top: none; width: 22%; font-weight: bold; text-align: left; font-size: 11px; padding: 4px;">Número do medidor<br>do vizinho:</td>
<td style="border: 1px solid black; border-top: none; width: 23%; font-size: 12px; text-align: center;">{medidor_viz}</td>
<td style="border: none; width: 10%;"></td>
<td style="border: 1px solid black; border-top: none; width: 22%; font-weight: bold; text-align: left; font-size: 11px; padding: 4px;">Número da estrutura<br>mais próxima:</td>
<td style="border: 1px solid black; border-top: none; width: 23%; font-size: 12px; text-align: center;">{estrutura}</td>
</tr>
</table>

<!-- FOTO -->
<div style="border: 1px solid black; width: 100%; height: 380px; display: flex; align-items: center; justify-content: center; background-color: #fafafa; overflow: hidden;">
{img_html}
</div>

</div>
"""

    # Camada extra de segurança: Remove qualquer espaço inicial persistente
    html_form = re.sub(r'^[ \t]+', '', html_form, flags=re.MULTILINE)
    
    st.markdown(html_form, unsafe_allow_html=True)
