import streamlit as st
import pandas as pd
import unicodedata
import re

# ==========================================
# 1. CONFIGURAÇÕES DA PÁGINA E CSS (VISUAL EXCEL)
# ==========================================
st.set_page_config(page_title="Gerador SGO & Nomes de Obra", page_icon="🏗️", layout="wide")

st.markdown("""
<style>
    [data-testid="stHeader"] { display: none !important; }
    
    .eh { background-color: #00b050; color: white; text-align: center; font-weight: 900; padding: 4px; border: 3px solid black; border-bottom: 0px; font-size: 14px; text-transform: uppercase; font-family: 'Arial Black', sans-serif;}
    .eh-yellow { background-color: #ffeb9c; color: #c00000; text-align: center; font-weight: 900; padding: 4px; border: 3px solid black; border-bottom: 0px; font-size: 14px;}
    .eh-dark { background-color: #00b050; color: white; font-weight: 900; padding: 4px; border: 3px solid black; border-bottom: 0px; font-size: 13px; text-transform: uppercase; font-family: 'Arial Black', sans-serif;}
    
    .et { width: 100%; border-collapse: collapse; border: 3px solid black; background-color: white; margin-bottom: 15px; }
    .et td { border: 2px solid black; padding: 2px 6px; font-weight: bold; font-family: Calibri, Arial, sans-serif; font-size: 13px; height: 24px; vertical-align: middle; }
    
    .lbl { width: 35%; background-color: #ffffff; color: black; }
    .val { width: 65%; background-color: #ffffff; color: black; text-transform: uppercase; }
    
    .text-blue { color: #0070c0 !important; }
    .text-red { color: #c00000 !important; }
    .text-green { color: #00b050 !important; }
    
    .obs-box { background-color: #595959; color: white; border: 3px solid black; padding: 8px; font-family: Calibri, Arial, sans-serif; font-size: 12px; font-weight: bold; font-style: italic; min-height: 250px; white-space: pre-wrap; line-height: 1.2; overflow-y: auto;}
    .desc-box { background-color: white; color: black; border: 3px solid black; padding: 4px 6px; font-family: 'Arial Black', sans-serif; font-size: 13px; font-weight: 900; text-transform: uppercase; margin-bottom: 15px; min-height: 24px;}
</style>
""", unsafe_allow_html=True)

def remover_acentos(texto):
    if pd.isna(texto) or texto == "": return ""
    texto = str(texto).upper().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

@st.cache_data(show_spinner=False)
def carregar_dados(file):
    xls = pd.ExcelFile(file)
    df_sisco = pd.read_excel(xls, sheet_name='Sisco')
    if 'Nota CCS' in df_sisco.columns:
        df_sisco['Nota CCS'] = df_sisco['Nota CCS'].astype(str).str.replace('.0', '', regex=False)
        
    try: df_notas = pd.read_excel(xls, sheet_name='NotasSisgb')
    except: df_notas = pd.DataFrame()
    if not df_notas.empty and 'PROTOCOLO' in df_notas.columns:
        df_notas['PROTOCOLO'] = df_notas['PROTOCOLO'].astype(str).str.replace('.0', '', regex=False)
        
    try: df_dados = pd.read_excel(xls, sheet_name='Dados', header=1)
    except: df_dados = pd.DataFrame()
    return df_sisco, df_notas, df_dados

# ==========================================
# 2. UPLOAD E PROCESSAMENTO DE ESTADOS
# ==========================================
arquivo_bd = st.sidebar.file_uploader("📥 Suba a planilha base (CRIAR NOME DA OBRA.xlsx)", type=["xlsx"])

# Variáveis em branco por padrão (Para mostrar a tela vazia como no Excel)
cc, instalacao, fase, tipo_obra_sisco, data_abertura, lat, lon = "", "", "", "", "", "", ""
cidade, cliente, endereco, area_resp, regional_formatado, obs = "", "", "", "", "", ""
obra_relampago, pi, gerente, executivo, empresa, contrato, tecnico, data_aprov = "", "", "", "", "", "", "", ""
responsavel_obra, descricao_sgo, tipo_nota_parceiro, valor_previsto = "", "", "", ""
obra_especial = "Normal"

df_sisco, df_notas, df_dados = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

if arquivo_bd:
    with st.spinner("Carregando banco de dados..."):
        df_sisco, df_notas, df_dados = carregar_dados(arquivo_bd)

# Layout: 4 Colunas simulando o Excel
c1, c2, c3, c4 = st.columns([0.8, 1.8, 2.5, 1.5])

with c1:
    st.markdown('<div class="eh">🎯 SOLICITAÇÃO</div>', unsafe_allow_html=True)
    solicitacao = st.text_input("", placeholder="Digite a Nota", label_visibility="collapsed")
    
    for _ in range(25):
        st.markdown('<div style="border: 1px solid #ccc; height: 23px; width: 100%;"></div>', unsafe_allow_html=True)

# Lógica de preenchimento caso digite algo
if solicitacao and not df_sisco.empty:
    solicitacao = solicitacao.strip()
    resultado = df_sisco[df_sisco['Nota CCS'] == solicitacao]
    resultado_notas = df_notas[df_notas['PROTOCOLO'] == solicitacao] if not df_notas.empty else pd.DataFrame()
    
    if not resultado.empty:
        row = resultado.iloc[0]
        row_notas = resultado_notas.iloc[0] if not resultado_notas.empty else None
        
        # --- MAPEAMENTO BASE ---
        cc = str(row.get('CC', '')).replace('.0', '')
        if cc.lower() == 'nan': cc = ""
            
        instalacao = str(row.get('INSTALACAO', '')).replace('.0', '')
        if instalacao.lower() == 'nan': instalacao = ""
            
        fase = str(row.get('Tipo de Carga', 'MO')).upper()
        if fase.lower() == 'NAN': fase = "MO"
            
        tipo_obra_sisco = str(row.get('Tipo de Projeto Descrição', ''))
        data_abertura = str(row.get('Data Abertura', ''))
        lat = str(row.get('Latitude', '')).replace('.', ',')
        lon = str(row.get('Longitude', '')).replace('.', ',')
        if lat.lower() == 'nan': lat = ""
        if lon.lower() == 'nan': lon = ""
        
        cidade = remover_acentos(row.get('Município', ''))
        cliente = str(row.get('Nome', '')).upper()
        if cliente.lower() == 'nan': cliente = ""
            
        # Endereço busca no NotasSisgb primeiro, fallback Sisco
        if row_notas is not None and pd.notna(row_notas.get('ENDEREÇO')):
            endereco = str(row_notas.get('ENDEREÇO'))
        else:
            endereco = str(row.get('Endereço', ''))
        if endereco.lower() == 'nan': endereco = ""
            
        area_resp = str(row.get('Descrição', 'EXPANSÃO')).upper()
        
        # PI: Busca NotasSisgb ('TIPO LIGAÇÃO'), fallback Sisco ('Tipo de Projeto(PI)')
        if row_notas is not None and pd.notna(row_notas.get('TIPO LIGAÇÃO')):
            pi = str(row_notas.get('TIPO LIGAÇÃO'))
        else:
            pi = str(row.get('Tipo de Projeto(PI)', ''))
        if pi.lower() == 'nan': pi = ""
            
        # Cruzamento com a aba DADOS usando o PI
        if pi and not df_dados.empty and 'PI' in df_dados.columns:
            dados_pi = df_dados[df_dados['PI'] == pi]
            if not dados_pi.empty:
                r_dados = dados_pi.iloc[0]
                tipo_nota_parceiro = str(r_dados.get('Tipo de NS|Parceiro', ''))
                responsavel_obra = str(r_dados.get('Resp. Obra', ''))
                data_aprov = f"{str(r_dados.get('Data final', ''))[:10]} ({str(r_dados.get('Qtd dias', ''))} DIAS)"
                
                # Se for certos PIs, valor é 7000, senão 30000
                if pi in ["UNP", "UNR", "UNI", "UNO", "UNU", "UNJ", "LPT", "MTP", "REG", "ASC", "SID"]:
                    valor_previsto = "R$ 7.000,00"
                else:
                    valor_previsto = "R$ 30.000,00"
        
        # Lógica de Regional -> Cruza para achar o Gerente/Executivo/Empresa corretos
        reg_raw = str(row.get('Regional', '')).upper()
        col_idx = 4 # Index das colunas de Regional na aba Dados (Norte, Noroeste, Sul, Centro, Leste)
        regional_formatado = "CM04-IMPERATRIZ"
        
        if "SUL" in reg_raw: 
            regional_formatado = "CM04-IMPERATRIZ"; col_idx = 14
        elif "CENTRO" in reg_raw: 
            regional_formatado = "CM03-BACABAL"; col_idx = 19
        elif "LESTE" in reg_raw: 
            regional_formatado = "CM02-TIMON"; col_idx = 24
        elif "NORTE" in reg_raw: 
            regional_formatado = "CM01-SAO LUIS"; col_idx = 4
        elif "NOROESTE" in reg_raw: 
            regional_formatado = "CM01-PINHEIRO"; col_idx = 9

        # Resgata equipe da regional respectiva baseada no PI
        if pi and not df_dados.empty:
            dados_pi = df_dados[df_dados['PI'] == pi]
            if not dados_pi.empty:
                gerente = str(dados_pi.iloc[0, col_idx]) # Gerente da regional correspondente
                executivo = str(dados_pi.iloc[0, col_idx + 1]) # Executivo
                empresa = str(dados_pi.iloc[0, col_idx + 2]) # Empresa
                contrato = str(dados_pi.iloc[0, col_idx + 3]) # Contrato
                tecnico = str(dados_pi.iloc[0, col_idx + 4]) # Técnico
                
                if gerente.lower() == 'nan': gerente = ""
                if executivo.lower() == 'nan': executivo = ""
                if empresa.lower() == 'nan': empresa = ""
                if contrato.lower() == 'nan': contrato = ""
                if tecnico.lower() == 'nan': tecnico = ""
        
        # Monta nomes
        cliente_curto = cliente.replace(" ", "-")[:15]
        sigla_mun = cidade[:3] if cidade else "XXX"
        descricao_sgo = f"{solicitacao}-{cliente}, CC-{cc}."
        obra_relampago = f"CT-UNR-{sigla_mun}-NS-{solicitacao}-{cliente_curto}"
        
        obs = str(row.get('Obs(última obs)', ''))
        if obs.lower() == 'nan': obs = ""
    else:
        st.toast("❌ Nota não encontrada no Sisco.")

# ==========================================
# 3. RENDERIZAÇÃO DAS COLUNAS FIXAS HTML
# ==========================================
with c2:
    st.markdown(f"""
    <div class="eh">🎲 DADOS</div>
    <table class="et">
        <tr><td class="lbl">Conta Contrato</td><td class="val">{cc}</td></tr>
        <tr><td class="lbl">Instalação CCS</td><td class="val">{instalacao}</td></tr>
        <tr><td class="lbl">FASE</td><td class="val">{fase}</td></tr>
        <tr><td class="lbl">Tipo de Obra</td><td class="val">{tipo_obra_sisco}</td></tr>
        <tr><td class="lbl">Data de Aber</td><td class="val">{data_abertura}</td></tr>
        <tr><td class="lbl">---</td><td class="val"></td></tr>
        <tr><td class="lbl text-blue">LAT</td><td class="val">{lat}</td></tr>
        <tr><td class="lbl text-blue">LONG</td><td class="val">{lon}</td></tr>
    </table>
    
    <div class="eh-yellow">🚧 Criar Nome da Obra 🚧</div>
    <table class="et">
        <tr><td class="lbl" style="font-style: italic;">Obra Especial ?</td><td class="val" style="text-align: right; text-transform: none; font-weight: 900;">{obra_especial}</td></tr>
        <tr><td class="lbl">Tipo de Obra</td><td class="val"></td></tr>
        <tr><td class="lbl">PI</td><td class="val"></td></tr>
        <tr><td class="lbl">Municipio</td><td class="val"></td></tr>
        <tr><td class="lbl">ID do Numero</td><td class="val"></td></tr>
        <tr><td class="lbl">Solicitação</td><td class="val"></td></tr>
        <tr><td class="lbl">Escrita Livre</td><td class="val"></td></tr>
    </table>
    """, unsafe_allow_html=True)

with c3:
    st.markdown(f"""
    <div class="eh">📝 CRIAÇÃO DA NOTA SGO 📝</div>
    <table class="et">
        <tr><td class="lbl">Tipo Nota | Parc</td><td class="val text-red" style="font-size: 11px;">{tipo_nota_parceiro}</td></tr>
        <tr><td class="lbl text-blue" style="font-size: 11px;">⚡ Obra Relampago ⚡</td><td class="val text-green">{obra_relampago}</td></tr>
        <tr><td class="lbl">Regional Sul</td><td class="val">{regional_formatado}</td></tr>
        <tr><td class="lbl">Cidade</td><td class="val">{cidade}</td></tr>
        <tr><td class="lbl">Área Responsá</td><td class="val">{area_resp}</td></tr>
        <tr><td class="lbl">Cliente</td><td class="val">{cliente}</td></tr>
        <tr><td class="lbl">Endereço</td><td class="val">{endereco}</td></tr>
        <tr><td class="lbl">PI</td><td class="val" style="font-style: italic;">{pi}</td></tr>
        <tr><td class="lbl">Gerente</td><td class="val">{gerente}</td></tr>
        <tr><td class="lbl">Executivo</td><td class="val">{executivo}</td></tr>
        <tr><td class="lbl">Responsavel O</td><td class="val text-blue" style="font-size: 11px;">{responsavel_obra}</td></tr>
        <tr><td class="lbl">Valor Previsto</td><td class="val text-green">{valor_previsto}</td></tr>
    </table>
    
    <div class="eh">👍 APROVAÇÃO DA NOTA SGO</div>
    <table class="et">
        <tr><td class="lbl">Empresa</td><td class="val">{empresa}</td></tr>
        <tr><td class="lbl">Contrato</td><td class="val">{contrato}</td></tr>
        <tr><td class="lbl">Técnico</td><td class="val">{tecnico}</td></tr>
        <tr><td class="lbl">Data</td><td class="val">{data_aprov}</td></tr>
    </table>
    
    <div class="eh-dark" style="text-align: left; padding-left: 10px;">"" MAIS OBSERVAÇÕES ABAIXO...</div>
    <div class="obs-box">{obs}</div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown(f"""
    <div class="eh">🖋 DESCRIÇÃO SGO 🖋</div>
    <div class="desc-box">{descricao_sgo}</div>
    """, unsafe_allow_html=True)
    
    for _ in range(35):
        st.markdown('<div style="border: 2px solid black; height: 18px; width: 100%; margin-bottom: 2px;"></div>', unsafe_allow_html=True)
