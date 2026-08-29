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
    
    .desc-row { border: 2px solid black; height: 22px; width: 100%; margin-bottom: 2px; padding: 2px 6px; font-weight: 900; font-family: 'Arial Black', sans-serif; font-size: 11px; text-transform: uppercase; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; background-color: white; color: black; line-height: 14px;}
    
    .stTextArea textarea { border: 3px solid black !important; border-radius: 0px !important; font-weight: bold; font-family: monospace; }
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

# Variáveis em branco por padrão
cc, instalacao, fase, tipo_obra_sisco, data_abertura, lat, lon = "", "", "", "", "", "", ""
cidade, cliente, endereco, area_resp, regional_formatado, obs = "", "", "", "", "", ""
obra_relampago, pi, gerente, executivo, empresa, contrato, tecnico, data_aprov = "", "", "", "", "", "", "", ""
responsavel_obra, tipo_nota_parceiro, valor_previsto = "", "", ""
obra_especial = "Normal"
descricoes_html = ""
nomes_obras_html = ""

df_sisco, df_notas, df_dados = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

if arquivo_bd:
    with st.spinner("Carregando banco de dados..."):
        df_sisco, df_notas, df_dados = carregar_dados(arquivo_bd)

# Layout: 4 Colunas simulando o Excel
c1, c2, c3, c4 = st.columns([0.8, 1.8, 2.5, 2.0])

with c1:
    st.markdown('<div class="eh">🎯 SOLICITAÇÕES</div>', unsafe_allow_html=True)
    sols_input = st.text_area("", height=600, placeholder="Cole as notas aqui\n(Uma por linha)", label_visibility="collapsed")
    
solicitacoes = [s.strip() for s in sols_input.split('\n') if s.strip()]

# Lógica de preenchimento caso digite algo
if solicitacoes and (not df_sisco.empty or not df_notas.empty):
    
    solicitacao_principal = solicitacoes[0]
    
    # Busca principal: tenta no Sisco, depois no NotasSisgb
    resultado_sisco = df_sisco[df_sisco['Nota CCS'] == solicitacao_principal] if not df_sisco.empty else pd.DataFrame()
    resultado_notas = df_notas[df_notas['PROTOCOLO'] == solicitacao_principal] if not df_notas.empty else pd.DataFrame()
    
    if not resultado_sisco.empty or not resultado_notas.empty:
        r_sisco = resultado_sisco.iloc[0] if not resultado_sisco.empty else None
        r_notas = resultado_notas.iloc[0] if not resultado_notas.empty else None
        
        # --- MAPEAMENTO INTELIGENTE (Cruza Sisco e NotasSisgb) ---
        cc = str(r_sisco.get('CC', '')) if r_sisco is not None else ""
        if not cc or cc.lower() == 'nan': cc = str(r_notas.get('CONTA CONTRATO', '')) if r_notas is not None else ""
        cc = cc.replace('.0', '')
            
        instalacao = str(r_sisco.get('INSTALACAO', '')) if r_sisco is not None else ""
        if not instalacao or instalacao.lower() == 'nan': instalacao = str(r_notas.get('INSTALAÇÃO', '')) if r_notas is not None else ""
        instalacao = instalacao.replace('.0', '')
            
        fase = str(r_sisco.get('Tipo de Carga', 'MO')).upper() if r_sisco is not None else "MO"
        if fase.lower() == 'NAN': fase = "MO"
            
        tipo_obra_sisco = str(r_sisco.get('Tipo de Projeto Descrição', '')) if r_sisco is not None else ""
        if tipo_obra_sisco.lower() == 'nan': tipo_obra_sisco = ""
            
        data_abertura = str(r_sisco.get('Data Abertura', '')) if r_sisco is not None else ""
        if not data_abertura or data_abertura.lower() == 'nan':
            data_abertura = str(r_notas.get('DATA DA SOLICITAÇÃO', ''))[:10] if r_notas is not None else ""
            
        lat = str(r_sisco.get('Latitude', '')) if r_sisco is not None else ""
        if not lat or lat.lower() == 'nan': lat = str(r_notas.get('LATITUDE', '')) if r_notas is not None else ""
        lat = lat.replace('.', ',')
        
        lon = str(r_sisco.get('Longitude', '')) if r_sisco is not None else ""
        if not lon or lon.lower() == 'nan': lon = str(r_notas.get('LONGITUDE', '')) if r_notas is not None else ""
        lon = lon.replace('.', ',')
        
        cidade_raw = str(r_sisco.get('Município', '')) if r_sisco is not None else ""
        if not cidade_raw or cidade_raw.lower() == 'nan': cidade_raw = str(r_notas.get('MUNICIPIO', '')) if r_notas is not None else ""
        cidade = remover_acentos(cidade_raw)
        
        cliente = str(r_sisco.get('Nome', '')) if r_sisco is not None else ""
        if not cliente or cliente.lower() == 'nan': cliente = str(r_notas.get('NOME DO SOLICITANTE', '')) if r_notas is not None else ""
        cliente = cliente.upper()
            
        endereco = str(r_notas.get('ENDEREÇO', '')) if r_notas is not None else ""
        if not endereco or endereco.lower() == 'nan': endereco = str(r_sisco.get('Endereço', '')) if r_sisco is not None else ""
            
        area_resp = str(r_sisco.get('Descrição', 'EXPANSÃO')).upper() if r_sisco is not None else "EXPANSÃO"
        if area_resp.lower() == 'nan': area_resp = "EXPANSÃO"
        
        pi = str(r_notas.get('TIPO LIGAÇÃO', '')) if r_notas is not None else ""
        if not pi or pi.lower() == 'nan': pi = str(r_sisco.get('Tipo de Projeto(PI)', '')) if r_sisco is not None else ""
        if pi.lower() == 'nan': pi = ""
            
        # Cruzamento com a aba DADOS
        if pi and not df_dados.empty and 'PI' in df_dados.columns:
            dados_pi = df_dados[df_dados['PI'] == pi]
            if not dados_pi.empty:
                r_dados = dados_pi.iloc[0]
                tipo_nota_parceiro = str(r_dados.get('Tipo de NS|Parceiro', ''))
                responsavel_obra = str(r_dados.get('Resp. Obra', ''))
                data_aprov = f"{str(r_dados.get('Data final', ''))[:10]} ({str(r_dados.get('Qtd dias', ''))} DIAS)"
                
                if pi in ["UNP", "UNR", "UNI", "UNO", "UNU", "UNJ", "LPT", "MTP", "REG", "ASC", "SID"]:
                    total_val = 7000 * len(solicitacoes)
                else:
                    total_val = 30000 * len(solicitacoes)
                valor_previsto = f"R$ {total_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        reg_raw = str(r_notas.get('REGIONAL', '')) if r_notas is not None else ""
        if not reg_raw or reg_raw.lower() == 'nan': reg_raw = str(r_sisco.get('Regional', '')) if r_sisco is not None else ""
        reg_raw = reg_raw.upper()
        
        col_idx = 4
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

        if pi and not df_dados.empty:
            dados_pi = df_dados[df_dados['PI'] == pi]
            if not dados_pi.empty:
                gerente = str(dados_pi.iloc[0, col_idx])
                executivo = str(dados_pi.iloc[0, col_idx + 1])
                empresa = str(dados_pi.iloc[0, col_idx + 2])
                contrato = str(dados_pi.iloc[0, col_idx + 3])
                tecnico = str(dados_pi.iloc[0, col_idx + 4])
                
                if gerente.lower() == 'nan': gerente = ""
                if executivo.lower() == 'nan': executivo = ""
                if empresa.lower() == 'nan': empresa = ""
                if contrato.lower() == 'nan': contrato = ""
                if tecnico.lower() == 'nan': tecnico = ""
        
        cliente_curto = cliente.replace(" ", "-")[:15]
        sigla_mun = cidade[:3] if cidade else "XXX"
        obra_relampago = f"CT-{pi if pi else 'UNR'}-{sigla_mun}-NS-{solicitacao_principal}-{cliente_curto}"
        
        obs = str(r_sisco.get('Obs(última obs)', '')) if r_sisco is not None else ""
        if not obs or obs.lower() == 'nan': obs = str(r_notas.get('PONTO DE REFERENCIA', '')) if r_notas is not None else ""
        if obs.lower() == 'nan': obs = ""
        
        # =========================================================
        # GERAÇÃO EM MASSA (DESCRIÇÕES E NOMES PARA TODAS AS NOTAS)
        # =========================================================
        for sol in solicitacoes:
            res_sol_sisco = df_sisco[df_sisco['Nota CCS'] == sol] if not df_sisco.empty else pd.DataFrame()
            res_sol_notas = df_notas[df_notas['PROTOCOLO'] == sol] if not df_notas.empty else pd.DataFrame()
            
            if not res_sol_sisco.empty or not res_sol_notas.empty:
                r_sol_sisco = res_sol_sisco.iloc[0] if not res_sol_sisco.empty else None
                r_sol_notas = res_sol_notas.iloc[0] if not res_sol_notas.empty else None
                
                cc_sol = str(r_sol_sisco.get('CC', '')) if r_sol_sisco is not None else ""
                if not cc_sol or cc_sol.lower() == 'nan':
                    cc_sol = str(r_sol_notas.get('CONTA CONTRATO', '')) if r_sol_notas is not None else ""
                cc_sol = cc_sol.replace('.0', '')
                
                cli_sol = str(r_sol_sisco.get('Nome', '')) if r_sol_sisco is not None else ""
                if not cli_sol or cli_sol.lower() == 'nan':
                    cli_sol = str(r_sol_notas.get('NOME DO SOLICITANTE', '')) if r_sol_notas is not None else ""
                cli_sol = cli_sol.upper()
                
                cid_sol = str(r_sol_sisco.get('Município', '')) if r_sol_sisco is not None else ""
                if not cid_sol or cid_sol.lower() == 'nan':
                    cid_sol = str(r_sol_notas.get('MUNICIPIO', '')) if r_sol_notas is not None else ""
                cid_sol = remover_acentos(cid_sol)
                
                sigla_mun_sol = cid_sol[:3] if cid_sol else "XXX"
                cli_curto_sol = cli_sol.replace(" ", "-")[:15]
                
                pi_sol = str(r_sol_notas.get('TIPO LIGAÇÃO', '')) if r_sol_notas is not None else ""
                if not pi_sol or pi_sol.lower() == 'nan': pi_sol = str(r_sol_sisco.get('Tipo de Projeto(PI)', '')) if r_sol_sisco is not None else ""
                if pi_sol.lower() == 'nan' or not pi_sol: pi_sol = "UNR"
                
                desc_str = f"{sol}-{cli_sol}, CC-{cc_sol}."
                nome_str = f"CT-{pi_sol}-{sigla_mun_sol}-NS-{sol}-{cli_curto_sol}"
            else:
                desc_str = f"{sol} - NÃO ENCONTRADO"
                nome_str = f"{sol} - NÃO ENCONTRADO"
                
            descricoes_html += f'<div class="desc-row">{desc_str}</div>\n'
            nomes_obras_html += f'<div class="desc-row">{nome_str}</div>\n'
            
    else:
        st.toast(f"❌ A nota principal '{solicitacao_principal}' não foi encontrada.")

linhas_restantes = 25 - len(solicitacoes)
for _ in range(max(linhas_restantes, 0)):
    descricoes_html += '<div class="desc-row"></div>\n'
    nomes_obras_html += '<div class="desc-row"></div>\n'

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
    st.markdown('<div class="eh">🖋 DESCRIÇÕES SGO 🖋</div>', unsafe_allow_html=True)
    st.markdown(descricoes_html, unsafe_allow_html=True)
    
    st.markdown('<div class="eh" style="margin-top: 15px;">🚧 NOMES DAS OBRAS 🚧</div>', unsafe_allow_html=True)
    st.markdown(nomes_obras_html, unsafe_allow_html=True)
