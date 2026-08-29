import streamlit as st
import pandas as pd
import unicodedata

# ==========================================
# 1. CONFIGURAÇÕES DA PÁGINA E CSS (VISUAL MODERNO E PROFISSIONAL)
# ==========================================
st.set_page_config(page_title="Gerador SGO & Nomes de Obra", page_icon="🏗️", layout="wide")

st.markdown("""
<style>
    [data-testid="stHeader"] { display: none !important; }
    html, body, [class*="css"] { font-size: 12px !important; }
    
    .eh { background-color: #059669; color: #f8fafc; text-align: center; font-weight: 700; padding: 6px; border: 1px solid #cbd5e1; border-bottom: none; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; border-radius: 4px 4px 0 0;}
    .eh-yellow { background-color: #fef08a; color: #991b1b; text-align: center; font-weight: 700; padding: 6px; border: 1px solid #cbd5e1; border-bottom: 1px solid #cbd5e1; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; border-radius: 4px 4px 0 0; margin-bottom: 5px;}
    .eh-dark { background-color: #047857; color: white; font-weight: 700; padding: 6px; border: 1px solid #cbd5e1; border-bottom: none; font-size: 11px; text-transform: uppercase; border-radius: 4px 4px 0 0;}
    
    .et { width: 100%; border-collapse: collapse; border: 1px solid #cbd5e1; background-color: white; margin-bottom: 15px; border-radius: 0 0 4px 4px; box-shadow: 0 1px 2px rgba(0,0,0,0.05);}
    .et td { border: 1px solid #e2e8f0; padding: 4px 8px; font-family: ui-sans-serif, system-ui, sans-serif; font-size: 11px; height: 26px; vertical-align: middle; }
    
    .lbl { width: 38%; background-color: #f8fafc; color: #475569; font-weight: 600; white-space: nowrap;}
    .val { width: 62%; background-color: #ffffff; color: #0f172a; font-weight: 700; text-transform: uppercase; }
    
    .text-blue { color: #2563eb !important; }
    .text-red { color: #dc2626 !important; }
    .text-green { color: #059669 !important; }
    
    .obs-box { background-color: #1e293b; color: #f8fafc; border: 1px solid #cbd5e1; padding: 10px; font-family: ui-sans-serif, system-ui, sans-serif; font-size: 11px; font-style: italic; min-height: 200px; white-space: pre-wrap; line-height: 1.4; overflow-y: auto; border-radius: 0 0 4px 4px;}
    .desc-row { border: 1px solid #cbd5e1; height: 26px; width: 100%; margin-bottom: 4px; padding: 4px 8px; font-weight: 600; font-family: ui-monospace, monospace; font-size: 11px; text-transform: uppercase; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; background-color: white; color: #0f172a; border-radius: 4px; box-shadow: 0 1px 2px rgba(0,0,0,0.02);}
    
    .lbl-box { background-color: #fef08a; border: 1px solid #cbd5e1; border-radius: 4px; padding: 0px 8px; font-size: 11px; font-weight: 700; color: #7f1d1d; height: 35px; display: flex; align-items: center; margin-bottom: 0px; margin-top: 2px;}
    div[data-baseweb="select"] > div { border: 1px solid #cbd5e1; border-radius: 4px; min-height: 35px !important; height: 35px !important; font-size: 11px; background-color: white;}
    input[data-testid="stTextInput"] { border: 1px solid #cbd5e1; border-radius: 4px; height: 35px !important; min-height: 35px !important; font-size: 11px; font-weight: bold; background-color: white;}
    .stSelectbox, .stTextInput { margin-bottom: -10px !important; }
    .stTextArea textarea { border: 1px solid #94a3b8 !important; border-radius: 4px !important; font-size: 11px; font-family: ui-monospace, monospace; }
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
# 2. UPLOAD E LISTAS SUSPENSAS
# ==========================================
arquivo_bd = st.sidebar.file_uploader("📥 Suba a planilha base (CRIAR NOME DA OBRA.xlsx)", type=["xlsx"])

df_sisco, df_notas, df_dados = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
lista_tipos_obra, lista_pi, lista_mun, lista_id = [], [], [], []

if arquivo_bd:
    with st.spinner("Carregando banco de dados..."):
        df_sisco, df_notas, df_dados = carregar_dados(arquivo_bd)
        
    if not df_dados.empty:
        if 'TIPO DE OBRA' in df_dados.columns: lista_tipos_obra = sorted(df_dados['TIPO DE OBRA'].dropna().unique().tolist())
        if 'PI' in df_dados.columns: lista_pi = sorted(df_dados['PI'].dropna().unique().tolist())
        if 'SIGLA-MUNICIPIO' in df_dados.columns: lista_mun = sorted(df_dados['SIGLA-MUNICIPIO'].dropna().unique().tolist())
        if 'ID DO NUMERO' in df_dados.columns: lista_id = sorted([str(x).replace('.0', '') for x in df_dados['ID DO NUMERO'].dropna().unique().tolist()])

c1, c2, c3, c4 = st.columns([0.8, 1.8, 2.5, 2.0])

with c1:
    st.markdown('<div class="eh">🎯 SOLICITAÇÕES</div>', unsafe_allow_html=True)
    sols_input = st.text_area("", height=600, placeholder="Cole as notas aqui\n(Uma por linha)", label_visibility="collapsed")
    
solicitacoes = [s.strip() for s in sols_input.split('\n') if s.strip()]

cc, instalacao, fase, tipo_obra_sisco, data_abertura, lat, lon = "", "", "", "", "", "", ""
cidade_auto, cliente_auto, endereco_auto, area_resp, reg_raw, obs = "", "", "", "", "", ""
pi_auto, gerente, executivo, empresa, contrato, tecnico, data_aprov = "", "", "", "", "", "", ""
responsavel_obra, tipo_nota_parceiro, valor_previsto = "", "", ""
obra_relampago_formatada, descricoes_html, nomes_obras_html = "", "", ""

if solicitacoes and (not df_sisco.empty or not df_notas.empty):
    solicitacao_principal = solicitacoes[0]
    resultado_sisco = df_sisco[df_sisco['Nota CCS'] == solicitacao_principal] if not df_sisco.empty else pd.DataFrame()
    resultado_notas = df_notas[df_notas['PROTOCOLO'] == solicitacao_principal] if not df_notas.empty else pd.DataFrame()
    
    if not resultado_sisco.empty or not resultado_notas.empty:
        r_sisco = resultado_sisco.iloc[0] if not resultado_sisco.empty else None
        r_notas = resultado_notas.iloc[0] if not resultado_notas.empty else None
        
        cc = str(r_sisco.get('CC', '')) if r_sisco is not None else ""
        if not cc or cc.lower() == 'nan': cc = str(r_notas.get('CONTA CONTRATO', '')) if r_notas is not None else ""
        cc = cc.replace('.0', '')
            
        instalacao = str(r_sisco.get('INSTALACAO', '')) if r_sisco is not None else ""
        if not instalacao or instalacao.lower() == 'nan': instalacao = str(r_notas.get('INSTALAÇÃO', '')) if r_notas is not None else ""
        instalacao = instalacao.replace('.0', '')
            
        fase = str(r_sisco.get('Tipo de Carga', 'MO')).upper() if r_sisco is not None else "MO"
        if fase.lower() == 'nan' or 'NÃO ESPECIFICADO' in fase or 'NAO ESPECIFICADO' in fase: fase = "MO"
            
        tipo_obra_raw = str(r_sisco.get('Detalhes', '')) if r_sisco is not None else ""
        if tipo_obra_raw.lower() == 'nan': tipo_obra_raw = ""
        parts = tipo_obra_raw.replace("-", " ").strip().split(" ")
        if len(parts) > 3:
            tipo_obra_sisco = " ".join(parts[:3])
        else:
            tipo_obra_sisco = ""
            
        data_abertura = str(r_sisco.get('Data Abertura', '')) if r_sisco is not None else ""
        if not data_abertura or data_abertura.lower() == 'nan': data_abertura = str(r_notas.get('DATA DA SOLICITAÇÃO', ''))[:10] if r_notas is not None else ""
            
        lat = str(r_sisco.get('Latitude', '')) if r_sisco is not None else ""
        if not lat or lat.lower() == 'nan': lat = str(r_notas.get('LATITUDE', '')) if r_notas is not None else ""
        lat = lat.replace('.', ',')
        
        lon = str(r_sisco.get('Longitude', '')) if r_sisco is not None else ""
        if not lon or lon.lower() == 'nan': lon = str(r_notas.get('LONGITUDE', '')) if r_notas is not None else ""
        lon = lon.replace('.', ',')
        
        cidade_raw = str(r_sisco.get('Município', '')) if r_sisco is not None else ""
        if not cidade_raw or cidade_raw.lower() == 'nan': cidade_raw = str(r_notas.get('MUNICIPIO', '')) if r_notas is not None else ""
        cidade_auto = remover_acentos(cidade_raw)
        
        cliente_auto = str(r_sisco.get('Nome', '')) if r_sisco is not None else ""
        if not cliente_auto or cliente_auto.lower() == 'nan': cliente_auto = str(r_notas.get('NOME DO SOLICITANTE', '')) if r_notas is not None else ""
        cliente_auto = cliente_auto.upper()
            
        endereco_auto = str(r_notas.get('ENDEREÇO', '')) if r_notas is not None else ""
        if not endereco_auto or endereco_auto.lower() == 'nan': endereco_auto = str(r_sisco.get('Endereço', '')) if r_sisco is not None else ""
        
        pi_auto = str(r_notas.get('TIPO LIGAÇÃO', '')) if r_notas is not None else ""
        if not pi_auto or pi_auto.lower() == 'nan': pi_auto = str(r_sisco.get('Tipo de Projeto(PI)', '')) if r_sisco is not None else ""
        if pi_auto.lower() == 'nan': pi_auto = ""
            
        reg_raw = str(r_notas.get('REGIONAL', '')) if r_notas is not None else ""
        if not reg_raw or reg_raw.lower() == 'nan': reg_raw = str(r_sisco.get('Regional', '')) if r_sisco is not None else ""
        reg_raw = reg_raw.upper()
        
        obs = str(r_sisco.get('Obs(última obs)', '')) if r_sisco is not None else ""
        if not obs or obs.lower() == 'nan': obs = str(r_notas.get('PONTO DE REFERENCIA', '')) if r_notas is not None else ""
        if obs.lower() == 'nan': obs = ""
    else:
        st.toast(f"❌ A nota principal '{solicitacao_principal}' não foi encontrada.")

# ==========================================
# 3. PAINEL DE DADOS E FORMULÁRIO DE OVERRIDE
# ==========================================
with c2:
    st.markdown(f"""
    <div class="eh">🎲 DADOS</div>
    <table class="et">
        <tr><td class="lbl">Conta Contrato</td><td class="val">{cc}</td></tr>
        <tr><td class="lbl">Instalação CCS</td><td class="val">{instalacao}</td></tr>
        <tr><td class="lbl">FASE</td><td class="val">{fase}</td></tr>
        <tr><td class="lbl">Tipo de Obra</td><td class="val">{tipo_obra_sisco}</td></tr>
        <tr><td class="lbl">Data Abertura</td><td class="val">{data_abertura}</td></tr>
        <tr><td class="lbl">---</td><td class="val"></td></tr>
        <tr><td class="lbl text-blue">LAT</td><td class="val">{lat}</td></tr>
        <tr><td class="lbl text-blue">LONG</td><td class="val">{lon}</td></tr>
    </table>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="eh-yellow" style="margin-bottom: 0px;">🚧 Criar Nome da Obra 🚧</div>', unsafe_allow_html=True)
    
    def criar_linha_input(label, widget_type, key, options=None):
        cA, cB = st.columns([1, 2.5], gap="small")
        with cA:
            st.markdown(f'<div class="lbl-box">{label}</div>', unsafe_allow_html=True)
        with cB:
            if widget_type == "select":
                return st.selectbox("", options, key=key, label_visibility="collapsed")
            else:
                return st.text_input("", key=key, label_visibility="collapsed")

    with st.container():
        man_especial = criar_linha_input("Obra Especial ?", "select", "i1", ["", "OE-Obras Juridicas e/ou Especiais", "EX-Exceções"])
        man_tipo_obra = criar_linha_input("Tipo de Obra", "select", "i2", [""] + lista_tipos_obra)
        man_pi = criar_linha_input("PI", "select", "i3", [""] + lista_pi)
        man_mun = criar_linha_input("Municipio", "select", "i4", [""] + lista_mun)
        man_id = criar_linha_input("ID do Numero", "select", "i5", [""] + lista_id)
        man_sol = criar_linha_input("Solicitação", "text", "i6")
        man_livre = criar_linha_input("Escrita Livre", "text", "i7")


# ==========================================
# 4. LÓGICA DE CRUZAMENTO DE DADOS (OVERRIDE MANUAL)
# ==========================================
pi_ativo = man_pi if man_pi else pi_auto

# Repassa Município Manual para a Tabela SGO
if man_mun and not cidade_auto:
    cidade_auto = man_mun.split('-', 1)[1] if '-' in man_mun else man_mun
    
# Repassa Cliente Manual para a Tabela SGO
if man_livre and not cliente_auto:
    cliente_auto = man_livre.upper()

if man_mun and not df_dados.empty and 'SIGLA-MUNICIPIO' in df_dados.columns:
    dados_mun = df_dados[df_dados['SIGLA-MUNICIPIO'] == man_mun]
    if not dados_mun.empty:
        reg_raw = str(dados_mun.iloc[0].get('REGIONAL', '')).upper()

regional_formatado = ""
regional_label = "Regional"
col_idx = 0

if reg_raw:
    if "SUL" in reg_raw: 
        regional_formatado = "CM04-IMPERATRIZ"; col_idx = 14; regional_label = "Regional Sul"
    elif "CENTRO" in reg_raw: 
        regional_formatado = "CM03-BACABAL"; col_idx = 19; regional_label = "Regional Centro"
    elif "LESTE" in reg_raw: 
        regional_formatado = "CM02-TIMON"; col_idx = 24; regional_label = "Regional Leste"
    elif "NORTE" in reg_raw: 
        regional_formatado = "CM01-SAO LUIS"; col_idx = 4; regional_label = "Regional Norte"
    elif "NOROESTE" in reg_raw: 
        regional_formatado = "CM01-PINHEIRO"; col_idx = 9; regional_label = "Regional Noroeste"

area_resp = ""
if pi_ativo and not df_dados.empty and 'PI' in df_dados.columns:
    dados_pi = df_dados[df_dados['PI'] == pi_ativo]
    if not dados_pi.empty:
        r_dados = dados_pi.iloc[0]
        
        area_resp_nova = str(r_dados.get('Tipo', ''))
        if area_resp_nova.lower() != 'nan' and area_resp_nova != "":
            area_resp = area_resp_nova.upper()
            
        tipo_nota_parceiro = str(r_dados.get('Tipo de NS|Parceiro', ''))
        responsavel_obra = str(r_dados.get('Resp. Obra', ''))
        
        qtd_dias = str(r_dados.get('Qtd dias', '')).replace('.0', '')
        data_aprov = f"{str(r_dados.get('Data final', ''))[:10]} ({qtd_dias} DIAS)"
        
        if pi_ativo in ["UNP", "UNR", "UNI", "UNO", "UNU", "UNJ", "LPT", "MTP", "REG", "ASC", "SID"]:
            total_val = 7000 * (len(solicitacoes) if solicitacoes else 1)
        else:
            total_val = 30000 * (len(solicitacoes) if solicitacoes else 1)
        valor_previsto = f"R$ {total_val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        if col_idx > 0:
            gerente = str(dados_pi.iloc[0, col_idx])
            executivo = str(dados_pi.iloc[0, col_idx + 1])
            empresa = str(dados_pi.iloc[0, col_idx + 2])
            contrato = str(dados_pi.iloc[0, col_idx + 3]).replace('.0', '')
            tecnico = str(dados_pi.iloc[0, col_idx + 4])
            
            if gerente.lower() == 'nan': gerente = ""
            if executivo.lower() == 'nan': executivo = ""
            if empresa.lower() == 'nan': empresa = ""
            if contrato.lower() == 'nan': contrato = ""
            if tecnico.lower() == 'nan': tecnico = ""

if not area_resp: area_resp = ""

# ==========================================
# 5. GERADOR EM MASSA DOS NOMES
# ==========================================
pref_especial = f"{man_especial.split('-')[0]}-" if man_especial else ""
pref_tipo = man_tipo_obra.split('-')[0] if man_tipo_obra else "CT"
pref_pi = pi_ativo if pi_ativo else "UNR"
pref_id = man_id.split('-')[0] if man_id else "NS"

if not solicitacoes and (man_tipo_obra or man_pi or man_mun or man_id or man_sol or man_livre):
    pref_mun = man_mun.split('-')[0] if man_mun else "XXX"
    val_sol_final = man_sol if man_sol else "0000000000"
    val_livre_final = man_livre if man_livre else "NOME"
    
    raw_name = f"{pref_especial}{pref_tipo}-{pref_pi}-{pref_mun}-{pref_id}-{val_sol_final}-{val_livre_final}"
    clean_name = raw_name.replace(".", "").replace("_", "").replace(" ", "-")
    obra_relampago_formatada = clean_name[:34].upper()
    
    nomes_obras_html += f'<div class="desc-row">{obra_relampago_formatada}</div>\n'
    descricoes_html += f'<div class="desc-row">CRIADO MANUALMENTE</div>\n'
else:
    for sol in solicitacoes:
        res_sol_sisco = df_sisco[df_sisco['Nota CCS'] == sol] if not df_sisco.empty else pd.DataFrame()
        res_sol_notas = df_notas[df_notas['PROTOCOLO'] == sol] if not df_notas.empty else pd.DataFrame()
        
        if not res_sol_sisco.empty or not res_sol_notas.empty:
            r_sol_sisco = res_sol_sisco.iloc[0] if not res_sol_sisco.empty else None
            r_sol_notas = res_sol_notas.iloc[0] if not res_sol_notas.empty else None
            
            cc_sol = str(r_sol_sisco.get('CC', '')) if r_sol_sisco is not None else ""
            if not cc_sol or cc_sol.lower() == 'nan': cc_sol = str(r_sol_notas.get('CONTA CONTRATO', '')) if r_sol_notas is not None else ""
            cc_sol = cc_sol.replace('.0', '')
            
            cli_sol = str(r_sol_sisco.get('Nome', '')) if r_sol_sisco is not None else ""
            if not cli_sol or cli_sol.lower() == 'nan': cli_sol = str(r_sol_notas.get('NOME DO SOLICITANTE', '')) if r_sol_notas is not None else ""
            cli_sol = cli_sol.upper()
            
            cid_sol = str(r_sol_sisco.get('Município', '')) if r_sol_sisco is not None else ""
            if not cid_sol or cid_sol.lower() == 'nan': cid_sol = str(r_sol_notas.get('MUNICIPIO', '')) if r_sol_notas is not None else ""
            
            pref_mun = man_mun.split('-')[0] if man_mun else (remover_acentos(cid_sol)[:3] if cid_sol else "XXX")
            val_sol_final = man_sol if man_sol else sol
            val_livre_final = man_livre if man_livre else cli_sol.replace(" ", "-")[:15]
            
            desc_str = f"{sol}-{cli_sol}, CC-{cc_sol}."
            
            raw_name = f"{pref_especial}{pref_tipo}-{pref_pi}-{pref_mun}-{pref_id}-{val_sol_final}-{val_livre_final}"
            clean_name = raw_name.replace(".", "").replace("_", "").replace(" ", "-")
            nome_str = clean_name[:34].upper()
            
            if sol == solicitacoes[0]:
                obra_relampago_formatada = nome_str
        else:
            desc_str = f"{sol} - NÃO ENCONTRADO"
            nome_str = f"{sol} - NÃO ENCONTRADO"
            if sol == solicitacoes[0]: obra_relampago_formatada = nome_str
            
        descricoes_html += f'<div class="desc-row">{desc_str}</div>\n'
        nomes_obras_html += f'<div class="desc-row">{nome_str}</div>\n'

linhas_restantes = 25 - max(len(solicitacoes), 1 if obra_relampago_formatada else 0)
for _ in range(max(linhas_restantes, 0)):
    descricoes_html += '<div class="desc-row"></div>\n'
    nomes_obras_html += '<div class="desc-row"></div>\n'

# ==========================================
# 6. RENDERIZAÇÃO DAS COLUNAS 3 E 4
# ==========================================
with c3:
    lbl_obra_estilo = 'class="lbl"' if not man_tipo_obra and not man_pi and not man_livre and not man_especial else 'class="lbl text-red" style="background-color: #fef08a;"'
    val_obra_estilo = 'class="val text-green"' if not man_tipo_obra and not man_pi and not man_livre and not man_especial else 'class="val text-red" style="background-color: #fef08a; font-style: italic;"'
    lbl_obra_texto = "⚡ Obra Relampago ⚡" if not man_tipo_obra and not man_pi and not man_livre and not man_especial else "🚧 Nome da Obra 🚧"
    
    st.markdown(f"""
    <div class="eh">📝 CRIAÇÃO DA NOTA SGO 📝</div>
    <table class="et">
        <tr><td class="lbl">Tipo Nota | Parceiro</td><td class="val text-red" style="font-size: 11px;">{tipo_nota_parceiro}</td></tr>
        <tr><td {lbl_obra_estilo} style="font-size: 11px;">{lbl_obra_texto}</td><td {val_obra_estilo}>{obra_relampago_formatada}</td></tr>
        <tr><td class="lbl">{regional_label}</td><td class="val">{regional_formatado}</td></tr>
        <tr><td class="lbl">Cidade</td><td class="val">{cidade_auto}</td></tr>
        <tr><td class="lbl">Área Responsável</td><td class="val">{area_resp}</td></tr>
        <tr><td class="lbl">Cliente</td><td class="val">{cliente_auto}</td></tr>
        <tr><td class="lbl">Endereço</td><td class="val">{endereco_auto}</td></tr>
        <tr><td class="lbl">PI</td><td class="val text-blue" style="font-style: italic;">{pi_ativo}</td></tr>
        <tr><td class="lbl">Gerente</td><td class="val">{gerente}</td></tr>
        <tr><td class="lbl">Executivo</td><td class="val">{executivo}</td></tr>
        <tr><td class="lbl">Responsável Obra</td><td class="val text-blue" style="font-size: 11px;">{responsavel_obra}</td></tr>
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
