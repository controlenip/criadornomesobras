import streamlit as st
import pandas as pd
import unicodedata
import os

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

# Função para resetar os campos manuais
def limpar_campos_manuais():
    for i in range(1, 10):
        chave = f"i{i}"
        if chave in st.session_state:
            st.session_state[chave] = ""

def remover_acentos(texto):
    if pd.isna(texto) or texto == "": return ""
    texto = str(texto).upper().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

# Função auxiliar para Extração Inteligente com Prioridade (BI -> Sisco -> NotasSisgb)
def get_val(r_bi, r_sisco, r_notas, col_bi, col_sisco, col_notas, default=""):
    val = ""
    if r_bi is not None and col_bi in r_bi: val = str(r_bi.get(col_bi, ''))
    if (not val or val.lower() == 'nan') and r_sisco is not None and col_sisco in r_sisco: val = str(r_sisco.get(col_sisco, ''))
    if (not val or val.lower() == 'nan') and r_notas is not None and col_notas in r_notas: val = str(r_notas.get(col_notas, ''))
    if not val or val.lower() == 'nan': return default
    return val

@st.cache_data(show_spinner=False)
def carregar_dados(file):
    xls = pd.ExcelFile(file)
    
    try: 
        df_sisco = pd.read_excel(xls, sheet_name='Sisco')
        if 'Nota CCS' in df_sisco.columns: df_sisco['Nota CCS'] = df_sisco['Nota CCS'].astype(str).str.replace('.0', '', regex=False)
    except: df_sisco = pd.DataFrame()
        
    try: 
        df_notas = pd.read_excel(xls, sheet_name='NotasSisgb')
        if 'PROTOCOLO' in df_notas.columns: df_notas['PROTOCOLO'] = df_notas['PROTOCOLO'].astype(str).str.replace('.0', '', regex=False)
    except: df_notas = pd.DataFrame()
        
    try: df_dados = pd.read_excel(xls, sheet_name='Dados', header=1)
    except: df_dados = pd.DataFrame()
    
    # Prioridade Master: Tenta ler a aba BI (Levantamentos)
    try:
        df_bi = pd.read_excel(xls, sheet_name='BI', header=1)
        if 'Nota CCS' not in df_bi.columns: df_bi = pd.read_excel(xls, sheet_name='BI', header=2)
        if 'Nota CCS' in df_bi.columns: df_bi['Nota CCS'] = df_bi['Nota CCS'].astype(str).str.replace('.0', '', regex=False)
    except: df_bi = pd.DataFrame()
        
    return df_sisco, df_notas, df_dados, df_bi

# ==========================================
# 2. MENU LATERAL, UPLOAD E LISTAS SUSPENSAS
# ==========================================

# Inserção da Logo de forma Robusta
logo_path = os.path.join(os.path.dirname(__file__), "LOGO_NIP.png")
if os.path.exists(logo_path):
    st.sidebar.image(logo_path, use_container_width=True)
    st.sidebar.markdown("---")
elif os.path.exists("LOGO_NIP.png"):
    st.sidebar.image("LOGO_NIP.png", use_container_width=True)
    st.sidebar.markdown("---")
else:
    st.sidebar.warning("⚠️ Logo 'LOGO_NIP.png' não encontrada na pasta. Verifique se o nome está correto (maiúsculas/minúsculas).")

arquivo_bd = st.sidebar.file_uploader("📥 Suba a planilha base (CRIAR NOME DA OBRA.xlsx)", type=["xlsx"])

df_sisco, df_notas, df_dados, df_bi = pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
lista_tipos_obra, lista_pi, lista_mun, lista_id = [], [], [], []
map_tipo_obra, map_mun = {}, {}

if arquivo_bd:
    with st.spinner("Carregando banco de dados com Inteligência BI..."):
        df_sisco, df_notas, df_dados, df_bi = carregar_dados(arquivo_bd)
        
    if not df_dados.empty:
        if 'TIPO DE OBRA' in df_dados.columns: lista_tipos_obra = sorted(df_dados['TIPO DE OBRA'].dropna().unique().tolist())
        if 'PI' in df_dados.columns: lista_pi = sorted(df_dados['PI'].dropna().unique().tolist())
        if 'SIGLA-MUNICIPIO' in df_dados.columns: lista_mun = sorted(df_dados['SIGLA-MUNICIPIO'].dropna().unique().tolist())
        if 'ID DO NUMERO' in df_dados.columns: lista_id = sorted([str(x).replace('.0', '') for x in df_dados['ID DO NUMERO'].dropna().unique().tolist()])
        
        if 'TIPO DE OBRA NO SISCO' in df_dados.columns and 'SIGLA' in df_dados.columns:
            df_to = df_dados.dropna(subset=['TIPO DE OBRA NO SISCO', 'SIGLA'])
            map_tipo_obra = dict(zip(df_to['TIPO DE OBRA NO SISCO'].astype(str).str.strip().str.upper(), df_to['SIGLA'].astype(str).str.strip().str.upper()))
            
        if 'MUNICIPIO' in df_dados.columns and 'SIGLA.1' in df_dados.columns:
            df_mu = df_dados.dropna(subset=['MUNICIPIO', 'SIGLA.1'])
            map_mun = dict(zip(df_mu['MUNICIPIO'].astype(str).apply(remover_acentos), df_mu['SIGLA.1'].astype(str).str.strip().str.upper()))

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

if solicitacoes:
    sol_princ = solicitacoes[0]
    
    r_bi = df_bi[df_bi['Nota CCS'] == sol_princ].iloc[0] if not df_bi.empty and not df_bi[df_bi['Nota CCS'] == sol_princ].empty else None
    r_sisco = df_sisco[df_sisco['Nota CCS'] == sol_princ].iloc[0] if not df_sisco.empty and not df_sisco[df_sisco['Nota CCS'] == sol_princ].empty else None
    r_notas = df_notas[df_notas['PROTOCOLO'] == sol_princ].iloc[0] if not df_notas.empty and not df_notas[df_notas['PROTOCOLO'] == sol_princ].empty else None
    
    if r_bi is not None or r_sisco is not None or r_notas is not None:
        cc = get_val(r_bi, r_sisco, r_notas, 'CC', 'CC', 'CONTA CONTRATO').replace('.0', '')
        instalacao = get_val(r_bi, r_sisco, r_notas, 'Instalação', 'INSTALACAO', 'INSTALAÇÃO').replace('.0', '')
        
        fase = get_val(r_bi, r_sisco, r_notas, 'FASE', 'Tipo de Carga', '', 'MO').upper()
        if 'NÃO ESPECIFICADO' in fase or 'NAO ESPECIFICADO' in fase or not fase: fase = "MO"
            
        tipo_obra_raw = get_val(r_bi, r_sisco, r_notas, 'Tipo de Projeto Descrição', 'Detalhes', '')
        parts = tipo_obra_raw.replace("-", " ").strip().split()
        tipo_obra_sisco = " ".join(parts[:3]) if parts else ""
        
        data_abertura = get_val(r_bi, r_sisco, r_notas, 'Data Abertura', 'Data Abertura', 'DATA DA SOLICITAÇÃO')[:10]
        lat = get_val(r_bi, r_sisco, r_notas, 'LATITUDE', 'Latitude', 'LATITUDE').replace('.', ',')
        lon = get_val(r_bi, r_sisco, r_notas, 'LONGITUDE', 'Longitude', 'LONGITUDE').replace('.', ',')
        
        cidade_auto = remover_acentos(get_val(r_bi, r_sisco, r_notas, 'Município', 'Município', 'MUNICIPIO'))
        cliente_auto = get_val(r_bi, r_sisco, r_notas, 'NOME CLIENTE', 'Nome', 'NOME DO SOLICITANTE').upper()
        endereco_auto = get_val(r_bi, r_sisco, r_notas, 'TEXTO_GERAL', 'Endereço', 'ENDEREÇO')
        pi_auto = get_val(r_bi, r_sisco, r_notas, 'Tipo de Projeto(PI)', 'Tipo de Projeto(PI)', 'TIPO LIGAÇÃO')
        reg_raw = get_val(r_bi, r_sisco, r_notas, 'Regional', 'Regional', 'REGIONAL').upper()
        obs = get_val(r_bi, r_sisco, r_notas, 'TEXTO', 'Obs(última obs)', 'PONTO DE REFERENCIA')
    else:
        st.toast(f"❌ A nota principal '{sol_princ}' não foi encontrada.")

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
    
    st.markdown('<div class="eh-yellow" style="margin-bottom: 0px;">🚧 Criar Nome da Obra Manual 🚧</div>', unsafe_allow_html=True)
    
    def criar_linha_input(label, widget_type, key, options=None):
        cA, cB = st.columns([1, 2.5], gap="small")
        with cA:
            st.markdown(f'<div class="lbl-box">{label}</div>', unsafe_allow_html=True)
        with cB:
            if widget_type == "select": return st.selectbox("", options, key=key, label_visibility="collapsed")
            else: return st.text_input("", key=key, label_visibility="collapsed")

    with st.container():
        man_especial = criar_linha_input("Obra Especial ?", "select", "i1", ["", "OE-Obras Juridicas e/ou Especiais", "EX-Exceções"])
        man_tipo_obra = criar_linha_input("Tipo de Obra", "select", "i2", [""] + lista_tipos_obra)
        man_pi = criar_linha_input("PI", "select", "i3", [""] + lista_pi)
        man_mun = criar_linha_input("Municipio", "select", "i4", [""] + lista_mun)
        man_id = criar_linha_input("ID do Numero", "select", "i5", [""] + lista_id)
        man_sol = criar_linha_input("Solicitação", "text", "i6")
        man_livre = criar_linha_input("Escrita Livre / Cliente", "text", "i7")
        man_endereco = criar_linha_input("Endereço", "text", "i8")
        man_cc = criar_linha_input("Conta Contrato", "text", "i9")
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("🧹 Limpar Campos Manuais", on_click=limpar_campos_manuais, use_container_width=True)

# ==========================================
# 4. LÓGICA DE CRUZAMENTO DE DADOS
# ==========================================
pi_ativo = man_pi if man_pi else pi_auto

if man_mun: cidade_auto = man_mun.split('-', 1)[1] if '-' in man_mun else man_mun
if man_livre: cliente_auto = man_livre.upper()
if man_endereco: endereco_auto = man_endereco.upper()
if man_cc: cc = man_cc

if man_mun and not df_dados.empty and 'SIGLA-MUNICIPIO' in df_dados.columns:
    dados_mun = df_dados[df_dados['SIGLA-MUNICIPIO'] == man_mun]
    if not dados_mun.empty: reg_raw = str(dados_mun.iloc[0].get('REGIONAL', '')).upper()

regional_formatado = ""
regional_label = "Regional"
col_idx = 0

if reg_raw:
    if "SUL" in reg_raw: regional_formatado = "CM04-IMPERATRIZ"; col_idx = 14; regional_label = "Regional Sul"
    elif "CENTRO" in reg_raw: regional_formatado = "CM03-BACABAL"; col_idx = 19; regional_label = "Regional Centro"
    elif "LESTE" in reg_raw: regional_formatado = "CM02-TIMON"; col_idx = 24; regional_label = "Regional Leste"
    elif "NORTE" in reg_raw: regional_formatado = "CM01-SAO LUIS"; col_idx = 4; regional_label = "Regional Norte"
    elif "NOROESTE" in reg_raw: regional_formatado = "CM01-PINHEIRO"; col_idx = 9; regional_label = "Regional Noroeste"

area_resp = ""
if pi_ativo and not df_dados.empty and 'PI' in df_dados.columns:
    dados_pi = df_dados[df_dados['PI'] == pi_ativo]
    if not dados_pi.empty:
        r_dados = dados_pi.iloc[0]
        
        area_resp_nova = str(r_dados.get('Tipo', ''))
        if area_resp_nova.lower() != 'nan' and area_resp_nova: area_resp = area_resp_nova.upper()
            
        tipo_nota_parceiro = str(r_dados.get('Tipo de NS|Parceiro', ''))
        responsavel_obra = str(r_dados.get('Resp. Obra', ''))
        qtd_dias = str(r_dados.get('Qtd dias', '')).replace('.0', '')
        data_aprov = f"{str(r_dados.get('Data final', ''))[:10]} ({qtd_dias} DIAS)"
        
        total_val = 7000 * max(len(solicitacoes), 1) if pi_ativo in ["UNP", "UNR", "UNI", "UNO", "UNU", "UNJ", "LPT", "MTP", "REG", "ASC", "SID"] else 30000 * max(len(solicitacoes), 1)
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
# 5. GERADOR EM MASSA DOS NOMES E DESCRIÇÕES
# ==========================================
pref_especial = f"{man_especial.split('-')[0]}-" if man_especial else ""
pref_id = man_id.split('-')[0] if man_id else "NS"
pref_pi = pi_ativo if pi_ativo else "UNR"

if not solicitacoes and (man_tipo_obra or man_pi or man_mun or man_id or man_sol or man_livre or man_endereco or man_cc):
    pref_tipo = man_tipo_obra.split('-')[0] if man_tipo_obra else "CT"
    pref_mun = man_mun.split('-')[0] if man_mun else "XXX"
    val_sol_final = man_sol if man_sol else "0000000000"
    val_livre_final_nome = man_livre.replace(" ", "-")[:15] if man_livre else "NOME"
    val_livre_final_desc = man_livre.upper() if man_livre else "NOME"
    val_cc_final = man_cc if man_cc else "0000000000"
    
    raw_name = f"{pref_especial}{pref_tipo}-{pref_pi}-{pref_mun}-{pref_id}-{val_sol_final}-{val_livre_final_nome}"
    obra_relampago_formatada = raw_name.replace(".", "").replace("_", "").replace(" ", "-")[:34].upper()
    desc_str = f"{val_sol_final}-{val_livre_final_desc}, CC-{val_cc_final}."
    
    nomes_obras_html += f'<div class="desc-row">{obra_relampago_formatada}</div>\n'
    descricoes_html += f'<div class="desc-row">{desc_str}</div>\n'
else:
    for sol in solicitacoes:
        r_bi_loop = df_bi[df_bi['Nota CCS'] == sol].iloc[0] if not df_bi.empty and not df_bi[df_bi['Nota CCS'] == sol].empty else None
        r_sisco_loop = df_sisco[df_sisco['Nota CCS'] == sol].iloc[0] if not df_sisco.empty and not df_sisco[df_sisco['Nota CCS'] == sol].empty else None
        r_notas_loop = df_notas[df_notas['PROTOCOLO'] == sol].iloc[0] if not df_notas.empty and not df_notas[df_notas['PROTOCOLO'] == sol].empty else None
        
        if r_bi_loop is not None or r_sisco_loop is not None or r_notas_loop is not None:
            cc_sol = get_val(r_bi_loop, r_sisco_loop, r_notas_loop, 'CC', 'CC', 'CONTA CONTRATO').replace('.0', '')
            cli_sol = get_val(r_bi_loop, r_sisco_loop, r_notas_loop, 'NOME CLIENTE', 'Nome', 'NOME DO SOLICITANTE').upper()
            cid_sol = remover_acentos(get_val(r_bi_loop, r_sisco_loop, r_notas_loop, 'Município', 'Município', 'MUNICIPIO'))
            
            to_raw = get_val(r_bi_loop, r_sisco_loop, r_notas_loop, 'Tipo de Projeto Descrição', 'Detalhes', '')
            to_parts = to_raw.replace("-", " ").strip().split()
            to_sisco = " ".join(to_parts[:3]) if to_parts else ""
            
            pref_tipo = man_tipo_obra.split('-')[0] if man_tipo_obra else map_tipo_obra.get(to_sisco.upper(), "CT")
            pref_mun = man_mun.split('-')[0] if man_mun else map_mun.get(cid_sol, cid_sol[:3] if cid_sol else "XXX")
            
            val_sol_final = man_sol if man_sol else sol
            val_livre_final_nome = man_livre.replace(" ", "-")[:15] if man_livre else cli_sol.replace(" ", "-")[:15]
            val_livre_final_desc = man_livre.upper() if man_livre else cli_sol
            val_cc_final = man_cc if man_cc else cc_sol
            
            desc_str = f"{val_sol_final}-{val_livre_final_desc}, CC-{val_cc_final}."
            raw_name = f"{pref_especial}{pref_tipo}-{pref_pi}-{pref_mun}-{pref_id}-{val_sol_final}-{val_livre_final_nome}"
            nome_str = raw_name.replace(".", "").replace("_", "").replace(" ", "-")[:34].upper()
            
            if sol == solicitacoes[0]: obra_relampago_formatada = nome_str
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
    lbl_obra_estilo = 'class="lbl"' if not (man_tipo_obra or man_pi or man_livre or man_especial) else 'class="lbl text-red" style="background-color: #fef08a;"'
    val_obra_estilo = 'class="val text-green"' if not (man_tipo_obra or man_pi or man_livre or man_especial) else 'class="val text-red" style="background-color: #fef08a; font-style: italic;"'
    lbl_obra_texto = "⚡ Obra Relampago ⚡" if not (man_tipo_obra or man_pi or man_livre or man_especial) else "🚧 Nome da Obra 🚧"
    
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
