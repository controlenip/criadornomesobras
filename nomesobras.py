import streamlit as st
import pandas as pd
import unicodedata
import os
import base64
import re
import random

# ==========================================
# 1. CONFIGURAÇÕES DA PÁGINA E CSS (VISUAL MODERNO E PROFISSIONAL)
# ==========================================
st.set_page_config(page_title="Gerador SGO & Nomes de Obra", page_icon="🏗️", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    .block-container { padding-top: 1rem !important; padding-bottom: 2rem !important; }
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
    
    .obs-box { background-color: #1e293b; color: #f8fafc; border: 1px solid #cbd5e1; padding: 10px; font-family: ui-sans-serif, system-ui, sans-serif; font-size: 11px; font-style: italic; min-height: 80px; height: auto; white-space: pre-wrap; line-height: 1.4; border-radius: 0 0 4px 4px;}
    
    .desc-row { border: 1px solid #cbd5e1; height: 26px; width: 100%; margin-bottom: 4px; padding: 4px 8px; font-weight: 600; font-family: ui-monospace, monospace; font-size: 11px; text-transform: uppercase; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; background-color: white; color: #0f172a; border-radius: 4px; box-shadow: 0 1px 2px rgba(0,0,0,0.02);}
    
    .lbl-box { background-color: #fef08a; border: 1px solid #cbd5e1; border-radius: 4px; padding: 0px 8px; font-size: 11px; font-weight: 700; color: #7f1d1d; height: 35px; display: flex; align-items: center; margin-bottom: 0px; margin-top: 2px;}
    div[data-baseweb="select"] > div { border: 1px solid #cbd5e1; border-radius: 4px; min-height: 35px !important; height: 35px !important; font-size: 11px; background-color: white;}
    input[data-testid="stTextInput"] { border: 1px solid #cbd5e1; border-radius: 4px; height: 35px !important; min-height: 35px !important; font-size: 11px; font-weight: bold; background-color: white;}
    .stSelectbox, .stTextInput { margin-bottom: -10px !important; }
    .stTextArea textarea { border: 1px solid #94a3b8 !important; border-radius: 4px !important; font-size: 11px; font-family: ui-monospace, monospace; }
</style>
""", unsafe_allow_html=True)

def limpar_campos_manuais():
    for i in range(1, 10):
        chave = f"i{i}"
        if chave in st.session_state:
            st.session_state[chave] = ""
    if "text_area_obras" in st.session_state:
        st.session_state["text_area_obras"] = ""

def remover_acentos(texto):
    if pd.isna(texto) or texto == "": return ""
    texto = str(texto).upper().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def formatar_data(data_raw):
    try:
        if pd.isna(data_raw) or str(data_raw).lower() == 'nan' or str(data_raw).strip() == "":
            return ""
        data_obj = pd.to_datetime(data_raw)
        return data_obj.strftime('%d/%m/%Y')
    except:
        return str(data_raw)[:10]

@st.cache_data(show_spinner=False)
def carregar_dados(file):
    xls = pd.ExcelFile(file)
    
    try:
        df_sisco = pd.read_excel(xls, sheet_name='Sisco')
        if 'Nota CCS' in df_sisco.columns:
            df_sisco['Nota CCS'] = df_sisco['Nota CCS'].astype(str).str.replace('.0', '', regex=False)
    except:
        df_sisco = pd.DataFrame()
        
    try: 
        df_notas = pd.read_excel(xls, sheet_name='NotasSisgb')
    except: 
        try: df_notas = pd.read_excel(xls, sheet_name='NOTAS')
        except: df_notas = pd.DataFrame()
        
    if not df_notas.empty and 'PROTOCOLO' in df_notas.columns:
        df_notas['PROTOCOLO'] = df_notas['PROTOCOLO'].astype(str).str.replace('.0', '', regex=False)
        
    try: 
        df_dados = pd.read_excel(xls, sheet_name='DADOS', header=1)
    except:
        try: df_dados = pd.read_excel(xls, sheet_name='Dados', header=1)
        except: df_dados = pd.DataFrame()
    
    return df_sisco, df_notas, df_dados

# ==========================================
# 2. LOGO NO TOPO E DADOS PADRÃO
# ==========================================

st.markdown("<br>", unsafe_allow_html=True) 
if os.path.exists("LOGO_NIP.png"):
    with open("LOGO_NIP.png", "rb") as image_file:
        b64_logo = base64.b64encode(image_file.read()).decode()
    
    st.markdown(f'''
        <div style="text-align: center; margin-bottom: 10px;">
            <img src="data:image/png;base64,{b64_logo}" style="max-width: 150px; width: 100%; height: auto; pointer-events: none;">
        </div>
    ''', unsafe_allow_html=True)

col_btn1, col_btn2, col_btn3 = st.columns([2, 1, 2])
with col_btn2:
    st.button("🧹 Limpar Campos Manuais", on_click=limpar_campos_manuais, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

lista_tipos_obra = ['AF-AMPLIAÇÃO DE FASE', 'AP-AMPLIAÇÃO DE POTENCIA', 'CA-CONSTRUÇÃO DE AL', 'CT-CONSTRUÇÃO DE RD', 'DV-DIVISÃO DE CIRCUITO', 'FC-FLEXIBILIZAÇÃO DE CIRCUITO', 'IE-INSTALAÇÃO DE EQUIPAMENTOS', 'IT-INSTALAÇÃO DE TRANSFORMADORES', 'ME-MELHORIA DE REDE DE DISTRIBUIÇÃO', 'MI-MICROSSISTEMA ISOLADO DE GERAÇÃO DE ENERGIA', 'MP-REALOCAÇÃO DE POSTE', 'MT-REALOCAÇÃO DE TRANSFORMADORES RD', 'RC-RECAPACITAÇÃO DE CONDUTORES', 'RE-RECAPACITAÇÃO DE EQUIPAMENTOS DE RD', 'RF-RECAPACITAÇÃO DE C,FA, C,FU E PR', 'RP-RECAPACITAÇÃO DE POSTES', 'RT-RECAPACITAÇÃO DE TRANSFORMADOR DE RD', 'SI-SISTEMA INDIVIDUAL DE GERAÇÃO DE ENERGIA', 'TE-REALOCAÇÃO DE EQUIPAMENTOS']
lista_pi = ['ASC', 'ATV', 'BCP', 'BRE', 'BRT', 'CCF', 'DIF', 'DIS', 'EME', 'ERD', 'EUR', 'FIM', 'INC', 'INR', 'LPT', 'MBT', 'MCJ', 'MCR', 'MEL', 'MGD', 'MMT', 'MRS', 'MSE', 'MTP', 'NIV', 'OCP', 'ODS', 'PMC', 'REF', 'REG', 'SEG', 'SEQ', 'SID', 'SLS', 'SMC', 'TRI', 'UNI', 'UNP', 'UNR']
lista_mun = ['AAM-ALTO ALEGRE DO MARANHAO', 'AAP-ALTO ALEGRE DO PINDARE', 'ACL-ACAILANDIA', 'ACT-ALCANTARA', 'ADA-ALDEIAS ALTAS', 'ADM-AGUA DOCE DO MARANHAO', 'AFC-AFONSO CUNHA', 'ALM-ALTAMIRA DO MARANHAO', 'ALP-ALTO PARNAIBA', 'AME-ARAME', 'AMM-AMAPA DO MARANHAO', 'AMO-AMARANTE DO MARANHAO', 'ANA-ANAJATUBA', 'ANS-ANAPURUS', 'API-APICUM-ACU', 'ARA-ARAGUANA', 'ARI-ARARI', 'ARS-ARAIOSES', 'AXX-AXIXA', 'BAC-BACURI', 'BBR-BURITI BRAVO', 'BCA-BACABEIRA', 'BCB-BACABAL', 'BCP-BURITICUPU', 'BCT-BACURITUBA', 'BDC-BARRA DO CORDA', 'BEM-BERNARDO DO MEARIM', 'BGU-BELAGUA', 'BIV-BURITI', 'BJA-BREJO DE AREIA', 'BJD-BOM JARDIM', 'BJO-BREJO', 'BJS-BOM JESUS DAS SELVAS', 'BJU-BARAO DE GRAJAU', 'BLE-BENEDITO LEITE', 'BLS-BALSAS', 'BLU-BOM LUGAR', 'BQM-BEQUIMAO', 'BRN-BARREIRINHAS', 'BUT-BURITIRANA', 'BVG-BOA VISTA DO GURUPI', 'BVM-BELA VISTA DO MARANHAO', 'CAM-CAMPESTRE DO MARANHAO', 'CAN-CANDIDO MENDES', 'CAR-CAROLINA', 'CDL-CEDRAL', 'CGE-CENTRO DO GUILHERME', 'CGR-CACHOEIRA GRANDE', 'CHA-CHAPADINHA', 'CHE-CANTANHEDE', 'CID-CIDELANDIA', 'CJI-CAJARI', 'CJO-CAJAPIO', 'CLA-CONCEICAO DO LAGO-ACU', 'CMA-CENTRAL DO MARANHAO', 'CNM-CENTRO NOVO DO MARANHAO', 'CNO-COELHO NETO', 'COL-COLINAS', 'COO-CODO', 'CPN-CAPINZAL DO NORTE', 'CRA-COROATA', 'CRP-CURURUPU', 'CTP-CARUTAPERA', 'CXS-CAXIAS', 'DAV-DAVINOPOLIS', 'DBA-DUQUE BACELAR', 'DPO-DOM PEDRO', 'ESP-ESPERANTINOPOLIS', 'ETE-ESTREITO', 'FFA-FERNANDO FALCAO', 'FNM-FEIRA NOVA DO MARANHAO', 'FOR-FORTUNA', 'FSN-FORMOSA DA SERRA NEGRA', 'FTN-FORTALEZA DOS NOGUEIRAS', 'GDV-GODOFREDO VIANA', 'GEB-GOVERNADOR EUGENIO BARROS', 'GEL-GOVERNADOR EDISON LOBAO', 'GJU-GRAJAU', 'GLR-GOVERNADOR LUIZ ROCHA', 'GNB-GOVERNADOR NEWTON BELLO', 'GNF-GOVERNADOR NUNES FREIRE', 'GOA-GOVERNADOR ARCHER', 'GOD-GONCALVES DIAS', 'GRA-GRACA ARANHA', 'GUI-GUIMARAES', 'HUC-HUMBERTO DE CAMPOS', 'ICT-ICATU', 'IGG-IGARAPE GRANDE', 'IGM-IGARAPE DO MEIO', 'IPG-ITAIPAVA DO GRAJAU', 'IPZ-IMPERATRIZ', 'ITG-ITINGA DO MARANHAO', 'ITM-ITAPECURU MIRIM', 'JAT-JATOBA', 'JEV-JENIPAPO DOS VIEIRAS', 'JLB-JOAO LISBOA', 'JOS-JOSELANDIA', 'JUM-JUNCO DO MARANHAO', 'LAM-LAGOA DO MATO', 'LAN-LAJEADO NOVO', 'LGJ-LAGO DO JUNCO', 'LGM-LAGOA GRANDE DO MARANHAO', 'LGR-LAGO DOS RODRIGUES', 'LGV-LAGO VERDE', 'LIC-LIMA CAMPOS', 'LPD-LAGO DA PEDRA', 'LRT-LORETO', 'LUD-LUIS DOMINGUES', 'MAA-MAGALHAES DE ALMEIDA', 'MAL-MONTES ALTOS', 'MHO-MARANHAOZINHO', 'MIL-MILAGRES DO MARANHAO', 'MIR-MIRINZAL', 'MJS-MARAJA DO SENA', 'MME-MARACACUME', 'MON-MONCAO', 'MRA-MIRANDA DO NORTE', 'MRD-MIRADOR', 'MRR-MORROS', 'MTA-MATINHA', 'MTN-MATOES DO NORTE', 'MTR-MATA ROMA', 'MTS-MATOES', 'NCO-NOVA COLINAS', 'NIO-NOVA IORQUE', 'NRO-NINA RODRIGUES', 'NVO-NOVA OLINDA DO MARANHAO', "ODC-OLHO D'AGUA DAS CUNHAS", 'ONO-OLINDA NOVA DO MARANHAO', 'PAB-PASTOS BONS', 'PAF-PASSAGEM FRANCA', 'PAR-PAULO RAMOS', 'PCL-PACO DO LUMIAR', 'PCZ-PRIMEIRA CRUZ', 'PDR-PEDRO DO ROSARIO', 'PDS-PEDREIRAS', 'PDT-PRESIDENTE DUTRA', 'PFO-PORTO FRANCO', 'PHO-PINHEIRO', 'PIO-PIO XII', 'PJU-PRESIDENTE JUSCELINO', 'PMA-PALMEIRANDIA', 'PME-PRESIDENTE MEDICI', 'PMI-PINDARE-MIRIM', 'PNA-PARNARAMA', 'PNL-PENALVA', 'PNV-PAULINO NEVES', 'PPE-PIRAPEMAS', 'PPS-POCAO DE PEDRAS', 'PRB-PARAIBANO', 'PRM-PERI MIRIM', 'PRO-PERITORO', 'PSY-PRESIDENTE SARNEY', 'PTR-PORTO RICO DO MARANHAO', 'PVA-PRESIDENTE VARGAS', 'RAP-RAPOSA', 'RCO-RIACHAO', 'RFQ-RIBAMAR FIQUENE', 'RSO-ROSARIO', 'SAL-SANTO ANTONIO DOS LOPES', 'SAM-SANTO AMARO DO MARANHAO', 'SAR-SAO ROBERTO', 'SBN-SAO BERNARDO', 'SBR-SAO BENEDITO DO RIO PRETO', 'SBT-SAO BENTO', 'SBZ-SAO RAIMUNDO DO DOCA BEZERRA', 'SDM-SAO DOMINGOS DO MARANHAO', 'SDZ-SAO DOMINGOS DO AZEITAO', 'SER-SERRANO DO MARANHAO', 'SFB-SAO FELIX DE BALSAS', 'SFH-SAO FRANCISCO DO MARANHAO', 'SFJ-SAO FRANCISCO DO BREJAO', 'SFM-SANTA FILOMENA DO MARANHAO', 'SGM-SAO LUIS GONZAGA DO MARANHAO', 'SHL-SANTA HELENA', 'SJA-SAO JOAO BATISTA', 'SJB-SAO JOSE DOS BASILIOS', 'SJC-SAO JOAO DO CARU', 'SJI-SAO JOAO DO PARAISO', 'SJP-SAO JOAO DOS PATOS', 'SJR-SAO JOSE DE RIBAMAR', 'SJS-SAO JOAO DO SOTER', 'SLR-SENADOR LA ROCQUE', 'SLS-SAO LUIS', 'SMB-SAMBAIBA', 'SMH-SANTANA DO MARANHAO', 'SMT-SAO MATEUS DO MARANHAO', 'SNO-SITIO NOVO', 'SPB-SAO PEDRO DA AGUA BRANCA', 'SPC-SAO PEDRO DOS CRENTES', 'SQM-SANTA QUITERIA DO MARANHAO', 'SRI-SANTA RITA', 'SRM-SAO RAIMUNDO DAS MANGABEIRAS', 'STH-SATUBINHA', 'STI-SANTA INES', 'STL-SANTA LUZIA', 'STP-SANTA LUZIA DO PARUA', 'SUN-SUCUPIRA DO NORTE', 'SUR-SUCUPIRA DO RIACHAO', 'SVF-SAO VICENTE FERRER', 'SXC-SENADOR ALEXANDRE COSTA', 'TBR-TIMBIRAS', 'TFG-TASSO FRAGOSO', 'TMO-TIMON', 'TRL-TURILANDIA', 'TTA-TUTOIA', 'TTM-TUNTUM', 'TUF-TUFILANDIA', 'TUR-TURIACU', 'TVA-TRIZIDELA DO VALE', 'UBS-URBANO SANTOS', 'VFR-VITORINO FREIRE', 'VGG-VARGEM GRANDE', 'VNA-VIANA', 'VNM-VILA NOVA DOS MARTIRIOS', 'VTM-VITORIA DO MEARIM', 'ZDC-ZE DOCA']
lista_id = ['AL-Alimentador Tronco', 'BA-Barramento', 'CC-Conta Contrato', 'CO-Numero Componente', 'ID-IDENTIFICADOR', 'NR-Nota de Reclamação', 'NS-Nota CCS', 'OC-Ocorrência', 'OS-Ordem de Serviço', 'PF-CPF do cliente', 'PG-Ponto Geográfico', 'PT-Parecer Técnico', 'TR-Tempo Real']

arquivo_bd = st.sidebar.file_uploader("📥 Suba a planilha base (CRIAR NOME DA OBRA.xlsx)", type=["xlsx"])

df_sisco, df_notas, df_dados = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
map_tipo_obra, map_mun = {}, {}

if arquivo_bd:
    with st.spinner("Carregando banco de dados..."):
        df_sisco, df_notas, df_dados = carregar_dados(arquivo_bd)
        
    if not df_dados.empty:
        if 'TIPO DE OBRA' in df_dados.columns: lista_tipos_obra = sorted(df_dados['TIPO DE OBRA'].dropna().unique().tolist())
        if 'PI' in df_dados.columns: lista_pi = sorted(df_dados['PI'].dropna().unique().tolist())
        if 'SIGLA-MUNICIPIO' in df_dados.columns: lista_mun = sorted(df_dados['SIGLA-MUNICIPIO'].dropna().unique().tolist())
        if 'ID DO NUMERO' in df_dados.columns: lista_id = sorted([str(x).replace('.0', '') for x in df_dados['ID DO NUMERO'].dropna().unique().tolist()])

        # Prepara Mapeamentos Inteligentes
        if 'TIPO DE OBRA NO SISCO' in df_dados.columns and 'SIGLA' in df_dados.columns:
            df_to = df_dados.dropna(subset=['TIPO DE OBRA NO SISCO', 'SIGLA'])
            map_tipo_obra = dict(zip(df_to['TIPO DE OBRA NO SISCO'].astype(str).apply(remover_acentos), df_to['SIGLA'].astype(str).str.strip().str.upper()))
            
        if 'MUNICIPIO' in df_dados.columns and 'SIGLA.1' in df_dados.columns:
            df_mu = df_dados.dropna(subset=['MUNICIPIO', 'SIGLA.1'])
            map_mun = dict(zip(df_mu['MUNICIPIO'].astype(str).apply(remover_acentos), df_mu['SIGLA.1'].astype(str).str.strip().str.upper()))

c1, c2, c3, c4 = st.columns([0.8, 1.8, 2.5, 2.0])

# ==========================================
# COLUNA 1 - SOLICITAÇÕES INDIVIDUAIS
# ==========================================
with c1:
    st.markdown('<div class="eh">🎯 SOLICITAÇÕES</div>', unsafe_allow_html=True)
    
    st.markdown("<div style='padding: 8px 0px;'>", unsafe_allow_html=True)
    notas_associadas = st.checkbox("NOTAS ASSOCIADAS", value=True)
    st.markdown("</div>", unsafe_allow_html=True)
    
    sols_input = st.text_area("Cole as notas", key="text_area_obras", height=300, placeholder="Cole as notas aqui...", label_visibility="collapsed")
    
    solicitacoes = []
    if sols_input and sols_input.strip():
        parts = [p.strip() for p in re.split(r'[\s,;]+', sols_input.strip()) if p.strip()]
        
        notas_processadas = []
        for sol in parts:
            fase_sol = "MO"
            if not df_sisco.empty:
                r_s = df_sisco[df_sisco['Nota CCS'] == sol]
                if not r_s.empty:
                    f_temp = str(r_s.iloc[0].get('FASE', 'MO')).upper()
                    if f_temp not in ['NAN', 'NÃO ESPECIFICADO', 'NAO ESPECIFICADO', '']:
                        fase_sol = f_temp
            if fase_sol == "MO" and not df_notas.empty:
                r_n = df_notas[df_notas['PROTOCOLO'] == sol]
                if not r_n.empty:
                    f_temp = str(r_n.iloc[0].get('FASE', 'MO')).upper()
                    if f_temp not in ['NAN', 'NÃO ESPECIFICADO', 'NAO ESPECIFICADO', '']:
                        fase_sol = f_temp
            notas_processadas.append({'sol': sol, 'fase': fase_sol})
            
        tr_notes = [n['sol'] for n in notas_processadas if n['fase'] == 'TR']
        outras_notes = [n['sol'] for n in notas_processadas if n['fase'] != 'TR']
        
        if tr_notes:
            escolhida_tr = random.choice(tr_notes)
            tr_notes.remove(escolhida_tr)
            solicitacoes = [escolhida_tr] + tr_notes + outras_notes
        else:
            solicitacoes = parts

cc, instalacao, fase, tipo_obra_sisco, data_abertura, lat, lon, status_sap = "", "", "", "", "", "", "", ""
cidade_auto, cliente_auto, endereco_auto, area_resp, reg_raw, obs = "", "", "", "", "", ""
obs_extra = ""
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
            
        fase = str(r_sisco.get('FASE', 'MO')).upper() if r_sisco is not None else "MO"
        if fase.lower() == 'nan' or 'NÃO ESPECIFICADO' in fase or 'NAO ESPECIFICADO' in fase:
            fase = str(r_notas.get('FASE', 'MO')).upper() if r_notas is not None else "MO"
        if fase.lower() == 'nan' or 'NÃO ESPECIFICADO' in fase or 'NAO ESPECIFICADO' in fase: fase = "MO"
            
        tipo_obra_raw = str(r_notas.get('TIPO NOTA', '')) if r_notas is not None else ""
        if not tipo_obra_raw or tipo_obra_raw.lower() == 'nan':
            tipo_obra_raw = str(r_sisco.get('Detalhes', '')) if r_sisco is not None else ""
            
        if tipo_obra_raw.lower() == 'nan': tipo_obra_raw = ""
        parts = tipo_obra_raw.replace("-", " ").strip().split(" ")
        if len(parts) > 3:
            tipo_obra_sisco = " ".join(parts[:3])
        else:
            tipo_obra_sisco = tipo_obra_raw
            
        data_abertura_raw = str(r_notas.get('DATA ABERTURA', '')) if r_notas is not None else ""
        if not data_abertura_raw or data_abertura_raw.lower() == 'nan':
            data_abertura_raw = str(r_sisco.get('Data Abertura', '')) if r_sisco is not None else ""
        if not data_abertura_raw or data_abertura_raw.lower() == 'nan': 
            data_abertura_raw = str(r_notas.get('DATA DA SOLICITAÇÃO', '')) if r_notas is not None else ""
        data_abertura = formatar_data(data_abertura_raw)
            
        status_sap = str(r_notas.get('STATUS SAP', '')) if r_notas is not None else ""
        if status_sap.lower() == 'nan': status_sap = ""
            
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
        if not cliente_auto or cliente_auto.lower() == 'nan': cliente_auto = str(r_notas.get('NOME DO SOLICITANTE', r_notas.get('NOME', ''))) if r_notas is not None else ""
        cliente_auto = cliente_auto.upper()
            
        endereco_auto = str(r_notas.get('ENDEREÇO', '')) if r_notas is not None else ""
        if not endereco_auto or endereco_auto.lower() == 'nan': endereco_auto = str(r_sisco.get('Endereço', '')) if r_sisco is not None else ""
        
        pi_auto = str(r_notas.get('TIPO LIGAÇÃO', r_notas.get('TIPO NOTA', ''))) if r_notas is not None else ""
        if not pi_auto or pi_auto.lower() == 'nan': pi_auto = str(r_sisco.get('Tipo de Projeto(PI)', '')) if r_sisco is not None else ""
        if pi_auto.lower() == 'nan': pi_auto = ""
            
        reg_raw = str(r_notas.get('REGIONAL', '')) if r_notas is not None else ""
        if not reg_raw or reg_raw.lower() == 'nan': reg_raw = str(r_sisco.get('Regional', '')) if r_sisco is not None else ""
        reg_raw = reg_raw.upper()
        
        obs = str(r_sisco.get('INFORMAÇÕES', '')) if r_sisco is not None else ""
        if not obs or obs.lower() == 'nan': obs = str(r_notas.get('INFORMAÇÕES', '')) if r_notas is not None else ""
        if not obs or obs.lower() == 'nan': obs = str(r_sisco.get('Obs(última obs)', '')) if r_sisco is not None else ""
        if not obs or obs.lower() == 'nan': obs = str(r_notas.get('PONTO DE REFERENCIA', '')) if r_notas is not None else ""
        if obs.lower() == 'nan': obs = ""

        obs_extra = str(r_sisco.get('INFORMAÇÕES EXTRAS', '')) if r_sisco is not None else ""
        if not obs_extra or obs_extra.lower() == 'nan': obs_extra = str(r_notas.get('INFORMAÇÕES EXTRAS', '')) if r_notas is not None else ""
        if obs_extra.lower() == 'nan': obs_extra = ""
        
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
        <tr><td class="lbl">Tipo de Obra</td><td class="val">{tipo_obra_sisco}</td></tr>
        <tr><td class="lbl">Status SAP</td><td class="val">{status_sap}</td></tr>
        <tr><td class="lbl">Data Abertura</td><td class="val">{data_abertura}</td></tr>
        <tr><td class="lbl">Fase</td><td class="val">{fase}</td></tr>
        <tr><td class="lbl text-blue">LATITUDE</td><td class="val">{lat}</td></tr>
        <tr><td class="lbl text-blue">LONGITUDE</td><td class="val">{lon}</td></tr>
        <tr><td class="lbl text-blue" style="background-color: #f0fdf4;">LAT / LONG</td><td class="val" style="background-color: #f0fdf4;">{lat},{lon}</td></tr>
    </table>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="eh-yellow" style="margin-bottom: 0px;">🚧 Criar Nome da Obra Manual 🚧</div>', unsafe_allow_html=True)
    
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
        man_endereco = criar_linha_input("Endereço", "text", "i8")
        man_cc = criar_linha_input("Conta Contrato", "text", "i9")

# ==========================================
# 4. LÓGICA DE CRUZAMENTO DE DADOS E OVERRIDES MANUAIS
# ==========================================
pi_ativo = man_pi if man_pi else pi_auto

# Repasses Manuais para a Tabela SGO
if man_mun:
    cidade_auto = man_mun.split('-', 1)[1] if '-' in man_mun else man_mun
if man_livre:
    cliente_auto = man_livre.upper()
if man_endereco:
    endereco_auto = man_endereco.upper()
if man_cc:
    cc = man_cc

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
# 5. GERADOR EM MASSA DOS NOMES E DESCRIÇÕES
# ==========================================
pref_especial = f"{man_especial.split('-')[0]}-" if man_especial else ""
pref_tipo = man_tipo_obra.split('-')[0] if man_tipo_obra else "CT"
pref_pi = pi_ativo if pi_ativo else "UNR"
pref_id = man_id.split('-')[0] if man_id else "NS"

if not solicitacoes and (man_tipo_obra or man_pi or man_mun or man_id or man_sol or man_livre or man_endereco or man_cc):
    pref_mun = man_mun.split('-')[0] if man_mun else "XXX"
    val_sol_final = man_sol if man_sol else "0000000000"
    val_livre_final_nome = man_livre.replace(" ", "-")[:15] if man_livre else "NOME"
    val_livre_final_desc = man_livre.upper() if man_livre else "NOME"
    val_cc_final = man_cc if man_cc else "0000000000"
    
    raw_name = f"{pref_especial}{pref_tipo}-{pref_pi}-{pref_mun}-{pref_id}-{val_sol_final}-{val_livre_final_nome}"
    clean_name = raw_name.replace(".", "").replace("_", "").replace(" ", "-")
    obra_relampago_formatada = clean_name[:34].upper()
    
    fase_formatada = "(LIGAÇÃO MONOFÁSICA)" if fase.upper() == "MO" else "(LIGAÇÃO TRIFÁSICA)" if fase.upper() == "TR" else "(LIGAÇÃO BIFÁSICA)" if fase.upper() in ["BI", "BT"] else f"(FASE {fase})"
    desc_str = f"{val_sol_final}-{val_livre_final_desc}, CC-{val_cc_final} {fase_formatada}."
    
    nomes_obras_html += f'<div class="desc-row">{obra_relampago_formatada}</div>\n'
    descricoes_html += f'<div class="desc-row">{desc_str}</div>\n'
else:
    for idx, sol in enumerate(solicitacoes):
        res_sol_sisco = df_sisco[df_sisco['Nota CCS'] == sol] if not df_sisco.empty else pd.DataFrame()
        res_sol_notas = df_notas[df_notas['PROTOCOLO'] == sol] if not df_notas.empty else pd.DataFrame()
        
        if not res_sol_sisco.empty or not res_sol_notas.empty:
            r_sol_sisco = res_sol_sisco.iloc[0] if not res_sol_sisco.empty else None
            r_sol_notas = res_sol_notas.iloc[0] if not res_sol_notas.empty else None
            
            cc_sol = str(r_sol_sisco.get('CC', '')) if r_sol_sisco is not None else ""
            if not cc_sol or cc_sol.lower() == 'nan': cc_sol = str(r_sol_notas.get('CONTA CONTRATO', '')) if r_sol_notas is not None else ""
            cc_sol = cc_sol.replace('.0', '')
            
            cli_sol = str(r_sol_sisco.get('Nome', '')) if r_sol_sisco is not None else ""
            if not cli_sol or cli_sol.lower() == 'nan': 
                cli_sol = str(r_sol_notas.get('NOME DO SOLICITANTE', r_sol_notas.get('NOME', ''))) if r_sol_notas is not None else ""
            cli_sol = cli_sol.upper()
            
            cid_sol = str(r_sol_sisco.get('Município', '')) if r_sol_sisco is not None else ""
            if not cid_sol or cid_sol.lower() == 'nan': cid_sol = str(r_sol_notas.get('MUNICIPIO', '')) if r_sol_notas is not None else ""
            
            fase_sol = str(r_sol_sisco.get('FASE', 'MO')).upper() if r_sol_sisco is not None else "MO"
            if fase_sol.lower() == 'nan' or 'NÃO ESPECIFICADO' in fase_sol or 'NAO ESPECIFICADO' in fase_sol:
                fase_sol = str(r_sol_notas.get('FASE', 'MO')).upper() if r_sol_notas is not None else "MO"
            if fase_sol.lower() == 'nan' or 'NÃO ESPECIFICADO' in fase_sol or 'NAO ESPECIFICADO' in fase_sol: fase_sol = "MO"
            
            tipo_obra_raw_loop = str(r_sol_notas.get('TIPO NOTA', '')) if r_sol_notas is not None else ""
            if not tipo_obra_raw_loop or tipo_obra_raw_loop.lower() == 'nan':
                tipo_obra_raw_loop = str(r_sol_sisco.get('Detalhes', '')) if r_sol_sisco is not None else ""
                
            if tipo_obra_raw_loop.lower() == 'nan': tipo_obra_raw_loop = ""
            parts_to = tipo_obra_raw_loop.replace("-", " ").strip().split(" ")
            if len(parts_to) > 3:
                tipo_obra_sisco_loop = " ".join(parts_to[:3])
            else:
                tipo_obra_sisco_loop = tipo_obra_raw_loop

            cid_sol_limpo = remover_acentos(cid_sol)
            pref_mun = man_mun.split('-')[0] if man_mun else map_mun.get(cid_sol_limpo, cid_sol_limpo[:3] if cid_sol_limpo else "XXX")
            pref_tipo_loop = man_tipo_obra.split('-')[0] if man_tipo_obra else map_tipo_obra.get(remover_acentos(tipo_obra_sisco_loop), "CT")
            
            val_sol_final = man_sol if man_sol else sol
            
            val_livre_final_nome = man_livre.replace(" ", "-")[:15] if man_livre else cli_sol.replace(" ", "-")[:15]
            val_livre_final_desc = man_livre.upper() if man_livre else cli_sol
            
            val_cc_final = man_cc if man_cc else cc_sol
            
            fase_formatada = "(LIGAÇÃO MONOFÁSICA)" if fase_sol.upper() == "MO" else "(LIGAÇÃO TRIFÁSICA)" if fase_sol.upper() == "TR" else "(LIGAÇÃO BIFÁSICA)" if fase_sol.upper() in ["BI", "BT"] else f"(FASE {fase_sol})"
            desc_str = f"{val_sol_final}-{val_livre_final_desc}, CC-{val_cc_final} {fase_formatada}."
            
            raw_name = f"{pref_especial}{pref_tipo_loop}-{pref_pi}-{pref_mun}-{pref_id}-{val_sol_final}-{val_livre_final_nome}"
            clean_name = raw_name.replace(".", "").replace("_", "").replace(" ", "-")
            nome_str = clean_name[:34].upper()
            
            if idx == 0:
                obra_relampago_formatada = nome_str
        else:
            desc_str = f"{sol} - NÃO ENCONTRADO"
            nome_str = f"{sol} - NÃO ENCONTRADO"
            if idx == 0:
                obra_relampago_formatada = nome_str
            
        descricoes_html += f'<div class="desc-row">{desc_str}</div>\n'
        
        # Lógica de NOTAS ASSOCIADAS: Apenas a primeira ganha nome SGO, as outras ficam vazias para alinhamento
        if notas_associadas and idx > 0:
            nomes_obras_html += f'<div class="desc-row">&nbsp;</div>\n'
        else:
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
    <div class="obs-box" style="margin-bottom: 15px;">{obs}</div>

    <div class="eh-dark" style="text-align: left; padding-left: 10px;">"" MAIS OBSERVAÇÕES ABAIXO...</div>
    <div class="obs-box">{obs_extra}</div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown('<div class="eh">🖋 DESCRIÇÕES SGO 🖋</div>', unsafe_allow_html=True)
    st.markdown(descricoes_html, unsafe_allow_html=True)
    
    st.markdown('<div class="eh" style="margin-top: 15px;">🚧 NOMES DAS OBRAS 🚧</div>', unsafe_allow_html=True)
    st.markdown(nomes_obras_html, unsafe_allow_html=True)
