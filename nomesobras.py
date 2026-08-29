import streamlit as st
import pandas as pd
import unicodedata
import re

# ==========================================
# 1. CONFIGURAÇÕES DA PÁGINA E CSS (VISUAL EXCEL)
# ==========================================
st.set_page_config(page_title="Gerador SGO & Nomes de Obra", page_icon="🏗️", layout="wide")

# Injeção de CSS para imitar a interface do Excel
st.markdown("""
<style>
    /* Esconde os menus nativos do Streamlit para parecer um sistema limpo */
    [data-testid="stHeader"] { display: none !important; }
    
    /* Estilos das Tabelas e Cabeçalhos imitando o Excel */
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
    
    .obs-box { background-color: #595959; color: white; border: 3px solid black; padding: 8px; font-family: Calibri, Arial, sans-serif; font-size: 12px; font-weight: bold; font-style: italic; min-height: 300px; white-space: pre-wrap; line-height: 1.2; overflow-y: auto;}
    .desc-box { background-color: white; color: black; border: 3px solid black; padding: 4px 6px; font-family: 'Arial Black', sans-serif; font-size: 13px; font-weight: 900; text-transform: uppercase; margin-bottom: 15px;}
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
    # Tenta carregar as outras abas (fallback caso existam)
    try: df_dados = pd.read_excel(xls, sheet_name='Dados', header=1)
    except: df_dados = pd.DataFrame()
    return df_sisco, df_dados

# ==========================================
# 2. MENU LATERAL E UPLOAD
# ==========================================
arquivo_bd = st.sidebar.file_uploader("📥 Suba a planilha base (CRIAR NOME DA OBRA.xlsx)", type=["xlsx"])

if arquivo_bd:
    with st.spinner("Carregando banco de dados..."):
        df_sisco, df_dados = carregar_dados(arquivo_bd)
    
    # ==========================================
    # 3. CONSTRUÇÃO DO GRID (4 COLUNAS COMO NO EXCEL)
    # ==========================================
    # Proporções simulando a largura das colunas A, B:E, F:I, J no Excel
    c1, c2, c3, c4 = st.columns([0.8, 1.8, 2.5, 1.5])
    
    with c1:
        st.markdown('<div class="eh">🎯 SOLICITAÇÃO</div>', unsafe_allow_html=True)
        solicitacao = st.text_input("", placeholder="1080317771", label_visibility="collapsed")
        
        # Simula as linhas vazias da coluna A do Excel
        for _ in range(25):
            st.markdown('<div style="border: 1px solid #ccc; height: 23px; width: 100%;"></div>', unsafe_allow_html=True)
            
    if solicitacao:
        solicitacao = solicitacao.strip()
        resultado = df_sisco[df_sisco['Nota CCS'] == solicitacao]
        
        if not resultado.empty:
            row = resultado.iloc[0]
            
            # --- MAPEAMENTO DE VARIÁVEIS ---
            cc = str(row.get('CC', '')).replace('.0', '')
            instalacao = str(row.get('INSTALACAO', '')).replace('.0', '')
            fase = str(row.get('Tipo de Carga', 'MO')).upper()
            tipo_obra_sisco = str(row.get('Tipo de Projeto Descrição', ''))
            data_abertura = str(row.get('Data Abertura', ''))
            lat = str(row.get('Latitude', '')).replace('.', ',')
            lon = str(row.get('Longitude', '')).replace('.', ',')
            
            cidade = remover_acentos(row.get('Município', ''))
            cliente = str(row.get('Nome', '')).upper()
            endereco = str(row.get('Endereço', ''))
            area_resp = str(row.get('Descrição', 'EXPANSÃO')).upper()
            
            # Lógica Condicional da Regional
            reg_raw = str(row.get('Regional', '')).upper()
            regional_formatado = "CM04-IMPERATRIZ"
            if "SUL" in reg_raw: regional_formatado = "CM04-IMPERATRIZ"
            elif "CENTRO" in reg_raw: regional_formatado = "CM03-BACABAL"
            elif "LESTE" in reg_raw: regional_formatado = "CM02-TIMON"
            elif "NORTE" in reg_raw: regional_formatado = "CM01-SAO LUIS"
            elif "NOROESTE" in reg_raw: regional_formatado = "CM01-PINHEIRO"
            
            obs = str(row.get('Obs(última obs)', ''))
            if obs.lower() == 'nan': obs = ""
            
            # Formatação de Nomes da Obra e Descrição
            cliente_curto = cliente.replace(" ", "-")[:15]
            sigla_mun = cidade[:3] if cidade else "XXX"
            
            descricao_sgo = f"{solicitacao}-{cliente}, CC-{cc}."
            obra_relampago = f"CT-UNR-{sigla_mun}-NS-{solicitacao}-{cliente_curto}"
            
            # Buscas simuladas da aba DADOS (Gerente, Executivo, etc)
            pi = "UNR"
            gerente = "RENAN ACCIOLY PIMENTEL"
            executivo = "LETICIA CRISTINA DA SILVA BARATA"
            empresa = "DPL - 3º"
            contrato = "4600024692"
            tecnico = "JOSÉ DE ARAUJO ANDRADE JUNIOR"
            data_aprov = "07/03/2027 (190 DIAS)"
            
            # --- COLUNA 2: DADOS ---
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
                    <tr><td class="lbl" style="font-style: italic;">Obra Especial ?</td><td class="val" style="text-align: right; text-transform: none; font-weight: 900;">Normal</td></tr>
                    <tr><td class="lbl">Tipo de Obra</td><td class="val"></td></tr>
                    <tr><td class="lbl">PI</td><td class="val"></td></tr>
                    <tr><td class="lbl">Municipio</td><td class="val"></td></tr>
                    <tr><td class="lbl">ID do Numero</td><td class="val"></td></tr>
                    <tr><td class="lbl">Solicitação</td><td class="val"></td></tr>
                    <tr><td class="lbl">Escrita Livre</td><td class="val"></td></tr>
                </table>
                """, unsafe_allow_html=True)

            # --- COLUNA 3: CRIAÇÃO DA NOTA E OBSERVAÇÕES ---
            with c3:
                st.markdown(f"""
                <div class="eh">📝 CRIAÇÃO DA NOTA SGO 📝</div>
                <table class="et">
                    <tr><td class="lbl">Tipo Nota | Parc</td><td class="val text-red" style="font-size: 11px;">SOLICITAÇÃO CLIENTE | CLIENTE</td></tr>
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
                    <tr><td class="lbl">Valor Previsto</td><td class="val text-green">R$ 7.000,00</td></tr>
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

            # --- COLUNA 4: DESCRIÇÃO SGO ---
            with c4:
                st.markdown(f"""
                <div class="eh">🖋 DESCRIÇÃO SGO 🖋</div>
                <div class="desc-box">{descricao_sgo}</div>
                """, unsafe_allow_html=True)
                
                # Linhas vazias imitando o lado direito da planilha
                for _ in range(35):
                    st.markdown('<div style="border: 2px solid black; height: 18px; width: 100%; margin-bottom: 2px;"></div>', unsafe_allow_html=True)
        else:
            with c2:
                st.error("❌ Nota não encontrada no Sisco.")
else:
    st.info("👈 Faça o upload da planilha base para começar.")
