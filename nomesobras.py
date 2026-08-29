import streamlit as st
import pandas as pd
import unicodedata
import re

# ==========================================
# 1. CONFIGURAÇÕES DA PÁGINA
# ==========================================
st.set_page_config(page_title="Plataforma de Levantamento e Geração SGO", page_icon="🏗️", layout="wide")

def remover_acentos(texto):
    if pd.isna(texto) or texto == "":
        return ""
    texto = str(texto).upper().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return texto

@st.cache_data(show_spinner=False)
def carregar_dados(file):
    """Carrega as abas do Excel. Foco na aba Sisco."""
    xls = pd.ExcelFile(file)
    df_sisco = pd.read_excel(xls, sheet_name='Sisco')
    
    # Padroniza a coluna Nota CCS para string (removendo .0)
    if 'Nota CCS' in df_sisco.columns:
        df_sisco['Nota CCS'] = df_sisco['Nota CCS'].astype(str).str.replace('.0', '', regex=False)
        
    # Carrega a aba Dados (onde estão os gerentes/executivos), pulando a 1ª linha de cabeçalho duplo
    df_dados = pd.read_excel(xls, sheet_name='Dados', header=1)
    
    return df_sisco, df_dados

# ==========================================
# 2. CABEÇALHO E UPLOAD
# ==========================================
st.markdown("<h1 style='color: #1f4e78;'>🏗️ Plataforma de Levantamento e Geração SGO</h1>", unsafe_allow_html=True)
st.markdown("Busca automática de dados na base **Sisco** e formatação de Nomes de Obra.")
st.markdown("---")

arquivo_bd = st.sidebar.file_uploader("📥 Suba a planilha base (CRIAR NOME DA OBRA.xlsx)", type=["xlsx"])

if arquivo_bd:
    with st.spinner("Carregando banco de dados..."):
        df_sisco, df_dados = carregar_dados(arquivo_bd)
    st.sidebar.success("Base carregada com sucesso!")
    
    # ==========================================
    # 3. BARRA DE BUSCA
    # ==========================================
    solicitacao = st.text_input("🔍 Digite o número da Solicitação / Nota:", placeholder="Ex: 1080317771")
    
    if solicitacao:
        solicitacao = solicitacao.strip()
        
        # Filtra a aba Sisco
        resultado = df_sisco[df_sisco['Nota CCS'] == solicitacao]
        
        if not resultado.empty:
            row = resultado.iloc[0]
            
            # Extraindo variáveis
            cc = str(row.get('CC', '')).replace('.0', '')
            instalacao = str(row.get('INSTALACAO', '')).replace('.0', '')
            fase = str(row.get('Tipo de Carga', 'MO')).upper()
            data_abertura = str(row.get('Data Abertura', ''))
            lat = str(row.get('Latitude', ''))
            lon = str(row.get('Longitude', ''))
            
            cidade = remover_acentos(row.get('Município', ''))
            cliente = str(row.get('Nome', '')).upper()
            endereco = str(row.get('Endereço', ''))
            area_resp = str(row.get('Descrição', 'EXPANSÃO'))
            regional = str(row.get('Regional', ''))
            
            observacoes = str(row.get('Obs(última obs)', ''))
            if observacoes.lower() == 'nan': observacoes = ""
            
            # Formatação de Nomes
            cliente_curto = cliente.replace(" ", "-")[:15]
            sigla_mun = cidade[:3] if cidade else "XXX"
            
            descricao_sgo = f"{solicitacao}-{cliente}, CC-{cc}."
            nome_obra_sugerido = f"CT-UNR-{sigla_mun}-NS-{solicitacao}-{cliente_curto}"
            
            # ==========================================
            # 4. RENDERIZAÇÃO IDÊNTICA AO EXCEL (HTML/CSS)
            # ==========================================
            col1, col2, col3 = st.columns([1, 1.3, 1])
            
            # --- COLUNA 1: DADOS ---
            with col1:
                st.markdown(f"""
                <div style="border: 3px solid black; margin-bottom: 10px;">
                    <div style="background-color: #00b050; color: white; text-align: center; font-weight: 900; padding: 5px; border-bottom: 3px solid black; font-size: 18px;">
                        🎲 DADOS
                    </div>
                    <table style="width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 14px;">
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold; width: 40%;">Conta Contrato</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">{cc}</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">Instalação CCS</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">{instalacao}</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">FASE</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">{fase}</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">Tipo de Obra</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;"></td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">Data de Abert.</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">{data_abertura}</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">---</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;"></td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold; color: #0070c0;">LAT</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">{lat}</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold; color: #0070c0;">LONG</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">{lon}</td></tr>
                    </table>
                </div>
                """, unsafe_allow_html=True)

            # --- COLUNA 2: CRIAÇÃO DA NOTA ---
            with col2:
                st.markdown(f"""
                <div style="border: 3px solid black; margin-bottom: 10px;">
                    <div style="background-color: #00b050; color: white; text-align: center; font-weight: 900; padding: 5px; border-bottom: 3px solid black; font-size: 18px;">
                        📝 CRIAÇÃO DA NOTA SGO
                    </div>
                    <table style="width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 14px;">
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold; width: 35%;">Tipo Nota | Parc.</td><td style="border: 1px solid black; padding: 4px; font-weight: bold; color: #c00000;">SOLICITAÇÃO CLIENTE | CLIENTE</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">Regional</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">{regional}</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">Cidade</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">{cidade}</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">Área Responsável</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">{area_resp}</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">Cliente</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">{cliente}</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">Endereço</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">{endereco}</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">PI</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">UNR</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">Responsável Obra</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">EXCLUSIVA DA DISTRIBUIDORA</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">Valor Previsto</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">R$ 7.000,00</td></tr>
                    </table>
                </div>
                """, unsafe_allow_html=True)
                
                # Quadro de Aprovação da Nota
                st.markdown(f"""
                <div style="border: 3px solid black; margin-bottom: 10px; margin-top: 20px;">
                    <div style="background-color: #00b050; color: white; text-align: center; font-weight: 900; padding: 5px; border-bottom: 3px solid black; font-size: 15px;">
                        👍 APROVAÇÃO DA NOTA SGO
                    </div>
                    <table style="width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 14px;">
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold; width: 35%;">Empresa</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">DPL - 3º</td></tr>
                        <tr><td style="border: 1px solid black; padding: 4px; font-weight: bold;">Técnico</td><td style="border: 1px solid black; padding: 4px; font-weight: bold;">Buscar...</td></tr>
                    </table>
                </div>
                """, unsafe_allow_html=True)

            # --- COLUNA 3: DESCRIÇÃO E NOME DA OBRA ---
            with col3:
                st.markdown(f"""
                <div style="border: 3px solid black; margin-bottom: 10px;">
                    <div style="background-color: #00b050; color: white; text-align: center; font-weight: 900; padding: 5px; border-bottom: 3px solid black; font-size: 18px;">
                        🖋 DESCRIÇÃO SGO
                    </div>
                    <div style="padding: 10px; font-weight: bold; font-family: monospace; font-size: 14px; text-align: center;">
                        {descricao_sgo}
                    </div>
                </div>
                """, unsafe_allow_html=True)
                # O botão nativo de copy do Streamlit
                st.code(descricao_sgo, language="text")
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                st.markdown(f"""
                <div style="border: 3px solid black; margin-bottom: 10px;">
                    <div style="background-color: white; color: #c00000; text-align: center; font-weight: 900; padding: 5px; border-bottom: 3px solid black; font-size: 18px;">
                        🚧 Criar Nome da Obra 🚧
                    </div>
                    <div style="padding: 10px; font-weight: bold; font-family: monospace; font-size: 14px; color: #00b050; text-align: center;">
                        {nome_obra_sugerido}
                    </div>
                </div>
                """, unsafe_allow_html=True)
                st.code(nome_obra_sugerido, language="text")

            # --- RODAPÉ: OBSERVAÇÕES ---
            st.markdown(f"""
            <div style="border: 3px solid black; margin-top: 20px;">
                <div style="background-color: #00b050; color: white; font-weight: 900; padding: 5px; border-bottom: 3px solid black; font-size: 18px;">
                    💬 MAIS OBSERVAÇÕES A ACRESCENTAR NA NOTA
                </div>
            </div>
            """, unsafe_allow_html=True)
            st.text_area("Observações extraídas (Backoffice/Sisco):", value=observacoes, height=180, label_visibility="collapsed")
            
        else:
            st.error(f"❌ A Solicitação '{solicitacao}' não foi encontrada na aba Sisco.")
else:
    st.info("👈 Por favor, faça o upload da planilha base no menu lateral esquerdo para começar.")
