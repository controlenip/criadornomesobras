import streamlit as st
import pandas as pd
import unicodedata

# ==========================================
# 1. CONFIGURAÇÕES E FUNÇÕES AUXILIARES
# ==========================================
st.set_page_config(page_title="Gerador SGO & Nomes de Obra", page_icon="🏗️", layout="wide")

def remover_acentos(texto):
    if pd.isna(texto) or texto == "":
        return ""
    texto = str(texto).upper().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return texto

@st.cache_data(show_spinner=False)
def carregar_bancos_de_dados(file):
    """Carrega as abas do Excel para a memória (Pandas) de forma ultrarrápida."""
    xls = pd.ExcelFile(file)
    
    # Pula as 2 primeiras linhas em branco da aba BI
    df_bi = pd.read_excel(xls, sheet_name='BI', header=2)
    
    # A aba Sisco já tem o cabeçalho na primeira linha (header=0 é o padrão)
    df_sisco = pd.read_excel(xls, sheet_name='Sisco')
    
    # A aba Dados tem o cabeçalho real na segunda linha (header=1)
    df_dados = pd.read_excel(xls, sheet_name='Dados', header=1)
    
    # Padroniza as colunas de chave (Nota CCS / Protocolo) para string para evitar erros de busca
    if 'Nota CCS' in df_bi.columns: 
        df_bi['Nota CCS'] = df_bi['Nota CCS'].astype(str).str.replace('.0', '', regex=False)
    if 'Nota CCS' in df_sisco.columns: 
        df_sisco['Nota CCS'] = df_sisco['Nota CCS'].astype(str).str.replace('.0', '', regex=False)
    
    return df_bi, df_sisco, df_dados

# ==========================================
# 2. INTERFACE PRINCIPAL
# ==========================================
st.title("🏗️ Plataforma de Levantamento e Geração SGO")
st.markdown("Busca automática de dados em bases `BI`, `Sisco` e formatação de Nomes de Obra.")

# UPLOAD DO ARQUIVO DE BANCO DE DADOS
arquivo_bd = st.sidebar.file_uploader("📥 Suba a planilha base (CRIAR NOME DA OBRA.xlsx)", type=["xlsx"])

if arquivo_bd:
    with st.spinner("Carregando bancos de dados..."):
        df_bi, df_sisco, df_dados = carregar_bancos_de_dados(arquivo_bd)
    st.sidebar.success("Bancos carregados com sucesso!")

    st.markdown("---")
    
    # CAMPO DE BUSCA
    solicitacao = st.text_input("🔍 Digite o número da Solicitação / Nota:", placeholder="Ex: 1080317771")
    
    if solicitacao:
        solicitacao = solicitacao.strip()
        
        # LÓGICA DE BUSCA (Substitui os XLOOKUPs aninhados do Excel)
        # Tenta achar na aba BI primeiro, se não achar, procura no Sisco
        resultado_bi = df_bi[df_bi['Nota CCS'] == solicitacao]
        resultado_sisco = df_sisco[df_sisco['Nota CCS'] == solicitacao]
        
        encontrado = False
        dados_extraidos = {}
        
        if not resultado_bi.empty:
            encontrado = True
            row = resultado_bi.iloc[0]
            dados_extraidos = {
                "Conta Contrato": row.get('CC', ''),
                "Instalação": row.get('Instalação', ''),
                "Fase": row.get('FASE', 'MO'),
                "Cliente": row.get('NOME CLIENTE', ''),
                "Municipio": row.get('Município', ''),
                "Lat": row.get('LATITUDE', ''),
                "Lon": row.get('LONGITUDE', ''),
                "Texto": row.get('TEXTO_GERAL', '')
            }
        elif not resultado_sisco.empty:
            encontrado = True
            row = resultado_sisco.iloc[0]
            dados_extraidos = {
                "Conta Contrato": row.get('CC', ''),
                "Instalação": row.get('INSTALACAO', ''),
                "Fase": row.get('Tipo de Carga', 'MO'),
                "Cliente": row.get('Nome', ''),
                "Municipio": row.get('Município', ''),
                "Lat": row.get('Latitude', ''),
                "Lon": row.get('Longitude', ''),
                "Texto": row.get('Obs(última obs)', '')
            }
            
        if encontrado:
            cc = str(dados_extraidos['Conta Contrato']).replace('.0','')
            cliente = dados_extraidos['Cliente']
            municipio_limpo = remover_acentos(dados_extraidos['Municipio'])
            
            # Geração das Strings SGO
            descricao_sgo = f"{solicitacao}-{cliente}, CC-{cc}."
            
            # Layout da Tela
            col1, col2, col3 = st.columns([1, 1.5, 1])
            
            with col1:
                st.markdown("### 🎲 DADOS")
                st.info(f"**Conta Contrato:** {cc}\n\n**Instalação:** {str(dados_extraidos['Instalação']).replace('.0','')}\n\n**FASE:** {dados_extraidos['Fase']}")
                st.warning(f"**LAT:** {dados_extraidos['Lat']}\n\n**LONG:** {dados_extraidos['Lon']}")
                
            with col2:
                st.markdown("### 📝 CRIAÇÃO DA NOTA SGO")
                st.success(f"**Tipo Nota | Parceiro:** SOLICITAÇÃO CLIENTE | CLIENTE")
                st.write(f"**Cidade:** {municipio_limpo}")
                st.write(f"**Cliente:** {cliente}")
                st.write(f"**Endereço:** Buscar na base...")
                
            with col3:
                st.markdown("### 🖋 DESCRIÇÃO SGO")
                st.code(descricao_sgo, language="text")
                
                st.markdown("### 🚧 NOME DA OBRA")
                # Montagem complexa do nome do projeto
                nome_cliente_curto = str(cliente).replace(" ", "-")[:15]
                sigla_mun = municipio_limpo[:3] if municipio_limpo else "XXX"
                nome_obra = f"CT-UNR-{sigla_mun}-NS-{solicitacao}-{nome_cliente_curto}"
                st.code(nome_obra, language="text")
                
            st.markdown("### 📜 MAIS OBSERVAÇÕES ACRESCENTAR NA NOTA")
            # Área de texto
            st.text_area("Observações extraídas (Backoffice/Sisco/BI):", value=dados_extraidos['Texto'], height=150)
            
        else:
            st.error(f"❌ A Solicitação '{solicitacao}' não foi encontrada nas bases BI nem Sisco.")
else:
    st.info("👈 Por favor, faça o upload da planilha base no menu lateral esquerdo para começar.")
