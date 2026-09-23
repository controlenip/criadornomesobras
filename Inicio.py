import streamlit as st
import os
import base64

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Sistema de Obras NIP",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==========================================
# ESTILO VISUAL
# ==========================================
st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.2rem !important;
            padding-bottom: 2rem !important;
            max-width: 1200px;
        }

        .hero {
            text-align: center;
            padding: 10px 10px 18px 10px;
            overflow: visible !important;
        }

        .logo-wrap {
            width: 100%;
            min-height: 105px;
            display: flex;
            align-items: center;
            justify-content: center;
            overflow: visible !important;
            padding: 6px 0 4px 0;
            box-sizing: border-box;
        }

        .logo-wrap img {
            display: block;
            width: auto !important;
            height: auto !important;
            max-width: 190px !important;
            max-height: 100px !important;
            object-fit: contain !important;
            object-position: center center;
            margin: 0 auto;
        }

        .hero-title {
            font-size: 30px;
            font-weight: 800;
            color: #0f172a;
            margin: 6px 0 2px 0;
        }

        .hero-subtitle {
            font-size: 14px;
            color: #64748b;
            margin: 0;
        }

        .section-title {
            font-size: 18px;
            font-weight: 800;
            color: #0f172a;
            margin: 20px 0 10px 0;
        }

        .page-card {
            background: #ffffff;
            border: 1px solid #dbe3ec;
            border-radius: 10px;
            padding: 18px 20px;
            min-height: 365px;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.05);
        }

        .page-card-header {
            background: #059669;
            color: #ffffff;
            font-size: 16px;
            font-weight: 800;
            padding: 10px 12px;
            border-radius: 7px;
            margin-bottom: 14px;
            text-align: center;
        }

        .page-card p {
            color: #334155;
            font-size: 13px;
            line-height: 1.55;
            margin-bottom: 10px;
        }

        .page-card ul {
            margin: 8px 0 0 20px;
            padding: 0;
            color: #334155;
            font-size: 13px;
            line-height: 1.55;
        }

        .tip-box {
            margin-top: 16px;
            background: #f8fafc;
            border-left: 4px solid #059669;
            border-radius: 6px;
            padding: 12px 14px;
            color: #334155;
            font-size: 13px;
            line-height: 1.5;
        }

        .flow-box {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 14px 16px;
            font-size: 13px;
            color: #334155;
            line-height: 1.55;
        }

        .footer-note {
            text-align: center;
            color: #94a3b8;
            font-size: 11px;
            margin-top: 28px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# LOGO / CABEÇALHO
# ==========================================
# A logo abaixo usa exatamente o mesmo padrão da página CRIAR SGO,
# evitando qualquer recorte causado por contêineres com altura limitada.
st.markdown("<br>", unsafe_allow_html=True)
if os.path.exists("LOGO_NIP.png"):
    with open("LOGO_NIP.png", "rb") as image_file:
        b64_logo = base64.b64encode(image_file.read()).decode()

    st.markdown(f'''
        <div style="text-align: center; margin-bottom: 10px;">
            <img src="data:image/png;base64,{b64_logo}" style="max-width: 150px; width: 100%; height: auto; pointer-events: none;">
        </div>
    ''', unsafe_allow_html=True)

st.markdown(
    """
    <div class="hero">
        <div class="hero-title">Sistema de Obras</div>
        <p class="hero-subtitle">Acesso rápido às ferramentas de criação de SGO e relatório de expurgo.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="section-title">📌 Como funciona cada página</div>',
    unsafe_allow_html=True,
)

# ==========================================
# RESUMO DAS PÁGINAS
# ==========================================
col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown(
        """
        <div class="page-card">
            <div class="page-card-header">🏗️ CRIAR SGO</div>
            <p>
                Página destinada à consulta das solicitações e à preparação dos dados necessários
                para criação da nota SGO e dos nomes das obras.
            </p>
            <ul>
                <li>Cole uma ou várias notas no campo <b>SOLICITAÇÕES</b>.</li>
                <li>O sistema consulta automaticamente a base <b>BASE_LEVANTAMENTO_ATUALIZADA.xlsx</b>.</li>
                <li>Notas canceladas ou finalizadas são identificadas e retiradas do processamento.</li>
                <li>Com <b>NOTAS ASSOCIADAS</b> marcada, a regra de priorização da obra trifásica pode ser aplicada.</li>
                <li>Com a opção desmarcada, as solicitações seguem a ordem informada pelo usuário.</li>
                <li>São exibidos dados da obra, cliente, endereço, fase, regional, PI, coordenadas e demais informações cadastradas.</li>
                <li>Os campos manuais permitem ajustar ou criar o nome da obra quando necessário.</li>
                <li>Ao final, a página gera <b>DESCRIÇÕES SGO</b>, <b>NOMES DAS OBRAS</b>, dados de criação e dados de aprovação da nota.</li>
            </ul>
            <div class="tip-box">
                <b>Quando utilizar:</b> preparação e conferência das informações antes da criação da nota no SGO,
                principalmente quando existem várias solicitações relacionadas à mesma obra.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col2:
    st.markdown(
        """
        <div class="page-card">
            <div class="page-card-header">📊 RELATÓRIO DE EXPURGO</div>
            <p>
                Página utilizada para montar o formulário de não atendimento/expurgo com os dados
                da solicitação e a evidência de campo.
            </p>
            <ul>
                <li>Informe o <b>Nº da Nota / Protocolo</b>.</li>
                <li>O sistema busca automaticamente regional, data da solicitação, conta contrato, parceiro e endereço na base.</li>
                <li>Anexe uma <b>foto da evidência</b> em PNG, JPG ou JPEG.</li>
                <li>Complete os dados da visita, como data, horário, latitude, longitude e identificação da equipe.</li>
                <li>Preencha a justificativa, descrição do expurgo e tratativa no sistema comercial.</li>
                <li>Informe, quando necessário, medidor do cliente, medidor do vizinho e estrutura mais próxima.</li>
                <li>A foto anexada é incluída diretamente no formulário.</li>
                <li>O formulário pode ser enviado para impressão ou salvo em PDF pelo recurso de impressão da página.</li>
            </ul>
            <div class="tip-box">
                <b>Quando utilizar:</b> registro formal de uma situação de não atendimento ou expurgo,
                reunindo dados da nota, informações de campo e evidência fotográfica em um único formulário.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ==========================================
# FLUXO RÁPIDO
# ==========================================
st.markdown(
    '<div class="section-title">🧭 Fluxo rápido de utilização</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="flow-box">
        <b>1.</b> Utilize o menu lateral para escolher a ferramenta desejada.<br>
        <b>2.</b> Para criar ou conferir uma SGO, acesse <b>CRIAR SGO</b> e informe as solicitações.<br>
        <b>3.</b> Para documentar um não atendimento/expurgo, acesse <b>RELATÓRIO DE EXPURGO</b>, informe a nota e anexe a evidência.<br>
        <b>4.</b> Confira os dados preenchidos automaticamente antes de utilizar as informações ou gerar o documento final.
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="footer-note">Sistema de apoio às rotinas de obras — NIP</div>',
    unsafe_allow_html=True,
)
