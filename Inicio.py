import streamlit as st
import os
import base64
from pathlib import Path

# ==========================================
# CONFIGURACAO DA PAGINA
# ==========================================
st.set_page_config(
    page_title="Sistema de Obras NIP",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

ROOT = Path(__file__).resolve().parent

# ==========================================
# FUNCOES DE APOIO
# ==========================================
def localizar_pagina(candidatos):
    """Localiza uma pagina existente dentro de /pages sem depender do nome exato."""
    for relativo in candidatos:
        caminho = ROOT / relativo
        if caminho.exists():
            return relativo.replace("\\", "/")
    return None


def contar_extensoes(pasta, extensoes):
    caminho = ROOT / pasta
    if not caminho.is_dir():
        return 0
    return sum(
        1 for arq in caminho.iterdir()
        if arq.is_file() and arq.suffix.lower() in extensoes
    )


def status_item(ok, titulo, detalhe=""):
    icone = "✅" if ok else "⚠️"
    classe = "status-ok" if ok else "status-warn"
    detalhe_html = f"<span>{detalhe}</span>" if detalhe else ""
    return f'<div class="status-row {classe}"><b>{icone} {titulo}</b>{detalhe_html}</div>'


PAGE_SGO = localizar_pagina([
    "pages/1_CRIAR_SGO.py",
    "pages/1_Criar_SGO.py",
])
PAGE_EXPURGO = localizar_pagina([
    "pages/2_RELATORIO_DE_EXPURGO.py",
    "pages/2_Relatorio_de_Expurgo.py",
])
PAGE_MAPA = localizar_pagina([
    "pages/3_MAPA.py",
    "pages/3_Mapa.py",
])

base_ok = (ROOT / "BASE_LEVANTAMENTO_ATUALIZADA.xlsx").exists()
municipios_ok = (ROOT / "MUNICIPIOS-REGIONAIS.xlsx").exists()
kmls_ok = (ROOT / "kmls").is_dir()
kmzs_ok = (ROOT / "kmzs").is_dir()
logo_ok = (ROOT / "LOGO_NIP.png").exists()
qtd_kml = contar_extensoes("kmls", {".kml"})
qtd_kmz = contar_extensoes("kmzs", {".kmz", ".kml"})

# ==========================================
# ESTILO VISUAL
# ==========================================
st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1.0rem !important;
            padding-bottom: 2rem !important;
            max-width: 1320px;
        }

        .hero {
            text-align: center;
            padding: 6px 10px 10px 10px;
        }
        .hero-title {
            font-size: 31px;
            font-weight: 850;
            color: #0f172a;
            margin: 3px 0 3px 0;
        }
        .hero-subtitle {
            font-size: 14px;
            color: #64748b;
            margin: 0;
        }
        .hero-badge {
            display: inline-block;
            margin-top: 9px;
            padding: 5px 10px;
            background: #ecfdf5;
            color: #047857;
            border: 1px solid #a7f3d0;
            border-radius: 999px;
            font-size: 11px;
            font-weight: 800;
        }

        .section-title {
            font-size: 18px;
            font-weight: 850;
            color: #0f172a;
            margin: 20px 0 10px 0;
        }

        .metric-card {
            background: #ffffff;
            border: 1px solid #dbe3ec;
            border-radius: 10px;
            padding: 13px 15px;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
            min-height: 80px;
        }
        .metric-label {
            color: #64748b;
            font-size: 11px;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: .4px;
        }
        .metric-value {
            color: #0f172a;
            font-size: 22px;
            font-weight: 850;
            margin-top: 5px;
        }
        .metric-note {
            color: #94a3b8;
            font-size: 10px;
            margin-top: 2px;
        }

        .page-card {
            background: #ffffff;
            border: 1px solid #dbe3ec;
            border-radius: 12px;
            padding: 16px 18px;
            min-height: 430px;
            box-shadow: 0 2px 10px rgba(15, 23, 42, 0.05);
        }
        .page-card-header {
            color: #ffffff;
            font-size: 16px;
            font-weight: 850;
            padding: 10px 12px;
            border-radius: 8px;
            margin-bottom: 13px;
            text-align: center;
        }
        .header-sgo { background: #059669; }
        .header-mapa { background: #0D256C; }
        .header-expurgo { background: #7c3aed; }

        .page-card p {
            color: #334155;
            font-size: 12.5px;
            line-height: 1.5;
            margin-bottom: 8px;
        }
        .page-card ul {
            margin: 7px 0 0 19px;
            padding: 0;
            color: #334155;
            font-size: 12px;
            line-height: 1.55;
        }
        .tip-box {
            margin-top: 13px;
            background: #f8fafc;
            border-left: 4px solid #059669;
            border-radius: 6px;
            padding: 10px 12px;
            color: #334155;
            font-size: 11.5px;
            line-height: 1.45;
        }

        .flow-box {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            padding: 14px 16px;
            font-size: 12.5px;
            color: #334155;
            line-height: 1.65;
        }
        .flow-highlight {
            background: #fff7ed;
            border: 1px solid #fed7aa;
            border-left: 5px solid #f97316;
            border-radius: 8px;
            padding: 12px 14px;
            color: #7c2d12;
            font-size: 12.5px;
            line-height: 1.55;
        }

        .legend-box {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            padding: 12px 14px;
            min-height: 120px;
            font-size: 12px;
            color: #334155;
            line-height: 1.65;
        }

        .status-panel {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            padding: 8px 12px;
        }
        .status-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 14px;
            padding: 8px 2px;
            border-bottom: 1px solid #f1f5f9;
            color: #334155;
            font-size: 11.5px;
        }
        .status-row:last-child { border-bottom: none; }
        .status-row span { color: #94a3b8; font-size: 10.5px; text-align: right; }

        .decision-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
        }
        .decision-table th {
            background: #0f172a;
            color: white;
            padding: 9px 10px;
            text-align: left;
        }
        .decision-table td {
            border: 1px solid #e2e8f0;
            padding: 9px 10px;
            color: #334155;
        }
        .decision-table tr:nth-child(even) td { background: #f8fafc; }

        .footer-note {
            text-align: center;
            color: #94a3b8;
            font-size: 11px;
            margin-top: 28px;
        }

        div.stButton > button {
            font-weight: 800;
            min-height: 40px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# LOGO / CABECALHO
# ==========================================
st.markdown("<br>", unsafe_allow_html=True)
if logo_ok:
    with open(ROOT / "LOGO_NIP.png", "rb") as image_file:
        b64_logo = base64.b64encode(image_file.read()).decode()
    st.markdown(
        f'''
        <div style="text-align:center;margin-bottom:6px;">
            <img src="data:image/png;base64,{b64_logo}" style="max-width:145px;width:100%;height:auto;pointer-events:none;">
        </div>
        ''',
        unsafe_allow_html=True,
    )

st.markdown(
    """
    <div class="hero">
        <div class="hero-title">Sistema de Obras NIP</div>
        <p class="hero-subtitle">Central integrada para criação de SGO, inteligência geográfica de obras e relatórios de expurgo.</p>
        <div class="hero-badge">SGO • MAPA • EXPURGO</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# INDICADORES DO HUB
# ==========================================
c1, c2, c3, c4 = st.columns(4, gap="small")
with c1:
    st.markdown(
        '<div class="metric-card"><div class="metric-label">Ferramentas</div><div class="metric-value">3</div><div class="metric-note">SGO, Mapa e Expurgo</div></div>',
        unsafe_allow_html=True,
    )
with c2:
    st.markdown(
        f'<div class="metric-card"><div class="metric-label">Base Levantamento</div><div class="metric-value">{"✅ OK" if base_ok else "⚠️ Falta"}</div><div class="metric-note">BASE_LEVANTAMENTO_ATUALIZADA.xlsx</div></div>',
        unsafe_allow_html=True,
    )
with c3:
    st.markdown(
        f'<div class="metric-card"><div class="metric-label">Areas Especiais</div><div class="metric-value">{qtd_kml}</div><div class="metric-note">arquivos KML detectados</div></div>',
        unsafe_allow_html=True,
    )
with c4:
    st.markdown(
        '<div class="metric-card"><div class="metric-label">Regra Geografica</div><div class="metric-value">50 m</div><div class="metric-note">STATUS LIST = CONCLUIDO</div></div>',
        unsafe_allow_html=True,
    )

# ==========================================
# CARDS DAS PAGINAS
# ==========================================
st.markdown('<div class="section-title">📌 Ferramentas do sistema</div>', unsafe_allow_html=True)
col_sgo, col_mapa, col_exp = st.columns(3, gap="medium")

with col_sgo:
    st.markdown(
        """
        <div class="page-card">
            <div class="page-card-header header-sgo">🏗️ CRIAR SGO</div>
            <p>Consulta as solicitações e prepara os dados necessários para criação da nota SGO e dos nomes das obras.</p>
            <ul>
                <li>Recebe uma ou várias notas no campo <b>SOLICITAÇÕES</b>.</li>
                <li>Consulta automaticamente a <b>BASE_LEVANTAMENTO_ATUALIZADA.xlsx</b>.</li>
                <li>Identifica notas <b>CANC/FINL</b> e permite incluí-las quando necessário.</li>
                <li>Trata <b>NOTAS ASSOCIADAS</b> e <b>NOTAS VU</b>.</li>
                <li>Preenche cliente, endereço, fase, PI, regional, coordenadas e demais dados.</li>
                <li>Gera <b>DESCRIÇÕES SGO</b>, <b>NOMES DAS OBRAS</b> e dados de criação/aprovação.</li>
                <li><b>Verifica automaticamente conflitos geográficos de até 50 m</b> com obras concluídas.</li>
                <li>Quando encontra conflito, disponibiliza acesso direto ao mapa.</li>
            </ul>
            <div class="tip-box"><b>Use quando:</b> precisar criar, conferir ou organizar uma SGO antes do lançamento definitivo.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if PAGE_SGO:
        if st.button("🏗️ Abrir CRIAR SGO", use_container_width=True, type="primary", key="abrir_sgo_inicio"):
            st.switch_page(PAGE_SGO)
    else:
        st.button("⚠️ Página CRIAR SGO não localizada", use_container_width=True, disabled=True)

with col_mapa:
    st.markdown(
        """
        <div class="page-card">
            <div class="page-card-header header-mapa">🗺️ MAPA / INTELIGÊNCIA GEOGRÁFICA</div>
            <p>Central geográfica para visualizar redes, obras, conflitos e áreas especiais do território.</p>
            <ul>
                <li>Exibe redes elétricas processadas a partir de arquivos <b>KMZ/KML</b>.</li>
                <li>Filtra por <b>Regional, Município e Alimentador</b>.</li>
                <li>Pesquisa por nome/ID ou por coordenada.</li>
                <li>Mostra <b>obras concluídas</b> e obras em andamento/conflito.</li>
                <li>Desenha o <b>raio de 50 m ao redor da obra concluída</b>.</li>
                <li>Recebe automaticamente o foco do CRIAR SGO e abre com zoom no conflito.</li>
                <li>Exibe Áreas Quilombolas, Terras Indígenas, Sítios Arqueológicos e UCs.</li>
                <li>Sincroniza novas redes existentes na pasta <b>kmzs</b>.</li>
            </ul>
            <div class="tip-box"><b>Use quando:</b> precisar conferir proximidade, sobreposição, malha elétrica ou restrições geográficas.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if PAGE_MAPA:
        if st.button("🗺️ Abrir MAPA", use_container_width=True, type="primary", key="abrir_mapa_inicio"):
            st.switch_page(PAGE_MAPA)
    else:
        st.button("⚠️ Página MAPA não localizada", use_container_width=True, disabled=True)

with col_exp:
    st.markdown(
        """
        <div class="page-card">
            <div class="page-card-header header-expurgo">📊 RELATÓRIO DE EXPURGO</div>
            <p>Monta o formulário de não atendimento/expurgo com dados da solicitação e evidência de campo.</p>
            <ul>
                <li>Busca automaticamente a nota/protocolo na base.</li>
                <li>Preenche regional, data da solicitação, conta contrato, parceiro e endereço.</li>
                <li>Permite anexar foto da evidência em PNG/JPG/JPEG.</li>
                <li>Registra data, horário, latitude, longitude e identificação da equipe.</li>
                <li>Inclui justificativa, descrição do expurgo e tratativa comercial.</li>
                <li>Permite informar medidor do cliente, medidor do vizinho e estrutura próxima.</li>
                <li>Integra a evidência ao formulário final.</li>
                <li>Permite impressão e geração de PDF pelo próprio navegador.</li>
            </ul>
            <div class="tip-box"><b>Use quando:</b> precisar formalizar um não atendimento/expurgo com dados e evidência em um único documento.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if PAGE_EXPURGO:
        if st.button("📊 Abrir RELATÓRIO DE EXPURGO", use_container_width=True, type="primary", key="abrir_expurgo_inicio"):
            st.switch_page(PAGE_EXPURGO)
    else:
        st.button("⚠️ Página de EXPURGO não localizada", use_container_width=True, disabled=True)

# ==========================================
# FLUXO INTEGRADO
# ==========================================
st.markdown('<div class="section-title">⚡ Fluxo integrado do sistema</div>', unsafe_allow_html=True)
fa, fb = st.columns([1.45, 1], gap="medium")
with fa:
    st.markdown(
        """
        <div class="flow-box">
            <b>1.</b> Acesse <b>CRIAR SGO</b> e informe uma ou várias solicitações.<br>
            <b>2.</b> O sistema consulta a base e organiza as informações da obra.<br>
            <b>3.</b> Cada solicitação é comparada com obras cujo <b>STATUS LIST = CONCLUIDO</b>.<br>
            <b>4.</b> Se não houver concluída em até <b>50 m</b>, o fluxo de criação da SGO segue normalmente.<br>
            <b>5.</b> Se houver conflito, o CRIAR SGO mostra a nota concluída, a distância e o botão <b>Abrir Conflito no Mapa</b>.<br>
            <b>6.</b> O mapa abre já focado no local, mostrando a obra consultada e a concluída próxima.
        </div>
        """,
        unsafe_allow_html=True,
    )
with fb:
    st.markdown(
        """
        <div class="flow-highlight">
            <b>📍 Regra oficial de conflito</b><br><br>
            Uma solicitação é sinalizada quando existe uma obra com <b>STATUS LIST contendo CONCLUIDO</b> dentro de um raio de <b>até 50 metros</b>.<br><br>
            🔵 <b>Ponto azul:</b> obra concluída<br>
            🔴 <b>Ponto vermelho:</b> obra consultada/conflitante<br>
            ⭕ <b>Círculo:</b> raio de 50 m ao redor da concluída
        </div>
        """,
        unsafe_allow_html=True,
    )

# ==========================================
# QUAL FERRAMENTA USAR
# ==========================================
st.markdown('<div class="section-title">🧭 Qual ferramenta devo usar?</div>', unsafe_allow_html=True)
st.markdown(
    """
    <table class="decision-table">
        <tr><th>Necessidade</th><th>Ferramenta indicada</th></tr>
        <tr><td>Criar, conferir ou organizar uma nota SGO</td><td>🏗️ CRIAR SGO</td></tr>
        <tr><td>Verificar se uma obra está próxima de outra concluída</td><td>🗺️ MAPA</td></tr>
        <tr><td>Visualizar redes, alimentadores e áreas especiais</td><td>🗺️ MAPA</td></tr>
        <tr><td>Pesquisar poste, transformador, rede ou coordenada</td><td>🗺️ MAPA</td></tr>
        <tr><td>Formalizar um não atendimento / expurgo</td><td>📊 RELATÓRIO DE EXPURGO</td></tr>
        <tr><td>Gerar documento com evidência fotográfica</td><td>📊 RELATÓRIO DE EXPURGO</td></tr>
    </table>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# STATUS DO SISTEMA / DIAGNOSTICO
# ==========================================
st.markdown('<div class="section-title">🧩 Status do sistema</div>', unsafe_allow_html=True)
s1, s2 = st.columns(2, gap="medium")

with s1:
    html_status_1 = "".join([
        status_item(base_ok, "BASE_LEVANTAMENTO_ATUALIZADA.xlsx", "Base principal de notas e coordenadas"),
        status_item(municipios_ok, "MUNICIPIOS-REGIONAIS.xlsx", "Vínculo Município x Regional"),
        status_item(logo_ok, "LOGO_NIP.png", "Identidade visual"),
        status_item(PAGE_SGO is not None, "Página CRIAR SGO", PAGE_SGO or "Não localizada"),
        status_item(PAGE_EXPURGO is not None, "Página RELATÓRIO DE EXPURGO", PAGE_EXPURGO or "Não localizada"),
        status_item(PAGE_MAPA is not None, "Página MAPA", PAGE_MAPA or "Não localizada"),
    ])
    st.markdown(f'<div class="status-panel">{html_status_1}</div>', unsafe_allow_html=True)

with s2:
    html_status_2 = "".join([
        status_item(kmls_ok, "Pasta kmls", f"{qtd_kml} arquivo(s) KML" if kmls_ok else "Pasta não localizada"),
        status_item(kmzs_ok, "Pasta kmzs", f"{qtd_kmz} arquivo(s) de rede" if kmzs_ok else "Pasta não localizada"),
        status_item(kmls_ok and qtd_kml > 0, "Áreas especiais", "Disponíveis no mapa" if qtd_kml > 0 else "Nenhum KML detectado"),
        status_item(kmzs_ok and qtd_kmz > 0, "Malha para sincronização", "Há arquivos disponíveis" if qtd_kmz > 0 else "Pasta vazia ou ausente"),
    ])
    st.markdown(f'<div class="status-panel">{html_status_2}</div>', unsafe_allow_html=True)

# ==========================================
# FONTES DE DADOS
# ==========================================
st.markdown('<div class="section-title">📚 Fontes de dados utilizadas</div>', unsafe_allow_html=True)
d1, d2, d3, d4 = st.columns(4, gap="small")
with d1:
    st.markdown('<div class="legend-box"><b>📄 BASE_LEVANTAMENTO</b><br>Notas, status, clientes, municípios e coordenadas usadas pelo CRIAR SGO, MAPA e EXPURGO.</div>', unsafe_allow_html=True)
with d2:
    st.markdown('<div class="legend-box"><b>🏙️ MUNICIPIOS-REGIONAIS</b><br>Relaciona os municípios às respectivas regionais para filtros e visualização geográfica.</div>', unsafe_allow_html=True)
with d3:
    st.markdown('<div class="legend-box"><b>🗺️ kmls/</b><br>Áreas Quilombolas, Terras Indígenas, Sítios Arqueológicos e Unidades de Conservação.</div>', unsafe_allow_html=True)
with d4:
    st.markdown('<div class="legend-box"><b>⚡ kmzs/</b><br>Arquivos de malha elétrica utilizados para alimentar e sincronizar as redes exibidas no mapa.</div>', unsafe_allow_html=True)

st.markdown(
    '<div class="footer-note">Sistema de Obras NIP • SGO • Inteligência Geográfica • Expurgo</div>',
    unsafe_allow_html=True,
)
