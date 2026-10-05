import streamlit as st
import os
import base64
from pathlib import Path

st.set_page_config(
    page_title="Sistema de Obras NIP",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

ROOT = Path(__file__).resolve().parent


def localizar_pagina(candidatos):
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
    detalhe_html = f"<span>{detalhe}</span>" if detalhe else ""
    return f'<div class="status-row"><b>{icone} {titulo}</b>{detalhe_html}</div>'


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
database_ok = (ROOT / "database" / "redes.db").exists()
backup_dir_ok = (ROOT / "database" / "backups").is_dir()
qtd_kml = contar_extensoes("kmls", {".kml"})
qtd_kmz = contar_extensoes("kmzs", {".kmz", ".kml"})

st.markdown(
    """
    <style>
        .block-container { padding-top: 1rem !important; padding-bottom: 2rem !important; max-width: 1360px; }
        .hero { text-align:center; padding:6px 10px 10px 10px; }
        .hero-title { font-size:31px; font-weight:850; color:#0f172a; margin:3px 0; }
        .hero-subtitle { font-size:14px; color:#64748b; margin:0; }
        .hero-badge { display:inline-block; margin-top:9px; padding:5px 10px; background:#ecfdf5; color:#047857; border:1px solid #a7f3d0; border-radius:999px; font-size:11px; font-weight:800; }
        .section-title { font-size:18px; font-weight:850; color:#0f172a; margin:20px 0 10px 0; }
        .metric-card { background:#fff; border:1px solid #dbe3ec; border-radius:10px; padding:13px 15px; box-shadow:0 2px 8px rgba(15,23,42,.04); min-height:80px; }
        .metric-label { color:#64748b; font-size:11px; font-weight:800; text-transform:uppercase; letter-spacing:.4px; }
        .metric-value { color:#0f172a; font-size:22px; font-weight:850; margin-top:5px; }
        .metric-note { color:#94a3b8; font-size:10px; margin-top:2px; }
        .page-card { background:#fff; border:1px solid #dbe3ec; border-radius:12px; padding:16px 18px; min-height:520px; box-shadow:0 2px 10px rgba(15,23,42,.05); }
        .page-card-header { color:#fff; font-size:16px; font-weight:850; padding:10px 12px; border-radius:8px; margin-bottom:13px; text-align:center; }
        .header-sgo { background:#059669; }
        .header-mapa { background:#0D256C; }
        .header-expurgo { background:#7c3aed; }
        .page-card p { color:#334155; font-size:12.5px; line-height:1.5; margin-bottom:8px; }
        .page-card ul { margin:7px 0 0 19px; padding:0; color:#334155; font-size:12px; line-height:1.5; }
        .tip-box { margin-top:13px; background:#f8fafc; border-left:4px solid #059669; border-radius:6px; padding:10px 12px; color:#334155; font-size:11.5px; line-height:1.45; }
        .flow-box { background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:14px 16px; font-size:12.5px; color:#334155; line-height:1.65; }
        .flow-highlight { background:#fff7ed; border:1px solid #fed7aa; border-left:5px solid #f97316; border-radius:8px; padding:12px 14px; color:#7c2d12; font-size:12.5px; line-height:1.55; }
        .legend-box { background:#fff; border:1px solid #e2e8f0; border-radius:10px; padding:12px 14px; min-height:145px; font-size:12px; color:#334155; line-height:1.65; }
        .status-panel { background:#fff; border:1px solid #e2e8f0; border-radius:10px; padding:8px 12px; }
        .status-row { display:flex; align-items:center; justify-content:space-between; gap:14px; padding:8px 2px; border-bottom:1px solid #f1f5f9; color:#334155; font-size:11.5px; }
        .status-row:last-child { border-bottom:none; }
        .status-row span { color:#94a3b8; font-size:10.5px; text-align:right; }
        .decision-table { width:100%; border-collapse:collapse; font-size:12px; }
        .decision-table th { background:#0f172a; color:white; padding:9px 10px; text-align:left; }
        .decision-table td { border:1px solid #e2e8f0; padding:9px 10px; color:#334155; }
        .decision-table tr:nth-child(even) td { background:#f8fafc; }
        .footer-note { text-align:center; color:#94a3b8; font-size:11px; margin-top:28px; }
        div.stButton > button { font-weight:800; min-height:40px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown("<br>", unsafe_allow_html=True)
if logo_ok:
    with open(ROOT / "LOGO_NIP.png", "rb") as image_file:
        b64_logo = base64.b64encode(image_file.read()).decode()
    st.markdown(
        f'<div style="text-align:center;margin-bottom:6px;"><img src="data:image/png;base64,{b64_logo}" style="max-width:145px;width:100%;height:auto;pointer-events:none;"></div>',
        unsafe_allow_html=True,
    )

st.markdown(
    """
    <div class="hero">
        <div class="hero-title">Sistema de Obras NIP</div>
        <p class="hero-subtitle">Central integrada para criação de SGO, análise de conflitos em 50 m, inteligência geográfica da malha, qualidade de dados e relatórios de expurgo.</p>
        <div class="hero-badge">SGO • MAPA INTELIGENTE • QUALIDADE DE DADOS • EXPURGO</div>
    </div>
    """,
    unsafe_allow_html=True,
)

c1, c2, c3, c4 = st.columns(4, gap="small")
with c1:
    st.markdown('<div class="metric-card"><div class="metric-label">Ferramentas</div><div class="metric-value">3</div><div class="metric-note">SGO, Mapa e Expurgo</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="metric-card"><div class="metric-label">Base Levantamento</div><div class="metric-value">{"✅ OK" if base_ok else "⚠️ Falta"}</div><div class="metric-note">BASE_LEVANTAMENTO_ATUALIZADA.xlsx</div></div>', unsafe_allow_html=True)
with c3:
    st.markdown(f'<div class="metric-card"><div class="metric-label">Áreas Especiais</div><div class="metric-value">{qtd_kml}</div><div class="metric-note">arquivos KML disponíveis</div></div>', unsafe_allow_html=True)
with c4:
    st.markdown('<div class="metric-card"><div class="metric-label">Regra Geográfica</div><div class="metric-value">50 m</div><div class="metric-note">conflito com STATUS LIST = CONCLUÍDO</div></div>', unsafe_allow_html=True)

st.markdown('<div class="section-title">📌 Ferramentas do sistema</div>', unsafe_allow_html=True)
col_sgo, col_mapa, col_exp = st.columns(3, gap="medium")

with col_sgo:
    st.markdown(
        """
        <div class="page-card">
            <div class="page-card-header header-sgo">🏗️ CRIAR SGO</div>
            <p>Consulta as solicitações e prepara os dados necessários para criação da nota SGO e dos nomes das obras.</p>
            <ul>
                <li>Recebe uma ou várias notas em <b>SOLICITAÇÕES</b>.</li>
                <li>Consulta automaticamente a base de levantamento.</li>
                <li>Identifica notas <b>CANC/FINL</b> e permite incluí-las quando necessário.</li>
                <li>Trata <b>NOTAS ASSOCIADAS</b> e <b>NOTAS VU</b>.</li>
                <li>Preenche cliente, endereço, fase, PI, regional e coordenadas.</li>
                <li>Gera <b>DESCRIÇÕES SGO</b>, <b>NOMES DAS OBRAS</b> e dados de criação/aprovação.</li>
                <li>Verifica automaticamente conflito com obra concluída em até <b>50 m</b>.</li>
                <li>Quando encontra conflito, abre o <b>MAPA</b> diretamente no ponto analisado.</li>
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
            <p>Central geográfica para análise de obras, conflitos, malha elétrica, qualidade das coordenadas e restrições territoriais.</p>
            <ul>
                <li>Filtra por <b>Regional, Município e Alimentador</b>.</li>
                <li>Pesquisa por nome/ID, coordenada ou <b>Protocolo/Nota</b>.</li>
                <li><b>Concluídas e conflitantes aparecem sem cluster</b>; somente Todas as Obras usa clusters.</li>
                <li>Desenha o <b>raio de 50 m</b> na concluída e classifica a severidade.</li>
                <li>Identifica <b>múltiplas concluídas</b> dentro do mesmo raio.</li>
                <li>Possui <b>Visualizações Rápidas</b>: Conflitos, Malha, Restrições e Todas as Obras.</li>
                <li>Mostra Áreas Quilombolas, Terras Indígenas, Sítios Arqueológicos e UCs.</li>
                <li>Permite filtrar obras com <b>restrição/proximidade especial</b>.</li>
                <li>Exibe dashboard de conflitos por cidade, status e severidade.</li>
                <li>Inclui <b>Monitor de Qualidade</b>, duplicidades, protocolos repetidos e qualidade por município.</li>
                <li>Exporta conflitos, qualidade de dados e a <b>visão filtrada</b>.</li>
                <li>Sincroniza KMZ/KML e cria <b>backup preventivo</b> antes de excluir alimentadores.</li>
            </ul>
            <div class="tip-box"><b>Use quando:</b> precisar validar conflitos, investigar uma nota, analisar a malha, conferir restrições ou sanear problemas geográficos.</div>
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
                <li>Permite impressão e geração de PDF pelo navegador.</li>
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

st.markdown('<div class="section-title">⚡ Fluxo integrado do sistema</div>', unsafe_allow_html=True)
fa, fb = st.columns([1.45, 1], gap="medium")
with fa:
    st.markdown(
        """
        <div class="flow-box">
            <b>1.</b> Acesse <b>CRIAR SGO</b> e informe uma ou várias solicitações.<br>
            <b>2.</b> O sistema consulta a base e organiza os dados da obra.<br>
            <b>3.</b> Cada solicitação é comparada com obras cujo <b>STATUS LIST = CONCLUÍDO</b>.<br>
            <b>4.</b> Se houver concluída em até <b>50 m</b>, o CRIAR SGO informa a distância e disponibiliza <b>Abrir Conflito no Mapa</b>.<br>
            <b>5.</b> O mapa abre focado no local, com concluída azul, nova obra vermelha e círculo de 50 m.<br>
            <b>6.</b> Use pesquisa por protocolo, filtros, visualizações rápidas, áreas especiais e relatórios para aprofundar a análise.<br>
            <b>7.</b> Exporte conflitos, qualidade de dados ou a visão filtrada quando precisar registrar a análise.
        </div>
        """,
        unsafe_allow_html=True,
    )
with fb:
    st.markdown(
        """
        <div class="flow-highlight">
            <b>📍 Regra oficial de conflito</b><br><br>
            Uma solicitação é sinalizada quando existe uma obra com <b>STATUS LIST contendo CONCLUÍDO</b> em um raio de <b>até 50 metros</b>.<br><br>
            🔵 <b>Ponto azul:</b> concluída, sem cluster<br>
            🔴 <b>Ponto vermelho:</b> conflitante, sem cluster<br>
            ⭕ <b>Círculo:</b> raio de 50 m da concluída<br>
            📌 <b>Mesmo ponto:</b> coordenadas praticamente coincidentes<br>
            ⚠️ <b>Crítico / Alto / Médio:</b> classificação pela distância
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown('<div class="section-title">🚀 Recursos atuais do MAPA</div>', unsafe_allow_html=True)
rf1, rf2, rf3 = st.columns(3, gap="medium")
with rf1:
    st.markdown(
        """
        <div class="legend-box">
            <b>🚨 Conflitos e decisão</b><br>
            • Regra oficial de 50 m<br>
            • Mesmo Ponto / Crítico / Alto / Médio<br>
            • Múltiplas concluídas no mesmo raio<br>
            • Concluídas e conflitantes sem cluster<br>
            • Dashboard e consolidado por concluída
        </div>
        """,
        unsafe_allow_html=True,
    )
with rf2:
    st.markdown(
        """
        <div class="legend-box">
            <b>🔎 Pesquisa e qualidade</b><br>
            • Busca por Protocolo/Nota<br>
            • Busca por coordenada e elemento da malha<br>
            • Coordenadas inválidas e duplicadas<br>
            • Protocolos repetidos<br>
            • Qualidade geográfica por município
        </div>
        """,
        unsafe_allow_html=True,
    )
with rf3:
    st.markdown(
        """
        <div class="legend-box">
            <b>⚙️ Operação e segurança</b><br>
            • Presets de visualização<br>
            • Resumo dos filtros ativos<br>
            • Exportação da visão filtrada<br>
            • Registro da sincronização da malha<br>
            • Backup antes de excluir alimentador
        </div>
        """,
        unsafe_allow_html=True,
    )
st.caption("💡 O INÍCIO apenas apresenta e valida a estrutura dos arquivos. Os cálculos geográficos permanecem dentro do MAPA para manter esta página leve.")

st.markdown('<div class="section-title">🧭 Qual ferramenta devo usar?</div>', unsafe_allow_html=True)
st.markdown(
    """
    <table class="decision-table">
        <tr><th>Necessidade</th><th>Ferramenta indicada</th></tr>
        <tr><td>Criar, conferir ou organizar uma nota SGO</td><td>🏗️ CRIAR SGO</td></tr>
        <tr><td>Verificar proximidade com uma obra concluída</td><td>🗺️ MAPA — Conflitos 50 m</td></tr>
        <tr><td>Localizar rapidamente uma nota/protocolo</td><td>🗺️ MAPA — Pesquisa por Protocolo</td></tr>
        <tr><td>Visualizar redes, alimentadores e áreas especiais</td><td>🗺️ MAPA — Malha / Restrições</td></tr>
        <tr><td>Pesquisar poste, transformador, rede ou coordenada</td><td>🗺️ MAPA — Pesquisas Inteligentes</td></tr>
        <tr><td>Encontrar coordenadas duplicadas ou protocolos repetidos</td><td>🗺️ MAPA — Diagnóstico Preventivo</td></tr>
        <tr><td>Analisar qualidade geográfica por município</td><td>🗺️ MAPA — Qualidade de Dados</td></tr>
        <tr><td>Exportar as obras atualmente filtradas</td><td>🗺️ MAPA — Exportar Visão Atual</td></tr>
        <tr><td>Formalizar um não atendimento / expurgo</td><td>📊 RELATÓRIO DE EXPURGO</td></tr>
        <tr><td>Gerar documento com evidência fotográfica</td><td>📊 RELATÓRIO DE EXPURGO</td></tr>
    </table>
    """,
    unsafe_allow_html=True,
)

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
        status_item(database_ok, "Banco SQLite da malha", "database/redes.db" if database_ok else "Criado pelo MAPA quando necessário"),
        status_item(backup_dir_ok, "Backups da malha", "database/backups disponível" if backup_dir_ok else "Criado automaticamente na primeira exclusão protegida"),
    ])
    st.markdown(f'<div class="status-panel">{html_status_2}</div>', unsafe_allow_html=True)

st.markdown('<div class="section-title">📚 Fontes de dados utilizadas</div>', unsafe_allow_html=True)
d1, d2, d3, d4 = st.columns(4, gap="small")
with d1:
    st.markdown('<div class="legend-box"><b>📄 BASE_LEVANTAMENTO</b><br>Notas, status, clientes, municípios e coordenadas usadas pelo CRIAR SGO, MAPA e EXPURGO.</div>', unsafe_allow_html=True)
with d2:
    st.markdown('<div class="legend-box"><b>🏙️ MUNICIPIOS-REGIONAIS</b><br>Relaciona municípios às regionais para filtros e visualização geográfica.</div>', unsafe_allow_html=True)
with d3:
    st.markdown('<div class="legend-box"><b>🗺️ kmls/</b><br>Áreas Quilombolas, Terras Indígenas, Sítios Arqueológicos e Unidades de Conservação.</div>', unsafe_allow_html=True)
with d4:
    st.markdown('<div class="legend-box"><b>⚡ kmzs/ + SQLite</b><br>A malha é sincronizada com <b>database/redes.db</b> para acelerar a visualização e a gestão local dos alimentadores.</div>', unsafe_allow_html=True)

st.markdown('<div class="footer-note">Sistema de Obras NIP • SGO • MAPA Inteligente • Qualidade de Dados • Expurgo</div>', unsafe_allow_html=True)
