import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
import json
import html
import math
import unicodedata
from pathlib import Path
from datetime import datetime, date

st.set_page_config(
    page_title="Relatório de Expurgo",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 0.55rem !important;
            padding-bottom: 1rem !important;
            max-width: 1160px !important;
        }
        [data-testid="stAppViewBlockContainer"] { overflow: visible; }
        iframe { border: 0 !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

# Logo extraída do próprio modelo Excel oficial.
LOGO_EQTL_EXCEL_B64 = "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"

# Relação exatamente igual à aba Apoio do relatório Excel (A2:B26).
TRATATIVAS_EXPURGO = {"ATENDIMENTO EM TERRENO DE TERCEIROS": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE ATENDIDO POR OUTRA ÁREA": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE ATENDIDO POR OUTRA SOLICITAÇÃO (EXPANSÃO)": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE CONSTRUIU A PRÓPRIA REDE SEM APRESENTAR PROJETO": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE DESISTIU DO SERVIÇO": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE EM ÁREA DE DOMÍNIO PÚBLICO/LITÍGIO": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE EM ÁREA DE LOTEAMENTO PARTICULAR/ CONDOMÍNIO": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE EM ÁREA DE PRESERVAÇÃO AMBIENTAL": "SUSPENSÃO", "CLIENTE EMUC": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE LIGADO A REVELIA (CLANDESTINO)": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE NÃO LOCALIZADO": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE URBANO SE ARRUAMENTO": "REJEIÇÃO PARA CANCELAMENTO", "ERRO NA ABERTURA - CLIENTE DESEJA OUTRO SERVIÇO": "REJEIÇÃO PARA CANCELAMENTO", "JÁ EXISTE REDE DE BT PARA ATENDER O CLIENTE": "DEVOLVER A GSTC", "OBSTACULO IMPEDE EXECUÇÃO/ SEM ACESSO": "SUSPENSÃO", "PADRÃO INCORRETO": "DEVOLVER A GSTC", "PADRÃO INEXISTENTE": "DEVOLVER A GSTC", "PADRÃO TRIFÁSICO INEXISTENTE": "DEVOLVER A GSTC", "PEDIDO EM DUPLICIDADE": "REJEIÇÃO PARA CANCELAMENTO", "RECLASSIFICAÇÃO RR PARA UB": "INFORMAR PARA AJUSTE NO INDICADOR", "RECLASSIFICAÇÃO UB PARA RR": "INFORMAR PARA AJUSTE NO INDICADOR", "SEM ACESSO PERIODO CHUVOSO": "SUSPENSÃO", "SOMENTE O TERRENO": "REJEIÇÃO PARA CANCELAMENTO", "UC DEMOLIDA/ ABANDONADA": "REJEIÇÃO PARA CANCELAMENTO", "CUSTEIO": "REJEIÇÃO PARA CANCELAMENTO"}


def _normalizar_nome(texto):
    texto = "" if texto is None else str(texto).strip().upper()
    texto = "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )
    return " ".join(texto.split())


def _texto(valor):
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass

    if isinstance(valor, (pd.Timestamp, datetime, date)):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, bool):
        return "SIM" if valor else "NÃO"
    if isinstance(valor, int):
        return str(valor)
    if isinstance(valor, float):
        if math.isnan(valor):
            return ""
        if valor.is_integer():
            return str(int(valor))
        return f"{valor:.12f}".rstrip("0").rstrip(".")

    txt = str(valor).strip().replace('"', '')
    if txt.lower() in {"nan", "none", "null"}:
        return ""
    if txt.endswith(".0") and txt[:-2].replace("-", "").isdigit():
        txt = txt[:-2]
    return txt


def _formatar_data(valor):
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass
    try:
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except Exception:
        return _texto(valor)[:10]


def _localizar_arquivo(nome):
    """Localiza arquivos tanto na raiz quanto ao lado da página em pages/."""
    aqui = Path(__file__).resolve().parent
    candidatos = [
        Path.cwd() / nome,
        aqui / nome,
        aqui.parent / nome,
    ]
    for caminho in candidatos:
        if caminho.exists():
            return caminho
    return candidatos[0]


@st.cache_data(show_spinner=False)
def carregar_base_notas(caminho_str, mtime):
    caminho = Path(caminho_str)
    xls = pd.ExcelFile(caminho)
    try:
        df = pd.read_excel(xls, sheet_name="NOTAS")
    except Exception:
        try:
            df = pd.read_excel(xls, sheet_name="NotasSisgb")
        except Exception:
            return pd.DataFrame()

    if df.empty:
        return df
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df


def montar_mapa_notas(df):
    """Prepara os dados que podem ser vinculados aos campos do formulário."""
    if df.empty:
        return {}

    colunas = {_normalizar_nome(c): c for c in df.columns}

    def valor(row, aliases):
        for alias in aliases:
            col = colunas.get(_normalizar_nome(alias))
            if col is not None:
                v = row.get(col, "")
                if _texto(v) != "":
                    return v
        return ""

    mapa = {}
    for _, row in df.iterrows():
        protocolo = _texto(valor(row, ["PROTOCOLO", "Nº DA NOTA", "NOTA"]))
        if not protocolo:
            continue

        endereco_bruto = _texto(valor(row, ["ENDEREÇO", "ENDERECO"]))
        municipio = _texto(valor(row, ["MUNICIPIO", "MUNICÍPIO"]))
        endereco = endereco_bruto
        if endereco_bruto and municipio:
            # Mesmo comportamento da página anterior: endereço + município.
            endereco = f"{endereco_bruto} - {municipio}".strip(" -")

        justificativa = _texto(valor(row, [
            "JUSTIFICATIVA", "MOTIVO DO EXPURGO", "MOTIVO EXPURGO"
        ])).upper()

        mapa[protocolo.upper()] = {
            "distribuidora": _texto(valor(row, ["DISTRIBUIDORA"])) or "EQTL MA",
            "regional": _texto(valor(row, ["REGIONAL"])).upper(),
            "data_solicitacao": _formatar_data(valor(row, [
                "DATA ABERTURA", "DATA DA SOLICITAÇÃO", "DATA DA SOLICITACAO"
            ])),
            "conta_contrato": _texto(valor(row, ["CONTA CONTRATO", "CONTA CONTRATO CCS"])),
            "parceiro": _texto(valor(row, [
                "NOME", "NOME DO SOLICITANTE", "PARCEIRO DE NEGÓCIOS", "PARCEIRO DE NEGOCIOS"
            ])).upper(),
            "endereco": endereco.upper(),
            "data_visita": _formatar_data(valor(row, [
                "DATA DA VISITA", "DATA VISITA", "DATA DO LEVANTAMENTO", "DATA LEVANTAMENTO"
            ])),
            "horario": _texto(valor(row, [
                "HORÁRIO", "HORARIO", "HORÁRIO DA VISITA", "HORARIO DA VISITA"
            ])),
            "latitude": _texto(valor(row, ["LATITUDE"])),
            "longitude": _texto(valor(row, ["LONGITUDE"])),
            "equipe": _texto(valor(row, [
                "IDENTIFICAÇÃO DA EQUIPE", "IDENTIFICACAO DA EQUIPE", "EQUIPE"
            ])).upper() or "EQP NIP",
            "justificativa": justificativa,
            "descricao_expurgo": _texto(valor(row, [
                "DESCRIÇÃO DO EXPURGO", "DESCRICAO DO EXPURGO"
            ])).upper(),
            "tratativa": _texto(valor(row, [
                "TRATATIVA NO SISTEMA COMERCIAL", "TRATATIVA"
            ])).upper(),
            "medidor_cliente": _texto(valor(row, [
                "NÚMERO DO MEDIDOR DO CLIENTE ATENDIDO",
                "NUMERO DO MEDIDOR DO CLIENTE ATENDIDO",
                "MEDIDOR CLIENTE"
            ])),
            "medidor_vizinho": _texto(valor(row, [
                "NÚMERO DO MEDIDOR DO VIZINHO",
                "NUMERO DO MEDIDOR DO VIZINHO",
                "MEDIDOR VIZINHO"
            ])),
            "estrutura_proxima": _texto(valor(row, [
                "NÚMERO DA ESTRUTURA MAIS PRÓXIMA",
                "NUMERO DA ESTRUTURA MAIS PROXIMA",
                "ESTRUTURA MAIS PRÓXIMA",
                "ESTRUTURA MAIS PROXIMA"
            ])),
        }
    return mapa


arquivo_bd = _localizar_arquivo("BASE_LEVANTAMENTO_ATUALIZADA.xlsx")
if arquivo_bd.exists():
    df_notas = carregar_base_notas(str(arquivo_bd), os.path.getmtime(arquivo_bd))
else:
    df_notas = pd.DataFrame()
    st.warning("⚠️ BASE_LEVANTAMENTO_ATUALIZADA.xlsx não encontrada no repositório.")

mapa_notas = montar_mapa_notas(df_notas)

# O mapa vai para o HTML para o preenchimento ocorrer no próprio campo Nº da nota,
# sem rerun do Streamlit e sem apagar os campos manuais ou a foto.
base_json = json.dumps(mapa_notas, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
trat_json = json.dumps(TRATATIVAS_EXPURGO, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")

opcoes_justificativa = ['<option value=""></option>']
for motivo in TRATATIVAS_EXPURGO.keys():
    esc = html.escape(motivo, quote=True)
    opcoes_justificativa.append(f'<option value="{esc}">{esc}</option>')
opcoes_justificativa_html = "".join(opcoes_justificativa)

html_doc = r'''<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>
<style>
*{box-sizing:border-box}html,body{margin:0;padding:0;background:#fff;font-family:Calibri,Arial,sans-serif;color:#000}
.actions{position:fixed;top:72px;right:10px;left:auto;transform:none;z-index:9999;width:122px;padding:8px;display:flex;flex-direction:column;align-items:stretch;gap:8px;font-family:Calibri,Arial,sans-serif;background:rgba(255,255,255,.97);border:1px solid #e2e8f0;border-radius:6px;box-shadow:0 2px 8px rgba(0,0,0,.12)}
.action-btn,.photo-btn{border:0;border-radius:4px;width:100%;min-width:0;height:34px;padding:0 6px;font-size:8.4pt;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;color:#fff;background:#17375E;box-shadow:0 1px 2px rgba(0,0,0,.12);white-space:nowrap}
.action-btn:hover,.photo-btn:hover{filter:brightness(.94)}
.excel-btn{background:#217346}.pdf-btn{background:#B42318}.clear-btn{background:#F79646}.photo-btn{background:#24465F}.photo-btn input{display:none}
.print-stage{padding-top:0;padding-right:140px}
.sheet-canvas{width:801px;min-height:1190px;margin:0 auto;background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.top-space,.bottom-space{height:12.75pt}
.report-table{width:762px;margin-left:20px;border-collapse:collapse;table-layout:fixed;font-family:Calibri,Arial,sans-serif;color:#000}
.report-table td{padding:0;vertical-align:middle;overflow:hidden}
.top-bg,.top-logo,.top-title{background:#17375E;color:#fff;border:0}
.top-logo{position:relative;text-align:center;border-right:2.25pt solid #fff}
.top-logo img{width:142px;height:48px;object-fit:contain;display:block;margin:0 auto}
.top-title{text-align:center;font-size:18pt;font-weight:700;line-height:1;white-space:nowrap}
.blank-row td{border:0!important;background:#fff}
.section-label,.section-fill{background:#24465F;color:#fff;border-top:1px solid #24465F;border-bottom:1px solid #24465F;font-size:10pt;font-weight:700}
.section-label{border-left:1px solid #24465F;text-align:center}.section-fill{border-right:1px solid #24465F}
.lbl-cell,.value-cell,.treat-cell{border:1px solid #000;background:#fff;text-align:center;font-size:10pt;font-weight:700}
.lbl-small{font-size:8pt;line-height:1.05}.lbl-wrap{white-space:normal;line-height:1.05}.spacer-cell{border:0!important;background:#fff}
.treat-cell{background:#C6D9F1}
.fi,.fs{width:100%;height:100%;min-height:100%;margin:0;padding:0 4px;border:0;outline:0;background:transparent;color:#000;font-family:Calibri,Arial,sans-serif;font-size:10pt;font-weight:700;text-align:center;text-transform:uppercase;line-height:1}
.fi:focus,.fs:focus{outline:1px solid #4F81BD;outline-offset:-1px}.fi[readonly]{cursor:default}
.fs{appearance:none;-webkit-appearance:none;text-align-last:center;cursor:pointer;padding:0 13px 0 4px}
.select-cell{position:relative}
.select-cell::after{content:"▾";position:absolute;right:4px;top:50%;transform:translateY(-50%);font-size:6.5pt;line-height:1;color:#64748b;opacity:.16;pointer-events:none;transition:opacity .15s ease}
.select-cell:hover::after,.select-cell:focus-within::after{opacity:.62}
.evidence-cell{height:249.2pt!important;min-height:249.2pt!important;max-height:249.2pt!important;border:1px solid #000;background:#fff;position:relative;overflow:hidden;padding:0!important}
.evidence-grid{width:100%;height:249.2pt!important;min-height:249.2pt!important;max-height:249.2pt!important;display:grid;gap:1px;background:#000;overflow:hidden}
.evidence-grid.empty{display:flex;background:#fff;align-items:center;justify-content:center}
.evidence-placeholder{width:100%;height:100%;display:flex;align-items:center;justify-content:center;color:#94a3b8;font-size:9pt;font-weight:600;cursor:pointer;text-align:center}
.evidence-item{min-width:0;min-height:0;background:#fff;overflow:hidden}
.evidence-item img{width:100%;height:100%;display:block;object-fit:cover;object-position:center}
.evidence-grid.photos-1{grid-template-columns:1fr;grid-template-rows:1fr}
.evidence-grid.photos-2{grid-template-columns:repeat(2,1fr);grid-template-rows:1fr}
.evidence-grid.photos-3{grid-template-columns:repeat(3,1fr);grid-template-rows:1fr}
.evidence-grid.photos-4{grid-template-columns:repeat(2,1fr);grid-template-rows:repeat(2,1fr)}
.evidence-grid.photos-5{grid-template-columns:repeat(6,1fr);grid-template-rows:repeat(2,1fr)}
.evidence-grid.photos-5 .evidence-item:nth-child(-n+3){grid-column:span 2}
.evidence-grid.photos-5 .evidence-item:nth-child(n+4){grid-column:span 3}
.photo-count{font-size:8pt;font-weight:700;margin-left:4px;opacity:.9}
@media print{
 @page{size:A4 portrait;margin:0}
 html,body{width:210mm;height:297mm;margin:0!important;padding:0!important;overflow:hidden}
 .actions{display:none!important}
 .print-stage{width:210mm;height:297mm;padding-top:0!important;padding-right:0!important;display:flex;align-items:center;justify-content:center;overflow:hidden}
 .sheet-canvas{margin:0;transform:scale(.88);transform-origin:center center;flex:0 0 auto}
 .fi:focus,.fs:focus{outline:none!important}.fs{appearance:none;-webkit-appearance:none;padding-right:4px}.select-cell::after{display:none!important}.evidence-placeholder{display:none!important}
}
</style>
</head>
<body>
<div class="actions">
<label class="photo-btn" title="Adicionar até 5 fotos à área de evidências">FOTOS 📷 <span id="photoCount" class="photo-count">0/5</span><input id="photoInput" type="file" accept="image/png,image/jpeg" multiple></label>
<button class="action-btn excel-btn" type="button" onclick="exportarExcel()" title="Baixar o relatório preenchido em Excel">EXPORTAR EXCEL 📊</button>
<button class="action-btn pdf-btn" type="button" onclick="exportarPDF()" title="Imprimir ou salvar o relatório preenchido em PDF">EXPORTAR PDF 📄</button>
<button class="action-btn clear-btn" type="button" onclick="limparTudo()" title="Limpar todos os dados e fotos">LIMPAR 🧹</button>
</div>
<div class="print-stage"><div class="sheet-canvas">
<div class="top-space"></div>
<table class="report-table" aria-label="Formulário de Não Atendimento Expansão">
<colgroup><col style="width:123px"><col style="width:83px"><col style="width:94px"><col style="width:93px"><col style="width:77px"><col style="width:140px"><col style="width:76px"><col style="width:76px"></colgroup>
<tr style="height:12.75pt"><td class="top-logo" rowspan="3" colspan="2"><img src="data:image/png;base64,__LOGO__" alt="Grupo Equatorial Energia"></td><td class="top-bg" colspan="6"></td></tr>
<tr style="height:23.25pt"><td class="top-title" colspan="5">Formulário de Não Atendimento Expansão</td><td class="top-bg"></td></tr>
<tr style="height:12.75pt"><td class="top-bg" colspan="6"></td></tr>
<tr class="blank-row" style="height:12.75pt"><td colspan="8"></td></tr>
<tr style="height:20.25pt"><td class="lbl-cell">Distribuidora:</td><td class="value-cell select-cell"><select id="distribuidora" class="fs"><option>EQTL MA</option><option>EQTL PA</option><option>EQTL PI</option><option>EQTL AL</option></select></td><td class="spacer-cell"></td><td class="lbl-cell">Regional:</td><td class="value-cell select-cell"><select id="regional" class="fs"><option value=""></option><option>CENTRO</option><option>LESTE</option><option>METROPOLITANA</option><option>NORDESTE</option><option>NOROESTE</option><option>NORTE</option><option>OESTE</option><option>SUL</option></select></td><td class="spacer-cell"></td><td class="lbl-cell lbl-small">Data da<br>solicitação:</td><td class="value-cell"><input id="data_solicitacao" class="fi" type="text"></td></tr>
<tr class="blank-row" style="height:12.75pt"><td colspan="8"></td></tr>
<tr style="height:16.5pt"><td class="section-label">Dados do Cliente:</td><td class="section-fill" colspan="7"></td></tr>
<tr class="blank-row" style="height:12.75pt"><td colspan="8"></td></tr>
<tr style="height:26.25pt"><td class="lbl-cell">Nº da nota:</td><td class="value-cell" colspan="3"><input id="nota" class="fi" type="text" autocomplete="off" spellcheck="false"></td><td class="spacer-cell"></td><td class="lbl-cell">Conta Contrato:</td><td class="value-cell" colspan="2"><input id="conta_contrato" class="fi" type="text"></td></tr>
<tr class="blank-row" style="height:15pt"><td colspan="8"></td></tr>
<tr style="height:25.5pt"><td class="lbl-cell lbl-wrap">Parceiro de Negócios:</td><td class="value-cell" colspan="7"><input id="parceiro" class="fi" type="text"></td></tr>
<tr class="blank-row" style="height:12.75pt"><td colspan="8"></td></tr>
<tr style="height:25.5pt"><td class="lbl-cell">Endereço:</td><td class="value-cell" colspan="7"><input id="endereco" class="fi" type="text"></td></tr>
<tr class="blank-row" style="height:12.75pt"><td colspan="8"></td></tr>
<tr style="height:16.5pt"><td class="section-label">Dados da Visita:</td><td class="section-fill" colspan="7"></td></tr>
<tr class="blank-row" style="height:12.75pt"><td colspan="8"></td></tr>
<tr style="height:26.25pt"><td class="lbl-cell">Data:</td><td class="value-cell" colspan="3"><input id="data_visita" class="fi" type="text"></td><td class="spacer-cell"></td><td class="lbl-cell">Latitude:</td><td class="value-cell" colspan="2"><input id="latitude" class="fi" type="text"></td></tr>
<tr class="blank-row" style="height:15pt"><td colspan="8"></td></tr>
<tr style="height:26.25pt"><td class="lbl-cell">Horário:</td><td class="value-cell" colspan="3"><input id="horario" class="fi" type="text"></td><td class="spacer-cell"></td><td class="lbl-cell">Longitude:</td><td class="value-cell" colspan="2"><input id="longitude" class="fi" type="text"></td></tr>
<tr class="blank-row" style="height:12.75pt"><td colspan="8"></td></tr>
<tr style="height:25.5pt"><td class="lbl-cell lbl-wrap">Identificação da equipe:</td><td class="value-cell" colspan="7"><input id="equipe" class="fi" type="text" value="EQP NIP"></td></tr>
<tr class="blank-row" style="height:15.6pt"><td colspan="8"></td></tr>
<tr style="height:16.5pt"><td class="section-label">Motivo do expurgo:</td><td class="section-fill" colspan="7"></td></tr>
<tr class="blank-row" style="height:9pt"><td colspan="8"></td></tr>
<tr style="height:25.5pt"><td class="lbl-cell">Justificativa:</td><td class="value-cell select-cell" colspan="7"><select id="justificativa" class="fs">__OPTIONS__</select></td></tr>
<tr class="blank-row" style="height:6pt"><td colspan="8"></td></tr>
<tr style="height:25.5pt"><td class="lbl-cell lbl-wrap">Descrição do Expurgo:</td><td class="value-cell" colspan="7"><input id="descricao_expurgo" class="fi" type="text"></td></tr>
<tr class="blank-row" style="height:6.95pt"><td colspan="8"></td></tr>
<tr style="height:25.5pt"><td class="treat-cell lbl-wrap">Tratativa no Sistema Comercial:</td><td class="treat-cell" colspan="7"><input id="tratativa" class="fi" type="text"></td></tr>
<tr class="blank-row" style="height:12.75pt"><td colspan="8"></td></tr>
<tr style="height:16.5pt"><td class="section-label">Evidências:</td><td class="section-fill" colspan="7"></td></tr>
<tr class="blank-row" style="height:12.75pt"><td colspan="8"></td></tr>
<tr style="height:25.5pt"><td class="lbl-cell lbl-wrap">Número do medidor<br>do cliente atendido:</td><td class="value-cell" colspan="3"><input id="medidor_cliente" class="fi" type="text"></td><td class="spacer-cell"></td><td class="lbl-cell lbl-wrap">Número da nota do<br>atendimento em campo:</td><td class="value-cell" colspan="2"><input id="nota_campo" class="fi" type="text" readonly></td></tr>
<tr class="blank-row" style="height:7.5pt"><td colspan="8"></td></tr>
<tr style="height:25.5pt"><td class="lbl-cell lbl-wrap">Número do medidor<br>do vizinho:</td><td class="value-cell" colspan="3"><input id="medidor_vizinho" class="fi" type="text"></td><td class="spacer-cell"></td><td class="lbl-cell lbl-wrap">Número da estrutura<br>mais próxima:</td><td class="value-cell" colspan="2"><input id="estrutura_proxima" class="fi" type="text"></td></tr>
<tr class="blank-row" style="height:9.95pt"><td colspan="8"></td></tr>
<tr style="height:249.2pt!important"><td class="evidence-cell" colspan="8"><div id="evidenceGrid" class="evidence-grid empty"><label for="photoInput" class="evidence-placeholder">Clique aqui ou em FOTOS 📷 para adicionar até 5 fotos</label></div></td></tr>
</table>
<div class="bottom-space"></div>
</div></div>
<script>
let FOTOS_ATUAIS=[];
const LOGO_B64="__LOGO__";
function valorCampo(id){const el=document.getElementById(id);return el?((el.value||"").toString().trim()):""}
function baixarBlob(blob,nome){const url=URL.createObjectURL(blob);const a=document.createElement("a");a.href=url;a.download=nome;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1500)}
function exportarPDF(){window.print()}
function excelColPos(frac){const widths=[17.5703125,11.85546875,13.42578125,13.28515625,11,20,10.85546875,10.85546875];const total=widths.reduce((a,b)=>a+b,0);let alvo=Math.max(0,Math.min(1,frac))*total;for(let i=0;i<widths.length;i++){if(alvo<=widths[i])return 1+i+(alvo/widths[i]);alvo-=widths[i]}return 9}
function excelRowPos(frac){const heights=[12.95,15,15,15,15,15,15,15,15,15,15,15,15,15,15,30.75,27];const total=heights.reduce((a,b)=>a+b,0);let alvo=Math.max(0,Math.min(1,frac))*total;for(let i=0;i<heights.length;i++){if(alvo<=heights[i])return 37+i+(alvo/heights[i]);alvo-=heights[i]}return 54}
async function exportarExcel(){
  if(typeof ExcelJS==="undefined"){alert("Não foi possível carregar o módulo de exportação Excel. Verifique a conexão com a internet e tente novamente.");return}
  const wb=new ExcelJS.Workbook();wb.creator="NIP";wb.created=new Date();
  const ws=wb.addWorksheet("Modelo",{pageSetup:{paperSize:9,orientation:"portrait",fitToPage:false,scale:88,horizontalCentered:true,verticalCentered:true,margins:{left:0,right:0,top:0,bottom:0,header:0,footer:0}}});
  ws.views=[{showGridLines:false,zoomScale:100}];
  const larguras={A:2.85546875,B:17.5703125,C:11.85546875,D:13.42578125,E:13.28515625,F:11,G:20,H:10.85546875,I:10.85546875};Object.entries(larguras).forEach(([c,w])=>ws.getColumn(c).width=w);
  const alturas={2:15,3:23.25,4:12.75,6:20.25,8:16.5,10:26.25,11:15,12:25.5,13:15,14:25.5,16:16.5,17:15,18:26.25,19:15,20:26.25,21:15,22:25.5,23:15.6,24:16.5,25:9,26:25.5,27:6,28:25.5,29:6.95,30:25.5,31:15,32:16.5,33:15,34:25.5,35:7.5,36:25.5,37:9.95,38:12.95,39:15,40:15,41:15,42:15,43:15,44:15,45:15,46:15,47:15,48:15,49:15,50:15,51:15,52:15,53:30.75,54:27};Object.entries(alturas).forEach(([r,h])=>ws.getRow(Number(r)).height=h);
  ws.pageSetup.printArea="B2:I54";
  const PRETO="FF000000",BRANCO="FFFFFFFF",AZUL="FF17375E",AZUL2="FF24465F",AZULCLARO="FFC6D9F1";
  const borda={top:{style:"thin",color:{argb:PRETO}},left:{style:"thin",color:{argb:PRETO}},bottom:{style:"thin",color:{argb:PRETO}},right:{style:"thin",color:{argb:PRETO}}};
  function fillRange(rng,argb){ws.getCells?null:null;const [a,b]=rng.split(":");const c1=ws.getCell(a),c2=ws.getCell(b);for(let r=c1.row;r<=c2.row;r++)for(let c=c1.col;c<=c2.col;c++)ws.getCell(r,c).fill={type:"pattern",pattern:"solid",fgColor:{argb}}}
  function borderRange(rng){const [a,b]=rng.split(":");const c1=ws.getCell(a),c2=ws.getCell(b);for(let r=c1.row;r<=c2.row;r++)for(let c=c1.col;c<=c2.col;c++)ws.getCell(r,c).border=borda}
  function styleCell(addr,{bold=true,size=10,fill=null,color=PRETO,wrap=false,align="center"}={}){const c=ws.getCell(addr);c.font={name:"Calibri",size,bold,color:{argb:color}};c.alignment={horizontal:align,vertical:"middle",wrapText:wrap};if(fill)c.fill={type:"pattern",pattern:"solid",fgColor:{argb:fill}};c.border=borda;return c}
  fillRange("B2:I4",AZUL);ws.mergeCells("B2:C4");ws.mergeCells("D3:H3");
  ws.getCell("D3").value="Formulário de Não Atendimento Expansão";ws.getCell("D3").font={name:"Calibri",size:18,bold:true,color:{argb:BRANCO}};ws.getCell("D3").alignment={horizontal:"center",vertical:"middle"};
  ws.getCell("B2").border={right:{style:"medium",color:{argb:BRANCO}}};
  try{const logo=wb.addImage({base64:`data:image/png;base64,${LOGO_B64}`,extension:"png"});ws.addImage(logo,{tl:{col:1.15,row:1.18},br:{col:2.85,row:3.86},editAs:"oneCell"})}catch(e){}
  const secoes=[["B8:I8","Dados do Cliente:"],["B16:I16","Dados da Visita:"],["B24:I24","Motivo do expurgo:"],["B32:I32","Evidências:"]];
  secoes.forEach(([rng,txt])=>{ws.mergeCells(rng);const c=ws.getCell(rng.split(":")[0]);c.value=txt;c.fill={type:"pattern",pattern:"solid",fgColor:{argb:AZUL2}};c.font={name:"Calibri",size:10,bold:true,color:{argb:BRANCO}};c.alignment={horizontal:"left",vertical:"middle",indent:1}});
  const merges=["C10:E10","H10:I10","C12:I12","C14:I14","C18:E18","H18:I18","C20:E20","H20:I20","C22:I22","C26:I26","C28:I28","C30:I30","C34:E34","H34:I34","C36:E36","H36:I36","B38:I54"];merges.forEach(r=>ws.mergeCells(r));
  [["B6","Distribuidora:"],["E6","Regional:"],["H6","Data da\nsolicitação:"],["B10","Nº da nota:"],["G10","Conta Contrato:"],["B12","Parceiro de Negócios:"],["B14","Endereço:"],["B18","Data:"],["G18","Latitude:"],["B20","Horário:"],["G20","Longitude:"],["B22","Identificação da\nequipe:"],["B26","Justificativa:"],["B28","Descrição do\nExpurgo:"],["B30","Tratativa no Sistema\nComercial:"],["B34","Número do medidor\ndo cliente atendido:"],["G34","Número da nota do\natendimento em campo:"],["B36","Número do medidor\ndo vizinho:"],["G36","Número da estrutura\nmais próxima:"]].forEach(([a,v])=>{const c=styleCell(a,{bold:true,size:(a==="H6"?8:10),wrap:v.includes("\n")});c.value=v});
  [["C6","distribuidora"],["F6","regional"],["I6","data_solicitacao"],["C10","nota"],["H10","conta_contrato"],["C12","parceiro"],["C14","endereco"],["C18","data_visita"],["H18","latitude"],["C20","horario"],["H20","longitude"],["C22","equipe"],["C26","justificativa"],["C28","descricao_expurgo"],["C34","medidor_cliente"],["C36","medidor_vizinho"],["H36","estrutura_proxima"]].forEach(([a,id])=>{const c=styleCell(a,{bold:true,size:10});c.value=valorCampo(id)});
  styleCell("B30",{bold:true,size:9,fill:AZULCLARO,wrap:true});styleCell("C30",{bold:true,size:10,fill:AZULCLARO});ws.getCell("C30").value={formula:'IF(C26="","",VLOOKUP(C26,Apoio!$A$1:$B$27,2,0))',result:valorCampo("tratativa")};
  styleCell("H34",{bold:true,size:10});ws.getCell("H34").value={formula:"C10",result:valorCampo("nota_campo")||valorCampo("nota")};
  ["C6","F6","I6","C10","H10","C12","C14","C18","H18","C20","H20","C22","C26","C28","C30","C34","H34","C36","H36"].forEach(a=>{ws.getCell(a).alignment={horizontal:"center",vertical:"middle",wrapText:false}});
  borderRange("B6:C6");borderRange("E6:F6");borderRange("H6:I6");borderRange("B10:E10");borderRange("G10:I10");borderRange("B12:I12");borderRange("B14:I14");borderRange("B18:E18");borderRange("G18:I18");borderRange("B20:E20");borderRange("G20:I20");borderRange("B22:I22");borderRange("B26:I26");borderRange("B28:I28");borderRange("B30:I30");borderRange("B34:E34");borderRange("G34:I34");borderRange("B36:E36");borderRange("G36:I36");borderRange("B38:I54");
  ws.getCell("C26").dataValidation={type:"list",allowBlank:true,formulae:["Apoio!$A$2:$A$27"]};
  const apoio=wb.addWorksheet("Apoio");apoio.state="hidden";apoio.getCell("A1").value="Motivo";apoio.getCell("B1").value="Tratativa";let rr=2;for(const [mot,tr] of Object.entries(TRATATIVAS)){apoio.getCell(rr,1).value=mot;apoio.getCell(rr,2).value=tr;rr++}
  const grade=[[0,0,1,1],[0,0,.5,1],[.5,0,1,1],[0,0,1/3,1],[1/3,0,2/3,1],[2/3,0,1,1],[0,0,.5,.5],[.5,0,1,.5],[0,.5,.5,1],[.5,.5,1,1],[0,0,1/3,.5],[1/3,0,2/3,.5],[2/3,0,1,.5],[0,.5,.5,1],[.5,.5,1,1]];
  const layouts={1:[grade[0]],2:[grade[1],grade[2]],3:[grade[3],grade[4],grade[5]],4:[grade[6],grade[7],grade[8],grade[9]],5:[grade[10],grade[11],grade[12],grade[13],grade[14]]};
  if(FOTOS_ATUAIS.length){const lay=layouts[Math.min(5,FOTOS_ATUAIS.length)]||[];for(let i=0;i<Math.min(5,FOTOS_ATUAIS.length);i++){const url=FOTOS_ATUAIS[i];const ext=url.startsWith("data:image/png")?"png":"jpeg";try{const id=wb.addImage({base64:url,extension:ext});const [x1,y1,x2,y2]=lay[i];ws.addImage(id,{tl:{col:excelColPos(x1),row:excelRowPos(y1)},br:{col:excelColPos(x2),row:excelRowPos(y2)},editAs:"oneCell"})}catch(e){}}}
  const buf=await wb.xlsx.writeBuffer();const nome=(chaveNota(valorCampo("nota"))||"SEM_NOTA").replace(/[^A-Z0-9_-]/g,"_");baixarBlob(new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),`Relatorio_Expurgo_${nome}.xlsx`)
}
const BASE_NOTAS=__BASE_JSON__;
const TRATATIVAS={"ATENDIMENTO EM TERRENO DE TERCEIROS": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE ATENDIDO POR OUTRA ÁREA": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE ATENDIDO POR OUTRA SOLICITAÇÃO (EXPANSÃO)": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE CONSTRUIU A PRÓPRIA REDE SEM APRESENTAR PROJETO": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE DESISTIU DO SERVIÇO": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE EM ÁREA DE DOMÍNIO PÚBLICO/LITÍGIO": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE EM ÁREA DE LOTEAMENTO PARTICULAR/ CONDOMÍNIO": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE EM ÁREA DE PRESERVAÇÃO AMBIENTAL": "SUSPENSÃO", "CLIENTE EMUC": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE LIGADO A REVELIA (CLANDESTINO)": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE NÃO LOCALIZADO": "REJEIÇÃO PARA CANCELAMENTO", "CLIENTE URBANO SE ARRUAMENTO": "REJEIÇÃO PARA CANCELAMENTO", "ERRO NA ABERTURA - CLIENTE DESEJA OUTRO SERVIÇO": "REJEIÇÃO PARA CANCELAMENTO", "JÁ EXISTE REDE DE BT PARA ATENDER O CLIENTE": "DEVOLVER A GSTC", "OBSTACULO IMPEDE EXECUÇÃO/ SEM ACESSO": "SUSPENSÃO", "PADRÃO INCORRETO": "DEVOLVER A GSTC", "PADRÃO INEXISTENTE": "DEVOLVER A GSTC", "PADRÃO TRIFÁSICO INEXISTENTE": "DEVOLVER A GSTC", "PEDIDO EM DUPLICIDADE": "REJEIÇÃO PARA CANCELAMENTO", "RECLASSIFICAÇÃO RR PARA UB": "INFORMAR PARA AJUSTE NO INDICADOR", "RECLASSIFICAÇÃO UB PARA RR": "INFORMAR PARA AJUSTE NO INDICADOR", "SEM ACESSO PERIODO CHUVOSO": "SUSPENSÃO", "SOMENTE O TERRENO": "REJEIÇÃO PARA CANCELAMENTO", "UC DEMOLIDA/ ABANDONADA": "REJEIÇÃO PARA CANCELAMENTO", "CUSTEIO": "REJEIÇÃO PARA CANCELAMENTO"};
const idsAuto=["data_solicitacao","conta_contrato","parceiro","endereco","data_visita","horario","latitude","longitude","equipe","descricao_expurgo","tratativa","medidor_cliente","medidor_vizinho","estrutura_proxima"];
function chaveNota(v){return(v||"").trim().toUpperCase().replace(/\.0$/,"")}
function setInput(id,valor){const el=document.getElementById(id);if(el)el.value=valor||""}
function setSelect(id,valor){const el=document.getElementById(id);if(!el)return;const v=(valor||"").toString().trim().toUpperCase();if(!v){el.value="";return}let achou=false;for(const op of el.options){if((op.value||op.text).toString().trim().toUpperCase()===v){el.value=op.value;achou=true;break}}if(!achou){const op=document.createElement("option");op.value=valor;op.textContent=valor;el.appendChild(op);el.value=valor}}
function atualizarTratativa(preferirBase=""){const j=document.getElementById("justificativa").value||"";setInput("tratativa",preferirBase||TRATATIVAS[j]||"")}
function atualizarContadorFotos(qtd){const c=document.getElementById("photoCount");if(c)c.textContent=`${qtd}/5`}
function placeholderFotos(){return '<label for="photoInput" class="evidence-placeholder">Clique aqui ou em FOTOS 📷 para adicionar até 5 fotos</label>'}
function limparFotos(){FOTOS_ATUAIS=[];const grade=document.getElementById("evidenceGrid");grade.className="evidence-grid empty";grade.innerHTML=placeholderFotos();document.getElementById("photoInput").value="";atualizarContadorFotos(0)}
function lerFoto(arquivo){return new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=e=>resolve(e.target.result);reader.onerror=reject;reader.readAsDataURL(arquivo)})}
async function carregarFotos(fileList){const todos=Array.from(fileList||[]);if(!todos.length){limparFotos();return}if(todos.length>5){alert("Selecione no máximo 5 fotos. As 5 primeiras serão utilizadas.")}const arquivos=todos.slice(0,5);const urls=await Promise.all(arquivos.map(lerFoto));FOTOS_ATUAIS=urls.slice();const grade=document.getElementById("evidenceGrid");grade.className=`evidence-grid photos-${urls.length}`;grade.innerHTML="";urls.forEach((url,i)=>{const item=document.createElement("div");item.className="evidence-item";const img=document.createElement("img");img.src=url;img.alt=`Evidência ${i+1}`;item.appendChild(img);grade.appendChild(item)});atualizarContadorFotos(urls.length)}
function limparCamposDaNota(){setSelect("regional","");setSelect("distribuidora","EQTL MA");for(const id of idsAuto)setInput(id,"");setInput("equipe","EQP NIP");setSelect("justificativa","");setInput("nota_campo",chaveNota(document.getElementById("nota").value))}
function preencherPelaNota(){const nota=chaveNota(document.getElementById("nota").value);limparCamposDaNota();setInput("nota_campo",nota);const d=BASE_NOTAS[nota];if(!d)return;setSelect("distribuidora",d.distribuidora||"EQTL MA");setSelect("regional",d.regional||"");setInput("data_solicitacao",d.data_solicitacao||"");setInput("conta_contrato",d.conta_contrato||"");setInput("parceiro",d.parceiro||"");setInput("endereco",d.endereco||"");setInput("data_visita",d.data_visita||"");setInput("horario",d.horario||"");setInput("latitude",d.latitude||"");setInput("longitude",d.longitude||"");setInput("equipe",d.equipe||"EQP NIP");setInput("descricao_expurgo",d.descricao_expurgo||"");setInput("medidor_cliente",d.medidor_cliente||"");setInput("medidor_vizinho",d.medidor_vizinho||"");setInput("estrutura_proxima",d.estrutura_proxima||"");if(d.justificativa){setSelect("justificativa",d.justificativa);atualizarTratativa(d.tratativa||"")}else{setSelect("justificativa","");setInput("tratativa",d.tratativa||"")}}
function limparTudo(){document.getElementById("nota").value="";limparCamposDaNota();setInput("nota_campo","");limparFotos();document.getElementById("nota").focus()}
document.getElementById("nota").addEventListener("input",preencherPelaNota);document.getElementById("nota").addEventListener("change",preencherPelaNota);document.getElementById("justificativa").addEventListener("change",()=>atualizarTratativa(""));
document.getElementById("photoInput").addEventListener("change",function(evt){carregarFotos(evt.target.files)});
document.getElementById("nota").focus();
</script>
</body></html>'''

html_doc = (
    html_doc
    .replace("__LOGO__", LOGO_EQTL_EXCEL_B64)
    .replace("__OPTIONS__", opcoes_justificativa_html)
    .replace("__BASE_JSON__", base_json)
)

components.html(html_doc, height=1320, scrolling=False)
