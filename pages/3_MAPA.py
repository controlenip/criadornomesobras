import streamlit as st
import pandas as pd
import io
import zipfile
import re
import time
import os
import xml.etree.ElementTree as ET
import math
import requests
import unicodedata
import folium
from folium.plugins import MarkerCluster, MeasureControl, Draw
import gc
from streamlit_folium import st_folium
import html
from concurrent.futures import ThreadPoolExecutor
from scipy.spatial import cKDTree
import plotly.express as px
import sqlite3
import json
import shutil
from datetime import datetime

st.set_page_config(page_title="Gestão de Malha e Projetos", page_icon="🗺️", layout="wide")

# ==========================================
# CONFIGURAÇÕES CENTRAIS - NÃO ALTERAM A REGRA OPERACIONAL
# ==========================================
RAIO_CONFLITO_M = 50.0
RAIO_ARQUEOLOGIA_M = 150.0
LIMITE_REDES_SIMULTANEAS_PADRAO = 15
LIMITE_COORD_LAT = (-35.0, 5.0)
LIMITE_COORD_LON = (-75.0, -30.0)
CORES_STATUS_LIST = {
    '0': '#cbd5e1',
    'EM LEVANTAMENTO': '#22c55e',
    'ANALISE DE LEVANTAMENTO': '#eab308',
    'IMPRODUTIVO': '#f97316',
    'CORRECAO DE LEVANTAMENTO': '#7c3aed',
    'CONCLUIDO': '#1f77b4',
}

def _mtime_seguro(caminho):
    try:
        return os.path.getmtime(caminho)
    except OSError:
        return 0.0

def classificar_severidade_distancia(dist_m):
    try:
        d = float(dist_m)
    except Exception:
        return 'Sem classificação'
    if d <= 0.5:
        return '📌 Mesmo ponto (≈0 m)'
    if d <= 10:
        return '🔴 Crítico (≤ 10 m)'
    if d <= 25:
        return '🟠 Alto (10–25 m)'
    if d <= RAIO_CONFLITO_M:
        return '🟡 Médio (25–50 m)'
    return 'Fora do raio'

def dataframe_para_excel_bytes(df_export, nome_aba='Dados'):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name=str(nome_aba)[:31])
    return buffer.getvalue()

# ==========================================
# 1. MOTOR DE BANCO DE DADOS (SQLITE MIGRATION)
# ==========================================
def init_db_and_migrate():
    if not os.path.exists("database"):
        os.makedirs("database", exist_ok=True)
    
    conn = sqlite3.connect("database/redes.db")
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS malha
                 (ALIMENTADOR TEXT, REGIONAL TEXT, MUNICIPIO TEXT,
                  TIPO_GEOMETRIA TEXT, TIPO_REDE TEXT, NOME TEXT,
                  COORDS TEXT, COR TEXT)''')
    conn.commit()
    
    if os.path.exists("database/redes"):
        pkl_files = [f for f in os.listdir("database/redes") if f.endswith('.pkl')]
        for f in pkl_files:
            try:
                caminho_pkl = f"database/redes/{f}"
                df_pkl = pd.read_pickle(caminho_pkl)
                df_pkl['COORDS'] = df_pkl['COORDS'].apply(json.dumps)
                df_pkl.to_sql('malha', conn, if_exists='append', index=False)
                os.remove(caminho_pkl)
            except Exception:
                pass
    conn.close()

init_db_and_migrate()

# ==========================================
# 2. FUNÇÕES BASE, PLANILHA E GEOLOCALIZAÇÃO
# ==========================================
def remove_accents(input_str):
    nfkd_form = unicodedata.normalize('NFKD', str(input_str))
    return u"".join([c for c in nfkd_form if not unicodedata.combining(c)])

def latlon_to_xyz(lat, lon):
    R = 6371000.0
    lat_rad, lon_rad = math.radians(lat), math.radians(lon)
    return R * math.cos(lat_rad) * math.cos(lon_rad), R * math.cos(lat_rad) * math.sin(lon_rad), R * math.sin(lat_rad)

def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat/2.0)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2.0)**2
    c = 2 * math.asin(math.sqrt(a))
    return R * c

@st.cache_data(show_spinner=False)
def load_base_mapping():
    mun_to_reg = {}
    file_path = "MUNICIPIOS-REGIONAIS.xlsx"
    if not os.path.exists(file_path):
        st.sidebar.error(f"🚨 Planilha '{file_path}' não encontrada!")
    else:
        try:
            df_base = pd.read_excel(file_path)
            for _, row in df_base.iterrows():
                mun = remove_accents(str(row.get('MunicIpio', ''))).upper().strip()
                reg = str(row.get('Regional', '')).strip().upper()
                if mun and reg and reg != 'NAN': mun_to_reg[mun] = reg
        except Exception as e: st.sidebar.error(f"🚨 Erro: {e}")
    
    overrides_centro = ['SANTA LUZIA', 'CONCEICAO DO LAGO-ACU', 'CONCEICAO DO LAGO ACU', 'PINDARE-MIRIM', 'PINDARE MIRIM', 'OLHO DAGUA DAS CUNHAS', 'OLHO D\'AGUA DAS CUNHAS', 'GOVERNADOR LUIZ ROCHA']
    for mun in overrides_centro:
        if mun not in mun_to_reg: mun_to_reg[mun] = 'CENTRO'
    return mun_to_reg

@st.cache_data(show_spinner=False)
def get_base_geojson():
    mun_to_reg = load_base_mapping()
    url_geojson = "https://raw.githubusercontent.com/tbrugz/geodata-br/master/geojson/geojs-21-mun.json"
    try:
        resp = requests.get(url_geojson, headers={'User-Agent': 'Mozilla/5.0'}, timeout=5)
        geo_data = resp.json()
    except: return None

    reg_colors = {'LESTE': '#1f77b4', 'CENTRO': '#d62728', 'NOROESTE': '#ffed6f', 'NORTE': '#ff7f0e', 'SUL': '#8fbc8f', 'DESCONHECIDO': '#cccccc'}
    for feature in geo_data['features']:
        mun_name = feature['properties']['name']
        mun_name_norm = remove_accents(mun_name).upper().strip()
        reg = mun_to_reg.get(mun_name_norm, "DESCONHECIDO")
        feature['properties']['REGIONAL'] = reg
        feature['properties']['MUNICIPIO'] = mun_name_norm
        feature['properties']['fillColor'] = reg_colors.get(reg, '#cccccc')
    return geo_data

def is_point_in_polygon(lon, lat, polygon):
    inside = False
    for i in range(len(polygon)):
        p1x, p1y = polygon[i]
        p2x, p2y = polygon[(i + 1) % len(polygon)]
        if ((p1y > lat) != (p2y > lat)) and (lon < (p2x - p1x) * (lat - p1y) / (p2y - p1y + 1e-9) + p1x): inside = not inside
    return inside

def get_municipio_by_coord(lon, lat, geo_data):
    if not geo_data: return "N/A", "N/A"
    for feature in geo_data['features']:
        geom = feature['geometry']
        mun = feature['properties'].get('MUNICIPIO', 'N/A')
        reg = feature['properties'].get('REGIONAL', 'N/A')
        if geom['type'] == 'Polygon':
            for ring in geom['coordinates']:
                if is_point_in_polygon(lon, lat, ring): return mun, reg
        elif geom['type'] == 'MultiPolygon':
            for poly in geom['coordinates']:
                for ring in poly:
                    if is_point_in_polygon(lon, lat, ring): return mun, reg
    return "N/A", "N/A"

def extrair_coordenadas_vis(texto_coords):
    pontos = []
    for coord in texto_coords.strip().split():
        partes = coord.split(',')
        if len(partes) >= 2:
            try:
                lon = float(partes[0].strip().replace(',', '.'))
                lat = float(partes[1].strip().replace(',', '.'))
                if lat != 0.0 and lon != 0.0 and LIMITE_COORD_LAT[0] <= lat <= LIMITE_COORD_LAT[1] and LIMITE_COORD_LON[0] <= lon <= LIMITE_COORD_LON[1]: pontos.append([lat, lon]) 
            except: continue
    return pontos

def _repo_root():
    """Retorna a raiz do repositório mesmo quando esta página está dentro de /pages."""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _normalizar_nome_arquivo(nome):
    """Normaliza acentos/capitalização para localizar arquivos no Linux/Streamlit Cloud."""
    texto = unicodedata.normalize('NFKD', str(nome))
    texto = ''.join(c for c in texto if not unicodedata.combining(c))
    return texto.casefold().strip()


def resolver_arquivo_kml(nome_esperado):
    """Localiza o KML dentro de /kmls mesmo com pequenas diferenças de acento/capitalização."""
    pasta = os.path.join(_repo_root(), 'kmls')
    if not os.path.isdir(pasta):
        return None

    alvo = _normalizar_nome_arquivo(nome_esperado)
    for nome_real in os.listdir(pasta):
        if _normalizar_nome_arquivo(nome_real) == alvo:
            return os.path.join(pasta, nome_real)
    return None


def ler_kml_para_geojson(caminho_arquivo, cor_hex):
    """Lê Polygon/Point de KML de forma tolerante a namespaces XML."""
    if not caminho_arquivo or not os.path.exists(caminho_arquivo):
        return None

    try:
        with open(caminho_arquivo, 'r', encoding='utf-8-sig', errors='ignore') as f:
            kml_str = f.read()

        root = ET.fromstring(kml_str)
        features = []

        # {*} funciona com KMLs com ou sem namespace.
        for placemark in root.findall('.//{*}Placemark'):
            name_tag = placemark.find('{*}name')
            nome = name_tag.text.strip() if name_tag is not None and name_tag.text else "Área Demarcada"

            for poly in placemark.findall('.//{*}Polygon//{*}coordinates'):
                if not poly.text:
                    continue
                coords = []
                for coord_str in poly.text.strip().split():
                    partes = coord_str.split(',')
                    if len(partes) >= 2:
                        try:
                            coords.append([float(partes[0]), float(partes[1])])
                        except Exception:
                            pass
                if coords:
                    feat = {
                        "type": "Feature",
                        "properties": {"NOME": nome, "COR": cor_hex},
                        "geometry": {"type": "Polygon", "coordinates": [coords]},
                    }
                    ring = feat['geometry']['coordinates'][0]
                    lons = [pt[0] for pt in ring]
                    lats = [pt[1] for pt in ring]
                    feat['bboxes'] = [(min(lons), max(lons), min(lats), max(lats))]
                    features.append(feat)

            for pt in placemark.findall('.//{*}Point/{*}coordinates'):
                if not pt.text:
                    continue
                partes = pt.text.strip().split(',')
                if len(partes) >= 2:
                    try:
                        features.append({
                            "type": "Feature",
                            "properties": {"NOME": nome, "COR": cor_hex},
                            "geometry": {"type": "Point", "coordinates": [float(partes[0]), float(partes[1])]},
                        })
                    except Exception:
                        pass

        if features:
            return {"type": "FeatureCollection", "features": features}
        return None
    except Exception as exc:
        st.session_state.setdefault('_kml_erros', {})[os.path.basename(caminho_arquivo)] = str(exc)
        return None


@st.cache_data(show_spinner=False)
def get_kml_cached(nome_arquivo, color):
    caminho = resolver_arquivo_kml(nome_arquivo)
    return ler_kml_para_geojson(caminho, color)

def processar_um_kmz(f_name, f_bytes, base_map, geo_data):
    dict_cores = {'REDE PRIMÁRIA': '#e6194b', 'REDE PRIMARIA': '#e6194b', 'REDE SECUNDÁRIA': '#4363d8', 'REDE SECUNDARIA': '#4363d8', 'POSTE': '#808080', 'TRANSFORMADOR': '#f58231', 'CHAVE': '#3cb44b', 'REGULADOR': '#911eb4', 'RELIGADOR': '#46f0f0', 'CAPACITOR': '#ffe119', 'SUBESTAÇÃO': '#000000', 'SUBESTACAO': '#000000'}
    nome_arquivo = f_name.upper().replace('.KMZ', '').replace('.KML', '')
    conteudo_kml = ""
    if f_name.lower().endswith('.kmz'):
        try:
            with zipfile.ZipFile(io.BytesIO(f_bytes), 'r') as z:
                for item in z.namelist():
                    if item.lower().endswith('.kml'):
                        conteudo_kml = z.read(item).decode('utf-8', errors='ignore')
                        break
        except Exception: return None
    else:
        try: conteudo_kml = f_bytes.decode('utf-8', errors='ignore')
        except: return None
    conteudo_kml = re.sub(r'\sxmlns(:\w+)?="[^"]+"', '', conteudo_kml)
    try: root = ET.fromstring(conteudo_kml)
    except Exception: return None
    municipio, regional = "N/A", "N/A"
    mun_match = re.search(r'name=["\'](?:MUNICIPIO|CIDADE)["\'][^>]*>(.*?)</', conteudo_kml, re.IGNORECASE)
    if mun_match: 
        municipio = mun_match.group(1).strip().upper()
        mun_norm = remove_accents(municipio)
        if mun_norm in base_map: regional = base_map[mun_norm]
    if regional == "N/A":
        reg_match = re.search(r'name=["\'](?:REGIONAL|REGIAO)["\'][^>]*>(.*?)</', conteudo_kml, re.IGNORECASE)
        if reg_match: regional = reg_match.group(1).strip().upper()
    if regional == "N/A":
        sigla_match = re.search(r'\[([A-Z]{3})\]', nome_arquivo)
        if sigla_match: regional = sigla_match.group(1)
    registros_flat = []
    primeira_coord = None
    for folder in root.findall('.//Folder'):
        name_tag = folder.find('name')
        if name_tag is not None and name_tag.text:
            nome_pasta = name_tag.text.strip().upper()
            if "CEMAR" in nome_pasta or nome_arquivo in nome_pasta: continue
            cor_elemento = dict_cores.get(nome_pasta, '#333333')
            for placemark in folder.findall('.//Placemark'):
                pm_name_tag = placemark.find('name')
                nome_elemento = pm_name_tag.text.strip() if pm_name_tag is not None and pm_name_tag.text else "S/N"
                for ls in placemark.findall('.//LineString/coordinates'):
                    if ls.text:
                        coords = extrair_coordenadas_vis(ls.text)
                        if len(coords) > 1:
                            if primeira_coord is None:
                                primeira_coord = coords[0]
                            registros_flat.append({'ALIMENTADOR': nome_arquivo, 'REGIONAL': regional, 'MUNICIPIO': municipio, 'TIPO_GEOMETRIA': 'Linha', 'TIPO_REDE': nome_pasta, 'NOME': nome_elemento, 'COORDS': coords, 'COR': cor_elemento})
                for pt in placemark.findall('.//Point/coordinates'):
                    if pt.text:
                        coords = extrair_coordenadas_vis(pt.text)
                        if len(coords) > 0:
                            if primeira_coord is None:
                                primeira_coord = coords[0]
                            registros_flat.append({'ALIMENTADOR': nome_arquivo, 'REGIONAL': regional, 'MUNICIPIO': municipio, 'TIPO_GEOMETRIA': 'Ponto', 'TIPO_REDE': nome_pasta, 'NOME': nome_elemento, 'COORDS': coords[0], 'COR': cor_elemento})
    if municipio == "N/A" and primeira_coord is not None and geo_data is not None:
        mun_descob, reg_descob = get_municipio_by_coord(primeira_coord[1], primeira_coord[0], geo_data)
        if mun_descob != "N/A":
            municipio = mun_descob
            regional = reg_descob if reg_descob != "N/A" else regional
            for r in registros_flat:
                r['MUNICIPIO'] = municipio
                r['REGIONAL'] = regional
    if registros_flat: return pd.DataFrame(registros_flat)
    return None

def processar_e_salvar_kmz_paralelo(arquivos):
    base_map = load_base_mapping()
    geo_data = get_base_geojson()
    novos_processados = 0
    df_lote = []
    
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {
            executor.submit(processar_um_kmz, f.name, f.getvalue(), base_map, geo_data): f.name
            for f in arquivos
        }
        for future, nome_arquivo_processado in futures.items():
            try:
                df_alimentador = future.result()
            except Exception as e:
                # Um arquivo com problema não deve interromper todo o lote.
                st.warning(f"⚠️ Falha ao processar '{nome_arquivo_processado}': {e}")
                continue

            if df_alimentador is not None and not df_alimentador.empty:
                df_alimentador['COORDS'] = df_alimentador['COORDS'].apply(json.dumps)
                df_lote.append(df_alimentador)
                novos_processados += 1
                
    if df_lote:
        df_final = pd.concat(df_lote, ignore_index=True)
        conn = sqlite3.connect("database/redes.db")
        c = conn.cursor()
        alimentadores_inseridos = df_final['ALIMENTADOR'].unique().tolist()
        for alim in alimentadores_inseridos:
            c.execute("DELETE FROM malha WHERE ALIMENTADOR = ?", (alim,))
        df_final.to_sql('malha', conn, if_exists='append', index=False)
        conn.commit()
        conn.close()
        
    return novos_processados

@st.cache_data(show_spinner=False)
def carregar_banco_redes():
    try:
        conn = sqlite3.connect("database/redes.db")
        df = pd.read_sql("SELECT * FROM malha", conn)
        conn.close()
        if not df.empty:
            def safe_json_loads(x):
                try: return json.loads(x)
                except: return None
            df['COORDS'] = df['COORDS'].apply(safe_json_loads)
            df = df.dropna(subset=['COORDS'])
        return df
    except:
        return pd.DataFrame()

@st.cache_data(show_spinner=False)
def carregar_e_cruzar_obras(file_mtime=None):
    # file_mtime participa da chave do cache para invalidar automaticamente ao substituir o Excel.
    file_path = "BASE_LEVANTAMENTO_ATUALIZADA.xlsx"
    if not os.path.exists(file_path): return "Arquivo 'BASE_LEVANTAMENTO_ATUALIZADA.xlsx' não encontrado.", None, None, None
        
    try:
        df_obras = pd.read_excel(file_path)
        status_sisco_col = next((c for c in df_obras.columns if 'STATUS SISCO' in str(c).upper()), None)
        status_list_col = next((c for c in df_obras.columns if 'STATUS LIST' in str(c).upper()), None)
        lat_col = next((c for c in df_obras.columns if 'LATITUDE' in str(c).upper() or 'LAT' == str(c).upper()), None)
        lon_col = next((c for c in df_obras.columns if 'LONGITUDE' in str(c).upper() or 'LON' == str(c).upper()), None)
        
        if not all([status_sisco_col, status_list_col, lat_col, lon_col]):
            return "Erro: Colunas obrigatórias ausentes na planilha (Status ou Coordenadas).", None, None, None
            
        mun_col = next((c for c in df_obras.columns if 'MUNICIPIO' in str(c).upper() or 'CIDADE' in str(c).upper()), None)
        if mun_col:
            df_obras['MUNICIPIO_NORM'] = df_obras[mun_col].apply(lambda x: remove_accents(str(x)).upper().strip() if pd.notnull(x) else "DESCONHECIDO")
        else:
            df_obras['MUNICIPIO_NORM'] = "DESCONHECIDO"
            
        base_map = load_base_mapping()
        df_obras['REGIONAL_NORM'] = df_obras['MUNICIPIO_NORM'].map(base_map).fillna("DESCONHECIDO")

        data_col = next((c for c in df_obras.columns if 'DATA ABERTURA' in str(c).upper()), None)
        if data_col: df_obras['DATA_DT'] = pd.to_datetime(df_obras[data_col], errors='coerce')
        else: df_obras['DATA_DT'] = pd.NaT

        if df_obras[lat_col].dtype == object: df_obras[lat_col] = df_obras[lat_col].astype(str).str.replace(',', '.')
        if df_obras[lon_col].dtype == object: df_obras[lon_col] = df_obras[lon_col].astype(str).str.replace(',', '.')
        df_obras['LAT_CLEAN'] = pd.to_numeric(df_obras[lat_col], errors='coerce')
        df_obras['LON_CLEAN'] = pd.to_numeric(df_obras[lon_col], errors='coerce')
        
        mask_valid_coords = (
            (df_obras['LAT_CLEAN'].notnull()) & (df_obras['LON_CLEAN'].notnull()) & 
            (df_obras['LAT_CLEAN'] != 0.0) & (df_obras['LON_CLEAN'] != 0.0) & 
            (df_obras['LAT_CLEAN'] >= LIMITE_COORD_LAT[0]) & (df_obras['LAT_CLEAN'] <= LIMITE_COORD_LAT[1]) & 
            (df_obras['LON_CLEAN'] >= LIMITE_COORD_LON[0]) & (df_obras['LON_CLEAN'] <= LIMITE_COORD_LON[1])
        )
        total_linhas_base = int(len(df_obras))
        total_validas_coord = int(mask_valid_coords.sum())
        df_invalidas = df_obras[~mask_valid_coords].copy()

        def motivo_coord_invalida(row):
            lat = row.get('LAT_CLEAN')
            lon = row.get('LON_CLEAN')
            if pd.isna(lat) and pd.isna(lon):
                return 'LATITUDE E LONGITUDE AUSENTES/INVÁLIDAS'
            if pd.isna(lat):
                return 'LATITUDE AUSENTE/INVÁLIDA'
            if pd.isna(lon):
                return 'LONGITUDE AUSENTE/INVÁLIDA'
            if float(lat) == 0.0 and float(lon) == 0.0:
                return 'COORDENADA ZERADA (0,0)'
            if float(lat) == 0.0:
                return 'LATITUDE ZERADA'
            if float(lon) == 0.0:
                return 'LONGITUDE ZERADA'
            if not (-35.0 <= float(lat) <= 5.0) or not (-75.0 <= float(lon) <= -30.0):
                return 'COORDENADA FORA DO TERRITÓRIO BRASILEIRO'
            return 'COORDENADA INVÁLIDA'

        if not df_invalidas.empty:
            df_invalidas['MOTIVO DA INCONSISTÊNCIA'] = df_invalidas.apply(motivo_coord_invalida, axis=1)
            df_invalidas['NÍVEL'] = 'ERRO FATAL'
            df_invalidas['_TOTAL_LINHAS_BASE'] = total_linhas_base
            df_invalidas['_TOTAL_VALIDAS_COORD'] = total_validas_coord

        df_obras = df_obras[mask_valid_coords]
        
        # Regra oficial deste fluxo: obra concluída é definida por STATUS LIST = CONCLUIDO.
        status_list_norm = df_obras[status_list_col].apply(lambda x: remove_accents(str(x)).upper().strip())
        mask_concluida = status_list_norm.str.contains('CONCLUID', case=False, na=False)
        df_concluidas = df_obras[mask_concluida].copy()
        
        def normalizar(x): return remove_accents(str(x)).upper().strip()
        df_obras['STATUS_LIST_NORM'] = df_obras[status_list_col].apply(normalizar)
        status_alvos = ['0', 'EM LEVANTAMENTO', 'ANALISE DE LEVANTAMENTO', 'IMPRODUTIVO', 'CORRECAO DE LEVANTAMENTO']
        mask_andamento = df_obras['STATUS_LIST_NORM'].isin(status_alvos)
        df_andamento = df_obras[mask_andamento & (~mask_concluida)].copy()
        
        df_andamento['CONFLITO'] = False
        df_andamento['PROTOCOLO_CONFLITO'] = ""
        df_andamento['DISTANCIA_CONFLITO'] = 0.0
        df_andamento['NOME_CONCLUIDA'] = ""
        df_andamento['LAT_CONCLUIDA_CONFLITO'] = float('nan')
        df_andamento['LON_CONCLUIDA_CONFLITO'] = float('nan')
        # Campos adicionais de diagnóstico. A regra principal continua usando a concluída mais próxima.
        df_andamento['QTD_CONCLUIDAS_50M'] = 0
        df_andamento['PROTOCOLOS_CONCLUIDOS_50M'] = ''
        
        if not df_concluidas.empty and not df_andamento.empty:
            pts_concluidas = [latlon_to_xyz(row['LAT_CLEAN'], row['LON_CLEAN']) for _, row in df_concluidas.iterrows()]
            arvore_kdtree = cKDTree(pts_concluidas)
            c_flags, c_protos, c_dists, c_nomes, c_lats, c_lons = [], [], [], [], [], []
            c_qtd_50m, c_lista_50m = [], []
            for _, row in df_andamento.iterrows():
                xyz = latlon_to_xyz(row['LAT_CLEAN'], row['LON_CLEAN'])
                _, idx_mais_proximo = arvore_kdtree.query(xyz)
                obra_concluida_proxima = df_concluidas.iloc[idx_mais_proximo]
                lat_conc = float(obra_concluida_proxima['LAT_CLEAN'])
                lon_conc = float(obra_concluida_proxima['LON_CLEAN'])
                distancia_exata_m = haversine(row['LAT_CLEAN'], row['LON_CLEAN'], lat_conc, lon_conc) * 1000

                # Diagnóstico complementar: quantas concluídas existem no mesmo raio.
                # Não altera PROTOCOLO_CONFLITO nem a regra da concluída mais próxima.
                candidatos = arvore_kdtree.query_ball_point(xyz, r=RAIO_CONFLITO_M)
                protocolos_no_raio = []
                for idx_cand in candidatos:
                    cand = df_concluidas.iloc[idx_cand]
                    d_cand = haversine(row['LAT_CLEAN'], row['LON_CLEAN'], float(cand['LAT_CLEAN']), float(cand['LON_CLEAN'])) * 1000
                    if d_cand <= RAIO_CONFLITO_M:
                        protocolos_no_raio.append(str(cand.get('PROTOCOLO', 'S/N')))
                c_qtd_50m.append(len(protocolos_no_raio))
                c_lista_50m.append(' | '.join(dict.fromkeys(protocolos_no_raio)))

                if distancia_exata_m <= RAIO_CONFLITO_M:
                    c_flags.append(True); c_protos.append(str(obra_concluida_proxima.get('PROTOCOLO', 'S/N')))
                    c_dists.append(distancia_exata_m); c_nomes.append(str(obra_concluida_proxima.get('NOME', 'S/N')))
                    c_lats.append(lat_conc); c_lons.append(lon_conc)
                else:
                    c_flags.append(False); c_protos.append(""); c_dists.append(0.0); c_nomes.append("")
                    c_lats.append(float('nan')); c_lons.append(float('nan'))
            df_andamento['CONFLITO'] = c_flags
            df_andamento['PROTOCOLO_CONFLITO'] = c_protos
            df_andamento['DISTANCIA_CONFLITO'] = c_dists
            df_andamento['NOME_CONCLUIDA'] = c_nomes
            df_andamento['LAT_CONCLUIDA_CONFLITO'] = c_lats
            df_andamento['LON_CONCLUIDA_CONFLITO'] = c_lons
            df_andamento['QTD_CONCLUIDAS_50M'] = c_qtd_50m
            df_andamento['PROTOCOLOS_CONCLUIDOS_50M'] = c_lista_50m
            
        return "OK", df_concluidas, df_andamento, df_invalidas
    except Exception as e: return f"Erro processando dados: {str(e)}", None, None, None

# ==========================================
# FUNÇÕES AUXILIARES DE DIAGNÓSTICO (ADITIVAS)
# ==========================================
@st.cache_data(show_spinner=False)
def carregar_base_obras_completa(file_mtime=None):
    caminho = "BASE_LEVANTAMENTO_ATUALIZADA.xlsx"
    if not os.path.exists(caminho):
        return pd.DataFrame()
    try:
        d = pd.read_excel(caminho)
        d.columns = [str(c).strip() for c in d.columns]
        return d
    except Exception:
        return pd.DataFrame()

@st.cache_data(show_spinner=False)
def buscar_protocolo_na_base(protocolo, file_mtime=None):
    termo = str(protocolo).strip().upper()
    if not termo:
        return None
    d = carregar_base_obras_completa(file_mtime)
    if d.empty:
        return None
    col_p = next((c for c in d.columns if 'PROTOCOLO' in str(c).upper() or str(c).upper() in ['NOTA', 'Nº DA NOTA', 'NUMERO DA NOTA']), None)
    lat_col = next((c for c in d.columns if 'LATITUDE' in str(c).upper() or str(c).upper() == 'LAT'), None)
    lon_col = next((c for c in d.columns if 'LONGITUDE' in str(c).upper() or str(c).upper() == 'LON'), None)
    if not all([col_p, lat_col, lon_col]):
        return None
    serie = d[col_p].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()
    achou = d[serie == termo]
    if achou.empty:
        achou = d[serie.str.contains(re.escape(termo), na=False)]
    if achou.empty:
        return None
    row = achou.iloc[0].copy()
    try:
        lat = float(str(row[lat_col]).replace(',', '.'))
        lon = float(str(row[lon_col]).replace(',', '.'))
    except Exception:
        return {'encontrado': True, 'coordenada_valida': False, 'row': row.to_dict()}
    valido = (lat != 0 and lon != 0 and LIMITE_COORD_LAT[0] <= lat <= LIMITE_COORD_LAT[1] and LIMITE_COORD_LON[0] <= lon <= LIMITE_COORD_LON[1])
    return {'encontrado': True, 'coordenada_valida': valido, 'lat': lat, 'lon': lon, 'row': row.to_dict()}

@st.cache_data(show_spinner=False)
def analisar_qualidade_base(file_mtime=None):
    d = carregar_base_obras_completa(file_mtime)
    if d.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    lat_col = next((c for c in d.columns if 'LATITUDE' in str(c).upper() or str(c).upper() == 'LAT'), None)
    lon_col = next((c for c in d.columns if 'LONGITUDE' in str(c).upper() or str(c).upper() == 'LON'), None)
    prot_col = next((c for c in d.columns if 'PROTOCOLO' in str(c).upper()), None)
    mun_col = next((c for c in d.columns if 'MUNICIPIO' in str(c).upper() or 'CIDADE' in str(c).upper()), None)
    if not lat_col or not lon_col:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    lat = pd.to_numeric(d[lat_col].astype(str).str.replace(',', '.'), errors='coerce')
    lon = pd.to_numeric(d[lon_col].astype(str).str.replace(',', '.'), errors='coerce')
    d2 = d.copy()
    d2['_LAT_Q'] = lat; d2['_LON_Q'] = lon
    valid = lat.notna() & lon.notna() & (lat != 0) & (lon != 0) & lat.between(*LIMITE_COORD_LAT) & lon.between(*LIMITE_COORD_LON)
    d2['_COORD_VALIDA'] = valid
    dup_coords = d2[valid & d2.duplicated(subset=['_LAT_Q','_LON_Q'], keep=False)].copy()
    dup_prot = pd.DataFrame()
    if prot_col:
        ps = d2[prot_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        d2['_PROTO_Q'] = ps
        dup_prot = d2[ps.ne('') & ps.ne('nan') & ps.duplicated(keep=False)].copy()
    qualidade_mun = pd.DataFrame()
    if mun_col:
        tmp = pd.DataFrame({'MUNICIPIO': d2[mun_col].fillna('DESCONHECIDO').astype(str), 'VALIDA': valid.astype(int)})
        qualidade_mun = tmp.groupby('MUNICIPIO', as_index=False).agg(TOTAL=('VALIDA','size'), VALIDAS=('VALIDA','sum'))
        qualidade_mun['QUALIDADE_%'] = (qualidade_mun['VALIDAS'] / qualidade_mun['TOTAL'] * 100).round(1)
    return dup_coords, dup_prot, qualidade_mun

def _distancia_aprox_poligono_m(lat, lon, ring):
    if is_point_in_polygon(lon, lat, ring):
        return 0.0
    menor = float('inf')
    # Distância aos vértices é usada apenas como diagnóstico aproximado; a regra de interseção existente não muda.
    for pt in ring[::max(1, len(ring)//200)]:
        try:
            menor = min(menor, haversine(lat, lon, float(pt[1]), float(pt[0])) * 1000)
        except Exception:
            pass
    return menor

def analisar_proximidade_areas_especiais(lat, lon, limite_m=500.0):
    resultados = []
    for categoria, geo_d in dict_areas_especiais.items():
        if not geo_d:
            continue
        melhor = float('inf'); melhor_nome = ''
        for feat in geo_d.get('features', []):
            geom = feat.get('geometry', {})
            nome = feat.get('properties', {}).get('NOME', 'Sem Nome')
            if geom.get('type') == 'Polygon':
                for ring in geom.get('coordinates', []):
                    d = _distancia_aprox_poligono_m(lat, lon, ring)
                    if d < melhor:
                        melhor, melhor_nome = d, nome
            elif geom.get('type') == 'Point':
                try:
                    pt_lon, pt_lat = geom.get('coordinates', [None, None])
                    d = haversine(lat, lon, float(pt_lat), float(pt_lon)) * 1000
                    if d < melhor:
                        melhor, melhor_nome = d, nome
                except Exception:
                    pass
        if melhor <= limite_m:
            faixa = 'Dentro da área' if melhor <= 1 else ('Até 100 m' if melhor <= 100 else '100–500 m')
            resultados.append({'categoria': categoria, 'nome': melhor_nome, 'distancia_m_aprox': melhor, 'faixa': faixa})
    return resultados

def formatar_resumo_filtros(regioes, municipios, alimentadores, statuses):
    partes = []
    if regioes: partes.append('Regional: ' + ', '.join(regioes))
    if municipios: partes.append('Município: ' + ', '.join(municipios[:4]) + ('…' if len(municipios) > 4 else ''))
    if alimentadores: partes.append(f'Alimentadores: {len(alimentadores)} selecionado(s)')
    if statuses: partes.append('Status: ' + ', '.join(statuses))
    return ' | '.join(partes) if partes else 'Nenhum filtro geográfico/status aplicado'

# ==========================================
# 2. ESTRUTURA DA TELA E CONTAINERS
# ==========================================
st.markdown("<h2 style='color: #0D256C;'>🗺️ Gestão de Malha Elétrica e Obras (Inteligência Geográfica)</h2>", unsafe_allow_html=True)

# Foco opcional recebido da página CRIAR SGO.
# Quando existe conflito em até 50 m, a página já abre enquadrando as obras envolvidas.
foco_mapa_sgo = st.session_state.get("foco_mapa_conflito")
if foco_mapa_sgo and foco_mapa_sgo.get('conflitos'):
    conflitos_recebidos = foco_mapa_sgo.get('conflitos', [])
    obras_recebidas = foco_mapa_sgo.get('obras_digitadas', [])
    menor_distancia = min(float(x.get('distancia_m', 999999)) for x in conflitos_recebidos)
    st.error(
        f"🚨 Conflito recebido do CRIAR SGO: {len(obras_recebidas)} solicitação(ões) "
        f"com obra concluída em até 50 m. Menor distância: {menor_distancia:.1f} m."
    )
    if st.button("↩️ Limpar foco do CRIAR SGO", key="limpar_foco_sgo_mapa"):
        st.session_state.pop("foco_mapa_conflito", None)
        st.rerun()

kpi_container = st.container()
map_container = st.container()
table_container = st.container()

df = carregar_banco_redes()
base_map = load_base_mapping()
geo_data_ibge = get_base_geojson()

# ==========================================
# PREPARAÇÃO DAS ÁREAS ESPECIAIS
# ==========================================
def preprocessar_bboxes_kml(geo_data):
    if not geo_data: return
    for feat in geo_data['features']:
        geom = feat['geometry']
        if geom['type'] == 'Polygon':
            bboxes = []
            for ring in geom['coordinates']:
                lons = [pt[0] for pt in ring]
                lats = [pt[1] for pt in ring]
                bboxes.append((min(lons), max(lons), min(lats), max(lats)))
            feat['bboxes'] = bboxes

geo_q = get_kml_cached("Áreas Quilombolas.kml", "#ff7f00"); preprocessar_bboxes_kml(geo_q)
geo_i = get_kml_cached("Terras Indigenas.kml", "#2ca02c"); preprocessar_bboxes_kml(geo_i)
geo_a = get_kml_cached("Sítios Arqueológicos.kml", "#8c564b"); preprocessar_bboxes_kml(geo_a)
geo_uc_fed = get_kml_cached("UC Federal.kml", "#e6b800"); preprocessar_bboxes_kml(geo_uc_fed)
geo_uc_est = get_kml_cached("UC Estadual.kml", "#ffff00"); preprocessar_bboxes_kml(geo_uc_est)
geo_uc_mun = get_kml_cached("UC Municipal.kml", "#ffff00"); preprocessar_bboxes_kml(geo_uc_mun)

dict_areas_especiais = {
    "Quilombo": geo_q, "Terra Indígena": geo_i, "Sítio Arqueológico": geo_a,
    "UC Federal": geo_uc_fed, "UC Estadual": geo_uc_est, "UC Municipal": geo_uc_mun
}

def verificar_areas_da_obra(lat, lon):
    encontradas = []
    for categoria, geo_data in dict_areas_especiais.items():
        if not geo_data: continue
        for feat in geo_data['features']:
            geom = feat['geometry']
            nome = feat['properties'].get('NOME', 'Sem Nome')
            if geom['type'] == 'Polygon':
                for i, ring in enumerate(geom['coordinates']):
                    if 'bboxes' in feat:
                        min_lon, max_lon, min_lat, max_lat = feat['bboxes'][i]
                        if (min_lon <= lon <= max_lon) and (min_lat <= lat <= max_lat):
                            if is_point_in_polygon(lon, lat, ring):
                                encontradas.append(f"<b>{categoria}:</b> {html.escape(nome)}")
                                break
            elif geom['type'] == 'Point':
                pt_lon, pt_lat = geom['coordinates']
                if haversine(lat, lon, pt_lat, pt_lon) * 1000 <= RAIO_ARQUEOLOGIA_M:
                    encontradas.append(f"<b>{categoria}:</b> {html.escape(nome)} (Raio {int(RAIO_ARQUEOLOGIA_M)}m)")
    return "<br>".join(encontradas) if encontradas else "Nenhuma restrição"

# ==========================================
# 3. INTERFACE E SINCRONIZAÇÃO VIA GITHUB
# ==========================================
with st.sidebar:
    # ------------------------------------------
    # CONTROLES RÁPIDOS / PRESETS
    # ------------------------------------------
    with st.expander("⚙️ Visualizações Rápidas", expanded=True):
        p1, p2 = st.columns(2)
        if p1.button("🚨 Conflitos", use_container_width=True, key="preset_conflitos"):
            st.session_state.update({
                'chk_todas_obras': False, 'chk_concluidas': True, 'chk_conflitantes': True,
                'mostrar_status_0': False, 'mostrar_status_em_levantamento': False,
                'mostrar_status_analise_levantamento': False, 'mostrar_status_improdutivo': False,
                'mostrar_status_correcao_levantamento': False,
            })
            st.rerun()
        if p2.button("⚡ Malha", use_container_width=True, key="preset_malha"):
            st.session_state.update({
                'chk_todas_obras': False, 'chk_concluidas': False, 'chk_conflitantes': False,
                'mostrar_status_0': False, 'mostrar_status_em_levantamento': False,
                'mostrar_status_analise_levantamento': False, 'mostrar_status_improdutivo': False,
                'mostrar_status_correcao_levantamento': False,
                'chk_quilombos': False, 'chk_indigenas': False, 'chk_arqueologia': False,
                'chk_uc_federal': False, 'chk_uc_estadual': False, 'chk_uc_municipal': False,
            })
            st.rerun()
        p3, p4 = st.columns(2)
        if p3.button("🌳 Restrições", use_container_width=True, key="preset_restricoes"):
            st.session_state.update({
                'chk_quilombos': True, 'chk_indigenas': True, 'chk_arqueologia': True,
                'chk_uc_federal': True, 'chk_uc_estadual': True, 'chk_uc_municipal': True,
                'chk_todas_obras': True,
            })
            st.rerun()
        if p4.button("📍 Todas Obras", use_container_width=True, key="preset_todas"):
            st.session_state.update({'chk_todas_obras': True, 'chk_concluidas': False, 'chk_conflitantes': False})
            st.rerun()

        if st.button("↩️ Restaurar visualização padrão", use_container_width=True, key="reset_visualizacao"):
            chaves_limpar = [
                'filtro_regionais','filtro_municipios','filtro_alimentadores','busca_nome_rede',
                'busca_lat','busca_lon','busca_protocolo','foco_mapa_conflito',
                'chk_todas_obras','chk_concluidas','chk_conflitantes','mostrar_status_0',
                'mostrar_status_em_levantamento','mostrar_status_analise_levantamento',
                'mostrar_status_improdutivo','mostrar_status_correcao_levantamento',
                'chk_quilombos','chk_indigenas','chk_arqueologia','chk_uc_federal',
                'chk_uc_estadual','chk_uc_municipal','somente_restricoes','limite_redes_mapa'
            ]
            for k in chaves_limpar:
                st.session_state.pop(k, None)
            st.rerun()

    with st.expander("📥 1. Banco de Dados e Sincronização", expanded=True):
        st.markdown("A ferramenta lê as redes automaticamente da pasta **`kmzs`** no repositório.")
        pasta_kmz = "kmzs"
        if not os.path.exists(pasta_kmz):
            os.makedirs(pasta_kmz, exist_ok=True)
            
        arquivos_repositorio = [f for f in os.listdir(pasta_kmz) if f.lower().endswith(('.kmz', '.kml'))]
        alims_no_banco = set(df['ALIMENTADOR'].tolist()) if not df.empty else set()
        
        arquivos_novos = []
        for f in arquivos_repositorio:
            nome_alim = f.upper().replace('.KMZ', '').replace('.KML', '')
            if nome_alim not in alims_no_banco:
                arquivos_novos.append(f)
                
        if arquivos_novos:
            st.info(f"📂 {len(arquivos_novos)} arquivo(s) novo(s) na pasta '{pasta_kmz}' aguardando processamento.")
            
            if st.button(f"🚀 Sincronizar {len(arquivos_novos)} Novas Redes", type="primary", use_container_width=True):
                class LocalFileAdapter:
                    def __init__(self, filepath):
                        self.name = os.path.basename(filepath)
                        self.filepath = filepath
                    def getvalue(self):
                        with open(self.filepath, 'rb') as f:
                            return f.read()
                            
                lista_adapters = [LocalFileAdapter(os.path.join(pasta_kmz, f)) for f in arquivos_novos]
                
                qtd_total_processados = 0
                tamanho_lote = 15
                total_lotes = math.ceil(len(lista_adapters) / tamanho_lote)
                barra_progresso = st.progress(0.0)
                texto_status = st.empty()
                inicio_sync = time.time()
                
                for i in range(0, len(lista_adapters), tamanho_lote):
                    lote_atual = (i // tamanho_lote) + 1
                    lote_arquivos = lista_adapters[i:i+tamanho_lote]
                    texto_status.text(f"⏳ Processando e Salvando Lote {lote_atual} de {total_lotes}...")
                    qtd_total_processados += processar_e_salvar_kmz_paralelo(lote_arquivos)
                    barra_progresso.progress(lote_atual / total_lotes)
                    gc.collect()

                duracao_sync = time.time() - inicio_sync
                qtd_falhas = max(0, len(lista_adapters) - qtd_total_processados)
                st.session_state['_ultima_sincronizacao'] = {
                    'quando': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                    'arquivos': len(lista_adapters),
                    'processados': qtd_total_processados,
                    'falhas': qtd_falhas,
                    'duracao_s': round(duracao_sync, 1),
                }
                if qtd_total_processados > 0:
                    st.success(f"✅ Sincronização finalizada! {qtd_total_processados} redes salvas em {duracao_sync:.1f}s. Falhas/ignorados: {qtd_falhas}.")
                    carregar_banco_redes.clear()
                    time.sleep(1)
                    st.rerun()
                else:
                    st.warning(f"⚠️ Nenhuma rede nova foi gravada. Verifique os {qtd_falhas} arquivo(s) processados/ignorados.")
        else:
            st.success(f"✅ O banco de dados está atualizado.")
        if st.session_state.get('_ultima_sincronizacao'):
            us = st.session_state['_ultima_sincronizacao']
            st.caption(f"Última sincronização: {us.get('quando')} | {us.get('processados', 0)}/{us.get('arquivos', 0)} processados | {us.get('falhas', 0)} falhas/ignorados | {us.get('duracao_s', 0)}s")

    with st.expander("🔎 2. Pesquisas Inteligentes", expanded=False):
        tab_nome, tab_coord, tab_obra = st.tabs(["📝 Por Nome/ID", "📍 Por Coordenada", "🏗️ Por Protocolo"])
        termo_pesquisa, busca_lat, busca_lon, protocolo_pesquisa = "", None, None, ""
        with tab_nome:
            termo_pesquisa = st.text_input("Nome/Num. Poste ou Trafo:", placeholder="Ex: 554930...", key="busca_nome_rede").strip().upper()
        with tab_coord:
            c_lat, c_lon = st.columns(2)
            with c_lat: lat_input = st.text_input("Latitude:", placeholder="Ex: -5.532", key="busca_lat")
            with c_lon: lon_input = st.text_input("Longitude:", placeholder="Ex: -47.432", key="busca_lon")
            if lat_input and lon_input:
                try:
                    b_lat, b_lon = float(lat_input.replace(',', '.').strip()), float(lon_input.replace(',', '.').strip())
                    if LIMITE_COORD_LAT[0] <= b_lat <= LIMITE_COORD_LAT[1] and LIMITE_COORD_LON[0] <= b_lon <= LIMITE_COORD_LON[1]: busca_lat, busca_lon = b_lat, b_lon
                    else: st.warning("⚠️ Coordenada fora do Brasil.")
                except: st.warning("⚠️ Formato inválido.")
        with tab_obra:
            protocolo_pesquisa = st.text_input("Nº da Nota / Protocolo:", placeholder="Ex: 430163831", key="busca_protocolo").strip().upper()
            st.caption("Ao localizar uma coordenada válida, o mapa centraliza e destaca a obra sem alterar os filtros existentes.")

    with st.expander("🔍 3. Filtros Geográficos", expanded=True):
        lista_regioes = sorted(list(set(base_map.values()))) if base_map else ["CENTRO", "LESTE", "NOROESTE", "NORTE", "SUL"]
        regioes_sel = st.multiselect("📍 Regional:", lista_regioes, key="filtro_regionais")
        
        lista_municipios = []
        for mun, reg in base_map.items():
            if not regioes_sel or reg in regioes_sel: lista_municipios.append(mun)
        municipios_sel = st.multiselect("🏙️ Município (Foco e Contorno):", sorted(lista_municipios), key="filtro_municipios")
        
        df_filt = df.copy()
        if not df.empty:
            if regioes_sel: df_filt = df_filt[df_filt['REGIONAL'].isin(regioes_sel)]
            if municipios_sel: df_filt = df_filt[df_filt['MUNICIPIO'].isin(municipios_sel)]
        
        lista_alimentadores = sorted(df_filt['ALIMENTADOR'].unique().tolist()) if not df_filt.empty else []
        alim_sel = st.multiselect("⚡ Alimentador:", lista_alimentadores, key="filtro_alimentadores")
        
        if 'limite_redes_mapa' not in st.session_state:
            st.session_state['limite_redes_mapa'] = LIMITE_REDES_SIMULTANEAS_PADRAO
        limite_redes_atual = int(st.session_state.get('limite_redes_mapa', LIMITE_REDES_SIMULTANEAS_PADRAO))
        if not alim_sel:
            if len(lista_alimentadores) > limite_redes_atual:
                st.warning(f"⚠️ **Proteção de Memória:** {len(lista_alimentadores)} redes detectadas. Exibindo {limite_redes_atual}. Use os filtros ou carregue mais redes gradualmente.")
                alimentadores_visiveis = lista_alimentadores[:limite_redes_atual]
                ca, cb = st.columns(2)
                if ca.button("➕ Carregar +15", use_container_width=True, key="mais_15_redes"):
                    st.session_state['limite_redes_mapa'] = min(len(lista_alimentadores), limite_redes_atual + 15)
                    st.rerun()
                if limite_redes_atual > LIMITE_REDES_SIMULTANEAS_PADRAO and cb.button("↩️ Voltar para 15", use_container_width=True, key="voltar_15_redes"):
                    st.session_state['limite_redes_mapa'] = LIMITE_REDES_SIMULTANEAS_PADRAO
                    st.rerun()
            else:
                alimentadores_visiveis = lista_alimentadores
            st.caption(f"Redes visíveis: {len(alimentadores_visiveis)} de {len(lista_alimentadores)}")
        else:
            alimentadores_visiveis = alim_sel

    camadas_ativas = {}
    if not df.empty and alimentadores_visiveis:
        with st.expander("🗂️ 4. Camadas (Desempenho)", expanded=False):
            for alim in alimentadores_visiveis:
                st.markdown(f"**{alim}**")
                lista_camadas_alim = sorted(df[df['ALIMENTADOR'] == alim]['TIPO_REDE'].unique().tolist())
                camadas_essenciais = ['REDE PRIMÁRIA', 'REDE PRIMARIA', 'REDE SECUNDÁRIA', 'REDE SECUNDARIA', 'TRANSFORMADOR', 'POSTE']
                camadas_default = [c for c in lista_camadas_alim if c in camadas_essenciais]
                camadas_ativas[alim] = st.multiselect("Visibilidade das Camadas:", lista_camadas_alim, default=camadas_default, key=f"ms_{alim}")
            
    with st.expander("🗺️ 5. Áreas Especiais", expanded=False):
        mostrar_quilombos = st.checkbox("🟠 Áreas Quilombolas", value=False, key="chk_quilombos")
        mostrar_indigenas = st.checkbox("🟢 Terras Indígenas", value=False, key="chk_indigenas")
        mostrar_arqueologia = st.checkbox("🟤 Sítios Arqueológicos", value=False, key="chk_arqueologia")
        mostrar_uc_federal = st.checkbox("🟡 UC Federal", value=False, key="chk_uc_federal")
        mostrar_uc_estadual = st.checkbox("🟡 UC Estadual", value=False, key="chk_uc_estadual")
        mostrar_uc_municipal = st.checkbox("🟡 UC Municipal", value=False, key="chk_uc_municipal")

        status_kml = {
            "Áreas Quilombolas": geo_q,
            "Terras Indígenas": geo_i,
            "Sítios Arqueológicos": geo_a,
            "UC Federal": geo_uc_fed,
            "UC Estadual": geo_uc_est,
            "UC Municipal": geo_uc_mun,
        }
        carregados = [nome for nome, geo in status_kml.items() if geo and geo.get('features')]
        ausentes = [nome for nome, geo in status_kml.items() if not geo or not geo.get('features')]
        if carregados:
            st.caption("✅ KML carregados: " + ", ".join(carregados))
        if ausentes:
            st.warning("⚠️ KML sem dados/carregamento: " + ", ".join(ausentes))
        if st.session_state.get('_kml_erros'):
            with st.expander("Detalhes técnicos dos KML"):
                for nome, erro in st.session_state['_kml_erros'].items():
                    st.code(f"{nome}: {erro}")
    
    # Evita cruzamentos geoespaciais caros quando nenhuma camada especial está ativa.
    alguma_area_especial_ativa = any([
        mostrar_quilombos, mostrar_indigenas, mostrar_arqueologia,
        mostrar_uc_federal, mostrar_uc_estadual, mostrar_uc_municipal
    ])

    with st.expander("🚧 6. Obras e Projetos", expanded=True):
        mostrar_todas_obras = st.checkbox("📍 TODAS AS OBRAS (Clusters)", value=False, key="chk_todas_obras")
        mostrar_concluidas = st.checkbox("🔵 OBRAS CONCLUÍDAS", value=False, key="chk_concluidas") 
        mostrar_conflitantes = st.checkbox(f"🚨 OBRAS CONFLITANTES (Raio {int(RAIO_CONFLITO_M)}m)", value=False, key="chk_conflitantes")

        st.markdown("---")
        mostrar_status_0 = st.checkbox("⚪ STATUS LIST: 0", value=False, key="mostrar_status_0")
        mostrar_em_levantamento = st.checkbox("🟢 STATUS LIST: EM LEVANTAMENTO", value=False, key="mostrar_status_em_levantamento")
        mostrar_analise_levantamento = st.checkbox("🟡 STATUS LIST: ANÁLISE DE LEVANTAMENTO", value=False, key="mostrar_status_analise_levantamento")
        mostrar_improdutivo = st.checkbox("🟠 STATUS LIST: IMPRODUTIVO", value=False, key="mostrar_status_improdutivo")
        mostrar_correcao_levantamento = st.checkbox("🟣 STATUS LIST: CORREÇÃO DE LEVANTAMENTO", value=False, key="mostrar_status_correcao_levantamento")
        somente_restricoes = st.checkbox("🌳 SOMENTE OBRAS COM RESTRIÇÃO/PROXIMIDADE ESPECIAL", value=False, key="somente_restricoes", help="Filtro adicional. Não altera a regra de conflito; apenas mantém obras dentro ou até 500 m das áreas especiais carregadas.")

        # Mantém uma lista interna apenas para o processamento do mapa.
        # Para o usuário, cada STATUS LIST aparece como uma caixa de marcação independente.
        status_list_sel = []
        if mostrar_status_0:
            status_list_sel.append("0")
        if mostrar_em_levantamento:
            status_list_sel.append("EM LEVANTAMENTO")
        if mostrar_analise_levantamento:
            status_list_sel.append("ANÁLISE DE LEVANTAMENTO")
        if mostrar_improdutivo:
            status_list_sel.append("IMPRODUTIVO")
        if mostrar_correcao_levantamento:
            status_list_sel.append("CORREÇÃO DE LEVANTAMENTO")
        
        msg_obras, df_concluidas, df_andamento, df_invalidas = "OK", None, None, None
        if mostrar_concluidas or mostrar_conflitantes or mostrar_todas_obras or bool(status_list_sel) or bool(foco_mapa_sgo) or bool(protocolo_pesquisa) or somente_restricoes:
            obras_mtime = _mtime_seguro("BASE_LEVANTAMENTO_ATUALIZADA.xlsx")
            msg_obras, df_concluidas, df_andamento, df_invalidas = carregar_e_cruzar_obras(obras_mtime)
            if msg_obras != "OK": st.sidebar.warning(f"⚠️ {msg_obras}")
            else:
                if regioes_sel:
                    if df_concluidas is not None and not df_concluidas.empty: df_concluidas = df_concluidas[df_concluidas['REGIONAL_NORM'].isin(regioes_sel)]
                    if df_andamento is not None and not df_andamento.empty: df_andamento = df_andamento[df_andamento['REGIONAL_NORM'].isin(regioes_sel)]
                    if df_invalidas is not None and not df_invalidas.empty: df_invalidas = df_invalidas[df_invalidas['REGIONAL_NORM'].isin(regioes_sel)]
                if municipios_sel:
                    if df_concluidas is not None and not df_concluidas.empty: df_concluidas = df_concluidas[df_concluidas['MUNICIPIO_NORM'].isin(municipios_sel)]
                    if df_andamento is not None and not df_andamento.empty: df_andamento = df_andamento[df_andamento['MUNICIPIO_NORM'].isin(municipios_sel)]
                    if df_invalidas is not None and not df_invalidas.empty: df_invalidas = df_invalidas[df_invalidas['MUNICIPIO_NORM'].isin(municipios_sel)]

                # Filtro opcional por STATUS LIST. A normalização remove acentos para
                # casar corretamente com a coluna STATUS_LIST_NORM gerada na leitura.
                if status_list_sel and df_andamento is not None and not df_andamento.empty:
                    status_norm_sel = [remove_accents(x).upper().strip() for x in status_list_sel]
                    df_andamento = df_andamento[df_andamento['STATUS_LIST_NORM'].isin(status_norm_sel)]

                # Filtro opcional de restrições: cálculo paralelo, sem interferir no motor de conflitos.
                if somente_restricoes:
                    def _filtrar_restricao(dfx):
                        if dfx is None or dfx.empty:
                            return dfx
                        mask = dfx.apply(lambda r: bool(analisar_proximidade_areas_especiais(float(r['LAT_CLEAN']), float(r['LON_CLEAN']), 500.0)), axis=1)
                        return dfx[mask].copy()
                    df_concluidas = _filtrar_restricao(df_concluidas)
                    df_andamento = _filtrar_restricao(df_andamento)

                qtd_conflitos = df_andamento['CONFLITO'].sum() if df_andamento is not None else 0
                st.caption(f"🔵 Concluídas: {len(df_concluidas) if df_concluidas is not None else 0} | 🟢/🟡/🟠/🟣 Em análise: {len(df_andamento) if df_andamento is not None else 0} | 🚨 Conflitos: {int(qtd_conflitos)}")
                
    with st.expander("🗑️ 7. Gerenciar Malha Local", expanded=False):
        alim_para_deletar = st.selectbox("Apagar Alimentador do Banco:", ["Selecione..."] + sorted(df['ALIMENTADOR'].unique().tolist()) if not df.empty else ["Selecione..."], key="alim_excluir")
        if alim_para_deletar != "Selecione...":
            qtd_elementos_excluir = int((df['ALIMENTADOR'] == alim_para_deletar).sum()) if not df.empty else 0
            st.warning(f"⚠️ {alim_para_deletar}: {qtd_elementos_excluir:,} elemento(s) serão removidos do SQLite.".replace(',', '.'))
            confirmar_exclusao = st.checkbox("Confirmo a exclusão deste alimentador", key="confirmar_exclusao_alim")
            if st.button("❌ Excluir Permanentemente", use_container_width=True, disabled=not confirmar_exclusao, key="btn_excluir_alim"):
                db_path = "database/redes.db"
                if os.path.exists(db_path):
                    os.makedirs("database/backups", exist_ok=True)
                    backup_path = os.path.join("database", "backups", f"redes_antes_exclusao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db")
                    shutil.copy2(db_path, backup_path)
                conn = sqlite3.connect(db_path)
                c = conn.cursor()
                c.execute("DELETE FROM malha WHERE ALIMENTADOR = ?", (alim_para_deletar,))
                conn.commit()
                conn.close()
                carregar_banco_redes.clear()
                st.success(f"✅ Alimentador excluído. Backup preventivo criado em {backup_path if os.path.exists(db_path) else 'database/backups' }.")
                time.sleep(1)
                st.rerun()

# ==========================================
# DIAGNÓSTICO OPERACIONAL E FILTROS ATIVOS
# ==========================================
obras_mtime_diag = _mtime_seguro("BASE_LEVANTAMENTO_ATUALIZADA.xlsx")
dup_coords_diag, dup_prot_diag, qualidade_mun_diag = analisar_qualidade_base(obras_mtime_diag)

with st.expander("🩺 Diagnóstico do Sistema", expanded=False):
    d1, d2, d3, d4 = st.columns(4)
    d1.metric("Base de obras", "OK" if os.path.exists("BASE_LEVANTAMENTO_ATUALIZADA.xlsx") else "AUSENTE")
    d2.metric("Municípios/Regionais", "OK" if os.path.exists("MUNICIPIOS-REGIONAIS.xlsx") else "AUSENTE")
    d3.metric("SQLite", "OK" if os.path.exists("database/redes.db") else "AUSENTE")
    d4.metric("Alimentadores", len(df['ALIMENTADOR'].unique()) if not df.empty else 0)
    kml_ok = sum(1 for g in [geo_q, geo_i, geo_a, geo_uc_fed, geo_uc_est, geo_uc_mun] if g and g.get('features'))
    kml_erro = 6 - kml_ok
    e1, e2, e3, e4 = st.columns(4)
    e1.metric("KML válidos", kml_ok)
    e2.metric("KML ausentes/erro", kml_erro)
    e3.metric("Coords. duplicadas", len(dup_coords_diag))
    e4.metric("Protocolos repetidos", len(dup_prot_diag))
    if os.path.exists("BASE_LEVANTAMENTO_ATUALIZADA.xlsx"):
        st.caption("Última alteração da base: " + datetime.fromtimestamp(obras_mtime_diag).strftime('%d/%m/%Y %H:%M:%S'))
    if st.session_state.get('_ultima_sincronizacao'):
        st.caption("Última sincronização da malha: " + str(st.session_state['_ultima_sincronizacao']))

resumo_filtros_ativos = formatar_resumo_filtros(regioes_sel, municipios_sel, alim_sel, status_list_sel)
st.info("🔎 **Filtros ativos:** " + resumo_filtros_ativos)

st.markdown(
    """
    <div style="display:flex;flex-wrap:wrap;gap:10px;margin:4px 0 14px 0;font-size:12px;">
      <span style="padding:6px 10px;border-radius:999px;background:#dbeafe;">🔵 Concluída</span>
      <span style="padding:6px 10px;border-radius:999px;background:#fee2e2;">🔴 Conflito / raio 50 m</span>
      <span style="padding:6px 10px;border-radius:999px;background:#dcfce7;">🟢 Em levantamento</span>
      <span style="padding:6px 10px;border-radius:999px;background:#fef9c3;">🟡 Análise de levantamento</span>
      <span style="padding:6px 10px;border-radius:999px;background:#ffedd5;">🟠 Improdutivo</span>
      <span style="padding:6px 10px;border-radius:999px;background:#ede9fe;">🟣 Correção de levantamento</span>
      <span style="padding:6px 10px;border-radius:999px;background:#f1f5f9;">⚪ Status 0</span>
    </div>
    """,
    unsafe_allow_html=True
)

# ==========================================
# DASHBOARD DE INDICADORES E GRÁFICOS
# ==========================================
def render_kpi(icone, titulo, valor, cor_borda):
    return f"""
    <div style="background-color: white; border-radius: 8px; padding: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-left: 6px solid {cor_borda}; text-align: left; height: 100%;">
        <p style="margin: 0; font-size: 14px; color: #666; font-weight: 600;">{icone} {titulo}</p>
        <p style="margin: 0; font-size: 32px; color: #222; font-weight: 800; padding-top: 5px;">{valor}</p>
    </div>
    """

with kpi_container:
    c1, c2, c3, c4 = st.columns(4)
    val_alim = len(df['ALIMENTADOR'].unique()) if not df.empty else 0
    val_conc = len(df_concluidas) if df_concluidas is not None else 0
    val_anda = len(df_andamento) if df_andamento is not None else 0
    val_conf = df_andamento['CONFLITO'].sum() if df_andamento is not None else 0
    
    c1.markdown(render_kpi("⚡", "ALIMENTADORES MAPEADOS", val_alim, "#808080"), unsafe_allow_html=True)
    c2.markdown(render_kpi("🔵", "OBRAS CONCLUÍDAS", val_conc, "#1f77b4"), unsafe_allow_html=True)
    c3.markdown(render_kpi("🟢", "OBRAS EM ANDAMENTO", val_anda, "#2ca02c"), unsafe_allow_html=True)
    c4.markdown(render_kpi("🚨", "CONFLITOS (50m)", val_conf, "#d62728"), unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    if msg_obras == "OK" and val_conf > 0:
        df_conf = df_andamento[df_andamento['CONFLITO']].copy()
        if not df_conf.empty:
            st.markdown("### 📊 Análise dos Conflitos")
            st.caption("Visão gerencial das cidades, status e criticidade das obras sobrepostas em até 50 metros.")

            # Padroniza os status usados nos gráficos para as mesmas nomenclaturas do mapa.
            def _status_dashboard(valor):
                s = remove_accents(str(valor)).upper().strip()
                if s in ['', 'NAN', 'NONE', 'SEM INFORMACAO']:
                    return '0'
                if s == '0':
                    return '0'
                if 'CORRECAO DE LEVANTAMENTO' in s:
                    return 'Correção de levantamento'
                if 'ANALISE DE LEVANTAMENTO' in s:
                    return 'Análise de levantamento'
                if 'IMPRODUTIVO' in s:
                    return 'Improdutivo'
                if 'EM LEVANTAMENTO' in s:
                    return 'Em levantamento'
                return str(valor).strip() or '0'

            df_conf['_STATUS_DASH'] = df_conf['STATUS LIST'].apply(_status_dashboard)
            df_conf['_DIST_M'] = pd.to_numeric(df_conf['DISTANCIA_CONFLITO'], errors='coerce').fillna(0.0)
            df_conf['_SEVERIDADE_DASH'] = df_conf['_DIST_M'].apply(classificar_severidade_distancia)

            # Controle do ranking sem poluir o gráfico quando há muitos municípios.
            top_opt = st.radio(
                "Quantidade de cidades no ranking:",
                ["Top 10", "Top 15", "Todos"],
                horizontal=True,
                index=0,
                key="ranking_cidades_conflito"
            )
            limite_top = 10 if top_opt == "Top 10" else (15 if top_opt == "Top 15" else None)

            # ---------- Linha 1: Cidades + Status ----------
            col_chart1, col_chart2 = st.columns(2, gap="large")
            with col_chart1:
                df_barras = (
                    df_conf['MUNICIPIO_NORM']
                    .fillna('SEM MUNICÍPIO')
                    .replace('', 'SEM MUNICÍPIO')
                    .value_counts()
                    .rename_axis('Município')
                    .reset_index(name='Conflitos')
                )
                if limite_top is not None:
                    df_barras = df_barras.head(limite_top)
                # Para barra horizontal, ordem crescente coloca o maior no topo visual.
                df_barras_plot = df_barras.sort_values('Conflitos', ascending=True)
                total_conf = max(1, int(len(df_conf)))
                df_barras_plot['Percentual'] = (df_barras_plot['Conflitos'] / total_conf * 100).round(1)
                df_barras_plot['Rótulo'] = df_barras_plot.apply(
                    lambda r: f"{int(r['Conflitos'])} ({r['Percentual']:.1f}%)", axis=1
                )

                fig1 = px.bar(
                    df_barras_plot,
                    x='Conflitos',
                    y='Município',
                    orientation='h',
                    title="📍 Cidades com Mais Conflitos",
                    text='Rótulo',
                    color_discrete_sequence=['#d62728']
                )
                fig1.update_traces(
                    textposition='outside',
                    hovertemplate='<b>%{y}</b><br>Conflitos: %{x}<extra></extra>'
                )
                fig1.update_layout(
                    xaxis_title="Quantidade de conflitos",
                    yaxis_title="",
                    showlegend=False,
                    height=max(360, 28 * len(df_barras_plot) + 120),
                    margin=dict(l=10, r=55, t=55, b=30)
                )
                st.plotly_chart(fig1, use_container_width=True)

            with col_chart2:
                ordem_status = ['0', 'Em levantamento', 'Análise de levantamento', 'Improdutivo', 'Correção de levantamento']
                cores_status = {
                    '0': '#94a3b8',
                    'Em levantamento': '#22c55e',
                    'Análise de levantamento': '#eab308',
                    'Improdutivo': '#f97316',
                    'Correção de levantamento': '#7c3aed'
                }
                df_status = (
                    df_conf['_STATUS_DASH']
                    .value_counts()
                    .rename_axis('Status')
                    .reset_index(name='Quantidade')
                )
                df_status['Status'] = pd.Categorical(df_status['Status'], categories=ordem_status, ordered=True)
                df_status = df_status.sort_values('Status').dropna(subset=['Status'])

                fig2 = px.pie(
                    df_status,
                    names='Status',
                    values='Quantidade',
                    title="📊 Status das Obras Sobrepostas",
                    hole=0.58,
                    color='Status',
                    color_discrete_map=cores_status,
                    category_orders={'Status': ordem_status}
                )
                fig2.update_traces(
                    textposition='inside',
                    textinfo='percent',
                    hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>'
                )
                fig2.add_annotation(
                    text=f"<b>{len(df_conf)}</b><br>Conflitos",
                    x=0.5, y=0.5, showarrow=False,
                    font=dict(size=17)
                )
                fig2.update_layout(
                    legend_title_text='STATUS LIST',
                    height=max(360, 28 * len(df_barras_plot) + 120),
                    margin=dict(l=10, r=10, t=55, b=20)
                )
                st.plotly_chart(fig2, use_container_width=True)

            # ---------- Linha 2: Severidade + Cidade x Severidade ----------
            col_chart3, col_chart4 = st.columns(2, gap="large")
            with col_chart3:
                ordem_sev = ['📌 Mesmo ponto (≈0 m)', '🔴 Crítico (≤ 10 m)', '🟠 Alto (10–25 m)', '🟡 Médio (25–50 m)']
                cores_sev = {
                    '📌 Mesmo ponto (≈0 m)': '#7f1d1d',
                    '🔴 Crítico (≤ 10 m)': '#dc2626',
                    '🟠 Alto (10–25 m)': '#f97316',
                    '🟡 Médio (25–50 m)': '#eab308'
                }
                df_sev = (
                    df_conf['_SEVERIDADE_DASH']
                    .value_counts()
                    .reindex(ordem_sev, fill_value=0)
                    .rename_axis('Severidade')
                    .reset_index(name='Quantidade')
                )
                fig3 = px.bar(
                    df_sev,
                    x='Severidade',
                    y='Quantidade',
                    title="🚨 Conflitos por Faixa de Distância",
                    color='Severidade',
                    color_discrete_map=cores_sev,
                    text='Quantidade',
                    category_orders={'Severidade': ordem_sev}
                )
                fig3.update_traces(
                    textposition='outside',
                    hovertemplate='<b>%{x}</b><br>Conflitos: %{y}<extra></extra>'
                )
                fig3.update_layout(
                    xaxis_title="",
                    yaxis_title="Quantidade de conflitos",
                    showlegend=False,
                    height=390,
                    margin=dict(l=10, r=20, t=55, b=70)
                )
                st.plotly_chart(fig3, use_container_width=True)

            with col_chart4:
                cidades_top = df_barras['Município'].tolist()
                df_city_sev = df_conf[df_conf['MUNICIPIO_NORM'].isin(cidades_top)].copy()
                df_city_sev['MUNICIPIO_NORM'] = df_city_sev['MUNICIPIO_NORM'].fillna('SEM MUNICÍPIO').replace('', 'SEM MUNICÍPIO')
                pivot = (
                    df_city_sev.groupby(['MUNICIPIO_NORM', '_SEVERIDADE_DASH'])
                    .size()
                    .unstack(fill_value=0)
                    .reindex(columns=ordem_sev, fill_value=0)
                )
                # Ordena pelo total de conflitos para manter leitura gerencial.
                pivot['_TOTAL'] = pivot.sum(axis=1)
                pivot = pivot.sort_values('_TOTAL', ascending=True).drop(columns=['_TOTAL']).reset_index()
                df_stack = pivot.melt(
                    id_vars='MUNICIPIO_NORM',
                    value_vars=ordem_sev,
                    var_name='Severidade',
                    value_name='Quantidade'
                )
                fig4 = px.bar(
                    df_stack,
                    x='Quantidade',
                    y='MUNICIPIO_NORM',
                    orientation='h',
                    color='Severidade',
                    title="🏙️ Cidade × Severidade",
                    color_discrete_map=cores_sev,
                    category_orders={'Severidade': ordem_sev}
                )
                fig4.update_layout(
                    barmode='stack',
                    xaxis_title="Quantidade de conflitos",
                    yaxis_title="",
                    legend_title_text='Severidade',
                    height=max(390, 28 * len(pivot) + 120),
                    margin=dict(l=10, r=20, t=55, b=30)
                )
                fig4.update_traces(hovertemplate='<b>%{y}</b><br>%{fullData.name}: %{x}<extra></extra>')
                st.plotly_chart(fig4, use_container_width=True)

# Busca direta de protocolo/nota. É independente do motor de conflitos.
resultado_protocolo = None
protocolo_zoom_lat = protocolo_zoom_lon = None
if protocolo_pesquisa:
    resultado_protocolo = buscar_protocolo_na_base(protocolo_pesquisa, _mtime_seguro("BASE_LEVANTAMENTO_ATUALIZADA.xlsx"))
    if not resultado_protocolo:
        st.sidebar.warning("⚠️ Protocolo/nota não localizado na base de obras.")
    elif not resultado_protocolo.get('coordenada_valida'):
        st.sidebar.warning("⚠️ Protocolo localizado, mas sem coordenada válida para posicionar no mapa.")
    else:
        protocolo_zoom_lat = float(resultado_protocolo['lat'])
        protocolo_zoom_lon = float(resultado_protocolo['lon'])
        st.sidebar.success(f"🎯 Obra localizada em {protocolo_zoom_lat:.6f}, {protocolo_zoom_lon:.6f}")

# ==========================================
# 4. CONSTRUÇÃO DO MAPA FOLIUM E SIMBOLOGIA
# ==========================================
mapa = folium.Map(location=[-5.2, -45.0], zoom_start=6, tiles=None, prefer_canvas=True)

if protocolo_zoom_lat is not None and protocolo_zoom_lon is not None:
    row_busca = (resultado_protocolo or {}).get('row', {})
    nome_busca = str(row_busca.get('NOME', row_busca.get('NOME DA OBRA', 'S/N')))
    status_busca = str(row_busca.get('STATUS LIST', 'S/N'))
    mun_busca = str(row_busca.get('MUNICIPIO', row_busca.get('MUNICÍPIO', 'S/N')))
    popup_busca = f"""
    <div style='min-width:260px;font-family:sans-serif;'>
      <h4 style='margin:0 0 8px 0;color:#0D256C;border-bottom:2px solid #0D256C;padding-bottom:5px;'>🏗️ OBRA LOCALIZADA</h4>
      <b>PROTOCOLO:</b> {html.escape(protocolo_pesquisa)}<br>
      <b>NOME:</b> {html.escape(nome_busca)}<br>
      <b>STATUS LIST:</b> {html.escape(status_busca)}<br>
      <b>MUNICÍPIO:</b> {html.escape(mun_busca)}<br>
      <b>COORDENADAS:</b> {protocolo_zoom_lat:.6f}, {protocolo_zoom_lon:.6f}
    </div>
    """
    folium.Marker(
        [protocolo_zoom_lat, protocolo_zoom_lon],
        tooltip=f"Obra: {html.escape(protocolo_pesquisa)}",
        popup=folium.Popup(popup_busca, max_width=340),
        icon=folium.Icon(color='cadetblue', icon='search', prefix='fa')
    ).add_to(mapa)

mapa.add_child(MeasureControl(position='topleft', primary_length_unit='meters', primary_area_unit='sqmeters'))
Draw(export=False, position='topleft').add_to(mapa)

js_draw_loc = """
<script>
    setTimeout(function() {
        try {
            if (typeof L !== 'undefined' && L.drawLocal && L.drawLocal.draw && L.drawLocal.draw.toolbar) {
                L.drawLocal.draw.toolbar.actions.title = 'Cancelar desenho';
                L.drawLocal.draw.toolbar.actions.text = 'Cancelar';
                L.drawLocal.draw.toolbar.finish.title = 'Finalizar desenho';
                L.drawLocal.draw.toolbar.finish.text = 'Finalizar';
                L.drawLocal.draw.toolbar.undo.title = 'Desfazer último ponto';
                L.drawLocal.draw.toolbar.undo.text = 'Desfazer';
                L.drawLocal.draw.toolbar.buttons.polygon = 'Desenhar um polígono';
                L.drawLocal.draw.toolbar.buttons.polyline = 'Desenhar uma linha';
                L.drawLocal.draw.toolbar.buttons.rectangle = 'Desenhar um retângulo';
                L.drawLocal.draw.toolbar.buttons.circle = 'Desenhar um círculo';
                L.drawLocal.draw.toolbar.buttons.marker = 'Adicionar um marcador';
                L.drawLocal.draw.toolbar.buttons.circlemarker = 'Adicionar marcador circular';
            }
            if (typeof L !== 'undefined' && L.drawLocal && L.drawLocal.edit && L.drawLocal.edit.toolbar) {
                L.drawLocal.edit.toolbar.actions.save.title = 'Salvar alterações';
                L.drawLocal.edit.toolbar.actions.save.text = 'Salvar';
                L.drawLocal.edit.toolbar.actions.cancel.title = 'Cancelar edição';
                L.drawLocal.edit.toolbar.actions.cancel.text = 'Cancelar';
                L.drawLocal.edit.toolbar.actions.clearAll.title = 'Apagar todos os desenhos';
                L.drawLocal.edit.toolbar.actions.clearAll.text = 'Apagar Tudo';
            }
        } catch (e) { console.log("Folium Draw Erro: ", e); }
    }, 500);
</script>
"""
mapa.get_root().html.add_child(folium.Element(js_draw_loc))

folium.TileLayer(tiles='https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}', attr='Google', name='Satélite (Google Maps)', overlay=False, control=True, max_zoom=20).add_to(mapa)
folium.TileLayer(tiles='OpenStreetMap', name='Mapa Base (Limpo)', overlay=False, control=True, max_zoom=20).add_to(mapa)
folium.TileLayer(
    tiles='https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Dark_Gray_Base/MapServer/tile/{z}/{y}/{x}',
    attr='Tiles &copy; Esri &mdash; Esri, DeLorme, NAVTEQ',
    name='Mapa Base (Escuro - Foco em Redes)',
    overlay=False,
    control=True,
    max_zoom=20
).add_to(mapa)

if geo_data_ibge:
    def style_function(feature):
        reg_mun = feature['properties'].get('MUNICIPIO', '')
        reg_name = feature['properties'].get('REGIONAL', '')
        cor_regiao = feature['properties']['fillColor']
        if municipios_sel:
            if reg_mun in municipios_sel: return {'fillColor': cor_regiao, 'color': '#FF00FF', 'weight': 3, 'fillOpacity': 0.1}
            else: return {'fillColor': 'transparent', 'color': 'transparent', 'weight': 0}
        elif regioes_sel:
            if reg_name in regioes_sel: return {'fillColor': cor_regiao, 'color': cor_regiao, 'weight': 1, 'fillOpacity': 0.1}
            else: return {'fillColor': 'transparent', 'color': 'transparent', 'weight': 0}
        return {'fillColor': 'transparent', 'color': cor_regiao, 'weight': 1, 'fillOpacity': 0.0}

    folium.GeoJson(geo_data_ibge, name="Divisão IBGE (Maranhão)", style_function=style_function, tooltip=folium.features.GeoJsonTooltip(fields=['name', 'REGIONAL'], aliases=['Município:', 'Regional:'], style="background-color: white; color: #333; font-family: arial; font-size: 12px; padding: 10px;"), zoom_on_click=False, show=True).add_to(mapa)

todas_lats, todas_lons = [], []
busca_lats, busca_lons = [], []

if not df.empty:
    df_mapa = df.copy()
    if regioes_sel: df_mapa = df_mapa[df_mapa['REGIONAL'].isin(regioes_sel)]
    if municipios_sel: df_mapa = df_mapa[df_mapa['MUNICIPIO'].isin(municipios_sel)]
    df_mapa = df_mapa[df_mapa['ALIMENTADOR'].isin(alimentadores_visiveis)]

    mask_camadas = pd.Series(False, index=df_mapa.index)
    for alim in alimentadores_visiveis:
        if alim in camadas_ativas: mask_camadas = mask_camadas | ((df_mapa['ALIMENTADOR'] == alim) & (df_mapa['TIPO_REDE'].isin(camadas_ativas[alim])))
    df_mapa = df_mapa[mask_camadas]

    grid_pts, grid_info = [], []
    if not df_mapa.empty:
        for idx, row in df_mapa.iterrows():
            if row['TIPO_GEOMETRIA'] == 'Ponto':
                pt_lat, pt_lon = row['COORDS'][0], row['COORDS'][1]
                grid_pts.append(latlon_to_xyz(pt_lat, pt_lon))
                grid_info.append((row['TIPO_REDE'], row['NOME'], pt_lat, pt_lon))
            else:
                for pt in row['COORDS']:
                    pt_lat, pt_lon = pt[0], pt[1]
                    grid_pts.append(latlon_to_xyz(pt_lat, pt_lon))
                    grid_info.append((row['TIPO_REDE'], row['NOME'], pt_lat, pt_lon))
    tree_grid = cKDTree(grid_pts) if grid_pts else None

    df_busca = pd.DataFrame()
    nearest_idx = None

    if busca_lat is not None and busca_lon is not None and not df_mapa.empty:
        pts, indices = [], []
        for idx, row in df_mapa.iterrows():
            if row['TIPO_GEOMETRIA'] == 'Ponto':
                pts.append(latlon_to_xyz(row['COORDS'][0], row['COORDS'][1]))
                indices.append(idx)
            else:
                for pt in row['COORDS']:
                    pts.append(latlon_to_xyz(pt[0], pt[1]))
                    indices.append(idx)
                    
        if pts:
            tree = cKDTree(pts)
            target_xyz = latlon_to_xyz(busca_lat, busca_lon)
            dist_3d, min_idx_in_pts = tree.query(target_xyz)
            nearest_idx = indices[min_idx_in_pts]
            
            elem_prox = df_mapa.loc[nearest_idx]
            if elem_prox['TIPO_GEOMETRIA'] == 'Ponto': dist_metros = haversine(busca_lat, busca_lon, elem_prox['COORDS'][0], elem_prox['COORDS'][1]) * 1000
            else: dist_metros = min([haversine(busca_lat, busca_lon, pt[0], pt[1]) for pt in elem_prox['COORDS']]) * 1000
            
            st.sidebar.success(f"🎯 **Alvo mais próximo:** {elem_prox['TIPO_REDE']} ({elem_prox['NOME']}) a {dist_metros:.1f} metros.")
            df_busca = df_mapa.loc[[nearest_idx]]
            df_mapa = df_mapa.drop(nearest_idx)

    elif termo_pesquisa != "":
        mask_nome = df_mapa['NOME'].astype(str).str.contains(termo_pesquisa, case=False, na=False)
        df_busca = df_mapa[mask_nome]
        df_mapa = df_mapa[~mask_nome]

    dict_cores_render = {
        'REDE PRIMÁRIA': '#0000FF', 
        'REDE PRIMARIA': '#0000FF', 
        'REDE SECUNDÁRIA': '#FF00FF', 
        'REDE SECUNDARIA': '#FF00FF', 
        'POSTE': '#808080', 
        'TRANSFORMADOR': '#FFFF00', 
    }

    features_linhas = []
    features_postes = []
    features_trafos = []
    features_outros = []

    for _, row in df_mapa.iterrows():
        tipo_rede = str(row['TIPO_REDE']).upper()
        
        cor_oficial = dict_cores_render.get(tipo_rede, row['COR'])
        if pd.isna(cor_oficial) or not cor_oficial:
            cor_oficial = '#333333'
            
        if row['TIPO_GEOMETRIA'] == 'Linha':
            coords = [[pt[1], pt[0]] for pt in row['COORDS']]
            if len(coords) > 1:
                features_linhas.append({
                    "type": "Feature",
                    "geometry": {"type": "LineString", "coordinates": coords},
                    "properties": {
                        "TIPO_REDE": html.escape(tipo_rede), "NOME": html.escape(str(row['NOME'])), "COR": cor_oficial, 
                        "ALIMENTADOR": html.escape(str(row['ALIMENTADOR'])), "LOC": f"{html.escape(str(row['MUNICIPIO']))} - {html.escape(str(row['REGIONAL']))}"
                    }
                })
                for pt in row['COORDS']: todas_lats.append(pt[0]); todas_lons.append(pt[1])
        else:
            lat, lon = row['COORDS'][0], row['COORDS'][1]
            todas_lats.append(lat); todas_lons.append(lon)
            
            feat = {
                "type": "Feature",
                "geometry": {"type": "Point", "coordinates": [lon, lat]},
                "properties": {
                    "TIPO_REDE": html.escape(tipo_rede), "NOME": html.escape(str(row['NOME'])), "COR": cor_oficial, 
                    "ALIMENTADOR": html.escape(str(row['ALIMENTADOR'])), "LOC": f"{html.escape(str(row['MUNICIPIO']))} - {html.escape(str(row['REGIONAL']))}",
                    "GPS": f"{lat:.5f}, {lon:.5f}"
                }
            }
            
            if 'POSTE' in tipo_rede:
                features_postes.append(feat)
            elif 'TRANSFORMADOR' in tipo_rede:
                features_trafos.append(row)
            else:
                features_outros.append(feat)

    if features_linhas:
        folium.GeoJson(
            {"type": "FeatureCollection", "features": features_linhas},
            name="Redes (Linhas)",
            style_function=lambda x: {'color': x['properties']['COR'], 'weight': 4 if 'PRIM' in x['properties']['TIPO_REDE'] else 3, 'opacity': 0.9},
            tooltip=folium.features.GeoJsonTooltip(fields=['TIPO_REDE', 'NOME'], aliases=['Rede:', 'Identificação:']),
            popup=folium.features.GeoJsonPopup(fields=['TIPO_REDE', 'NOME', 'ALIMENTADOR', 'LOC'], aliases=['Rede:', 'Identificação:', 'Alimentador:', 'Localização:'])
        ).add_to(mapa)

    if features_postes:
        folium.GeoJson(
            {"type": "FeatureCollection", "features": features_postes},
            name="Postes",
            marker=folium.CircleMarker(radius=4, color='black', weight=1, fillColor='gray', fillOpacity=1.0),
            tooltip=folium.features.GeoJsonTooltip(fields=['TIPO_REDE', 'NOME'], aliases=['Rede:', 'Identificação:']),
            popup=folium.features.GeoJsonPopup(fields=['TIPO_REDE', 'NOME', 'ALIMENTADOR', 'LOC', 'GPS'], aliases=['Rede:', 'Identificação:', 'Alimentador:', 'Localização:', 'Coordenadas:'])
        ).add_to(mapa)

    if features_trafos:
        fg_trafos = folium.FeatureGroup(name="Transformadores")
        for row in features_trafos:
            lat, lon = row['COORDS'][0], row['COORDS'][1]
            html_popup = f"""
            <div style="font-family: sans-serif; font-size: 13px; min-width: 250px;">
                <table style="width:100%;">
                    <tr><td><b>Rede:</b></td><td>{html.escape(str(row['TIPO_REDE']))}</td></tr>
                    <tr><td><b>Identificação:</b></td><td>{html.escape(str(row['NOME']))}</td></tr>
                    <tr><td><b>Alimentador:</b></td><td>{html.escape(str(row['ALIMENTADOR']))}</td></tr>
                    <tr><td><b>Localização:</b></td><td>{html.escape(str(row['MUNICIPIO']))} - {html.escape(str(row['REGIONAL']))}</td></tr>
                    <tr><td><b>Coordenadas:</b></td><td>{lat:.5f}, {lon:.5f}</td></tr>
                </table>
            </div>
            """
            folium.RegularPolygonMarker(
                location=[lat, lon],
                number_of_sides=3,
                radius=8,
                color='#b8860b', 
                fillColor='yellow',
                fillOpacity=0.9,
                weight=1,
                tooltip=f"{html.escape(str(row['TIPO_REDE']))}: {html.escape(str(row['NOME']))}",
                popup=folium.Popup(html_popup, max_width=300)
            ).add_to(fg_trafos)
        fg_trafos.add_to(mapa)

    if features_outros:
        folium.GeoJson(
            {"type": "FeatureCollection", "features": features_outros},
            name="Outros Equipamentos",
            style_function=lambda x: {'color': x['properties']['COR'], 'fillColor': x['properties']['COR'], 'radius': 6, 'weight': 2, 'fillOpacity': 1.0},
            marker=folium.CircleMarker(radius=6, fill=True, fillOpacity=1.0),
            tooltip=folium.features.GeoJsonTooltip(fields=['TIPO_REDE', 'NOME'], aliases=['Rede:', 'Identificação:']),
            popup=folium.features.GeoJsonPopup(fields=['TIPO_REDE', 'NOME', 'ALIMENTADOR', 'LOC', 'GPS'], aliases=['Rede:', 'Identificação:', 'Alimentador:', 'Localização:', 'Coordenadas:'])
        ).add_to(mapa)

    fg_busca = folium.FeatureGroup(name="Resultado da Pesquisa", show=True)
    for _, row in df_busca.iterrows():
        coord_txt = f"{row['COORDS'][0]:.5f}, {row['COORDS'][1]:.5f}" if row['TIPO_GEOMETRIA'] == 'Ponto' else "Linha de Múltiplos Pontos"
        sv_url = f"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={row['COORDS'][0]},{row['COORDS'][1]}"
        html_popup = f"""
        <div style="min-width: 250px; font-family: sans-serif;">
            <h4 style="margin-top: 0; color: #FF00FF; border-bottom: 2px solid #FF00FF; padding-bottom: 5px;">{html.escape(str(row['TIPO_REDE']))}</h4>
            <table style="width:100%;">
                <tr><td style="color: #555; padding: 2px;"><b>IDENTIFICAÇÃO:</b></td><td>{html.escape(str(row['NOME']))}</td></tr>
                <tr><td style="color: #555; padding: 2px;"><b>LOCAL:</b></td><td>{html.escape(str(row['MUNICIPIO']))}</td></tr>
                <tr><td colspan='2' style='padding-top:10px;'><a href="{sv_url}" target="_blank" style="color: #0066cc; font-weight: bold; text-decoration: none;">👁️ Abrir Street View</a></td></tr>
            </table>
        </div>
        """
        popup = folium.Popup(html_popup, max_width=350)
        if row['TIPO_GEOMETRIA'] == 'Linha':
            folium.PolyLine(locations=row['COORDS'], color='#FF00FF', weight=8, opacity=1.0, popup=popup, tooltip=f"ALVO ENCONTRADO: {html.escape(str(row['NOME']))}").add_to(fg_busca)
            for pt in row['COORDS']: busca_lats.append(pt[0]); busca_lons.append(pt[1])
        else:
            folium.Marker(location=row['COORDS'], icon=folium.Icon(color='purple', icon='star'), popup=popup, tooltip=f"ALVO ENCONTRADO: {html.escape(str(row['NOME']))}").add_to(fg_busca)
            busca_lats.append(row['COORDS'][0]); busca_lons.append(row['COORDS'][1])
    if busca_lat is not None and busca_lon is not None: folium.Marker(location=[busca_lat, busca_lon], icon=folium.Icon(color='orange', icon='map-pin', prefix='fa'), tooltip="Sua Pesquisa GPS").add_to(fg_busca)
    fg_busca.add_to(mapa)

else:
    tree_grid = None

def calcular_rede_proxima(lat, lon):
    if not tree_grid: return "<span style='color:gray'>Ative um alimentador no filtro para calcular</span>"
    xyz = latlon_to_xyz(lat, lon)
    _, idx = tree_grid.query(xyz)
    tipo, nome, g_lat, g_lon = grid_info[idx]
    dist_m = haversine(lat, lon, g_lat, g_lon) * 1000
    return f"<b>{html.escape(str(tipo))}</b> {html.escape(str(nome))} ({dist_m:.1f}m)"


def verificar_areas_popup_otimizado(lat, lon):
    """Só executa o cruzamento pesado com KML quando alguma camada especial está ativa."""
    if not alguma_area_especial_ativa:
        return "<span style='color:gray'>Ative uma camada de Área Especial para analisar</span>"
    return verificar_areas_da_obra(lat, lon)

# ==========================================
# RENDERIZAÇÃO DAS ÁREAS ESPECIAIS (COM POPUPS)
# ==========================================
def adicionar_camada_area(geo_data, nome_camada, mapa_obj, cor, is_ponto=False):
    if geo_data:
        estilo = lambda x: {'fillColor': cor, 'color': cor, 'weight': 2, 'fillOpacity': 0.4}
        marcador = folium.CircleMarker(radius=6, fill=True, fillOpacity=1, color=cor) if is_ponto else None
        
        folium.GeoJson(
            geo_data, 
            name=nome_camada, 
            style_function=estilo if not is_ponto else None,
            marker=marcador,
            tooltip=folium.features.GeoJsonTooltip(fields=['NOME'], aliases=['Área Específica:']),
            popup=folium.features.GeoJsonPopup(fields=['NOME'], aliases=['Nome do Local:'], style="font-family: sans-serif; font-size: 14px; min-width: 200px;")
        ).add_to(mapa_obj)

if mostrar_quilombos: adicionar_camada_area(geo_q, "Áreas Quilombolas", mapa, "#ff7f00")
if mostrar_indigenas: adicionar_camada_area(geo_i, "Terras Indígenas", mapa, "#2ca02c")
if mostrar_arqueologia: adicionar_camada_area(geo_a, "Sítios Arqueológicos", mapa, "#8c564b", is_ponto=True)
if mostrar_uc_federal: adicionar_camada_area(geo_uc_fed, "UC Federal", mapa, "#e6b800")
if mostrar_uc_estadual: adicionar_camada_area(geo_uc_est, "UC Estadual", mapa, "#ffff00")
if mostrar_uc_municipal: adicionar_camada_area(geo_uc_mun, "UC Municipal", mapa, "#ffff00")


# ==========================================
# CAMADAS DE OBRAS E CRUZAMENTOS COM ÁREAS
# ==========================================
dados_tabela_conflito = []

if (mostrar_concluidas or mostrar_conflitantes or mostrar_todas_obras or bool(status_list_sel)) and msg_obras == "OK":
    
    if mostrar_todas_obras or bool(status_list_sel):
        nome_cluster_obras = "Todas as Obras (Geral)" if mostrar_todas_obras else "Obras por STATUS LIST"
        cluster_todas = MarkerCluster(name=nome_cluster_obras)
        if mostrar_todas_obras and df_concluidas is not None:
            for _, row in df_concluidas.iterrows():
                lat, lon = row['LAT_CLEAN'], row['LON_CLEAN']
                sv_url = f"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={lat},{lon}"
                areas_especiais = verificar_areas_popup_otimizado(lat, lon) 
                rede_prox = calcular_rede_proxima(lat, lon)
                
                html_popup = f"""<div style="min-width: 250px; font-family: sans-serif;"><h4 style="margin-top: 0; color: #1f77b4; border-bottom: 2px solid #1f77b4; padding-bottom: 5px;">✅ OBRA CONCLUÍDA</h4><table style="width:100%;"><tr><td style="color: #555; padding: 2px;"><b>PROTOCOLO:</b></td><td>{html.escape(str(row.get('PROTOCOLO', 'S/N')))}</td></tr><tr><td style="color: #555; padding: 2px;"><b>NOME:</b></td><td>{html.escape(str(row.get('NOME', 'S/N')))}</td></tr><tr><td style="color: #555; padding: 2px;"><b>REDE ELÉTRICA:</b></td><td>{rede_prox}</td></tr><tr><td style="color: #555; padding: 2px;"><b>ÁREAS:</b></td><td>{areas_especiais}</td></tr><tr><td colspan='2' style='padding-top:10px;'><a href="{sv_url}" target="_blank" style="color: #0066cc; font-weight: bold; text-decoration: none;">👁️ Abrir Street View</a></td></tr></table></div>"""
                
                folium.CircleMarker(location=[lat, lon], radius=5, color='black', weight=1, fill=True, fillColor='#1f77b4', fillOpacity=0.9, tooltip=f"Concluída: {html.escape(str(row.get('PROTOCOLO', 'S/N')))}", popup=folium.Popup(html_popup, max_width=350)).add_to(cluster_todas)
        
        if df_andamento is not None:
            # Nesta camada genérica, a aparência é definida SOMENTE pelo STATUS LIST.
            # A sinalização vermelha de conflito pertence exclusivamente à opção
            # 'OBRAS CONFLITANTES (Raio 50m)' abaixo. Assim, marcar um STATUS LIST
            # não ativa visualmente conflitos por conta própria.
            cores_status_list = {
                '0': '#cbd5e1',
                'EM LEVANTAMENTO': '#22c55e',
                'ANALISE DE LEVANTAMENTO': '#eab308',
                'IMPRODUTIVO': '#f97316',
                'CORRECAO DE LEVANTAMENTO': '#7c3aed',
            }
            for _, row in df_andamento.iterrows():
                lat, lon = row['LAT_CLEAN'], row['LON_CLEAN']
                status_list_atual = str(row.get('STATUS_LIST_NORM', '')).strip() or 'SEM STATUS'
                status_norm_cor = remove_accents(status_list_atual).upper().strip()
                cor = cores_status_list.get(status_norm_cor, '#2ca02c')
                titulo = f"🚧 {status_list_atual}"
                sv_url = f"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={lat},{lon}"
                areas_especiais = verificar_areas_popup_otimizado(lat, lon) 
                rede_prox = calcular_rede_proxima(lat, lon)
                
                html_popup = f"""<div style="min-width: 250px; font-family: sans-serif;"><h4 style="margin-top: 0; color: {cor}; border-bottom: 2px solid {cor}; padding-bottom: 5px;">{titulo}</h4><table style="width:100%;"><tr><td style="color: #555; padding: 2px;"><b>PROTOCOLO:</b></td><td>{html.escape(str(row.get('PROTOCOLO', 'S/N')))}</td></tr><tr><td style="color: #555; padding: 2px;"><b>NOME:</b></td><td>{html.escape(str(row.get('NOME', 'S/N')))}</td></tr><tr><td style="color: #555; padding: 2px;"><b>REDE ELÉTRICA:</b></td><td>{rede_prox}</td></tr><tr><td style="color: #555; padding: 2px;"><b>ÁREAS:</b></td><td>{areas_especiais}</td></tr><tr><td colspan='2' style='padding-top:10px;'><a href="{sv_url}" target="_blank" style="color: #0066cc; font-weight: bold; text-decoration: none;">👁️ Abrir Street View</a></td></tr></table></div>"""
                
                folium.CircleMarker(location=[lat, lon], radius=5, color='black', weight=1, fill=True, fillColor=cor, fillOpacity=0.9, tooltip=f"{titulo}: {html.escape(str(row.get('PROTOCOLO', 'S/N')))}", popup=folium.Popup(html_popup, max_width=350)).add_to(cluster_todas)
        cluster_todas.add_to(mapa)

    if mostrar_concluidas and df_concluidas is not None:
        fg_concluidas = MarkerCluster(name="Obras Concluídas", show=True)
        for _, row in df_concluidas.iterrows():
            protocolo = str(row.get('PROTOCOLO', 'S/N'))
            lat, lon = row['LAT_CLEAN'], row['LON_CLEAN']
            cor_concluida = '#1f77b4'
            sv_url = f"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={lat},{lon}"
            areas_especiais = verificar_areas_popup_otimizado(lat, lon)
            rede_prox = calcular_rede_proxima(lat, lon)
            
            html_popup = f"""<div style="min-width: 250px; font-family: sans-serif;"><h4 style="margin-top: 0; color: {cor_concluida}; border-bottom: 2px solid {cor_concluida}; padding-bottom: 5px;">✅ OBRA CONCLUÍDA</h4><table style="width:100%;"><tr><td style="color: #555; padding: 2px;"><b>PROTOCOLO:</b></td><td>{html.escape(protocolo)}</td></tr><tr><td style="color: #555; padding: 2px;"><b>NOME:</b></td><td>{html.escape(str(row.get('NOME', 'S/N')))}</td></tr><tr><td style="color: #555; padding: 2px;"><b>REDE ELÉTRICA:</b></td><td>{rede_prox}</td></tr><tr><td style="color: #555; padding: 2px;"><b>ÁREAS:</b></td><td>{areas_especiais}</td></tr><tr><td colspan='2' style='padding-top:10px;'><a href="{sv_url}" target="_blank" style="color: #0066cc; font-weight: bold; text-decoration: none;">👁️ Abrir Street View</a></td></tr></table></div>"""
            
            folium.CircleMarker(
                location=[lat, lon], radius=6, color='black', weight=1, fill=True, 
                fillColor=cor_concluida, fillOpacity=1, popup=folium.Popup(html_popup, max_width=350),
                tooltip=f"Obra Concluída: {html.escape(protocolo)}"
            ).add_to(fg_concluidas)
            
        fg_concluidas.add_to(mapa)
            
    if mostrar_conflitantes and df_andamento is not None:
        fg_andamento = folium.FeatureGroup(name="Obras Conflitantes", show=True)
        for _, row in df_andamento.iterrows():
            if not row['CONFLITO']: continue
            lat, lon = row['LAT_CLEAN'], row['LON_CLEAN']
            protocolo = str(row.get('PROTOCOLO', 'S/N'))
            nome_nova = str(row.get('NOME', 'S/N'))
            nome_alvo = str(row.get('NOME_CONCLUIDA', 'S/N'))
            sv_url = f"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={lat},{lon}"
            areas_especiais = verificar_areas_popup_otimizado(lat, lon)
            rede_prox = calcular_rede_proxima(lat, lon)
            
            dist_conflito = float(row['DISTANCIA_CONFLITO'])
            if dist_conflito <= 0.5:
                severidade = '📌 MESMO PONTO'
            elif dist_conflito <= 10:
                severidade = '🔴 CRÍTICO'
            elif dist_conflito <= 25:
                severidade = '🟠 ALTO'
            else:
                severidade = '🟡 MÉDIO'

            dados_tabela_conflito.append({
                "Severidade": severidade,
                "Protocolo (Nova)": protocolo,
                "Nome (Nova)": nome_nova,
                "STATUS LIST (Nova)": str(row.get('STATUS_LIST_NORM', '')).strip(),
                "Município": str(row.get('MUNICIPIO_NORM', '')).strip(),
                "Regional": str(row.get('REGIONAL_NORM', '')).strip(),
                "Tipo Nota": str(row.get('TIPO NOTA', '')).strip(),
                "Conflito (Concluída)": row['PROTOCOLO_CONFLITO'],
                "Nome (Concluída)": nome_alvo,
                "STATUS LIST (Concluída)": 'CONCLUÍDO',
                "Concluídas no raio": int(row.get('QTD_CONCLUIDAS_50M', 0) or 0),
                "Protocolos concluídos no raio": str(row.get('PROTOCOLOS_CONCLUIDOS_50M', '')).strip(),
                "Distância (m)": round(dist_conflito, 2),
                "Latitude": lat,
                "Longitude": lon,
                "Google Maps": f"https://www.google.com/maps/search/?api=1&query={lat},{lon}"
            })
                
            html_popup = f"""
            <div style="min-width:300px;font-family:sans-serif;line-height:1.45;">
              <h4 style="margin:0 0 8px 0;color:#dc2626;border-bottom:2px solid #dc2626;padding-bottom:5px;">🚨 CONFLITO DETECTADO</h4>
              <div style="font-weight:700;color:#334155;margin-top:4px;">🏗️ OBRA NOVA</div>
              <b>Protocolo:</b> {html.escape(protocolo)}<br>
              <b>Nome:</b> {html.escape(nome_nova)}<br>
              <b>Status:</b> {html.escape(str(row.get('STATUS_LIST_NORM','')))}<br>
              <b>Município / Regional:</b> {html.escape(str(row.get('MUNICIPIO_NORM','')))} / {html.escape(str(row.get('REGIONAL_NORM','')))}<br>
              <div style="font-weight:700;color:#dc2626;margin-top:8px;">🔵 CONCLUÍDA DE REFERÊNCIA</div>
              <b>Protocolo:</b> {html.escape(str(row['PROTOCOLO_CONFLITO']))}<br>
              <b>Nome:</b> {html.escape(nome_alvo)}<br>
              <b>Distância:</b> {dist_conflito:.2f} m — {html.escape(severidade)}<br>
              <b>Concluídas no raio:</b> {int(row.get('QTD_CONCLUIDAS_50M', 0) or 0)}<br>
              <div style="font-weight:700;color:#334155;margin-top:8px;">🧭 CONTEXTO</div>
              <b>Rede elétrica:</b> {rede_prox}<br>
              <b>Restrições:</b> {areas_especiais}<br>
              <div style="padding-top:9px;"><a href="{sv_url}" target="_blank" style="color:#0066cc;font-weight:bold;text-decoration:none;">👁️ Abrir Street View</a></div>
            </div>
            """

            # O raio de 50 m pertence à OBRA CONCLUÍDA (azul), pois é ela
            # que define a área onde uma nova solicitação gera conflito.
            lat_conc = row.get('LAT_CONCLUIDA_CONFLITO')
            lon_conc = row.get('LON_CONCLUIDA_CONFLITO')
            if pd.notna(lat_conc) and pd.notna(lon_conc):
                lat_conc, lon_conc = float(lat_conc), float(lon_conc)
                protocolo_conc = str(row.get('PROTOCOLO_CONFLITO', 'S/N'))

                # Círculo não interativo: não bloqueia o clique no marcador azul.
                folium.Circle(
                    location=[lat_conc, lon_conc],
                    radius=RAIO_CONFLITO_M,
                    color='#ff0000',
                    weight=4,
                    opacity=1.0,
                    fill=True,
                    fill_color='#ff0000',
                    fill_opacity=0.10,
                    interactive=False
                ).add_to(fg_andamento)

                popup_concluida = f"""
                <div style='min-width:250px;font-family:sans-serif;'>
                    <h4 style='margin-top:0;color:#1f77b4;border-bottom:2px solid #1f77b4;padding-bottom:5px;'>✅ OBRA CONCLUÍDA</h4>
                    <b>PROTOCOLO:</b> {html.escape(protocolo_conc)}<br>
                    <b>NOME:</b> {html.escape(nome_alvo)}<br>
                    <b>STATUS LIST:</b> CONCLUÍDO<br>
                    <b>COORDENADAS:</b> {lat_conc:.6f}, {lon_conc:.6f}<br>
                    <b>RAIO DE CONFLITO:</b> 50 m
                </div>
                """
                folium.CircleMarker(
                    location=[lat_conc, lon_conc], radius=7, color='black', weight=1,
                    fill=True, fill_color='#1f77b4', fill_opacity=1.0,
                    tooltip=f"Concluída: {html.escape(protocolo_conc)}",
                    popup=folium.Popup(popup_concluida, max_width=320)
                ).add_to(fg_andamento)

            # A solicitação em conflito permanece vermelha e clicável.
            folium.CircleMarker(
                location=[lat, lon], radius=6, color='black', weight=1,
                fill=True, fill_color='red', fill_opacity=0.9,
                tooltip=f"Conflito: {html.escape(protocolo)}",
                popup=folium.Popup(html_popup, max_width=350)
            ).add_to(fg_andamento)
        fg_andamento.add_to(mapa)

# ==========================================
# FOCO AUTOMÁTICO RECEBIDO DO CRIAR SGO
# ==========================================
coords_foco_sgo = []
if foco_mapa_sgo and foco_mapa_sgo.get('conflitos'):
    fg_foco = folium.FeatureGroup(name="🚨 Conflito vindo do CRIAR SGO", show=True)
    conflitos_foco = foco_mapa_sgo.get('conflitos', [])

    # Evita repetir a mesma solicitação quando ela possui mais de uma concluída próxima.
    novas_desenhadas = set()
    concluidas_desenhadas = set()

    for item in conflitos_foco:
        obra_nova = str(item.get('obra_nova', ''))
        lat_nova = float(item.get('lat_nova'))
        lon_nova = float(item.get('lon_nova'))
        obra_conc = str(item.get('obra_concluida', ''))
        lat_conc = float(item.get('lat_concluida'))
        lon_conc = float(item.get('lon_concluida'))
        dist_m = float(item.get('distancia_m', 0.0))
        nome_nova = html.escape(str(item.get('nome_nova', '')))
        nome_conc = html.escape(str(item.get('nome_concluida', '')))
        mun_nova = html.escape(str(item.get('municipio_nova', '')))
        mun_conc = html.escape(str(item.get('municipio_concluida', '')))
        status_conc = html.escape(str(item.get('status_list_concluida', 'CONCLUIDO')))

        chave_nova = (obra_nova, round(lat_nova, 7), round(lon_nova, 7))
        if chave_nova not in novas_desenhadas:
            novas_desenhadas.add(chave_nova)
            coords_foco_sgo.append([lat_nova, lon_nova])
            popup_nova = f"""
            <div style='min-width:260px;font-family:sans-serif;'>
                <h4 style='margin:0 0 8px;color:#f97316;'>🟠 SOLICITAÇÃO CONSULTADA</h4>
                <b>Obra:</b> {html.escape(obra_nova)}<br>
                <b>Nome:</b> {nome_nova or '-'}<br>
                <b>Município:</b> {mun_nova or '-'}<br>
                <b>Coordenadas:</b> {lat_nova:.6f}, {lon_nova:.6f}
            </div>
            """
            folium.Marker(
                [lat_nova, lon_nova],
                icon=folium.Icon(color='orange', icon='map-pin', prefix='fa'),
                tooltip=f"Solicitação {obra_nova}",
                popup=folium.Popup(popup_nova, max_width=320)
            ).add_to(fg_foco)
        chave_conc = (obra_conc, round(lat_conc, 7), round(lon_conc, 7))
        if chave_conc not in concluidas_desenhadas:
            concluidas_desenhadas.add(chave_conc)
            coords_foco_sgo.append([lat_conc, lon_conc])
            popup_conc = f"""
            <div style='min-width:260px;font-family:sans-serif;'>
                <h4 style='margin:0 0 8px;color:#2563eb;'>🔵 OBRA CONCLUÍDA PRÓXIMA</h4>
                <b>Obra:</b> {html.escape(obra_conc)}<br>
                <b>Status LIST:</b> {status_conc}<br>
                <b>Nome:</b> {nome_conc or '-'}<br>
                <b>Município:</b> {mun_conc or '-'}<br>
                <b>Coordenadas:</b> {lat_conc:.6f}, {lon_conc:.6f}<br>
                <b>Raio de conflito:</b> 50 m
            </div>
            """

            # O círculo é centrado na obra concluída e não captura cliques.
            folium.Circle(
                location=[lat_conc, lon_conc],
                radius=RAIO_CONFLITO_M,
                color='#ff0000',
                weight=4,
                opacity=1.0,
                fill=True,
                fill_color='#ff0000',
                fill_opacity=0.10,
                interactive=False
            ).add_to(fg_foco)

            # Marcador azul adicionado depois do círculo para permanecer acima e clicável.
            folium.Marker(
                [lat_conc, lon_conc],
                icon=folium.Icon(color='blue', icon='check', prefix='fa'),
                tooltip=f"Concluída {obra_conc}",
                popup=folium.Popup(popup_conc, max_width=320),
                z_index_offset=1000
            ).add_to(fg_foco)

        folium.PolyLine(
            [[lat_nova, lon_nova], [lat_conc, lon_conc]],
            color='#dc2626', weight=3, opacity=0.9, dash_array='6,5',
            tooltip=f"{obra_nova} ↔ {obra_conc}: {dist_m:.1f} m"
        ).add_to(fg_foco)

        # Distância no meio da ligação.
        lat_meio = (lat_nova + lat_conc) / 2.0
        lon_meio = (lon_nova + lon_conc) / 2.0
        folium.Marker(
            [lat_meio, lon_meio],
            icon=folium.DivIcon(html=f"<div style='background:white;border:1px solid #dc2626;border-radius:5px;padding:2px 5px;color:#991b1b;font-weight:bold;font-size:11px;white-space:nowrap;'>{dist_m:.1f} m</div>")
        ).add_to(fg_foco)

    fg_foco.add_to(mapa)

folium.LayerControl(position='topright').add_to(mapa)

# Helper global para cards KPI.
# Deve ficar fora dos blocos condicionais para ser reutilizado em Conflitos e Qualidade.
def render_kpi_card(col, css_class, title, value, subtitle):
    with col:
        st.markdown(
            f"<div class='nip-kpi-card {css_class}'><div class='nip-kpi-title'>{title}</div><div class='nip-kpi-value'>{value}</div><div class='nip-kpi-sub'>{subtitle}</div></div>",
            unsafe_allow_html=True
        )

# -------------------------------------------------------------
# 5. TABELA INTELIGENTE E BOTÃO DE EXPORTAÇÃO
# -------------------------------------------------------------
zoom_lat, zoom_lon = None, None

with table_container:
    if mostrar_conflitantes and msg_obras == "OK" and len(dados_tabela_conflito) > 0:
        st.markdown("---")
        st.markdown("### 🚨 Conflitos Geográficos — Obras em até 50 m de Concluídas")
        st.caption("Priorize os conflitos mais próximos. Selecione uma linha da tabela para centralizar a obra correspondente no mapa.")

        df_tabela = pd.DataFrame(dados_tabela_conflito)
        df_tabela['Distância (m)'] = pd.to_numeric(df_tabela['Distância (m)'], errors='coerce').fillna(0.0)
        df_tabela = df_tabela.sort_values(['Distância (m)', 'Protocolo (Nova)'], ascending=[True, True]).reset_index(drop=True)

        total_conflitos = int(len(df_tabela))
        qtd_criticos = int((df_tabela['Severidade'] == '🔴 CRÍTICO').sum())
        qtd_concluidas_distintas = int(df_tabela['Conflito (Concluída)'].astype(str).nunique())
        menor_dist = float(df_tabela['Distância (m)'].min()) if total_conflitos else 0.0

        st.markdown("""
        <style>
        .nip-kpi-card {
            border-radius: 18px;
            padding: 18px 18px 16px 18px;
            color: #ffffff;
            box-shadow: 0 10px 24px rgba(0,0,0,0.14);
            border: 1px solid rgba(255,255,255,0.12);
            min-height: 118px;
            margin-bottom: 8px;
        }
        .nip-kpi-title {
            font-size: 0.88rem;
            font-weight: 700;
            opacity: 0.95;
            margin-bottom: 10px;
            letter-spacing: 0.2px;
        }
        .nip-kpi-value {
            font-size: 2.15rem;
            font-weight: 800;
            line-height: 1.0;
            margin-bottom: 8px;
        }
        .nip-kpi-sub {
            font-size: 0.82rem;
            opacity: 0.92;
            line-height: 1.25;
        }
        .nip-kpi-red {background: linear-gradient(135deg, #ff4d4f 0%, #c81e1e 100%);}
        .nip-kpi-orange {background: linear-gradient(135deg, #ff9f43 0%, #ff6b00 100%);}
        .nip-kpi-blue {background: linear-gradient(135deg, #4096ff 0%, #1d4ed8 100%);}
        .nip-kpi-purple {background: linear-gradient(135deg, #8b5cf6 0%, #5b21b6 100%);}
        .nip-kpi-amber {background: linear-gradient(135deg, #f59e0b 0%, #b45309 100%);}
        .nip-kpi-green {background: linear-gradient(135deg, #22c55e 0%, #15803d 100%);}
        .nip-kpi-cyan {background: linear-gradient(135deg, #06b6d4 0%, #0f766e 100%);}
        .nip-chip-wrap {display:flex; flex-wrap:wrap; gap:10px; margin: 8px 0 14px 0;}
        .nip-chip {
            display:inline-flex; align-items:center; gap:8px;
            padding:10px 14px; border-radius:999px; font-size:0.88rem; font-weight:700;
            color:#132238; background:#eef2ff; border:1px solid #d9e2ff;
        }
        .nip-chip strong {font-size:0.95rem;}
        </style>
        """, unsafe_allow_html=True)

        k1, k2, k3, k4 = st.columns(4)
        render_kpi_card(k1, 'nip-kpi-red', '🚨 Conflitos encontrados', f"{total_conflitos:,}".replace(',', '.'), 'Obras novas dentro de 50 m de uma concluída.')
        render_kpi_card(k2, 'nip-kpi-orange', '⚠️ Críticos (≤ 10 m)', f"{qtd_criticos:,}".replace(',', '.'), 'Casos mais urgentes para verificação imediata.')
        render_kpi_card(k3, 'nip-kpi-blue', '🔵 Concluídas envolvidas', f"{qtd_concluidas_distintas:,}".replace(',', '.'), 'Quantidade de obras concluídas impactando novos registros.')
        render_kpi_card(k4, 'nip-kpi-purple', '📏 Menor distância', f"{menor_dist:.2f} m", 'Menor separação geográfica encontrada entre as obras.')

        f1, f2, f3 = st.columns([1, 1, 1])
        with f1:
            op_reg = sorted([x for x in df_tabela['Regional'].dropna().astype(str).unique() if x and x != 'DESCONHECIDO'])
            filtro_reg_confl = st.multiselect("Regional", op_reg, key="filtro_reg_conflitos")
        with f2:
            base_mun = df_tabela[df_tabela['Regional'].isin(filtro_reg_confl)] if filtro_reg_confl else df_tabela
            op_mun = sorted([x for x in base_mun['Município'].dropna().astype(str).unique() if x and x != 'DESCONHECIDO'])
            filtro_mun_confl = st.multiselect("Município", op_mun, key="filtro_mun_conflitos")
        with f3:
            filtro_sev = st.multiselect(
                "Severidade",
                ['📌 MESMO PONTO', '🔴 CRÍTICO', '🟠 ALTO', '🟡 MÉDIO'],
                key="filtro_sev_conflitos"
            )

        df_conf_view = df_tabela.copy()
        if filtro_reg_confl:
            df_conf_view = df_conf_view[df_conf_view['Regional'].isin(filtro_reg_confl)]
        if filtro_mun_confl:
            df_conf_view = df_conf_view[df_conf_view['Município'].isin(filtro_mun_confl)]
        if filtro_sev:
            df_conf_view = df_conf_view[df_conf_view['Severidade'].isin(filtro_sev)]

        st.caption(f"Exibindo {len(df_conf_view)} de {total_conflitos} conflito(s).")

        cols_conf = [
            'Severidade', 'Protocolo (Nova)', 'Nome (Nova)', 'STATUS LIST (Nova)',
            'Município', 'Regional', 'Tipo Nota', 'Conflito (Concluída)',
            'Nome (Concluída)', 'STATUS LIST (Concluída)', 'Concluídas no raio',
            'Protocolos concluídos no raio', 'Distância (m)', 'Latitude', 'Longitude', 'Google Maps'
        ]
        df_conf_view = df_conf_view[[c for c in cols_conf if c in df_conf_view.columns]].copy()

        try:
            event = st.dataframe(
                df_conf_view,
                use_container_width=True,
                height=min(520, 80 + max(1, len(df_conf_view)) * 35),
                on_select="rerun",
                selection_mode="single_row",
                hide_index=True,
                column_config={
                    "Distância (m)": st.column_config.NumberColumn("Distância (m)", format="%.2f m"),
                    "Latitude": st.column_config.NumberColumn("Latitude", format="%.6f"),
                    "Longitude": st.column_config.NumberColumn("Longitude", format="%.6f"),
                    "Google Maps": st.column_config.LinkColumn("📍 Google Maps", display_text="Abrir Maps"),
                }
            )
            if hasattr(event, 'selection') and event.selection.rows:
                idx = event.selection.rows[0]
                zoom_lat = float(df_conf_view.iloc[idx]['Latitude'])
                zoom_lon = float(df_conf_view.iloc[idx]['Longitude'])
        except Exception:
            st.dataframe(df_conf_view, use_container_width=True, hide_index=True)

        cdl1, cdl2 = st.columns([1, 1])
        with cdl1:
            csv = df_conf_view.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 Baixar conflitos filtrados (CSV)",
                data=csv,
                file_name="conflitos_geograficos_filtrados.csv",
                mime="text/csv",
                use_container_width=True
            )
        with cdl2:
            try:
                buffer_xlsx = io.BytesIO()
                with pd.ExcelWriter(buffer_xlsx, engine='openpyxl') as writer:
                    df_conf_view.to_excel(writer, sheet_name='Conflitos', index=False)
                    resumo_conf = pd.DataFrame({
                        'Indicador': ['Conflitos', 'Críticos até 10 m', 'Concluídas distintas', 'Menor distância (m)'],
                        'Valor': [total_conflitos, qtd_criticos, qtd_concluidas_distintas, round(menor_dist, 2)]
                    })
                    resumo_conf.to_excel(writer, sheet_name='Resumo', index=False)
                st.download_button(
                    label="📊 Baixar conflitos (Excel)",
                    data=buffer_xlsx.getvalue(),
                    file_name="conflitos_geograficos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as exc:
                st.caption(f"Excel indisponível: {exc}")

        with st.expander("🔗 Consolidado por obra concluída", expanded=False):
            consolidado = (
                df_tabela.groupby(['Conflito (Concluída)', 'Nome (Concluída)'], dropna=False)
                .agg(
                    NOVAS_EM_CONFLITO=('Protocolo (Nova)', 'nunique'),
                    MENOR_DISTANCIA_M=('Distância (m)', 'min'),
                    MAIOR_DISTANCIA_M=('Distância (m)', 'max'),
                    MUNICIPIOS=('Município', lambda x: ' | '.join(sorted(set(str(v) for v in x if str(v).strip()))))
                )
                .reset_index()
                .sort_values(['NOVAS_EM_CONFLITO','MENOR_DISTANCIA_M'], ascending=[False, True])
            )
            st.caption("Mostra quais obras concluídas concentram mais novas solicitações dentro do raio operacional.")
            st.dataframe(consolidado, use_container_width=True, hide_index=True)
            st.download_button(
                "📥 Baixar consolidado por concluída (Excel)",
                data=dataframe_para_excel_bytes(consolidado, 'Consolidado'),
                file_name="conflitos_por_obra_concluida.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="download_consolidado_concluida"
            )

    if (mostrar_concluidas or mostrar_conflitantes or mostrar_todas_obras or bool(status_list_sel)) and msg_obras == "OK":
        if df_invalidas is not None and not df_invalidas.empty:
            st.markdown("---")

            total_base = int(df_invalidas['_TOTAL_LINHAS_BASE'].iloc[0]) if '_TOTAL_LINHAS_BASE' in df_invalidas.columns else len(df_invalidas)
            total_validas = int(df_invalidas['_TOTAL_VALIDAS_COORD'].iloc[0]) if '_TOTAL_VALIDAS_COORD' in df_invalidas.columns else max(0, total_base - len(df_invalidas))
            perc_validas = (total_validas / total_base * 100.0) if total_base else 0.0

            with st.expander(f"⚠️ Monitor de Qualidade de Dados — {len(df_invalidas):,} inconsistência(s)".replace(',', '.'), expanded=False):
                st.markdown("#### Qualidade das coordenadas da base")
                st.caption("Registros desta seção foram ignorados no mapa porque a coordenada não permite posicionamento geográfico confiável.")

                q1, q2, q3 = st.columns(3)
                render_kpi_card(q1, 'nip-kpi-amber', '⚠️ Inconsistências', f"{len(df_invalidas):,}".replace(',', '.'), 'Registros ignorados no mapa por coordenada inválida ou não confiável.')
                render_kpi_card(q2, 'nip-kpi-green', '✅ Coordenadas válidas', f"{total_validas:,}".replace(',', '.'), 'Linhas aptas para posicionamento geográfico no mapa.')
                render_kpi_card(q3, 'nip-kpi-cyan', '🌎 Qualidade geográfica', f"{perc_validas:.1f}%", 'Percentual da base com latitude e longitude utilizáveis.')

                resumo_motivos = (
                    df_invalidas['MOTIVO DA INCONSISTÊNCIA']
                    .fillna('NÃO IDENTIFICADO')
                    .value_counts()
                    .rename_axis('Motivo')
                    .reset_index(name='Quantidade')
                )

                st.markdown("##### 📌 Resumo por motivo")
                chips = []
                for _, row in resumo_motivos.head(6).iterrows():
                    motivo = str(row['Motivo'])
                    qtd = int(row['Quantidade'])
                    chips.append(f"<div class='nip-chip'>📍 <span>{motivo}</span> <strong>{qtd:,}</strong></div>".replace(',', '.'))
                if chips:
                    st.markdown(f"<div class='nip-chip-wrap'>{''.join(chips)}</div>", unsafe_allow_html=True)
                st.dataframe(resumo_motivos, use_container_width=True, hide_index=True)

                fq1, fq2, fq3 = st.columns(3)
                with fq1:
                    op_mot = resumo_motivos['Motivo'].tolist()
                    filtro_motivo = st.multiselect("Motivo", op_mot, key="filtro_motivo_qd")
                with fq2:
                    mun_col_q = next((c for c in ['MUNICIPIO', 'MUNICÍPIO', 'MUNICIPIO_NORM'] if c in df_invalidas.columns), None)
                    op_mun_q = sorted(df_invalidas[mun_col_q].dropna().astype(str).unique().tolist()) if mun_col_q else []
                    filtro_mun_q = st.multiselect("Município", op_mun_q, key="filtro_mun_qd")
                with fq3:
                    status_col_q = next((c for c in df_invalidas.columns if 'STATUS SISCO' in str(c).upper()), None)
                    op_status_q = sorted(df_invalidas[status_col_q].dropna().astype(str).unique().tolist()) if status_col_q else []
                    filtro_status_q = st.multiselect("STATUS SISCO", op_status_q, key="filtro_status_qd")

                df_q = df_invalidas.copy()
                if filtro_motivo:
                    df_q = df_q[df_q['MOTIVO DA INCONSISTÊNCIA'].isin(filtro_motivo)]
                if filtro_mun_q and mun_col_q:
                    df_q = df_q[df_q[mun_col_q].astype(str).isin(filtro_mun_q)]
                if filtro_status_q and status_col_q:
                    df_q = df_q[df_q[status_col_q].astype(str).isin(filtro_status_q)]

                cols_preferidas = [
                    'PROTOCOLO', 'TIPO NOTA', 'MUNICIPIO', 'MUNICÍPIO', 'REGIONAL_NORM',
                    'LATITUDE', 'LONGITUDE', 'STATUS SISCO', 'STATUS LIST',
                    'MOTIVO DA INCONSISTÊNCIA', 'NÍVEL'
                ]
                cols_to_show = []
                for c in cols_preferidas:
                    if c in df_q.columns and c not in cols_to_show:
                        cols_to_show.append(c)

                st.caption(f"Exibindo {len(df_q)} de {len(df_invalidas)} inconsistência(s).")
                st.dataframe(df_q[cols_to_show], use_container_width=True, height=420, hide_index=True)

                qd1, qd2 = st.columns(2)
                with qd1:
                    csv_q = df_q[cols_to_show].to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        "📥 Baixar inconsistências filtradas (CSV)",
                        data=csv_q,
                        file_name="inconsistencias_coordenadas.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                with qd2:
                    try:
                        buffer_q = io.BytesIO()
                        with pd.ExcelWriter(buffer_q, engine='openpyxl') as writer:
                            df_q[cols_to_show].to_excel(writer, sheet_name='Inconsistencias', index=False)
                            resumo_motivos.to_excel(writer, sheet_name='Resumo_por_Motivo', index=False)
                        st.download_button(
                            "📊 Baixar inconsistências (Excel)",
                            data=buffer_q.getvalue(),
                            file_name="monitor_qualidade_dados.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    except Exception as exc:
                        st.caption(f"Excel indisponível: {exc}")

    # Diagnósticos adicionais de qualidade, sem remover registros nem alterar o mapa.
    if msg_obras == "OK" and (mostrar_concluidas or mostrar_conflitantes or mostrar_todas_obras or bool(status_list_sel) or bool(protocolo_pesquisa)):
        dup_coords, dup_prot, qualidade_mun = analisar_qualidade_base(_mtime_seguro("BASE_LEVANTAMENTO_ATUALIZADA.xlsx"))
        if not dup_coords.empty or not dup_prot.empty or not qualidade_mun.empty:
            st.markdown("---")
            with st.expander("🔎 Diagnóstico Preventivo — duplicidades e qualidade por município", expanded=False):
                dqa, dqb = st.columns(2)
                dqa.metric("Linhas com coordenada repetida", len(dup_coords))
                dqb.metric("Linhas com protocolo repetido", len(dup_prot))
                if not qualidade_mun.empty:
                    st.markdown("##### 🌎 Qualidade geográfica por município")
                    qm = qualidade_mun.sort_values(['QUALIDADE_%','TOTAL'], ascending=[True, False]).head(25).copy()
                    fig_qm = px.bar(
                        qm.sort_values('QUALIDADE_%', ascending=True),
                        x='QUALIDADE_%', y='MUNICIPIO', orientation='h',
                        text='QUALIDADE_%',
                        title='Municípios com menor percentual de coordenadas válidas'
                    )
                    fig_qm.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                    fig_qm.update_layout(xaxis_title='% de coordenadas válidas', yaxis_title='', height=max(380, 24*len(qm)+100), showlegend=False)
                    st.plotly_chart(fig_qm, use_container_width=True)
                if not dup_coords.empty:
                    st.markdown("##### 📌 Coordenadas repetidas")
                    cols_dup = [c for c in ['PROTOCOLO','NOME','MUNICIPIO','MUNICÍPIO','STATUS LIST','_LAT_Q','_LON_Q'] if c in dup_coords.columns]
                    st.dataframe(dup_coords[cols_dup].head(500), use_container_width=True, hide_index=True)
                if not dup_prot.empty:
                    st.markdown("##### 🔁 Protocolos repetidos")
                    cols_prot = [c for c in ['PROTOCOLO','NOME','MUNICIPIO','MUNICÍPIO','STATUS LIST','_LAT_Q','_LON_Q'] if c in dup_prot.columns]
                    st.dataframe(dup_prot[cols_prot].head(500), use_container_width=True, hide_index=True)
                pacote_diag = io.BytesIO()
                with pd.ExcelWriter(pacote_diag, engine='openpyxl') as writer:
                    if not dup_coords.empty: dup_coords.to_excel(writer, sheet_name='Coords_Duplicadas', index=False)
                    if not dup_prot.empty: dup_prot.to_excel(writer, sheet_name='Protocolos_Repetidos', index=False)
                    if not qualidade_mun.empty: qualidade_mun.to_excel(writer, sheet_name='Qualidade_Municipio', index=False)
                st.download_button(
                    "📊 Baixar diagnóstico preventivo (Excel)", data=pacote_diag.getvalue(),
                    file_name="diagnostico_preventivo_mapa.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, key="download_diag_preventivo"
                )

    # Exportação da visão de obras atualmente filtrada.
    # IMPORTANTE: análises geoespaciais de proximidade são calculadas SOMENTE
    # quando o usuário solicita a geração do arquivo. Isso evita travar o mapa
    # em todo rerun do Streamlit.
    if msg_obras == "OK" and ((df_concluidas is not None and not df_concluidas.empty) or (df_andamento is not None and not df_andamento.empty)):
        with st.expander("📥 Exportar visão atual de obras", expanded=False):
            partes_export = []
            if df_concluidas is not None and not df_concluidas.empty:
                dc = df_concluidas.copy(); dc['_CAMADA_EXPORT'] = 'CONCLUÍDA'; partes_export.append(dc)
            if df_andamento is not None and not df_andamento.empty:
                da = df_andamento.copy(); da['_CAMADA_EXPORT'] = 'EM ANÁLISE'; partes_export.append(da)
            df_export_atual = pd.concat(partes_export, ignore_index=True, sort=False) if partes_export else pd.DataFrame()

            if not df_export_atual.empty:
                st.caption(f"{len(df_export_atual)} obra(s) na visão filtrada atual.")
                incluir_restricoes_export = st.checkbox(
                    "Incluir análise de proximidade com Áreas Especiais (mais demorado)",
                    value=False,
                    key="export_incluir_restricoes"
                )

                chave_export = (
                    len(df_export_atual),
                    tuple(sorted(regioes_sel)),
                    tuple(sorted(municipios_sel)),
                    tuple(sorted(status_list_sel)),
                    bool(incluir_restricoes_export),
                    _mtime_seguro("BASE_LEVANTAMENTO_ATUALIZADA.xlsx")
                )

                if st.button("⚙️ Gerar arquivo da visão filtrada", use_container_width=True, key="gerar_visao_filtrada"):
                    with st.spinner("Preparando arquivo filtrado..."):
                        df_exp = df_export_atual.copy()
                        if 'DISTANCIA_CONFLITO' in df_exp.columns:
                            df_exp['SEVERIDADE_CONFLITO'] = df_exp['DISTANCIA_CONFLITO'].apply(classificar_severidade_distancia)

                        if incluir_restricoes_export:
                            def _restr_export(r):
                                try:
                                    itens = analisar_proximidade_areas_especiais(float(r['LAT_CLEAN']), float(r['LON_CLEAN']), 500.0)
                                    return ' | '.join(
                                        f"{i['categoria']}: {i['faixa']} ({i['distancia_m_aprox']:.0f}m aprox.)"
                                        for i in itens
                                    )
                                except Exception:
                                    return ''
                            df_exp['RESTRICOES_PROXIMIDADE'] = df_exp.apply(_restr_export, axis=1)

                        st.session_state['visao_filtrada_excel_bytes'] = dataframe_para_excel_bytes(df_exp, 'Obras_Filtradas')
                        st.session_state['visao_filtrada_excel_chave'] = chave_export

                if st.session_state.get('visao_filtrada_excel_bytes') is not None:
                    if st.session_state.get('visao_filtrada_excel_chave') == chave_export:
                        st.download_button(
                            "📥 Baixar visão filtrada (Excel)",
                            data=st.session_state['visao_filtrada_excel_bytes'],
                            file_name="obras_visao_filtrada.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key="download_visao_filtrada"
                        )
                    else:
                        st.caption("Os filtros mudaram. Clique em **Gerar arquivo da visão filtrada** novamente.")

# -------------------------------------------------------------
# 6. GERENCIAMENTO DE ZOOM E RENDERIZAÇÃO FINAL DO MAPA
# -------------------------------------------------------------
if coords_foco_sgo:
    lats_foco = [p[0] for p in coords_foco_sgo]
    lons_foco = [p[1] for p in coords_foco_sgo]
    # Pequena margem garante que os dois pontos e o círculo de 50 m fiquem visíveis.
    min_lat, max_lat = min(lats_foco), max(lats_foco)
    min_lon, max_lon = min(lons_foco), max(lons_foco)
    margem_lat = max(0.0006, (max_lat - min_lat) * 0.35)
    margem_lon = max(0.0006, (max_lon - min_lon) * 0.35)
    mapa.fit_bounds(
        [[min_lat - margem_lat, min_lon - margem_lon], [max_lat + margem_lat, max_lon + margem_lon]],
        padding_top_left=[35, 35],
        padding_bottom_right=[35, 35],
        max_zoom=19
    )
elif zoom_lat is not None and zoom_lon is not None:
    mapa.fit_bounds([[zoom_lat - 0.001, zoom_lon - 0.001], [zoom_lat + 0.001, zoom_lon + 0.001]])
elif protocolo_zoom_lat is not None and protocolo_zoom_lon is not None:
    mapa.fit_bounds([[protocolo_zoom_lat - 0.001, protocolo_zoom_lon - 0.001], [protocolo_zoom_lat + 0.001, protocolo_zoom_lon + 0.001]])
elif busca_lat is not None and busca_lon is not None:
    mapa.fit_bounds([[busca_lat - 0.001, busca_lon - 0.001], [busca_lat + 0.001, busca_lon + 0.001]])
elif busca_lats and busca_lons: 
    mapa.fit_bounds([[min(busca_lats), min(busca_lons)], [max(busca_lats), max(busca_lons)]])
elif municipios_sel and geo_data_ibge:
    mun_foco_lats, mun_foco_lons = [], []
    for feature in geo_data_ibge['features']:
        if feature['properties'].get('MUNICIPIO') in municipios_sel:
            geom = feature['geometry']
            if geom['type'] == 'Polygon':
                for pt in geom['coordinates'][0]: mun_foco_lats.append(pt[1]); mun_foco_lons.append(pt[0])
            elif geom['type'] == 'MultiPolygon':
                for poly in geom['coordinates']:
                    for pt in poly[0]: mun_foco_lats.append(pt[1]); mun_foco_lons.append(pt[0])
    if mun_foco_lats and mun_foco_lons: mapa.fit_bounds([[min(mun_foco_lats), min(mun_foco_lons)], [max(mun_foco_lats), max(mun_foco_lons)]])
    elif todas_lats and todas_lons: mapa.fit_bounds([[min(todas_lats), min(todas_lons)], [max(todas_lats), max(todas_lons)]])
elif todas_lats and todas_lons: 
    mapa.fit_bounds([[min(todas_lats), min(todas_lons)], [max(todas_lats), max(todas_lons)]])

with map_container:
    st.markdown("### Mapa de Redes e Obras")
    st_folium(mapa, use_container_width=True, height=850, returned_objects=[])
