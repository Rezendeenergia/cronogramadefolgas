import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import numpy as np
from typing import Dict, List, Tuple
import requests
from msal import ConfidentialClientApplication
import io

st.set_page_config(
    page_title="Sistema de Cronograma de Equipes - Pará",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #F7931E 0%, #000000 100%);
        color: white; padding: 20px; border-radius: 10px;
        text-align: center; margin-bottom: 30px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
    }
    .metric-card {
        background: white; padding: 20px; border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border-left: 4px solid #F7931E; margin-bottom: 15px;
    }
    .stButton > button {
        background-color: #F7931E; color: #000000;
        border: 2px solid #000000; font-weight: bold;
    }
    .stButton > button:hover {
        background-color: #000000; color: #F7931E;
        border: 2px solid #F7931E;
    }
    .stMetric > div {
        background-color: white; border: 1px solid #F7931E;
        border-radius: 10px; padding: 15px;
    }
    .stExpander { border: 1px solid #F7931E; }
</style>
""", unsafe_allow_html=True)

CIDADES_PARA = {
    'Belém': {'lat': -1.4558, 'lon': -48.4902},
    'Ananindeua': {'lat': -1.3656, 'lon': -48.3739},
    'Santarém': {'lat': -2.4426, 'lon': -54.7085},
    'Santarem': {'lat': -2.4426, 'lon': -54.7085},
    'Marabá': {'lat': -5.3686, 'lon': -49.1178},
    'Parauapebas': {'lat': -6.0675, 'lon': -49.9024},
    'Castanhal': {'lat': -1.2939, 'lon': -47.9261},
    'Abaetetuba': {'lat': -1.7218, 'lon': -48.8788},
    'Canaã dos Carajás': {'lat': -6.4969, 'lon': -49.8771},
    'Marituba': {'lat': -1.3473, 'lon': -48.3439},
    'Barcarena': {'lat': -1.6155, 'lon': -48.6289},
    'Altamira': {'lat': -3.2039, 'lon': -52.2094},
    'Paragominas': {'lat': -2.9977, 'lon': -47.3548},
    'Tucuruí': {'lat': -3.7661, 'lon': -49.6725},
    'Bragança': {'lat': -1.0534, 'lon': -46.7655},
    'Itaituba': {'lat': -4.2761, 'lon': -55.9836},
    'Oriximiná': {'lat': -1.7653, 'lon': -55.8661},
    'Redenção': {'lat': -8.0273, 'lon': -50.0305},
    'Capanema': {'lat': -1.1944, 'lon': -47.1808},
    'Conceição do Araguaia': {'lat': -8.2578, 'lon': -49.2644},
    'Tailândia': {'lat': -2.9496, 'lon': -48.3458},
    'Juruti': {'lat': -2.1440, 'lon': -56.0891},
    'Vila Gorete': {'lat': -2.4256, 'lon': -55.2365},
    'Mojui dos Campos': {'lat': -2.6824, 'lon': -54.6418},
    'Menbeca': {'lat': -2.2196, 'lon': -54.9899},
    'Barreiras': {'lat': -4.0902, 'lon': -55.6892},
    'Almeirim': {'lat': -1.5276, 'lon': -52.5775},
    'Rurópolis': {'lat': -4.0941, 'lon': -54.9106},
    'Vista Alegre': {'lat': -2.6100, 'lon': -54.9200},
}

ABA_GERAL    = 'FOLGAS DAS EQUIPES GERAL'
ABA_NORDESTE = 'FOLGA DA NORDESTE'
ABA_P1       = 'Planilha1'

# Mapeamento padrão (GERAL e NORDESTE) — colunas com possíveis espaços
MAPPING_PADRAO = {
    'COLABORADOR': 'colaborador',
    'INICIO':      'inicio',
    'TERMINO':     'termino',
    'BASE/CAMPO':  'base_campo',
    'ORIGEM':      'origem',
    'DESTINO':     'destino',
    'FUNÇAO':      'funcao',
    'SUPERVISOR':  'supervisor',
    'MÊS':         'mes',
    'OBSERVAÇÃO':  'observacao',
}

# Mapeamento Planilha1
MAPPING_P1 = {
    'NOME':             'colaborador',
    'INICIO DA FOLGA':  'inicio',
    'FIM DA FOLGA':     'termino',
    'RETORNO DA FOLGA': 'base_campo',
    'ORIGEM':           'origem',
    'DESTINO':          'destino',
    'REGIONAL':         'regional',
    'SUPERVISOR':       'supervisor',
    'STATUS':           'status_planilha',
    'OBSERVAÇÃO':       'observacao',
}


def normalize_columns(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """Strip espaços nos headers e renomeia conforme mapping."""
    df.columns = [str(c).strip() for c in df.columns]
    rename = {}
    for src, dst in mapping.items():
        src_s = src.strip()
        if src_s in df.columns:
            rename[src_s] = dst
        else:
            for col in df.columns:
                if col.upper() == src_s.upper():
                    rename[col] = dst
                    break
    return df.rename(columns=rename)


def safe_to_timestamp(val):
    """Converte para Timestamp; retorna NaT se não for data válida (ex: 'ATESTADO')."""
    if val is None:
        return pd.NaT
    if isinstance(val, float) and np.isnan(val):
        return pd.NaT
    if isinstance(val, (datetime,)):
        return pd.Timestamp(val)
    try:
        return pd.Timestamp(val)
    except Exception:
        return pd.NaT


class SharePointConnector:
    def __init__(self):
        self.client_id     = st.secrets["sharepoint"]["client_id"]
        self.client_secret = st.secrets["sharepoint"]["client_secret"]
        self.tenant_id     = st.secrets["sharepoint"]["tenant_id"]

    def _headers(self):
        app = ConfidentialClientApplication(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            client_credential=self.client_secret,
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in result:
            raise Exception(f"Token falhou: {result.get('error_description','?')}")
        return {"Authorization": f"Bearer {result['access_token']}"}

    def _site_id(self, h):
        url = "https://graph.microsoft.com/v1.0/sites/rezendeenergia.sharepoint.com:/sites/Intranet"
        r = requests.get(url, headers=h); r.raise_for_status()
        return r.json()['id']

    def _download(self, h, site_id, filename):
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='{filename}')"
        r = requests.get(url, headers=h); r.raise_for_status()
        for item in r.json().get('value', []):
            if item['name'] == filename:
                dl = requests.get(
                    f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item['id']}/content",
                    headers=h
                )
                dl.raise_for_status()
                return dl.content
        raise FileNotFoundError(f"'{filename}' não encontrado no SharePoint.")

    @st.cache_data(ttl=300)
    def get_all_sheets(_self) -> dict:
        """Retorna {'geral': df|None, 'nordeste': df|None, 'planilha1': df|None}."""
        try:
            h = _self._headers()
            sid = _self._site_id(h)
            content = _self._download(h, sid, 'FOLGA DAS EQUIPES GERAL.xlsx')
            xls = pd.ExcelFile(io.BytesIO(content))
            sheets = {}

            # ── GERAL ─────────────────────────────────────────────────────
            if ABA_GERAL in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=ABA_GERAL)
                df = normalize_columns(df, MAPPING_PADRAO)
                sheets['geral'] = df
            else:
                sheets['geral'] = None

            # ── NORDESTE ──────────────────────────────────────────────────
            if ABA_NORDESTE in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=ABA_NORDESTE)
                df = normalize_columns(df, MAPPING_PADRAO)
                tem_dados = 'colaborador' in df.columns and df['colaborador'].dropna().shape[0] > 0
                sheets['nordeste'] = df if tem_dados else None
            else:
                sheets['nordeste'] = None

            # ── PLANILHA1 ─────────────────────────────────────────────────
            if ABA_P1 in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=ABA_P1)
                df = normalize_columns(df, MAPPING_P1)
                # Garantir colunas de compatibilidade
                for col in ['funcao', 'mes']:
                    if col not in df.columns:
                        df[col] = None
                tem_dados = 'colaborador' in df.columns and df['colaborador'].dropna().shape[0] > 0
                sheets['planilha1'] = df if tem_dados else None
            else:
                sheets['planilha1'] = None

            return sheets

        except Exception as e:
            st.error(f"Erro ao conectar com SharePoint: {e}")
            return None


class CronogramaAnalyzer:
    def __init__(self, df: pd.DataFrame, regional_label: str = "Geral"):
        self.df = df.copy()
        self.regional_label = regional_label
        self._process()

    def _process(self):
        if 'colaborador' not in self.df.columns:
            st.error("Coluna 'colaborador' não encontrada.")
            return

        # Datas seguras
        for col in ['inicio', 'termino', 'base_campo']:
            if col in self.df.columns:
                self.df[col] = self.df[col].apply(safe_to_timestamp)

        # Strip em strings
        str_cols = ['colaborador','origem','destino','supervisor','funcao',
                    'observacao','regional','status_planilha']
        for col in str_cols:
            if col in self.df.columns:
                self.df[col] = self.df[col].apply(
                    lambda x: str(x).strip()
                    if pd.notna(x) and str(x).strip() not in ('nan','None','')
                    else None
                )

        self.df = self.df[self.df['colaborador'].notna()].reset_index(drop=True)
        self._add_coords()

    def _add_coords(self):
        def coords(cidade):
            if pd.isna(cidade): return None, None
            c = CIDADES_PARA.get(str(cidade).strip().title(), {'lat': None,'lon': None})
            return c['lat'], c['lon']

        for prefix in ['origem', 'destino']:
            if prefix in self.df.columns:
                self.df[[f'{prefix}_lat', f'{prefix}_lon']] = self.df[prefix].apply(
                    lambda x: pd.Series(coords(x))
                )

    def fmt_date(self, v):
        if pd.isna(v): return 'N/A'
        if hasattr(v, 'strftime'): return v.strftime("%d/%m/%Y")
        return str(v)

    def audit_folgas(self) -> List[Dict]:
        problemas = []
        for col in self.df['colaborador'].unique():
            sub = self.df[self.df['colaborador'] == col].dropna(subset=['inicio','termino'])
            if len(sub) < 2: continue
            sub = sub.sort_values('inicio')
            for i in range(len(sub) - 1):
                t = sub.iloc[i]['termino']
                s = sub.iloc[i+1]['inicio']
                if pd.notna(t) and pd.notna(s):
                    diff = (s - t).days
                    if diff < 30:
                        problemas.append({
                            'colaborador':   col,
                            'folga1_termino': self.fmt_date(t),
                            'folga2_inicio':  self.fmt_date(s),
                            'dias_intervalo': diff,
                            'supervisor':    sub.iloc[i].get('supervisor',''),
                            'funcao':        sub.iloc[i].get('funcao',''),
                        })
        return problemas


def _build_cronograma_rows(analyzer: CronogramaAnalyzer):
    hoje = datetime.now().date()
    rows = []
    for _, row in analyzer.df.iterrows():
        col  = row['colaborador']
        sup  = row.get('supervisor') or 'N/A'
        dst  = row.get('destino')    or 'N/A'
        org  = row.get('origem')     or 'N/A'
        fnc  = row.get('funcao')     or 'N/A'
        obs  = row.get('observacao') or ''

        if pd.notna(row.get('inicio')) and pd.notna(row.get('termino')):
            ini = row['inicio'].date()
            ter = row['termino'].date()
            if ter < hoje:
                continue
            dias = (ini - hoje).days
            dur  = (ter - ini).days + 1

            if dias <= 0 and ter >= hoje:
                status, cor, pri = "EM_FOLGA",      "🟢", 2
                dias = 0
            elif dias <= 7:
                status, cor, pri = "URGENTE",        "🔴", 1
            elif dias <= 15:
                status, cor, pri = "ATENÇÃO",        "🟡", 3
            elif dias <= 30:
                status, cor, pri = "PROGRAMADO",     "🟢", 4
            else:
                status, cor, pri = "DISTANTE",       "🔵", 5

            rows.append(dict(colaborador=col, supervisor=sup, destino=dst, origem=org,
                             funcao=fnc, observacao=obs, inicio=ini, termino=ter,
                             dias=dias, duracao=dur, status=status, cor=cor, pri=pri))
        else:
            rows.append(dict(colaborador=col, supervisor=sup, destino='N/A', origem=org,
                             funcao=fnc, observacao=obs, inicio=None, termino=None,
                             dias=999, duracao=0, status="SEM_PROGRAMAÇÃO", cor="⚫", pri=6))

    return sorted(rows, key=lambda x: (x['pri'], x['dias']))


def show_cronograma(analyzer: CronogramaAnalyzer):
    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("🔄 Atualizar Dados", type="primary"):
            st.cache_data.clear(); st.rerun()

    st.subheader(f"📊 Cronograma — {analyzer.regional_label} — Ordenado por Urgência")
    rows = _build_cronograma_rows(analyzer)

    if not rows:
        st.warning("Nenhum dado encontrado."); return

    df_disp = pd.DataFrame([{
        '👤 COLABORADOR':    r['colaborador'],
        '🔧 FUNÇÃO':         r['funcao'],
        '🎯 DIAS P/ FOLGA':  r['dias'] if r['dias'] != 999 else 'N/A',
        '📅 INÍCIO':         analyzer.fmt_date(r['inicio']) if r['inicio'] else 'N/A',
        '📅 FIM':            analyzer.fmt_date(r['termino']) if r['termino'] else 'N/A',
        '⏱️ DURAÇÃO':        f"{r['duracao']} dias" if r['duracao'] > 0 else 'N/A',
        '🏠 ORIGEM':         r['origem'],
        '🏙️ DESTINO':        r['destino'],
        '👨‍💼 SUPERVISOR':     r['supervisor'],
        '⚡ STATUS':          f"{r['cor']} {r['status']}",
        '📝 OBS':            r['observacao'],
    } for r in rows])

    def hl(row):
        d = row['🎯 DIAS P/ FOLGA']
        if d == 'N/A': return ['background-color:#f8f9fa']*len(row)
        if isinstance(d,(int,float)) and d <= 3: return ['background-color:#ffebee']*len(row)
        if isinstance(d,(int,float)) and d <= 7: return ['background-color:#fff3e0']*len(row)
        return ['']*len(row)

    def cs(val):
        if '🔴' in str(val): return 'background-color:#ffebee'
        if '🟡' in str(val): return 'background-color:#fff8e1'
        if '🟢' in str(val): return 'background-color:#e8f5e8'
        if '⚫' in str(val): return 'background-color:#f5f5f5'
        return ''

    styled = df_disp.style.apply(hl, axis=1).applymap(cs, subset=['⚡ STATUS'])
    st.dataframe(styled, use_container_width=True, height=600)


def show_planilha1(analyzer: CronogramaAnalyzer):
    col1, col2 = st.columns([3,1])
    with col2:
        if st.button("🔄 Atualizar", type="primary", key="btn_p1"):
            st.cache_data.clear(); st.rerun()

    st.subheader("📋 Planilha1")
    df = analyzer.df.copy()

    col_f1, col_f2 = st.columns(2)
    with col_f1:
        opts_reg = ['Todos'] + sorted(df['regional'].dropna().unique().tolist()) \
                   if 'regional' in df.columns else ['Todos']
        reg = st.selectbox("Regional:", opts_reg)
    with col_f2:
        opts_st = ['Todos'] + sorted(df['status_planilha'].dropna().unique().tolist()) \
                  if 'status_planilha' in df.columns else ['Todos']
        sts = st.selectbox("Status:", opts_st)

    if reg != 'Todos' and 'regional' in df.columns:
        df = df[df['regional'] == reg]
    if sts != 'Todos' and 'status_planilha' in df.columns:
        df = df[df['status_planilha'] == sts]

    col_map = {
        'colaborador':   '👤 NOME',
        'inicio':        '📅 INÍCIO',
        'termino':       '📅 FIM',
        'base_campo':    '📅 RETORNO',
        'origem':        '🏠 ORIGEM',
        'destino':       '🏙️ DESTINO',
        'regional':      '🌎 REGIONAL',
        'supervisor':    '👨‍💼 SUPERVISOR',
        'status_planilha':'⚡ STATUS',
        'observacao':    '📝 OBS',
    }
    rows = []
    for _, row in df.iterrows():
        r = {}
        for k, label in col_map.items():
            val = row.get(k, None)
            r[label] = analyzer.fmt_date(val) if k in ('inicio','termino','base_campo') \
                       else (val if pd.notna(val) else 'N/A')
        rows.append(r)

    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, height=500)
    else:
        st.warning("Nenhum dado encontrado.")


def show_map(analyzer: CronogramaAnalyzer):
    st.header("🗺️ Mapa das Equipes")
    col1, col2 = st.columns(2)
    with col1:
        sups = ['Todos'] + list(analyzer.df['supervisor'].dropna().unique())
        sup_f = st.selectbox("Supervisor:", sups)
    with col2:
        cols = ['Todos'] + list(analyzer.df['colaborador'].dropna().unique())
        col_f = st.selectbox("Colaborador:", cols)

    df = analyzer.df.copy()
    if sup_f != 'Todos': df = df[df['supervisor'] == sup_f]
    if col_f != 'Todos': df = df[df['colaborador'] == col_f]

    if col_f != 'Todos':
        st.subheader(f"📋 Resumo: {col_f}")
        sub = df[df['colaborador'] == col_f]
        for idx, (_, row) in enumerate(sub.iterrows()):
            with st.expander(f"Programação {idx+1}"):
                c1,c2,c3,c4 = st.columns(4)
                c1.write(f"**🏠 Origem:** {row.get('origem','N/A')}")
                c2.write(f"**🎯 Destino:** {row.get('destino','N/A')}")
                c3.write(f"**👨‍💼 Supervisor:** {row.get('supervisor','N/A')}")
                c4.write(f"**🔧 Função:** {row.get('funcao','N/A') or 'N/A'}")
                d1,d2,d3 = st.columns(3)
                d1.write(f"**📅 Início:** {analyzer.fmt_date(row.get('inicio'))}")
                d2.write(f"**📅 Fim:** {analyzer.fmt_date(row.get('termino'))}")
                d3.write(f"**📅 Retorno:** {analyzer.fmt_date(row.get('base_campo'))}")
                obs = row.get('observacao','')
                if obs and obs != 'N/A': st.info(f"📝 {obs}")
                if pd.notna(row.get('inicio')) and pd.notna(row.get('termino')):
                    st.info(f"⏰ Duração: {(row['termino']-row['inicio']).days+1} dias")

    if df.empty:
        st.warning("Nenhum dado para os filtros."); return

    colors = ['#F7931E','#000000','red','green','purple','orange','darkred',
              'lightred','beige','darkblue','darkgreen','cadetblue','darkpurple',
              'white','pink','lightblue','lightgreen','gray','lightgray']

    if col_f != 'Todos':
        cidades_rel = set()
        for _, row in df.iterrows():
            for c in ['origem','destino']:
                if pd.notna(row.get(c)): cidades_rel.add(str(row[c]).strip().title())
        mapa = folium.Map(location=[-3.7,-52.0], zoom_start=6, tiles='OpenStreetMap')
        for cidade in cidades_rel:
            if cidade in CIDADES_PARA:
                cd = CIDADES_PARA[cidade]
                folium.Marker([cd['lat'],cd['lon']],
                    popup=f"<b>{cidade}</b>", tooltip=cidade,
                    icon=folium.Icon(color='orange',icon='home')).add_to(mapa)
        for idx,(_, row) in enumerate(df.iterrows()):
            if all(pd.notna(row.get(k)) for k in ['origem_lat','origem_lon','destino_lat','destino_lon']):
                clr = colors[idx % len(colors)]
                i_br = analyzer.fmt_date(row.get('inicio'))
                t_br = analyzer.fmt_date(row.get('termino'))
                b_br = analyzer.fmt_date(row.get('base_campo'))
                folium.PolyLine(
                    [[row['origem_lat'],row['origem_lon']],[row['destino_lat'],row['destino_lon']]],
                    color=clr, weight=3, opacity=0.8,
                    popup=f"<b>{row['colaborador']}</b><br>De:{row.get('origem','N/A')}<br>Para:{row.get('destino','N/A')}<br>{i_br}-{t_br}<br>Retorno:{b_br}"
                ).add_to(mapa)
                folium.Marker([row['destino_lat'],row['destino_lon']],
                    popup=f"<b>{row['colaborador']}</b><br>Destino:{row.get('destino','N/A')}<br>Supervisor:{row.get('supervisor','N/A')}<br>{i_br}-{t_br}<br>Retorno:{b_br}",
                    tooltip=f"{row['colaborador']} - {row.get('destino','N/A')}",
                    icon=folium.Icon(color=clr,icon='user')).add_to(mapa)
    else:
        mapa = folium.Map(location=[-3.7,-52.0], zoom_start=6, tiles='OpenStreetMap')
        for cidade, cd in CIDADES_PARA.items():
            folium.Marker([cd['lat'],cd['lon']], popup=f"<b>{cidade}</b>", tooltip=cidade,
                icon=folium.Icon(color='orange',icon='home')).add_to(mapa)
        for idx,(_, row) in enumerate(df.iterrows()):
            if all(pd.notna(row.get(k)) for k in ['origem_lat','origem_lon','destino_lat','destino_lon']):
                clr = colors[idx % len(colors)]
                folium.PolyLine(
                    [[row['origem_lat'],row['origem_lon']],[row['destino_lat'],row['destino_lon']]],
                    color=clr, weight=3, opacity=0.8,
                    popup=f"<b>{row['colaborador']}</b><br>De:{row.get('origem','?')}<br>Para:{row.get('destino','?')}"
                ).add_to(mapa)
                folium.Marker([row['destino_lat'],row['destino_lon']],
                    popup=f"<b>{row['colaborador']}</b><br>{row.get('destino','?')}<br>{row.get('supervisor','?')}",
                    tooltip=f"{row['colaborador']} - {row.get('destino','?')}",
                    icon=folium.Icon(color=clr,icon='user')).add_to(mapa)

    st_folium(mapa, width=1000, height=600)
    if col_f != 'Todos':
        st.info("💡 🏠 Laranja = Bases | 👤 Colorido = Destino | Linhas = Rota")
    else:
        st.info("💡 🏠 Laranja = Bases Operacionais | 👤 Colorido = Destinos | Linhas = Rotas")


def show_audit(analyzer: CronogramaAnalyzer):
    st.header("🔍 Auditoria de Folgas")
    st.info("📋 Regra: mínimo 30 dias de intervalo entre folgas.")
    problemas = analyzer.audit_folgas()

    if problemas:
        st.error(f"⚠️ {len(problemas)} problema(s) encontrado(s):")
        for i, p in enumerate(problemas, 1):
            with st.expander(f"Problema {i}: {p['colaborador']} — {p['dias_intervalo']} dias"):
                c1,c2 = st.columns(2)
                c1.write(f"**Término folga 1:** {p['folga1_termino']}")
                c2.write(f"**Início folga 2:** {p['folga2_inicio']}")
                st.write(f"**Supervisor:** {p['supervisor']} | **Função:** {p.get('funcao','N/A') or 'N/A'}")
                st.write(f"**Intervalo:** {p['dias_intervalo']} dias")
                if p['dias_intervalo'] < 15:   st.error("🚨 Crítico: < 15 dias")
                elif p['dias_intervalo'] < 30:  st.warning("⚠️ Atenção: < 30 dias")
        st.subheader("📊 Estatísticas")
        c1,c2,c3 = st.columns(3)
        c1.metric("Total Problemas", len(problemas))
        c2.metric("Casos Críticos", len([p for p in problemas if p['dias_intervalo'] < 15]))
        c3.metric("Intervalo Médio", f"{sum(p['dias_intervalo'] for p in problemas)/len(problemas):.1f} dias")
    else:
        st.success("✅ Todas as folgas em conformidade!")


def show_reports(analyzer: CronogramaAnalyzer):
    st.header("📊 Relatórios")

    st.subheader("👨‍💼 Por Supervisor")
    sup_stats = analyzer.df.groupby('supervisor').agg(colaborador=('colaborador','count')).reset_index()
    sup_stats.columns = ['Supervisor','Colaboradores']
    st.dataframe(sup_stats, use_container_width=True)
    fig = px.bar(sup_stats, x='Supervisor', y='Colaboradores',
                 title='Colaboradores por Supervisor',
                 color='Colaboradores',
                 color_continuous_scale=[[0,'#000000'],[1,'#F7931E']])
    st.plotly_chart(fig, use_container_width=True)

    # Por Função
    if 'funcao' in analyzer.df.columns and analyzer.df['funcao'].notna().any():
        st.subheader("🔧 Por Função")
        fnc = analyzer.df['funcao'].value_counts().reset_index()
        fnc.columns = ['Função','Qtd']
        fig2 = px.bar(fnc.head(15), x='Função', y='Qtd', title='Distribuição por Função',
                      color='Qtd', color_continuous_scale=[[0,'#000000'],[1,'#F7931E']])
        st.plotly_chart(fig2, use_container_width=True)

    st.subheader("🏙️ Movimentações")
    c1,c2 = st.columns(2)
    with c1:
        st.write("**Origens:**")
        if 'origem' in analyzer.df.columns:
            o = analyzer.df['origem'].value_counts().head(10)
            if not o.empty: st.bar_chart(o)
    with c2:
        st.write("**Destinos:**")
        if 'destino' in analyzer.df.columns:
            d = analyzer.df['destino'].value_counts().head(10)
            if not d.empty: st.bar_chart(d)

    st.subheader("💾 Exportar")
    if st.button("📥 Baixar CSV"):
        st.download_button("📥 Download CSV", analyzer.df.to_csv(index=False),
            file_name=f"cronograma_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv")


def main():
    st.markdown("""
    <div class="main-header">
        <h1>🗺️ Sistema de Cronograma de Equipes - Pará</h1>
        <p>Gestão inteligente de folgas e movimentação das equipes</p>
    </div>""", unsafe_allow_html=True)

    st.sidebar.title("📋 Menu de Navegação")
    page = st.sidebar.selectbox("Selecione:", [
        "📋 Cronograma Geral",
        "📋 Cronograma Nordeste",
        "📋 Planilha1",
        "🗺️ Mapa das Equipes",
        "🔍 Auditoria de Folgas",
        "📊 Relatórios",
    ])

    with st.spinner("Carregando dados do SharePoint..."):
        connector = SharePointConnector()
        sheets = connector.get_all_sheets()

    if sheets is None:
        st.error("❌ Não foi possível carregar os dados. Verifique a conexão com o SharePoint.")
        st.stop()

    if page == "📋 Cronograma Geral":
        df = sheets.get('geral')
        if df is None or df.empty:
            st.warning("Aba GERAL vazia ou não encontrada.")
        else:
            show_cronograma(CronogramaAnalyzer(df, "Geral"))

    elif page == "📋 Cronograma Nordeste":
        df = sheets.get('nordeste')
        if df is None or df.empty:
            st.info("A aba 'FOLGA DA NORDESTE' está vazia ou sem dados.")
        else:
            show_cronograma(CronogramaAnalyzer(df, "Nordeste"))

    elif page == "📋 Planilha1":
        df = sheets.get('planilha1')
        if df is None or df.empty:
            st.info("A 'Planilha1' está vazia ou sem dados.")
        else:
            show_planilha1(CronogramaAnalyzer(df, "Planilha1"))

    elif page == "🗺️ Mapa das Equipes":
        df = sheets.get('geral')
        if df is None or df.empty:
            st.warning("Sem dados para o mapa.")
        else:
            show_map(CronogramaAnalyzer(df, "Geral"))

    elif page == "🔍 Auditoria de Folgas":
        df = sheets.get('geral')
        if df is None or df.empty:
            st.warning("Sem dados para auditar.")
        else:
            show_audit(CronogramaAnalyzer(df, "Geral"))

    elif page == "📊 Relatórios":
        df = sheets.get('geral')
        if df is None or df.empty:
            st.warning("Sem dados para relatório.")
        else:
            show_reports(CronogramaAnalyzer(df, "Geral"))


if __name__ == "__main__":
    main()
