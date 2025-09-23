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

# Configuração da página
st.set_page_config(
    page_title="Sistema de Cronograma de Equipes - Pará",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado para UX/UI avançado
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #F7931E 0%, #000000 100%);
        color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 30px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
    }

    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        border-left: 4px solid #F7931E;
        margin-bottom: 15px;
    }

    .status-ativo {
        background-color: #F7931E;
        color: #000000;
        padding: 5px 10px;
        border-radius: 15px;
        font-size: 12px;
        font-weight: bold;
    }

    .status-folga {
        background-color: #000000;
        color: #F7931E;
        padding: 5px 10px;
        border-radius: 15px;
        font-size: 12px;
        font-weight: bold;
    }

    .status-indefinido {
        background-color: #f8f9fa;
        color: #000000;
        padding: 5px 10px;
        border-radius: 15px;
        font-size: 12px;
        font-weight: bold;
        border: 1px solid #F7931E;
    }

    .audit-warning {
        background-color: #fff3cd;
        border: 1px solid #F7931E;
        border-radius: 5px;
        padding: 10px;
        margin: 10px 0;
    }

    .audit-error {
        background-color: #000000;
        border: 1px solid #F7931E;
        border-radius: 5px;
        padding: 10px;
        margin: 10px 0;
        color: #F7931E;
    }

    .sidebar .sidebar-content {
        background: linear-gradient(180deg, #F7931E 0%, #000000 100%);
    }

    /* Estilização adicional para elementos Streamlit */
    .stSelectbox > div > div {
        border-color: #F7931E;
    }

    .stButton > button {
        background-color: #F7931E;
        color: #000000;
        border: 2px solid #000000;
        font-weight: bold;
    }

    .stButton > button:hover {
        background-color: #000000;
        color: #F7931E;
        border: 2px solid #F7931E;
    }

    .stMetric > div {
        background-color: white;
        border: 1px solid #F7931E;
        border-radius: 10px;
        padding: 15px;
    }

    .stTab {
        background-color: #F7931E;
        color: #000000;
    }

    .stExpander {
        border: 1px solid #F7931E;
    }
</style>
""", unsafe_allow_html=True)

# Coordenadas das principais cidades do Pará
CIDADES_PARA = {
    'Belém': {'lat': -1.4558, 'lon': -48.4902},
    'Ananindeua': {'lat': -1.3656, 'lon': -48.3739},
    'Santarém': {'lat': -2.4426, 'lon': -54.7085},
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
    'Almeirim': {'lat': -1.5276090427351592, 'lon': -52.577482130144006},
    'Rurópolis': {'lat': -4.094116299218916, 'lon': -54.91062274171425}
}


class SharePointConnector:
    def __init__(self):
        self.client_id = st.secrets["sharepoint"]["client_id"]
        self.client_secret = st.secrets["sharepoint"]["client_secret"]
        self.tenant_id = st.secrets["sharepoint"]["tenant_id"]

    @st.cache_data(ttl=300)  # Cache por 5 minutos
    def get_data(_self):
        try:
            app = ConfidentialClientApplication(
                _self.client_id,
                authority=f"https://login.microsoftonline.com/{_self.tenant_id}",
                client_credential=_self.client_secret,
            )

            result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            if "access_token" in result:
                headers = {"Authorization": f"Bearer {result['access_token']}"}

                # Obter site_id
                site_url = "https://graph.microsoft.com/v1.0/sites/rezendeenergia.sharepoint.com:/sites/Intranet"
                site_response = requests.get(site_url, headers=headers)

                if site_response.status_code == 200:
                    site_data = site_response.json()
                    site_id = site_data['id']

                    # Buscar arquivo
                    search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='FOLGA DAS EQUIPES GERAL.xlsx')"
                    search_response = requests.get(search_url, headers=headers)

                    if search_response.status_code == 200:
                        search_data = search_response.json()
                        files_found = search_data.get('value', [])

                        for item in files_found:
                            if item['name'] == 'FOLGA DAS EQUIPES GERAL.xlsx':
                                download_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item['id']}/content"
                                download_response = requests.get(download_url, headers=headers)

                                if download_response.status_code == 200:
                                    df = pd.read_excel(io.BytesIO(download_response.content))
                                    return df
            return None
        except Exception as e:
            st.error(f"Erro ao conectar com SharePoint: {e}")
            return None


class CronogramaAnalyzer:
    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()
        self.process_data()

    def process_data(self):
        """Processa e limpa os dados"""
        # Renomear colunas para padrão
        column_mapping = {
            'COLABORADOR': 'colaborador',
            'INICIO': 'inicio',
            'TERMINO': 'termino',
            'BASE/CAMPO': 'base_campo',
            'ORIGEM': 'origem',
            'DESTINO': 'destino',
            'SUPERVISOR': 'supervisor',
            'MÊS': 'mes'
        }

        # Usar os nomes atuais das colunas se existirem
        available_columns = self.df.columns.tolist()
        for old_name, new_name in column_mapping.items():
            if old_name in available_columns:
                self.df.rename(columns={old_name: new_name}, inplace=True)
            elif len(available_columns) >= len(column_mapping):
                # Se não encontrar pelo nome, usa a posição
                idx = list(column_mapping.keys()).index(old_name)
                if idx < len(available_columns):
                    self.df.rename(columns={available_columns[idx]: new_name}, inplace=True)

        # Converter datas
        date_columns = ['inicio', 'termino', 'base_campo']
        for col in date_columns:
            if col in self.df.columns:
                self.df[col] = pd.to_datetime(self.df[col], errors='coerce')

        # Limpar dados vazios
        self.df = self.df.dropna(subset=['colaborador'])

        # Adicionar coordenadas
        self.add_coordinates()

    def add_coordinates(self):
        """Adiciona coordenadas das cidades"""

        def get_coords(cidade):
            if pd.isna(cidade) or cidade == '':
                return None, None

            cidade_clean = str(cidade).strip().title()
            coords = CIDADES_PARA.get(cidade_clean, {'lat': None, 'lon': None})
            return coords['lat'], coords['lon']

        if 'origem' in self.df.columns:
            self.df[['origem_lat', 'origem_lon']] = self.df['origem'].apply(
                lambda x: pd.Series(get_coords(x))
            )

        if 'destino' in self.df.columns:
            self.df[['destino_lat', 'destino_lon']] = self.df['destino'].apply(
                lambda x: pd.Series(get_coords(x))
            )

    def format_date_br(self, date_value):
        """Formatar data para padrão brasileiro (dd/mm/aaaa)"""
        if pd.isna(date_value):
            return 'N/A'

        if hasattr(date_value, 'strftime'):
            return date_value.strftime("%d/%m/%Y")

        return str(date_value)

    def get_status_atual(self) -> Dict:
        """Retorna status atual das equipes"""
        hoje = datetime.now().date()
        status = {
            'em_folga': [],
            'ativo': [],
            'sem_programacao': []
        }

        for _, row in self.df.iterrows():
            colaborador = row['colaborador']

            if pd.notna(row.get('inicio')) and pd.notna(row.get('termino')):
                inicio = row['inicio'].date() if hasattr(row['inicio'], 'date') else row['inicio']
                termino = row['termino'].date() if hasattr(row['termino'], 'date') else row['termino']

                if inicio <= hoje <= termino:
                    status['em_folga'].append({
                        'colaborador': colaborador,
                        'inicio': self.format_date_br(inicio),
                        'termino': self.format_date_br(termino),
                        'origem': row.get('origem', ''),
                        'destino': row.get('destino', ''),
                        'supervisor': row.get('supervisor', '')
                    })
                else:
                    status['ativo'].append({
                        'colaborador': colaborador,
                        'origem': row.get('origem', ''),
                        'supervisor': row.get('supervisor', '')
                    })
            else:
                status['sem_programacao'].append({
                    'colaborador': colaborador,
                    'supervisor': row.get('supervisor', '')
                })

        return status

    def audit_folgas(self) -> List[Dict]:
        """Audita intervalos entre folgas (mínimo 30 dias)"""
        problemas = []

        # Agrupar por colaborador
        colaboradores = self.df['colaborador'].unique()

        for colaborador in colaboradores:
            folgas_colaborador = self.df[self.df['colaborador'] == colaborador].copy()
            folgas_colaborador = folgas_colaborador.dropna(subset=['inicio', 'termino'])

            if len(folgas_colaborador) > 1:
                # Ordenar por data de início
                folgas_colaborador = folgas_colaborador.sort_values('inicio')

                for i in range(len(folgas_colaborador) - 1):
                    termino_atual = folgas_colaborador.iloc[i]['termino']
                    inicio_proximo = folgas_colaborador.iloc[i + 1]['inicio']

                    if pd.notna(termino_atual) and pd.notna(inicio_proximo):
                        diferenca = (inicio_proximo - termino_atual).days

                        if diferenca < 30:
                            problemas.append({
                                'colaborador': colaborador,
                                'folga1_termino': self.format_date_br(termino_atual),
                                'folga2_inicio': self.format_date_br(inicio_proximo),
                                'dias_intervalo': diferenca,
                                'supervisor': folgas_colaborador.iloc[i]['supervisor']
                            })

        return problemas


def create_map(df: pd.DataFrame) -> folium.Map:
    """Cria mapa com as movimentações das equipes"""
    # Centro do Pará
    m = folium.Map(
        location=[-3.7, -52.0],
        zoom_start=6,
        tiles='OpenStreetMap'
    )

    # Adicionar cidades base
    for cidade, coords in CIDADES_PARA.items():
        folium.Marker(
            location=[coords['lat'], coords['lon']],
            popup=f"<b>{cidade}</b><br>Base Operacional",
            tooltip=cidade,
            icon=folium.Icon(color='orange', icon='home')
        ).add_to(m)

        # Adicionar rotas de folga
    colors = ['#F7931E', '#000000', 'red', 'green', 'purple', 'orange', 'darkred', 'lightred',
              'beige', 'darkblue', 'darkgreen', 'cadetblue', 'darkpurple',
              'white', 'pink', 'lightblue', 'lightgreen', 'gray', 'lightgray']

    for idx, (_, row) in enumerate(df.iterrows()):
        if (pd.notna(row.get('origem_lat')) and pd.notna(row.get('origem_lon')) and
                pd.notna(row.get('destino_lat')) and pd.notna(row.get('destino_lon'))):
            color = colors[idx % len(colors)]

            # Linha da rota
            folium.PolyLine(
                locations=[
                    [row['origem_lat'], row['origem_lon']],
                    [row['destino_lat'], row['destino_lon']]
                ],
                color=color,
                weight=3,
                opacity=0.8,
                popup=f"<b>{row['colaborador']}</b><br>"
                      f"De: {row.get('origem', 'N/A')}<br>"
                      f"Para: {row.get('destino', 'N/A')}<br>"
                      f"Período: {row.get('inicio', 'N/A')} - {row.get('termino', 'N/A')}"
            ).add_to(m)

            # Marcador de destino
            folium.Marker(
                location=[row['destino_lat'], row['destino_lon']],
                popup=f"<b>{row['colaborador']}</b><br>"
                      f"Destino: {row.get('destino', 'N/A')}<br>"
                      f"Supervisor: {row.get('supervisor', 'N/A')}<br>"
                      f"Período: {row.get('inicio', 'N/A')} - {row.get('termino', 'N/A')}",
                tooltip=f"{row['colaborador']} - {row.get('destino', 'N/A')}",
                icon=folium.Icon(color=color, icon='user')
            ).add_to(m)

    return m


def show_cronograma_encarregado(analyzer):
    """Página executiva mostrando cronograma ordenado por proximidade de folga"""

    # Botão de atualização no topo
    col1, col2 = st.columns([3, 1])

    with col2:
        if st.button("🔄 Atualizar Dados", type="primary"):
            st.cache_data.clear()  # Limpa o cache
            st.rerun()  # Recarrega a página

    # Calcular dados para dashboard
    hoje = datetime.now().date()
    dados_folgas = []

    for _, row in analyzer.df.iterrows():
        colaborador = row['colaborador']
        supervisor = row.get('supervisor', 'N/A')
        destino = row.get('destino', 'N/A')
        origem = row.get('origem', 'N/A')

        if pd.notna(row.get('inicio')):
            inicio = row['inicio'].date() if hasattr(row['inicio'], 'date') else row['inicio']
            termino = row['termino'].date() if hasattr(row['termino'], 'date') else row['termino']

            # Se já voltou da folga (termino < hoje), não incluir na tabela
            if termino < hoje:
                continue

            dias_para_folga = (inicio - hoje).days
            duracao = (termino - inicio).days + 1

            # Determinar status
            if dias_para_folga <= 0:
                if termino >= hoje:
                    status = "EM_FOLGA"
                    status_cor = "🟢"
                    dias_para_folga = 0
                    prioridade = 2  # Em folga tem prioridade 2
                else:
                    continue  # Não deveria chegar aqui devido ao filtro acima
            elif dias_para_folga <= 7:
                status = "URGENTE"
                status_cor = "🔴"
                prioridade = 1  # Urgente tem prioridade 1 (mais alta)
            elif dias_para_folga <= 15:
                status = "ATENÇÃO"
                status_cor = "🟡"
                prioridade = 3
            elif dias_para_folga <= 30:
                status = "PROGRAMADO"
                status_cor = "🟢"
                prioridade = 4
            else:
                status = "DISTANTE"
                status_cor = "🔵"
                prioridade = 5

            dados_folgas.append({
                'colaborador': colaborador,
                'supervisor': supervisor,
                'destino': destino,
                'origem': origem,
                'inicio': inicio,
                'termino': termino,
                'dias_para_folga': dias_para_folga,
                'duracao': duracao,
                'status': status,
                'status_cor': status_cor,
                'prioridade': prioridade
            })
        else:
            # Sem programação
            dados_folgas.append({
                'colaborador': colaborador,
                'supervisor': supervisor,
                'destino': 'N/A',
                'origem': origem,
                'inicio': None,
                'termino': None,
                'dias_para_folga': 999,  # Colocar no final
                'duracao': 0,
                'status': "SEM_PROGRAMAÇÃO",
                'status_cor': "⚫",
                'prioridade': 6
            })

    # Ordenar por prioridade (urgência) e depois por dias para folga
    dados_folgas = sorted(dados_folgas, key=lambda x: (x['prioridade'], x['dias_para_folga']))

    # TABELA EXECUTIVA
    st.subheader("📊 Cronograma Detalhado - Ordenado por Urgência")

    if dados_folgas:
        # Criar DataFrame para exibição
        df_display = pd.DataFrame([{
            '👤 COLABORADOR': dado['colaborador'],
            '🎯 DIAS PARA FOLGA': dado['dias_para_folga'] if dado['dias_para_folga'] != 999 else 'N/A',
            '📅 INÍCIO': analyzer.format_date_br(dado['inicio']) if dado['inicio'] else 'N/A',
            '📅 FIM': analyzer.format_date_br(dado['termino']) if dado['termino'] else 'N/A',
            '⏱️ DURAÇÃO': f"{dado['duracao']} dias" if dado['duracao'] > 0 else 'N/A',
            '🏠 ORIGEM': dado['origem'],
            '🏙️ DESTINO': dado['destino'],
            '👨‍💼 SUPERVISOR': dado['supervisor'],
            '⚡ STATUS': f"{dado['status_cor']} {dado['status']}"
        } for dado in dados_folgas])

        # Aplicar estilo baseado no status
        def color_status(val):
            if '🔴' in str(val):
                return 'background-color: #ffebee'
            elif '🟡' in str(val):
                return 'background-color: #fff8e1'
            elif '🟢' in str(val):
                return 'background-color: #e8f5e8'
            elif '⚫' in str(val):
                return 'background-color: #f5f5f5'
            return ''

        # Colorir linha inteira baseado nos dias para folga
        def highlight_urgency(row):
            dias = row['🎯 DIAS PARA FOLGA']
            if dias == 'N/A':
                return ['background-color: #f8f9fa'] * len(row)
            elif isinstance(dias, (int, float)) and dias <= 3:
                return ['background-color: #ffebee'] * len(row)
            elif isinstance(dias, (int, float)) and dias <= 7:
                return ['background-color: #fff3e0'] * len(row)
            else:
                return [''] * len(row)

        styled_df = df_display.style.apply(highlight_urgency, axis=1) \
            .applymap(color_status, subset=['⚡ STATUS'])

        st.dataframe(styled_df, use_container_width=True, height=600)

    else:
        st.warning("Nenhum dado encontrado.")


def show_map_page(analyzer):
    st.header("🗺️ Mapa das Equipes")

    # Filtros
    col1, col2 = st.columns(2)

    with col1:
        supervisores = ['Todos'] + list(analyzer.df['supervisor'].dropna().unique())
        supervisor_filtro = st.selectbox("Filtrar por Supervisor:", supervisores)

    with col2:
        colaboradores = ['Todos'] + list(analyzer.df['colaborador'].dropna().unique())
        colaborador_filtro = st.selectbox("Filtrar por Colaborador:", colaboradores)

    # Aplicar filtros
    df_filtered = analyzer.df.copy()

    if supervisor_filtro != 'Todos':
        df_filtered = df_filtered[df_filtered['supervisor'] == supervisor_filtro]

    if colaborador_filtro != 'Todos':
        df_filtered = df_filtered[df_filtered['colaborador'] == colaborador_filtro]

    # Mostrar resumo se colaborador específico foi selecionado
    if colaborador_filtro != 'Todos':
        st.subheader(f"📋 Resumo: {colaborador_filtro}")

        colaborador_data = df_filtered[df_filtered['colaborador'] == colaborador_filtro]

        if not colaborador_data.empty:
            for idx, (_, row) in enumerate(colaborador_data.iterrows()):
                with st.expander(f"Programação {idx + 1}"):
                    col1, col2, col3 = st.columns(3)

                    with col1:
                        st.write("**🏠 Origem:**")
                        st.write(row.get('origem', 'N/A'))

                    with col2:
                        st.write("**🎯 Destino:**")
                        st.write(row.get('destino', 'N/A'))

                    with col3:
                        st.write("**👨‍💼 Supervisor:**")
                        st.write(row.get('supervisor', 'N/A'))

                    # Datas
                    col4, col5, col6 = st.columns(3)

                    with col4:
                        st.write("**📅 Início da Folga:**")
                        st.write(analyzer.format_date_br(row.get('inicio')))

                    with col5:
                        st.write("**📅 Fim da Folga:**")
                        st.write(analyzer.format_date_br(row.get('termino')))

                    with col6:
                        st.write("**📅 Retorno Base/Campo:**")
                        st.write(analyzer.format_date_br(row.get('base_campo')))

                    # Calcular duração da folga
                    if pd.notna(row.get('inicio')) and pd.notna(row.get('termino')):
                        inicio = row['inicio']
                        termino = row['termino']
                        duracao = (termino - inicio).days + 1
                        st.info(f"⏰ **Duração da folga:** {duracao} dias")
        else:
            st.warning("Nenhuma programação encontrada para este colaborador.")

    # Criar e exibir mapa
    if not df_filtered.empty:
        # Lógica para mostrar apenas cidades relevantes quando filtrar colaborador
        if colaborador_filtro != 'Todos':
            # Coletar cidades origem e destino do colaborador filtrado
            cidades_relevantes = set()
            for _, row in df_filtered.iterrows():
                if pd.notna(row.get('origem')):
                    cidades_relevantes.add(str(row['origem']).strip().title())
                if pd.notna(row.get('destino')):
                    cidades_relevantes.add(str(row['destino']).strip().title())

            # Criar mapa personalizado
            mapa = folium.Map(
                location=[-3.7, -52.0],
                zoom_start=6,
                tiles='OpenStreetMap'
            )

            # Adicionar apenas cidades relevantes
            for cidade in cidades_relevantes:
                if cidade in CIDADES_PARA:
                    coords = CIDADES_PARA[cidade]
                    folium.Marker(
                        location=[coords['lat'], coords['lon']],
                        popup=f"<b>{cidade}</b><br>Base Operacional",
                        tooltip=cidade,
                        icon=folium.Icon(color='orange', icon='home')
                    ).add_to(mapa)

            # Adicionar rotas do colaborador
            colors = ['#F7931E', '#000000', 'red', 'green', 'purple']
            for idx, (_, row) in enumerate(df_filtered.iterrows()):
                if (pd.notna(row.get('origem_lat')) and pd.notna(row.get('origem_lon')) and
                        pd.notna(row.get('destino_lat')) and pd.notna(row.get('destino_lon'))):
                    color = colors[idx % len(colors)]

                    # Formatação de datas
                    inicio_br = analyzer.format_date_br(row.get('inicio'))
                    termino_br = analyzer.format_date_br(row.get('termino'))
                    base_campo_br = analyzer.format_date_br(row.get('base_campo'))

                    # Linha da rota
                    folium.PolyLine(
                        locations=[
                            [row['origem_lat'], row['origem_lon']],
                            [row['destino_lat'], row['destino_lon']]
                        ],
                        color=color,
                        weight=3,
                        opacity=0.8,
                        popup=f"<b>{row['colaborador']}</b><br>"
                              f"De: {row.get('origem', 'N/A')}<br>"
                              f"Para: {row.get('destino', 'N/A')}<br>"
                              f"Folga: {inicio_br} - {termino_br}<br>"
                              f"Retorno Base/Campo: {base_campo_br}"
                    ).add_to(mapa)

                    # Marcador de destino
                    folium.Marker(
                        location=[row['destino_lat'], row['destino_lon']],
                        popup=f"<b>{row['colaborador']}</b><br>"
                              f"Destino: {row.get('destino', 'N/A')}<br>"
                              f"Supervisor: {row.get('supervisor', 'N/A')}<br>"
                              f"Folga: {inicio_br} - {termino_br}<br>"
                              f"Retorno Base/Campo: {base_campo_br}",
                        tooltip=f"{row['colaborador']} - {row.get('destino', 'N/A')}",
                        icon=folium.Icon(color=color, icon='user')
                    ).add_to(mapa)
        else:
            # Usar função original para visão geral
            mapa = create_map(df_filtered)

        st_folium(mapa, width=1000, height=600)

        # Legenda
        if colaborador_filtro != 'Todos':
            st.info(
                "💡 **Legenda:** 🏠 Laranja = Bases Origem/Destino do colaborador | 👤 Colorido = Destino da folga | "
                "Linhas coloridas = Rota da movimentação")
        else:
            st.info("💡 **Legenda:** 🏠 Laranja = Bases Operacionais | 👤 Colorido = Colaboradores em destino | "
                    "Linhas coloridas = Rotas de movimentação")
    else:
        st.warning("Nenhum dado encontrado para os filtros selecionados.")


def show_audit_page(analyzer):
    st.header("🔍 Auditoria de Folgas")

    st.info("📋 **Regra:** Deve haver pelo menos 30 dias de intervalo entre uma folga e outra.")

    problemas = analyzer.audit_folgas()

    if problemas:
        st.error(f"⚠️ Encontrados {len(problemas)} problema(s) de conformidade:")

        for i, problema in enumerate(problemas, 1):
            with st.expander(f"Problema {i}: {problema['colaborador']} - {problema['dias_intervalo']} dias"):
                col1, col2 = st.columns(2)

                with col1:
                    st.write("**Primeira Folga (Término):**")
                    st.write(problema['folga1_termino'])

                with col2:
                    st.write("**Segunda Folga (Início):**")
                    st.write(problema['folga2_inicio'])

                st.write(f"**Supervisor:** {problema['supervisor']}")
                st.write(f"**Intervalo:** {problema['dias_intervalo']} dias")

                if problema['dias_intervalo'] < 15:
                    st.error("🚨 Crítico: Menos de 15 dias de intervalo")
                elif problema['dias_intervalo'] < 30:
                    st.warning("⚠️ Atenção: Menos de 30 dias de intervalo")
    else:
        st.success("✅ Todas as folgas estão em conformidade com a regra de 30 dias!")

    # Estatísticas de auditoria
    st.subheader("📊 Estatísticas de Auditoria")

    if problemas:
        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("Total de Problemas", len(problemas))

        with col2:
            criticos = len([p for p in problemas if p['dias_intervalo'] < 15])
            st.metric("Casos Críticos", criticos)

        with col3:
            intervalo_medio = sum([p['dias_intervalo'] for p in problemas]) / len(problemas)
            st.metric("Intervalo Médio", f"{intervalo_medio:.1f} dias")


def show_reports_page(analyzer):
    st.header("📊 Relatórios")

    # Relatório por supervisor
    st.subheader("👨‍💼 Relatório por Supervisor")

    supervisor_stats = analyzer.df.groupby('supervisor').agg({
        'colaborador': 'count',
        'origem': lambda x: x.nunique(),
        'destino': lambda x: x.nunique()
    }).reset_index()

    supervisor_stats.columns = ['Supervisor', 'Colaboradores', 'Origens', 'Destinos']
    st.dataframe(supervisor_stats, use_container_width=True)

    # Gráfico de distribuição
    if not supervisor_stats.empty:
        fig = px.bar(
            supervisor_stats,
            x='Supervisor',
            y='Colaboradores',
            title='Distribuição de Colaboradores por Supervisor',
            color='Colaboradores',
            color_continuous_scale=[[0, '#000000'], [1, '#F7931E']]
        )
        st.plotly_chart(fig, use_container_width=True)

    # Relatório de movimentações por cidade
    st.subheader("🏙️ Movimentações por Cidade")

    col1, col2 = st.columns(2)

    with col1:
        st.write("**Origens mais frequentes:**")
        origens = analyzer.df['origem'].value_counts().head(10)
        if not origens.empty:
            st.bar_chart(origens)

    with col2:
        st.write("**Destinos mais frequentes:**")
        destinos = analyzer.df['destino'].value_counts().head(10)
        if not destinos.empty:
            st.bar_chart(destinos)

    # Exportar dados
    st.subheader("💾 Exportar Dados")

    if st.button("📥 Baixar Relatório Completo (CSV)"):
        csv = analyzer.df.to_csv(index=False)
        st.download_button(
            label="📥 Download CSV",
            data=csv,
            file_name=f"cronograma_equipes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )


def main():
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>🗺️ Sistema de Cronograma de Equipes - Pará</h1>
        <p>Gestão inteligente de folgas e movimentação das equipes</p>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar
    st.sidebar.title("📋 Menu de Navegação")
    page = st.sidebar.selectbox(
        "Selecione a página:",
        ["📋 Cronograma por Encarregado", "🗺️ Mapa das Equipes", "🔍 Auditoria de Folgas", "📊 Relatórios"]
    )

    # Carregar dados
    with st.spinner("Carregando dados do SharePoint..."):
        connector = SharePointConnector()
        df = connector.get_data()

    if df is None:
        st.error("❌ Não foi possível carregar os dados. Verifique a conexão com o SharePoint.")
        st.stop()

    analyzer = CronogramaAnalyzer(df)

    if page == "📋 Cronograma por Encarregado":
        show_cronograma_encarregado(analyzer)
    elif page == "🗺️ Mapa das Equipes":
        show_map_page(analyzer)
    elif page == "🔍 Auditoria de Folgas":
        show_audit_page(analyzer)
    elif page == "📊 Relatórios":
        show_reports_page(analyzer)


if __name__ == "__main__":
    main()
