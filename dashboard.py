import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import requests
from msal import ConfidentialClientApplication
import io

# Configurações da página
st.set_page_config(
    page_title="Controle de Contratação - Rezende Energia",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado com cores da Rezende Energia
st.markdown("""
    <style>
    .main {
        background-color: #FFFFFF;
    }
    .stMetric {
        background-color: #F7931E;
        padding: 15px;
        border-radius: 5px;
        color: #000000;
    }
    .stMetric label {
        color: #000000 !important;
        font-weight: bold;
    }
    .stMetric .metric-value {
        color: #000000 !important;
    }
    h1, h2, h3 {
        color: #000000;
    }
    .alerta-30dias {
        background-color: #FF4444;
        color: white;
        padding: 10px;
        border-radius: 5px;
        font-weight: bold;
        text-align: center;
        margin: 10px 0;
    }
    .sidebar .sidebar-content {
        background-color: #000000;
    }
    </style>
    """, unsafe_allow_html=True)

# Credenciais do Azure AD (carregadas do secrets.toml)
try:
    CLIENT_ID = st.secrets["azure"]["client_id"]
    CLIENT_SECRET = st.secrets["azure"]["client_secret"]
    TENANT_ID = st.secrets["azure"]["tenant_id"]
except KeyError:
    st.error("❌ Credenciais não encontradas. Configure o arquivo .streamlit/secrets.toml")
    st.stop()

# Funções específicas para análise
FUNCOES_ANALISE = [
    "AJUDANTE DE SERVIÇOS GERAIS",
    "ELETRICISTA",
    "O.P DE RETROESCAVADEIRA",
    "OP. DE MOTOSSERA",
    "MOTORISTA OPERADOR DE MUNCK"
]


@st.cache_data(ttl=300)  # Cache por 5 minutos
def carregar_dados_sharepoint():
    """Carrega dados do SharePoint"""
    try:
        # Configurar autenticação
        app = ConfidentialClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            client_credential=CLIENT_SECRET,
        )

        # Obter token
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

        if "access_token" in result:
            headers = {"Authorization": f"Bearer {result['access_token']}"}

            # Obter o site_id
            site_url = "https://graph.microsoft.com/v1.0/sites/rezendeenergia.sharepoint.com:/sites/Intranet"
            site_response = requests.get(site_url, headers=headers)

            if site_response.status_code == 200:
                site_data = site_response.json()
                site_id = site_data['id']

                # Buscar o arquivo
                search_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/search(q='CONTROLE CONTRATAÇÃO')"
                search_response = requests.get(search_url, headers=headers)

                if search_response.status_code == 200:
                    search_data = search_response.json()
                    files_found = search_data.get('value', [])

                    for item in files_found:
                        if 'CONTROLE CONTRATAÇÃO' in item['name'] and (
                                item['name'].endswith('.xlsx') or item['name'].endswith('.xlsb')):
                            # Baixar o arquivo
                            download_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item['id']}/content"
                            download_response = requests.get(download_url, headers=headers)

                            if download_response.status_code == 200:
                                df = pd.read_excel(io.BytesIO(download_response.content))
                                return df

        return None
    except Exception as e:
        st.error(f"Erro ao carregar dados: {str(e)}")
        return None


def processar_dados(df):
    """Processa e limpa os dados"""
    # Renomear colunas para facilitar (assumindo ordem correta)
    colunas_esperadas = [
        'Nome', 'Função', 'Regional', 'Cidade',
        'Data Abertura', 'Data Doc Recebida', 'Data ASO',
        'Data Doc Admissao', 'Data Final NRs', 'Data Inclusao Bernhoeft',
        'Data Aprovacao Bernhoeft', 'Data POP Seguranca', 'Data Integracao Equatorial',
        'Status Atual', 'Liberado Campo'
    ]

    # Renomear colunas baseado na ordem
    if len(df.columns) >= len(colunas_esperadas):
        # Usar apenas as primeiras 15 colunas
        df = df.iloc[:, :15].copy()
        df.columns = colunas_esperadas
    else:
        st.error(f"❌ Número de colunas incorreto. Esperado: {len(colunas_esperadas)}, Encontrado: {len(df.columns)}")
        return df

    # Converter colunas de data
    colunas_data = ['Data Abertura', 'Data Doc Recebida', 'Data ASO', 'Data Doc Admissao',
                    'Data Final NRs', 'Data Inclusao Bernhoeft', 'Data Aprovacao Bernhoeft',
                    'Data POP Seguranca', 'Data Integracao Equatorial']

    for col in colunas_data:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)

    # Filtrar apenas funções de interesse
    if 'Função' in df.columns:
        df = df[df['Função'].isin(FUNCOES_ANALISE)].copy()

    # Remover não contratados
    if 'Status Atual' in df.columns:
        df = df[df['Status Atual'] != 'NÃO CONTRATADO'].copy()
        
    if 'Liberado Campo' in df.columns:
        df = df[df['Liberado Campo'] != 'NÃO CONTRATADO'].copy()
    return df


def calcular_ultima_data(row):
    """Calcula a última data registrada para um colaborador"""
    colunas_data = ['Data Doc Recebida', 'Data ASO', 'Data Doc Admissao',
                    'Data Final NRs', 'Data Inclusao Bernhoeft', 'Data Aprovacao Bernhoeft',
                    'Data POP Seguranca', 'Data Integracao Equatorial']

    datas = [row[col] for col in colunas_data if col in row.index and pd.notna(row[col])]
    return max(datas) if datas else None


def calcular_data_ajustada_nrs(row):
    """Retorna a data mais recente entre NRs e Doc Admissão"""
    if pd.notna(row['Data Final NRs']) and pd.notna(row['Data Doc Admissao']):
        return max(row['Data Final NRs'], row['Data Doc Admissao'])
    elif pd.notna(row['Data Final NRs']):
        return row['Data Final NRs']
    elif pd.notna(row['Data Doc Admissao']):
        return row['Data Doc Admissao']
    return None


def calcular_data_ajustada_integracao(row):
    """Retorna a data mais recente entre Integração e POP"""
    if pd.notna(row['Data Integracao Equatorial']) and pd.notna(row['Data POP Seguranca']):
        return max(row['Data Integracao Equatorial'], row['Data POP Seguranca'])
    elif pd.notna(row['Data Integracao Equatorial']):
        return row['Data Integracao Equatorial']
    elif pd.notna(row['Data POP Seguranca']):
        return row['Data POP Seguranca']
    return None


def calcular_metricas(df):
    """Calcula todas as métricas necessárias"""
    df = df.copy()

    # Última data registrada
    df['Ultima Data'] = df.apply(calcular_ultima_data, axis=1)

    # Tempo total de mobilização
    df['Tempo Total Mobilizacao'] = (df['Ultima Data'] - df['Data Abertura']).dt.days

    # Tempo entre abertura e doc recebida
    df['Tempo Abertura_DocRecebida'] = (df['Data Doc Recebida'] - df['Data Abertura']).dt.days

    # Tempo entre doc admissão e última data
    df['Tempo DocAdmissao_Ultima'] = (df['Ultima Data'] - df['Data Doc Admissao']).dt.days

    # Duração de cada etapa
    df['Tempo DocRecebida_ASO'] = (df['Data ASO'] - df['Data Doc Recebida']).dt.days
    df['Tempo ASO_DocAdmissao'] = (df['Data Doc Admissao'] - df['Data ASO']).dt.days

    # Ajustar NRs - diferença absoluta entre Doc Admissão e NRs
    df['Data Ajustada NRs'] = df.apply(calcular_data_ajustada_nrs, axis=1)
    df['Tempo DocAdmissao_NRs'] = (df['Data Ajustada NRs'] - df['Data Doc Admissao']).dt.days.abs()

    df['Tempo NRs_InclusaoBernhoeft'] = (df['Data Inclusao Bernhoeft'] - df['Data Ajustada NRs']).dt.days
    df['Tempo InclusaoBernhoeft_AprovacaoBernhoeft'] = (
                df['Data Aprovacao Bernhoeft'] - df['Data Inclusao Bernhoeft']).dt.days
    df['Tempo AprovacaoBernhoeft_POP'] = (df['Data POP Seguranca'] - df['Data Aprovacao Bernhoeft']).dt.days

    # Ajustar Integração - diferença absoluta entre POP e Integração
    df['Data Ajustada Integracao'] = df.apply(calcular_data_ajustada_integracao, axis=1)
    df['Tempo POP_Integracao'] = (df['Data Ajustada Integracao'] - df['Data POP Seguranca']).dt.days.abs()

    # Mês de referência
    df['Mes Abertura'] = df['Data Abertura'].dt.to_period('M')

    # Alerta > 30 dias
    df['Alerta_30dias'] = df['Tempo Total Mobilizacao'] > 30

    return df


def criar_grafico_barras(df, coluna_metrica, titulo, cor='#F7931E'):
    """Cria gráfico de barras"""
    fig = px.bar(df, x=df.index, y=coluna_metrica,
                 title=titulo,
                 color_discrete_sequence=[cor],
                 text=coluna_metrica)
    fig.update_traces(texttemplate='%{text:.1f}', textposition='outside')
    fig.update_layout(
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(color='#000000')
    )
    return fig


def criar_grafico_linha_temporal(df):
    """Cria gráfico de linha temporal por mês"""
    dados_mes = df.groupby('Mes Abertura')['Tempo Total Mobilizacao'].mean().reset_index()
    dados_mes['Mes Abertura'] = dados_mes['Mes Abertura'].astype(str)

    fig = px.line(dados_mes, x='Mes Abertura', y='Tempo Total Mobilizacao',
                  title='Média de Tempo Total de Mobilização por Mês',
                  markers=True,
                  color_discrete_sequence=['#F7931E'],
                  text='Tempo Total Mobilizacao')
    fig.update_traces(texttemplate='%{text:.1f}', textposition='top center')
    fig.update_layout(
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(color='#000000'),
        xaxis_title='Mês de Abertura da Vaga',
        yaxis_title='Tempo Médio (dias)'
    )
    return fig


# Interface Principal
def main():
    # Logo e Título
    st.markdown("""
        <h1 style='text-align: center; color: #000000;'>
        📊 Dashboard de Controle de Contratação
        </h1>
        <h3 style='text-align: center; color: #F7931E;'>Rezende Energia</h3>
        <hr style='border: 2px solid #F7931E;'>
    """, unsafe_allow_html=True)

    # Carregar dados
    with st.spinner('🔄 Carregando dados do SharePoint...'):
        df = carregar_dados_sharepoint()

    if df is None:
        st.error("❌ Não foi possível carregar os dados do SharePoint. Verifique as credenciais.")
        return

    # Processar dados
    df = processar_dados(df)
    df = calcular_metricas(df)

    if df.empty:
        st.warning("⚠️ Nenhum dado encontrado para as funções especificadas.")
        return

    # Sidebar - Filtros
    st.sidebar.header("🔍 Filtros")

    regionais = ['Todas'] + sorted(df['Regional'].dropna().unique().tolist())
    regional_selecionada = st.sidebar.selectbox('Regional', regionais)

    if regional_selecionada != 'Todas':
        df_filtrado = df[df['Regional'] == regional_selecionada]
    else:
        df_filtrado = df.copy()

    cidades = ['Todas'] + sorted(df_filtrado['Cidade'].dropna().unique().tolist())
    cidade_selecionada = st.sidebar.selectbox('Cidade', cidades)

    if cidade_selecionada != 'Todas':
        df_filtrado = df_filtrado[df_filtrado['Cidade'] == cidade_selecionada]

    funcoes = ['Todas'] + FUNCOES_ANALISE
    funcao_selecionada = st.sidebar.selectbox('Função', funcoes)

    if funcao_selecionada != 'Todas':
        df_filtrado = df_filtrado[df_filtrado['Função'] == funcao_selecionada]

    # Tabs principais
    tab1, tab2 = st.tabs(["📈 Dashboard Geral", "👥 Detalhamento Individual"])

    with tab1:
        # Alertas de 30 dias
        alertas = df_filtrado[df_filtrado['Alerta_30dias'] == True]
        if not alertas.empty:
            st.markdown(f"""
                <div class='alerta-30dias'>
                ⚠️ ATENÇÃO: {len(alertas)} colaborador(es) com tempo de mobilização superior a 30 dias!
                </div>
            """, unsafe_allow_html=True)

        # KPIs principais
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            media_total = df_filtrado['Tempo Total Mobilizacao'].mean()
            st.metric("📅 Tempo Médio Total", f"{media_total:.1f} dias")

        with col2:
            media_abertura_doc = df_filtrado['Tempo Abertura_DocRecebida'].mean()
            st.metric("📄 Abertura → Doc. Recebida", f"{media_abertura_doc:.1f} dias")

        with col3:
            media_doc_final = df_filtrado['Tempo DocAdmissao_Ultima'].mean()
            st.metric("📋 Doc. Admissão → Final", f"{media_doc_final:.1f} dias")

        with col4:
            total_colaboradores = len(df_filtrado)
            st.metric("👥 Total de Colaboradores", total_colaboradores)

        st.markdown("<hr style='border: 1px solid #F7931E;'>", unsafe_allow_html=True)

        # Gráficos
        col1, col2 = st.columns(2)

        with col1:
            # Média por Regional
            media_regional = df.groupby('Regional')['Tempo Total Mobilizacao'].mean().sort_values(ascending=False)
            fig1 = criar_grafico_barras(media_regional, 'Tempo Total Mobilizacao',
                                        'Tempo Médio de Mobilização por Regional')
            st.plotly_chart(fig1, use_container_width=True)

        with col2:
            # Média por Cidade
            media_cidade = df.groupby('Cidade')['Tempo Total Mobilizacao'].mean().sort_values(ascending=False).head(10)
            fig2 = criar_grafico_barras(media_cidade, 'Tempo Total Mobilizacao',
                                        'Top 10 Cidades - Tempo Médio de Mobilização')
            st.plotly_chart(fig2, use_container_width=True)

        # Gráfico temporal
        st.markdown("### 📊 Evolução Temporal")
        fig_temporal = criar_grafico_linha_temporal(df)
        st.plotly_chart(fig_temporal, use_container_width=True)

        # Tabela de médias por etapa
        st.markdown("### ⏱️ Tempo Médio por Etapa do Processo")

        etapas_media = {
            'Abertura → Doc. Recebida': df_filtrado['Tempo Abertura_DocRecebida'].mean(),
            'Doc. Recebida → ASO': df_filtrado['Tempo DocRecebida_ASO'].mean(),
            'ASO → Doc. Admissão': df_filtrado['Tempo ASO_DocAdmissao'].mean(),
            'Doc. Admissão → NRs': df_filtrado['Tempo DocAdmissao_NRs'].mean(),
            'NRs → Inclusão Bernhoeft': df_filtrado['Tempo NRs_InclusaoBernhoeft'].mean(),
            'Inclusão → Aprovação Bernhoeft': df_filtrado['Tempo InclusaoBernhoeft_AprovacaoBernhoeft'].mean(),
            'Aprovação Bernhoeft → POP': df_filtrado['Tempo AprovacaoBernhoeft_POP'].mean(),
            'POP → Integração': df_filtrado['Tempo POP_Integracao'].mean()
        }

        df_etapas = pd.DataFrame(list(etapas_media.items()), columns=['Etapa', 'Tempo Médio (dias)'])
        df_etapas['Tempo Médio (dias)'] = df_etapas['Tempo Médio (dias)'].round(1)
        st.dataframe(df_etapas, use_container_width=True, hide_index=True)

    with tab2:
        st.markdown("### 👥 Situação Detalhada por Colaborador")

        # Preparar dados para exibição
        colunas_exibir = ['Nome', 'Função', 'Regional', 'Cidade', 'Data Abertura',
                          'Tempo Total Mobilizacao', 'Status Atual', 'Alerta_30dias']

        df_detalhado = df_filtrado[colunas_exibir].copy()
        df_detalhado['Data Abertura'] = df_detalhado['Data Abertura'].dt.strftime('%d/%m/%Y')
        df_detalhado.columns = ['Nome', 'Função', 'Regional', 'Cidade', 'Data Abertura',
                                'Tempo Total (dias)', 'Status', 'Alerta > 30 dias']

        # Destacar alertas
        def highlight_alertas(row):
            if row['Alerta > 30 dias']:
                return ['background-color: #FFE6E6'] * len(row)
            return [''] * len(row)

        st.dataframe(
            df_detalhado.style.apply(highlight_alertas, axis=1),
            use_container_width=True,
            hide_index=True
        )

        # Detalhes individuais
        st.markdown("### 🔍 Detalhes Completos")
        colaborador_selecionado = st.selectbox('Selecione um colaborador:', df_filtrado['Nome'].tolist())

        if colaborador_selecionado:
            dados_colab = df_filtrado[df_filtrado['Nome'] == colaborador_selecionado].iloc[0]

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("#### 📋 Informações Gerais")
                st.write(f"**Nome:** {dados_colab['Nome']}")
                st.write(f"**Função:** {dados_colab['Função']}")
                st.write(f"**Regional:** {dados_colab['Regional']}")
                st.write(f"**Cidade:** {dados_colab['Cidade']}")
                st.write(f"**Status:** {dados_colab['Status Atual']}")

            with col2:
                st.markdown("#### ⏱️ Tempos")
                st.write(f"**Tempo Total:** {dados_colab['Tempo Total Mobilizacao']:.0f} dias")
                st.write(f"**Abertura → Doc. Recebida:** {dados_colab['Tempo Abertura_DocRecebida']:.0f} dias")
                st.write(f"**Doc. Admissão → Final:** {dados_colab['Tempo DocAdmissao_Ultima']:.0f} dias")

                if dados_colab['Alerta_30dias']:
                    st.markdown("⚠️ **ALERTA: Mobilização > 30 dias**")


if __name__ == "__main__":
    main()
