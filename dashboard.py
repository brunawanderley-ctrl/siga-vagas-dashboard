import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import json
import sqlite3
import subprocess
import os
from pathlib import Path
from datetime import datetime

# Configuração da página
st.set_page_config(
    page_title="Vagas Colégio Elo",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

BASE_DIR = Path(__file__).parent

# CSS Corporate SaaS Dark Mode - Navy Blue
st.markdown("""
<style>
    /* Dark theme - Navy blue base */
    .stApp {
        background: linear-gradient(180deg, #0a1628 0%, #0f2137 50%, #132743 100%);
    }

    /* Main container */
    .main .block-container {
        padding: 2rem 3rem;
        max-width: 1400px;
    }

    /* Headers */
    h1, h2, h3 {
        color: #ffffff !important;
        font-weight: 600 !important;
    }

    h1 {
        font-size: 2.5rem !important;
        color: #ffffff !important;
    }

    /* Metric cards */
    [data-testid="stMetric"] {
        background: linear-gradient(145deg, #0d1f35 0%, #142d4c 100%);
        border: 1px solid rgba(59, 130, 246, 0.2);
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.4);
    }

    [data-testid="stMetric"] label {
        color: #94a3b8 !important;
        font-size: 0.85rem !important;
        text-transform: uppercase;
        letter-spacing: 1px;
    }

    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #ffffff !important;
        font-size: 2rem !important;
        font-weight: 700 !important;
    }

    [data-testid="stMetric"] [data-testid="stMetricDelta"] {
        color: #22c55e !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        background: rgba(15, 33, 55, 0.9);
        border: 1px solid rgba(59, 130, 246, 0.15);
        border-radius: 12px;
        padding: 0.5rem;
        gap: 0.5rem;
    }

    .stTabs [data-baseweb="tab"] {
        background: transparent;
        color: #94a3b8;
        border-radius: 8px;
        padding: 0.75rem 1.5rem;
        font-weight: 500;
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #1e4976 0%, #2563eb 100%);
        color: white !important;
    }

    /* Expander */
    .streamlit-expanderHeader {
        background: rgba(15, 33, 55, 0.9) !important;
        border: 1px solid rgba(59, 130, 246, 0.15);
        border-radius: 12px;
        color: #ffffff !important;
    }

    .streamlit-expanderHeader p,
    .streamlit-expanderHeader span,
    [data-testid="stExpander"] summary,
    [data-testid="stExpander"] summary span,
    [data-testid="stExpander"] summary p {
        color: #ffffff !important;
    }

    [data-testid="stExpander"] {
        background: rgba(15, 33, 55, 0.9);
        border: 1px solid rgba(59, 130, 246, 0.15);
        border-radius: 12px;
    }

    [data-testid="stExpander"] details {
        background: rgba(15, 33, 55, 0.9);
        border-radius: 12px;
    }

    /* Dataframe */
    .stDataFrame {
        background: rgba(15, 33, 55, 0.9);
        border-radius: 12px;
    }

    /* Divider */
    hr {
        border-color: rgba(59, 130, 246, 0.15);
    }

    /* Caption */
    .stCaption {
        color: #64748b !important;
    }

    /* Button */
    .stButton > button {
        background: linear-gradient(135deg, #1e4976 0%, #2563eb 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(37, 99, 235, 0.3);
    }

    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.4);
    }

    /* Info box */
    .stAlert {
        background: rgba(37, 99, 235, 0.1);
        border: 1px solid rgba(37, 99, 235, 0.25);
        border-radius: 12px;
    }

    /* Plotly charts */
    .js-plotly-plot {
        border-radius: 16px;
    }

    /* Card styling */
    .premium-card {
        background: linear-gradient(145deg, #0d1f35 0%, #142d4c 100%);
        border: 1px solid rgba(59, 130, 246, 0.2);
        border-radius: 16px;
        padding: 2rem;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.4);
    }
</style>
""", unsafe_allow_html=True)

# Layout do tema para gráficos Plotly - Navy Blue
PLOTLY_LAYOUT = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    font=dict(color='#94a3b8', family='Inter, sans-serif'),
    title=dict(font=dict(color='#ffffff', size=18)),
    xaxis=dict(
        gridcolor='rgba(59, 130, 246, 0.1)',
        linecolor='rgba(59, 130, 246, 0.15)',
        tickfont=dict(color='#94a3b8')
    ),
    yaxis=dict(
        gridcolor='rgba(59, 130, 246, 0.1)',
        linecolor='rgba(59, 130, 246, 0.15)',
        tickfont=dict(color='#94a3b8')
    ),
    legend=dict(
        bgcolor='rgba(0,0,0,0)',
        font=dict(color='#94a3b8')
    ),
    margin=dict(t=60, b=40, l=40, r=40)
)

# Cores Corporate SaaS - Navy Blue
COLORS = {
    'primary': '#2563eb',      # Azul principal
    'secondary': '#1e4976',    # Azul marinho escuro
    'accent': '#3b82f6',       # Azul claro
    'success': '#22c55e',
    'warning': '#f59e0b',
    'danger': '#ef4444',
    'info': '#0ea5e9',
    'muted': '#64748b',
    'gradient': ['#1e4976', '#2563eb', '#3b82f6', '#60a5fa'],
    # Termômetro de ocupação (quanto maior, melhor)
    'hot': '#22c55e',      # 90-100% - Verde intenso (excelente)
    'warm': '#84cc16',     # 80-89% - Verde claro (muito bom)
    'mild': '#fbbf24',     # 70-79% - Amarelo (bom)
    'cool': '#f97316',     # 60-69% - Laranja (atenção)
    'cold': '#ef4444',     # <60% - Vermelho (crítico)
}

def get_ocupacao_color(ocupacao):
    """Retorna cor baseada na ocupação - quanto maior, mais quente/verde"""
    if ocupacao >= 90:
        return COLORS['hot']       # Verde intenso
    elif ocupacao >= 80:
        return COLORS['warm']      # Verde claro
    elif ocupacao >= 70:
        return COLORS['mild']      # Amarelo
    elif ocupacao >= 50:
        return COLORS['cool']      # Laranja
    else:
        return COLORS['cold']      # Vermelho (crítico)

BASE_PATH = Path(__file__).parent / "output"

# Carrega dados atuais
@st.cache_data(ttl=60)
def carregar_dados():
    with open(BASE_PATH / "resumo_ultimo.json") as f:
        resumo = json.load(f)
    with open(BASE_PATH / "vagas_ultimo.json") as f:
        vagas = json.load(f)
    return resumo, vagas

# Carrega histórico do banco
@st.cache_data(ttl=60)
def carregar_historico():
    db_path = BASE_PATH / "vagas.db"
    if not db_path.exists():
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), 0

    conn = sqlite3.connect(db_path)

    query_unidades = """
    SELECT e.data_extracao, v.unidade_codigo, v.unidade_nome,
           SUM(v.vagas) as vagas, SUM(v.matriculados) as matriculados,
           SUM(v.novatos) as novatos, SUM(v.veteranos) as veteranos,
           SUM(v.disponiveis) as disponiveis
    FROM vagas v JOIN 'extrações' e ON v.extracao_id = e.id
    GROUP BY e.id, v.unidade_codigo ORDER BY e.data_extracao
    """
    df_unidades = pd.read_sql_query(query_unidades, conn)

    query_total = """
    SELECT e.data_extracao, SUM(v.vagas) as vagas, SUM(v.matriculados) as matriculados,
           SUM(v.novatos) as novatos, SUM(v.veteranos) as veteranos, SUM(v.disponiveis) as disponiveis
    FROM vagas v JOIN 'extrações' e ON v.extracao_id = e.id
    GROUP BY e.id ORDER BY e.data_extracao
    """
    df_total = pd.read_sql_query(query_total, conn)

    query_segmento = """
    SELECT e.data_extracao, v.segmento, SUM(v.vagas) as vagas,
           SUM(v.matriculados) as matriculados, SUM(v.disponiveis) as disponiveis
    FROM vagas v JOIN 'extrações' e ON v.extracao_id = e.id
    GROUP BY e.id, v.segmento ORDER BY e.data_extracao
    """
    df_segmento = pd.read_sql_query(query_segmento, conn)

    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM 'extrações'")
    num_extracoes = cursor.fetchone()[0]
    conn.close()

    for df in [df_unidades, df_total, df_segmento]:
        if not df.empty:
            df['data_extracao'] = pd.to_datetime(df['data_extracao'])
            df['data_formatada'] = df['data_extracao'].dt.strftime('%d/%m %H:%M')

    return df_unidades, df_total, df_segmento, num_extracoes

try:
    resumo, vagas = carregar_dados()
    df_hist_unidades, df_hist_total, df_hist_segmento, num_extracoes = carregar_historico()
except FileNotFoundError:
    st.error("Arquivos de dados não encontrados. Execute a extração primeiro.")
    st.stop()

# Header Premium
col_title, col_btn = st.columns([5, 1])

with col_title:
    st.markdown("""
        <h1 style='margin-bottom: 0;'>Dashboard de Vagas</h1>
        <p style='color: #3b82f6; font-size: 1.2rem; margin-top: 0.5rem;'>Colégio Elo • Visão Executiva</p>
    """, unsafe_allow_html=True)

with col_btn:
    st.write("")
    if st.button("🔄 Atualizar", use_container_width=True):
        with st.spinner("Extraindo dados..."):
            try:
                result = subprocess.run(
                    ["bash", str(BASE_DIR / "cron_extrator.sh")],
                    capture_output=True, text=True, timeout=600, cwd=str(BASE_DIR)
                )
                if result.returncode == 0:
                    st.success("✅ Dados atualizados!")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error(f"❌ Erro: {result.stderr}")
            except Exception as e:
                st.error(f"❌ Erro: {str(e)}")

# Info bar
st.markdown(f"""
    <div style='display: flex; gap: 2rem; color: #64748b; font-size: 0.85rem; margin-bottom: 2rem;'>
        <span>📅 Última atualização: <strong style='color: #94a3b8;'>{resumo['data_extracao'][:16].replace('T', ' ')}</strong></span>
        <span>📊 Período: <strong style='color: #94a3b8;'>{resumo['periodo']}</strong></span>
        <span>🔢 Extrações: <strong style='color: #94a3b8;'>{num_extracoes}</strong></span>
    </div>
""", unsafe_allow_html=True)

# Métricas principais
total = resumo['total_geral']
ocupacao = round(total['matriculados'] / total['vagas'] * 100, 1)

col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    st.metric("OCUPAÇÃO", f"{ocupacao}%", delta=None)
with col2:
    st.metric("MATRICULADOS", f"{total['matriculados']:,}".replace(",", "."))
with col3:
    st.metric("VAGAS TOTAIS", f"{total['vagas']:,}".replace(",", "."))
with col4:
    st.metric("DISPONÍVEIS", f"{total['disponiveis']:,}".replace(",", "."))
with col5:
    st.metric("NOVATOS", f"{total['novatos']:,}".replace(",", "."))
with col6:
    st.metric("VETERANOS", f"{total['veteranos']:,}".replace(",", "."))

st.markdown("<br>", unsafe_allow_html=True)

# Gráficos principais
col_left, col_right = st.columns(2)

with col_left:
    st.markdown("### 🌡️ Ocupação por Unidade")
    st.markdown("""
        <div style='display: flex; gap: 1rem; font-size: 0.75rem; margin-bottom: 0.5rem;'>
            <span style='color: #22c55e;'>🔥 90-100% Excelente</span>
            <span style='color: #84cc16;'>✨ 80-89% Muito Bom</span>
            <span style='color: #fbbf24;'>⚡ 70-79% Bom</span>
            <span style='color: #f97316;'>⚠️ 50-69% Atenção</span>
            <span style='color: #ef4444;'>❄️ &lt;50% Crítico</span>
        </div>
    """, unsafe_allow_html=True)

    df_unidades = pd.DataFrame([
        {
            'Unidade': u['nome'].split('(')[1].replace(')', '') if '(' in u['nome'] else u['nome'],
            'Ocupação': round(u['total']['matriculados'] / u['total']['vagas'] * 100, 1),
            'Matriculados': u['total']['matriculados'],
            'Vagas': u['total']['vagas']
        }
        for u in resumo['unidades']
    ])

    fig1 = go.Figure()

    # Barra de fundo (vagas totais)
    fig1.add_trace(go.Bar(
        name='Capacidade',
        x=df_unidades['Unidade'],
        y=[100] * len(df_unidades),
        marker_color='rgba(59, 130, 246, 0.12)',
        hoverinfo='skip'
    ))

    # Barra de ocupação - quanto maior, mais quente (verde)
    colors = [get_ocupacao_color(o) for o in df_unidades['Ocupação']]

    fig1.add_trace(go.Bar(
        name='Ocupação',
        x=df_unidades['Unidade'],
        y=df_unidades['Ocupação'],
        marker_color=colors,
        text=df_unidades['Ocupação'].apply(lambda x: f'{x}%'),
        textposition='outside',
        textfont=dict(color='#ffffff', size=14, family='Inter')
    ))

    fig1.update_layout(
        paper_bgcolor=PLOTLY_LAYOUT['paper_bgcolor'],
        plot_bgcolor=PLOTLY_LAYOUT['plot_bgcolor'],
        font=PLOTLY_LAYOUT['font'],
        margin=PLOTLY_LAYOUT['margin'],
        barmode='overlay',
        showlegend=False,
        height=350,
        yaxis=dict(**PLOTLY_LAYOUT['yaxis'], range=[0, 110], title=''),
        xaxis=dict(**PLOTLY_LAYOUT['xaxis'], title='')
    )

    st.plotly_chart(fig1, use_container_width=True)

with col_right:
    st.markdown("### Distribuição por Segmento")

    segmentos_total = {}
    for unidade in resumo['unidades']:
        for seg, vals in unidade['segmentos'].items():
            if seg not in segmentos_total:
                segmentos_total[seg] = {'matriculados': 0, 'vagas': 0}
            segmentos_total[seg]['matriculados'] += vals['matriculados']
            segmentos_total[seg]['vagas'] += vals['vagas']

    df_seg = pd.DataFrame([
        {'Segmento': seg, 'Matriculados': v['matriculados'], 'Vagas': v['vagas']}
        for seg, v in segmentos_total.items()
    ])

    ordem = ['Ed. Infantil', 'Fund. I', 'Fund. II', 'Ens. Médio']
    df_seg['ordem'] = df_seg['Segmento'].map({s: i for i, s in enumerate(ordem)})
    df_seg = df_seg.sort_values('ordem')

    fig2 = go.Figure()

    fig2.add_trace(go.Bar(
        name='Vagas',
        x=df_seg['Segmento'],
        y=df_seg['Vagas'],
        marker_color='rgba(59, 130, 246, 0.25)',
        text=df_seg['Vagas'],
        textposition='outside',
        textfont=dict(color='#3b82f6')
    ))

    fig2.add_trace(go.Bar(
        name='Matriculados',
        x=df_seg['Segmento'],
        y=df_seg['Matriculados'],
        marker=dict(
            color=df_seg['Matriculados'],
            colorscale=[[0, '#1e4976'], [1, '#2563eb']]
        ),
        text=df_seg['Matriculados'],
        textposition='outside',
        textfont=dict(color='#ffffff')
    ))

    fig2.update_layout(
        paper_bgcolor=PLOTLY_LAYOUT['paper_bgcolor'],
        plot_bgcolor=PLOTLY_LAYOUT['plot_bgcolor'],
        font=PLOTLY_LAYOUT['font'],
        margin=PLOTLY_LAYOUT['margin'],
        xaxis=PLOTLY_LAYOUT['xaxis'],
        yaxis=PLOTLY_LAYOUT['yaxis'],
        barmode='group',
        height=350,
        legend=dict(
            orientation='h',
            yanchor='bottom',
            y=1.02,
            xanchor='right',
            x=1,
            bgcolor='rgba(0,0,0,0)',
            font=dict(color='#94a3b8')
        )
    )

    st.plotly_chart(fig2, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# Seção de histórico
if num_extracoes >= 2:
    st.markdown("### 📈 Evolução Histórica")

    # Filtro de período
    col_periodo1, col_periodo2, col_periodo3 = st.columns([1, 1, 2])

    with col_periodo1:
        if not df_hist_total.empty:
            data_min = df_hist_total['data_extracao'].min().date()
            data_max = df_hist_total['data_extracao'].max().date()
            data_inicio = st.date_input("Data Início", value=data_min, min_value=data_min, max_value=data_max, key="hist_inicio")

    with col_periodo2:
        if not df_hist_total.empty:
            data_fim = st.date_input("Data Fim", value=data_max, min_value=data_min, max_value=data_max, key="hist_fim")

    with col_periodo3:
        st.markdown(f"<p style='color: #64748b; margin-top: 2rem;'>📅 Período: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}</p>", unsafe_allow_html=True)

    # Filtra dados pelo período selecionado
    df_hist_total_filtrado = df_hist_total[
        (df_hist_total['data_extracao'].dt.date >= data_inicio) &
        (df_hist_total['data_extracao'].dt.date <= data_fim)
    ]
    df_hist_unidades_filtrado = df_hist_unidades[
        (df_hist_unidades['data_extracao'].dt.date >= data_inicio) &
        (df_hist_unidades['data_extracao'].dt.date <= data_fim)
    ]

    tab1, tab2 = st.tabs(["Visão Geral", "Por Unidade"])

    with tab1:
        df_hist_total_filtrado['ocupacao'] = round(df_hist_total_filtrado['matriculados'] / df_hist_total_filtrado['vagas'] * 100, 1)

        fig_hist = go.Figure()

        fig_hist.add_trace(go.Scatter(
            x=df_hist_total_filtrado['data_formatada'],
            y=df_hist_total_filtrado['ocupacao'],
            mode='lines+markers',
            name='Ocupação',
            line=dict(color=COLORS['primary'], width=3),
            marker=dict(size=10, color=COLORS['primary']),
            fill='tozeroy',
            fillcolor='rgba(37, 99, 235, 0.1)'
        ))

        fig_hist.update_layout(
            paper_bgcolor=PLOTLY_LAYOUT['paper_bgcolor'],
            plot_bgcolor=PLOTLY_LAYOUT['plot_bgcolor'],
            font=PLOTLY_LAYOUT['font'],
            margin=PLOTLY_LAYOUT['margin'],
            xaxis=PLOTLY_LAYOUT['xaxis'],
            height=300,
            yaxis=dict(**PLOTLY_LAYOUT['yaxis'], title='Ocupação %', range=[0, 100])
        )

        st.plotly_chart(fig_hist, use_container_width=True)

    with tab2:
        fig_unid = go.Figure()
        cores_unid = [COLORS['primary'], COLORS['accent'], COLORS['info'], '#60a5fa']

        for i, unidade in enumerate(df_hist_unidades_filtrado['unidade_nome'].unique()):
            df_u = df_hist_unidades_filtrado[df_hist_unidades_filtrado['unidade_nome'] == unidade]
            nome = unidade.split('(')[1].replace(')', '') if '(' in unidade else unidade

            fig_unid.add_trace(go.Scatter(
                x=df_u['data_formatada'],
                y=df_u['matriculados'],
                mode='lines+markers',
                name=nome,
                line=dict(color=cores_unid[i % len(cores_unid)], width=2),
                marker=dict(size=8)
            ))

        fig_unid.update_layout(**PLOTLY_LAYOUT, height=300, hovermode='x unified')
        st.plotly_chart(fig_unid, use_container_width=True)

st.markdown("<br>", unsafe_allow_html=True)

# Filtro por Segmento (todas unidades)
st.markdown("### 🎯 Filtro por Segmento")

col_seg_filter, col_search = st.columns([1, 2])

with col_seg_filter:
    segmentos_disponiveis = ['Ed. Infantil', 'Fund. I', 'Fund. II', 'Ens. Médio']
    segmento_filtro = st.selectbox("Selecione o Segmento", segmentos_disponiveis, key="filtro_segmento")

with col_search:
    busca_turma = st.text_input("🔍 Buscar Turma", placeholder="Digite o nome da turma...", key="busca_turma")

# Dados do segmento selecionado em todas as unidades
dados_segmento_todas = []
for unidade in resumo['unidades']:
    nome_unidade = unidade['nome'].split('(')[1].replace(')', '') if '(' in unidade['nome'] else unidade['nome']
    if segmento_filtro in unidade['segmentos']:
        seg_data = unidade['segmentos'][segmento_filtro]
        disponíveis = seg_data['vagas'] - seg_data['matriculados']
        ocup = round(seg_data['matriculados'] / seg_data['vagas'] * 100, 1) if seg_data['vagas'] > 0 else 0

        if ocup >= 90: status = '🔥 Excelente'
        elif ocup >= 80: status = '✨ Muito Bom'
        elif ocup >= 70: status = '⚡ Bom'
        elif ocup >= 50: status = '⚠️ Atenção'
        else: status = '❄️ Crítico'

        # Calcula pré-matriculados
        unidade_vagas_data = next((u for u in vagas['unidades'] if u['codigo'] == unidade['codigo']), None)
        pre_matr = sum(t['pre_matriculados'] for t in unidade_vagas_data['turmas'] if t['segmento'] == segmento_filtro) if unidade_vagas_data else 0

        dados_segmento_todas.append({
            'Unidade': nome_unidade,
            'Vagas': seg_data['vagas'],
            'Novatos': seg_data['novatos'],
            'Veteranos': seg_data['veteranos'],
            'Matriculados': seg_data['matriculados'],
            'Disponíveis': disponíveis,
            'Ocupação %': ocup,
            'Status': status,
            'Pré-Matr.': pre_matr
        })

df_seg_todas = pd.DataFrame(dados_segmento_todas)

# Estilização
def barra_ocup_todas(val):
    if val >= 90: cor = '#22c55e'
    elif val >= 80: cor = '#84cc16'
    elif val >= 70: cor = '#fbbf24'
    elif val >= 50: cor = '#f97316'
    else: cor = '#ef4444'
    return f'background: linear-gradient(90deg, {cor} {val}%, transparent {val}%); color: white; font-weight: bold;'

def colorir_status_todas(val):
    base = 'font-weight: 600; font-family: "SF Pro Display", system-ui, sans-serif; letter-spacing: 0.5px; text-transform: uppercase; font-size: 11px;'
    if 'Excelente' in val: return f'{base} color: #22c55e;'
    elif 'Muito Bom' in val: return f'{base} color: #84cc16;'
    elif 'Bom' in val: return f'{base} color: #fbbf24;'
    elif 'Atenção' in val: return f'{base} color: #f97316;'
    else: return f'{base} color: #ef4444;'

st.markdown(f"**{segmento_filtro}** em todas as unidades:")
styled_seg_todas = df_seg_todas.style.map(barra_ocup_todas, subset=['Ocupação %']).map(colorir_status_todas, subset=['Status'])
st.dataframe(styled_seg_todas, use_container_width=True, hide_index=True)

# Busca de turmas
if busca_turma:
    st.markdown(f"### 🔍 Resultados para: *{busca_turma}*")
    turmas_encontradas = []
    for unidade in vagas['unidades']:
        nome_unidade = unidade['nome'].split('(')[1].replace(')', '') if '(' in unidade['nome'] else unidade['nome']
        for turma in unidade['turmas']:
            if busca_turma.lower() in turma['turma'].lower():
                disponíveis = turma['vagas'] - turma['matriculados']
                ocup = round(turma['matriculados'] / turma['vagas'] * 100, 1) if turma['vagas'] > 0 else 0

                if ocup >= 90: status = '🔥 Excelente'
                elif ocup >= 80: status = '✨ Muito Bom'
                elif ocup >= 70: status = '⚡ Bom'
                elif ocup >= 50: status = '⚠️ Atenção'
                else: status = '❄️ Crítico'

                turmas_encontradas.append({
                    'Unidade': nome_unidade,
                    'Segmento': turma['segmento'],
                    'Turma': turma['turma'],
                    'Vagas': turma['vagas'],
                    'Matriculados': turma['matriculados'],
                    'Disponíveis': disponíveis,
                    'Ocupação %': ocup,
                    'Status': status
                })

    if turmas_encontradas:
        df_busca = pd.DataFrame(turmas_encontradas)
        styled_busca = df_busca.style.map(barra_ocup_todas, subset=['Ocupação %']).map(colorir_status_todas, subset=['Status'])
        st.dataframe(styled_busca, use_container_width=True, hide_index=True)
    else:
        st.info(f"Nenhuma turma encontrada com '{busca_turma}'")

st.markdown("<br>", unsafe_allow_html=True)

# Análise por Segmento (por unidade)
st.markdown("### 📊 Análise por Unidade")

col_select, col_empty = st.columns([1, 2])
with col_select:
    unidades_nomes = [u['nome'].split('(')[1].replace(')', '') if '(' in u['nome'] else u['nome'] for u in resumo['unidades']]
    unidade_selecionada = st.selectbox("Selecione a Unidade", unidades_nomes, key="seg_unidade")

# Encontra índice da unidade selecionada
idx_unidade = unidades_nomes.index(unidade_selecionada)
unidade_dados = resumo['unidades'][idx_unidade]
unidade_vagas_seg = vagas['unidades'][idx_unidade]

# Monta dados por segmento
segmentos_data = []
for seg, vals in unidade_dados['segmentos'].items():
    # Calcula pré-matriculados do segmento
    pre_matr_seg = sum(t['pre_matriculados'] for t in unidade_vagas_seg['turmas'] if t['segmento'] == seg)
    disponíveis = vals['vagas'] - vals['matriculados']
    ocup_seg = round(vals['matriculados'] / vals['vagas'] * 100, 1) if vals['vagas'] > 0 else 0

    if ocup_seg >= 90: status = '🔥 Excelente'
    elif ocup_seg >= 80: status = '✨ Muito Bom'
    elif ocup_seg >= 70: status = '⚡ Bom'
    elif ocup_seg >= 50: status = '⚠️ Atenção'
    else: status = '❄️ Crítico'

    segmentos_data.append({
        'Segmento': seg,
        'Vagas': vals['vagas'],
        'Novatos': vals['novatos'],
        'Veteranos': vals['veteranos'],
        'Matriculados': vals['matriculados'],
        'Disponíveis': disponíveis,
        'Ocupação %': ocup_seg,
        'Status': status,
        'Pré-Matr.': pre_matr_seg
    })

df_segmentos = pd.DataFrame(segmentos_data)

# Ordena por segmento
ordem_seg = {'Ed. Infantil': 0, 'Fund. I': 1, 'Fund. II': 2, 'Ens. Médio': 3}
df_segmentos['ordem'] = df_segmentos['Segmento'].map(ordem_seg)
df_segmentos = df_segmentos.sort_values('ordem').drop('ordem', axis=1)

# Estilização
def barra_ocupacao_seg(val):
    if val >= 90: cor = '#22c55e'
    elif val >= 80: cor = '#84cc16'
    elif val >= 70: cor = '#fbbf24'
    elif val >= 50: cor = '#f97316'
    else: cor = '#ef4444'
    return f'background: linear-gradient(90deg, {cor} {val}%, transparent {val}%); color: white; font-weight: bold;'

def colorir_status_seg(val):
    base_style = 'font-weight: 600; font-family: "SF Pro Display", "Segoe UI", system-ui, sans-serif; letter-spacing: 0.5px; text-transform: uppercase; font-size: 11px;'
    if 'Excelente' in val: return f'{base_style} color: #22c55e;'
    elif 'Muito Bom' in val: return f'{base_style} color: #84cc16;'
    elif 'Bom' in val: return f'{base_style} color: #fbbf24;'
    elif 'Atenção' in val: return f'{base_style} color: #f97316;'
    else: return f'{base_style} color: #ef4444;'

styled_seg = df_segmentos.style.map(barra_ocupacao_seg, subset=['Ocupação %']).map(colorir_status_seg, subset=['Status'])

col_tabela, col_grafico = st.columns([2, 1])

with col_tabela:
    st.dataframe(styled_seg, use_container_width=True, hide_index=True)

with col_grafico:
    # Gráfico de Status
    status_counts = df_segmentos['Status'].value_counts()
    cores_status = {
        '🔥 Excelente': '#22c55e',
        '✨ Muito Bom': '#84cc16',
        '⚡ Bom': '#fbbf24',
        '⚠️ Atenção': '#f97316',
        '❄️ Crítico': '#ef4444'
    }

    fig_status = go.Figure(data=[go.Pie(
        labels=status_counts.index,
        values=status_counts.values,
        hole=0.5,
        marker_colors=[cores_status.get(s, '#94a3b8') for s in status_counts.index],
        textinfo='label+value',
        textfont=dict(size=12, color='white')
    )])

    fig_status.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#94a3b8'),
        showlegend=False,
        height=250,
        margin=dict(t=30, b=10, l=10, r=10),
        title=dict(text='Status por Segmento', font=dict(color='#ffffff', size=14))
    )

    st.plotly_chart(fig_status, use_container_width=True)

# Legenda de Status destacada
st.markdown("""
    <div style='display: flex; justify-content: center; gap: 1.5rem; padding: 1rem; background: rgba(15, 33, 55, 0.5); border-radius: 12px; margin: 1rem 0;'>
        <span style='color: #22c55e; font-weight: 600;'>🔥 EXCELENTE (90-100%)</span>
        <span style='color: #84cc16; font-weight: 600;'>✨ MUITO BOM (80-89%)</span>
        <span style='color: #fbbf24; font-weight: 600;'>⚡ BOM (70-79%)</span>
        <span style='color: #f97316; font-weight: 600;'>⚠️ ATENÇÃO (50-69%)</span>
        <span style='color: #ef4444; font-weight: 600;'>❄️ CRÍTICO (&lt;50%)</span>
    </div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Detalhamento por unidade
st.markdown("### 🏫 Detalhamento por Unidade")

tabs = st.tabs([u['nome'].split('(')[1].replace(')', '') if '(' in u['nome'] else u['nome']
                for u in resumo['unidades']])

for i, tab in enumerate(tabs):
    with tab:
        unidade = resumo['unidades'][i]
        unidade_vagas = vagas['unidades'][i]

        t = unidade['total']
        ocup = round(t['matriculados'] / t['vagas'] * 100, 1)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Ocupação", f"{ocup}%")
        c2.metric("Matriculados", t['matriculados'])
        c3.metric("Disponíveis", t['disponiveis'])
        c4.metric("Novatos / Veteranos", f"{t['novatos']} / {t['veteranos']}")

        col_a, col_b = st.columns([2, 1])

        with col_a:
            df_seg_u = pd.DataFrame([
                {'Segmento': seg, **vals}
                for seg, vals in unidade['segmentos'].items()
            ])

            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=df_seg_u['Segmento'],
                y=df_seg_u['vagas'],
                name='Vagas',
                marker_color='rgba(59, 130, 246, 0.25)'
            ))
            fig.add_trace(go.Bar(
                x=df_seg_u['Segmento'],
                y=df_seg_u['matriculados'],
                name='Matriculados',
                marker_color=COLORS['primary']
            ))
            fig.update_layout(**PLOTLY_LAYOUT, height=280, barmode='group')
            st.plotly_chart(fig, use_container_width=True)

        with col_b:
            fig_pie = go.Figure(data=[go.Pie(
                labels=['Novatos', 'Veteranos'],
                values=[t['novatos'], t['veteranos']],
                hole=.6,
                marker_colors=[COLORS['info'], COLORS['primary']]
            )])
            fig_pie.update_layout(
                paper_bgcolor=PLOTLY_LAYOUT['paper_bgcolor'],
                plot_bgcolor=PLOTLY_LAYOUT['plot_bgcolor'],
                font=PLOTLY_LAYOUT['font'],
                margin=PLOTLY_LAYOUT['margin'],
                height=280,
                showlegend=True,
                legend=dict(orientation='h', yanchor='bottom', y=-0.2, bgcolor='rgba(0,0,0,0)', font=dict(color='#94a3b8'))
            )
            st.plotly_chart(fig_pie, use_container_width=True)

        with st.expander("📋 Ver todas as turmas"):
            df_turmas = pd.DataFrame(unidade_vagas['turmas'])
            df_turmas = df_turmas[['segmento', 'turma', 'vagas', 'novatos', 'veteranos', 'matriculados', 'pre_matriculados']]
            # Calcula disponíveis corretamente
            df_turmas['disponiveis'] = df_turmas['vagas'] - df_turmas['matriculados']
            df_turmas['ocupacao'] = round(df_turmas['matriculados'] / df_turmas['vagas'] * 100, 1)

            # Status baseado na ocupação
            def get_status(ocup):
                if ocup >= 90: return '🔥 Excelente'
                elif ocup >= 80: return '✨ Muito Bom'
                elif ocup >= 70: return '⚡ Bom'
                elif ocup >= 50: return '⚠️ Atenção'
                else: return '❄️ Crítico'

            df_turmas['status'] = df_turmas['ocupacao'].apply(get_status)
            df_turmas.columns = ['Segmento', 'Turma', 'Vagas', 'Novatos', 'Veteranos', 'Matriculados', 'Pré-Matr.', 'Disponíveis', 'Ocupação %', 'Status']

            # Reordenar colunas
            df_turmas = df_turmas[['Segmento', 'Turma', 'Vagas', 'Novatos', 'Veteranos', 'Matriculados', 'Disponíveis', 'Ocupação %', 'Status', 'Pré-Matr.']]

            # Função para criar barra de ocupação
            def barra_ocupacao(val):
                if val >= 90:
                    cor = '#22c55e'
                elif val >= 80:
                    cor = '#84cc16'
                elif val >= 70:
                    cor = '#fbbf24'
                elif val >= 50:
                    cor = '#f97316'
                else:
                    cor = '#ef4444'
                return f'background: linear-gradient(90deg, {cor} {val}%, transparent {val}%); color: white; font-weight: bold;'

            # Função para colorir status
            def colorir_status(val):
                base_style = 'font-weight: 600; font-family: "SF Pro Display", "Segoe UI", system-ui, sans-serif; letter-spacing: 0.5px; text-transform: uppercase; font-size: 11px;'
                if 'Excelente' in val:
                    return f'{base_style} color: #22c55e;'
                elif 'Muito Bom' in val:
                    return f'{base_style} color: #84cc16;'
                elif 'Bom' in val:
                    return f'{base_style} color: #fbbf24;'
                elif 'Atenção' in val:
                    return f'{base_style} color: #f97316;'
                else:
                    return f'{base_style} color: #ef4444;'

            styled_df = df_turmas.style.map(barra_ocupacao, subset=['Ocupação %']).map(colorir_status, subset=['Status'])
            st.dataframe(styled_df, use_container_width=True, hide_index=True)

# Footer
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown(f"""
    <div style='text-align: center; color: #64748b; font-size: 0.8rem; padding: 2rem 0;'>
        <p>Dashboard atualizado automaticamente às 6h • Última extração: {resumo['data_extracao'][:16].replace('T', ' ')}</p>
        <p style='color: #475569;'>Colégio Elo © 2026</p>
    </div>
""", unsafe_allow_html=True)
