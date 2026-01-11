import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import json
import sqlite3
from pathlib import Path
from datetime import datetime

# Configuração da página
st.set_page_config(
    page_title="Vagas Colégio Elo",
    page_icon="🎓",
    layout="wide"
)

# CSS customizado
st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .stMetric {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
    }
</style>
""", unsafe_allow_html=True)

BASE_PATH = Path(__file__).parent / "output"

# Carrega dados atuais
@st.cache_data(ttl=300)
def carregar_dados():
    with open(BASE_PATH / "resumo_ultimo.json") as f:
        resumo = json.load(f)
    with open(BASE_PATH / "vagas_ultimo.json") as f:
        vagas = json.load(f)
    return resumo, vagas

# Carrega histórico do banco
@st.cache_data(ttl=300)
def carregar_historico():
    db_path = BASE_PATH / "vagas.db"
    if not db_path.exists():
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), 0

    conn = sqlite3.connect(db_path)

    # Query para histórico agregado por extração e unidade
    query_unidades = """
    SELECT
        e.data_extracao,
        v.unidade_codigo,
        v.unidade_nome,
        SUM(v.vagas) as vagas,
        SUM(v.matriculados) as matriculados,
        SUM(v.novatos) as novatos,
        SUM(v.veteranos) as veteranos,
        SUM(v.disponiveis) as disponiveis
    FROM vagas v
    JOIN 'extrações' e ON v.extracao_id = e.id
    GROUP BY e.id, v.unidade_codigo
    ORDER BY e.data_extracao
    """
    df_unidades = pd.read_sql_query(query_unidades, conn)

    # Query para histórico total
    query_total = """
    SELECT
        e.data_extracao,
        SUM(v.vagas) as vagas,
        SUM(v.matriculados) as matriculados,
        SUM(v.novatos) as novatos,
        SUM(v.veteranos) as veteranos,
        SUM(v.disponiveis) as disponiveis
    FROM vagas v
    JOIN 'extrações' e ON v.extracao_id = e.id
    GROUP BY e.id
    ORDER BY e.data_extracao
    """
    df_total = pd.read_sql_query(query_total, conn)

    # Query para histórico por segmento
    query_segmento = """
    SELECT
        e.data_extracao,
        v.segmento,
        SUM(v.vagas) as vagas,
        SUM(v.matriculados) as matriculados,
        SUM(v.disponiveis) as disponiveis
    FROM vagas v
    JOIN 'extrações' e ON v.extracao_id = e.id
    GROUP BY e.id, v.segmento
    ORDER BY e.data_extracao
    """
    df_segmento = pd.read_sql_query(query_segmento, conn)

    # Número de extrações
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM 'extrações'")
    num_extracoes = cursor.fetchone()[0]

    conn.close()

    # Converte datas
    for df in [df_unidades, df_total, df_segmento]:
        if not df.empty:
            df['data_extracao'] = pd.to_datetime(df['data_extracao'])
            df['data_formatada'] = df['data_extracao'].dt.strftime('%d/%m %H:%M')

    return df_unidades, df_total, df_segmento, num_extracoes

try:
    resumo, vagas = carregar_dados()
    df_hist_unidades, df_hist_total, df_hist_segmento, num_extracoes = carregar_historico()
except FileNotFoundError:
    st.error("Arquivos de dados não encontrados.")
    st.stop()

# Título
st.title("🎓 Dashboard de Vagas - Colégio Elo")
st.caption(f"Última atualização: {resumo['data_extracao'][:16].replace('T', ' ')} | Período: {resumo['periodo']} | Extrações: {num_extracoes}")

st.divider()

# Métricas gerais
col1, col2, col3, col4, col5 = st.columns(5)

total = resumo['total_geral']
ocupacao = round(total['matriculados'] / total['vagas'] * 100, 1)

col1.metric("Total de Vagas", f"{total['vagas']:,}".replace(",", "."))
col2.metric("Matriculados", f"{total['matriculados']:,}".replace(",", "."))
col3.metric("Disponíveis", f"{total['disponiveis']:,}".replace(",", "."))
col4.metric("Novatos", f"{total['novatos']:,}".replace(",", "."))
col5.metric("Ocupação", f"{ocupacao}%")

st.divider()

# Gráficos principais
col_left, col_right = st.columns(2)

with col_left:
    st.subheader("📊 Vagas x Matriculados por Unidade")

    df_unidades = pd.DataFrame([
        {
            'Unidade': u['nome'].split('(')[1].replace(')', '') if '(' in u['nome'] else u['nome'],
            'Vagas': u['total']['vagas'],
            'Matriculados': u['total']['matriculados'],
            'Disponíveis': u['total']['disponiveis']
        }
        for u in resumo['unidades']
    ])

    fig1 = go.Figure()
    fig1.add_trace(go.Bar(name='Vagas', x=df_unidades['Unidade'], y=df_unidades['Vagas'], marker_color='#4472C4'))
    fig1.add_trace(go.Bar(name='Matriculados', x=df_unidades['Unidade'], y=df_unidades['Matriculados'], marker_color='#70AD47'))
    fig1.update_layout(barmode='group', height=400)
    st.plotly_chart(fig1, use_container_width=True)

with col_right:
    st.subheader("📈 Taxa de Ocupação por Unidade")

    df_ocupacao = pd.DataFrame([
        {
            'Unidade': u['nome'].split('(')[1].replace(')', '') if '(' in u['nome'] else u['nome'],
            'Ocupação': round(u['total']['matriculados'] / u['total']['vagas'] * 100, 1)
        }
        for u in resumo['unidades']
    ])

    fig2 = px.bar(df_ocupacao, x='Unidade', y='Ocupação', text='Ocupação',
                  color='Ocupação', color_continuous_scale=['#ff6b6b', '#ffd93d', '#6bcb77'])
    fig2.update_traces(texttemplate='%{text}%', textposition='outside')
    fig2.update_layout(height=400, showlegend=False)
    st.plotly_chart(fig2, use_container_width=True)

st.divider()

# SEÇÃO DE HISTÓRICO
st.subheader("📈 Histórico de Evolução")

if num_extracoes < 2:
    st.info("⏳ O histórico será exibido após 2 ou mais extrações.")
else:
    tab_hist1, tab_hist2, tab_hist3 = st.tabs(["📊 Total Geral", "🏫 Por Unidade", "📚 Por Segmento"])

    with tab_hist1:
        col_h1, col_h2 = st.columns(2)

        with col_h1:
            fig_hist = go.Figure()
            fig_hist.add_trace(go.Scatter(
                x=df_hist_total['data_formatada'],
                y=df_hist_total['matriculados'],
                mode='lines+markers',
                name='Matriculados',
                line=dict(color='#70AD47', width=3),
                marker=dict(size=8)
            ))
            fig_hist.add_trace(go.Scatter(
                x=df_hist_total['data_formatada'],
                y=df_hist_total['vagas'],
                mode='lines',
                name='Vagas (capacidade)',
                line=dict(color='#4472C4', width=2, dash='dash')
            ))
            fig_hist.update_layout(
                title='Evolução de Matriculados',
                xaxis_title='Data/Hora',
                yaxis_title='Quantidade',
                height=400,
                hovermode='x unified'
            )
            st.plotly_chart(fig_hist, use_container_width=True)

        with col_h2:
            df_hist_total['ocupacao'] = round(df_hist_total['matriculados'] / df_hist_total['vagas'] * 100, 1)

            fig_ocup = go.Figure()
            fig_ocup.add_trace(go.Scatter(
                x=df_hist_total['data_formatada'],
                y=df_hist_total['ocupacao'],
                mode='lines+markers+text',
                name='Ocupação %',
                line=dict(color='#ED7D31', width=3),
                marker=dict(size=8),
                text=df_hist_total['ocupacao'].apply(lambda x: f'{x}%'),
                textposition='top center'
            ))
            fig_ocup.update_layout(
                title='Evolução da Taxa de Ocupação',
                xaxis_title='Data/Hora',
                yaxis_title='Ocupação (%)',
                height=400,
                yaxis=dict(range=[0, 100])
            )
            st.plotly_chart(fig_ocup, use_container_width=True)

        if len(df_hist_total) >= 2:
            st.markdown("**📊 Variação desde a primeira extração:**")
            primeiro = df_hist_total.iloc[0]
            ultimo = df_hist_total.iloc[-1]

            var_cols = st.columns(4)
            var_matriculados = ultimo['matriculados'] - primeiro['matriculados']
            var_novatos = ultimo['novatos'] - primeiro['novatos']
            var_veteranos = ultimo['veteranos'] - primeiro['veteranos']
            var_disponiveis = ultimo['disponiveis'] - primeiro['disponiveis']

            var_cols[0].metric("Matriculados", ultimo['matriculados'], f"{'+' if var_matriculados >= 0 else ''}{var_matriculados}")
            var_cols[1].metric("Novatos", ultimo['novatos'], f"{'+' if var_novatos >= 0 else ''}{var_novatos}")
            var_cols[2].metric("Veteranos", ultimo['veteranos'], f"{'+' if var_veteranos >= 0 else ''}{var_veteranos}")
            var_cols[3].metric("Disponíveis", ultimo['disponiveis'], f"{'+' if var_disponiveis >= 0 else ''}{var_disponiveis}")

    with tab_hist2:
        fig_unid = go.Figure()
        cores = {'Boa Viagem': '#4472C4', 'Jaboatão': '#70AD47', 'Paulista': '#ED7D31', 'Cordeiro': '#9E480E'}

        for unidade in df_hist_unidades['unidade_nome'].unique():
            df_u = df_hist_unidades[df_hist_unidades['unidade_nome'] == unidade]
            nome_curto = unidade.split('(')[1].replace(')', '') if '(' in unidade else unidade
            cor = cores.get(nome_curto, '#666666')

            fig_unid.add_trace(go.Scatter(
                x=df_u['data_formatada'],
                y=df_u['matriculados'],
                mode='lines+markers',
                name=nome_curto,
                line=dict(color=cor, width=2),
                marker=dict(size=6)
            ))

        fig_unid.update_layout(
            title='Evolução de Matriculados por Unidade',
            xaxis_title='Data/Hora',
            yaxis_title='Matriculados',
            height=450,
            hovermode='x unified'
        )
        st.plotly_chart(fig_unid, use_container_width=True)

    with tab_hist3:
        fig_seg = go.Figure()
        cores_seg = {'Ed. Infantil': '#4472C4', 'Fund. I': '#70AD47', 'Fund. II': '#ED7D31', 'Ens. Médio': '#9E480E'}
        ordem_seg = ['Ed. Infantil', 'Fund. I', 'Fund. II', 'Ens. Médio']

        for segmento in ordem_seg:
            df_s = df_hist_segmento[df_hist_segmento['segmento'] == segmento]
            if not df_s.empty:
                fig_seg.add_trace(go.Scatter(
                    x=df_s['data_formatada'],
                    y=df_s['matriculados'],
                    mode='lines+markers',
                    name=segmento,
                    line=dict(color=cores_seg.get(segmento, '#666666'), width=2),
                    marker=dict(size=6)
                ))

        fig_seg.update_layout(
            title='Evolução de Matriculados por Segmento',
            xaxis_title='Data/Hora',
            yaxis_title='Matriculados',
            height=450,
            hovermode='x unified'
        )
        st.plotly_chart(fig_seg, use_container_width=True)

st.divider()

# Detalhamento por unidade
st.subheader("🏫 Detalhamento por Unidade")

tabs = st.tabs([u['nome'].split('(')[1].replace(')', '') if '(' in u['nome'] else u['nome'] for u in resumo['unidades']])

for i, tab in enumerate(tabs):
    with tab:
        unidade = resumo['unidades'][i]
        unidade_vagas = vagas['unidades'][i]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Vagas", unidade['total']['vagas'])
        c2.metric("Matriculados", unidade['total']['matriculados'])
        c3.metric("Disponíveis", unidade['total']['disponiveis'])
        ocup = round(unidade['total']['matriculados'] / unidade['total']['vagas'] * 100, 1)
        c4.metric("Ocupação", f"{ocup}%")

        col_a, col_b = st.columns(2)

        with col_a:
            df_seg = pd.DataFrame([
                {'Segmento': seg, **vals}
                for seg, vals in unidade['segmentos'].items()
            ])

            fig_seg = go.Figure()
            fig_seg.add_trace(go.Bar(name='Vagas', x=df_seg['Segmento'], y=df_seg['vagas'], marker_color='#4472C4'))
            fig_seg.add_trace(go.Bar(name='Matriculados', x=df_seg['Segmento'], y=df_seg['matriculados'], marker_color='#70AD47'))
            fig_seg.update_layout(barmode='group', title='Por Segmento', height=350)
            st.plotly_chart(fig_seg, use_container_width=True)

        with col_b:
            fig_pizza = px.pie(
                values=[unidade['total']['novatos'], unidade['total']['veteranos']],
                names=['Novatos', 'Veteranos'],
                title='Novatos vs Veteranos',
                color_discrete_sequence=['#ED7D31', '#4472C4']
            )
            fig_pizza.update_layout(height=350)
            st.plotly_chart(fig_pizza, use_container_width=True)

        with st.expander("📋 Ver todas as turmas"):
            df_turmas = pd.DataFrame(unidade_vagas['turmas'])
            df_turmas = df_turmas[['segmento', 'turma', 'vagas', 'novatos', 'veteranos', 'matriculados', 'disponiveis']]
            df_turmas.columns = ['Segmento', 'Turma', 'Vagas', 'Novatos', 'Veteranos', 'Matriculados', 'Disponíveis']
            st.dataframe(df_turmas, use_container_width=True, hide_index=True)

st.divider()

# Comparativo geral por segmento
st.subheader("📊 Comparativo por Segmento (Todas as Unidades)")

segmentos_total = {}
for unidade in resumo['unidades']:
    for seg, vals in unidade['segmentos'].items():
        if seg not in segmentos_total:
            segmentos_total[seg] = {'vagas': 0, 'matriculados': 0, 'novatos': 0, 'veteranos': 0, 'disponiveis': 0}
        for k in segmentos_total[seg]:
            segmentos_total[seg][k] += vals[k]

df_seg_total = pd.DataFrame([
    {'Segmento': seg, **vals, 'Ocupação': round(vals['matriculados']/vals['vagas']*100, 1)}
    for seg, vals in segmentos_total.items()
])

ordem = ['Ed. Infantil', 'Fund. I', 'Fund. II', 'Ens. Médio']
df_seg_total['ordem'] = df_seg_total['Segmento'].map({s: i for i, s in enumerate(ordem)})
df_seg_total = df_seg_total.sort_values('ordem')

col_x, col_y = st.columns(2)

with col_x:
    fig_comp = go.Figure()
    fig_comp.add_trace(go.Bar(name='Vagas', x=df_seg_total['Segmento'], y=df_seg_total['vagas'], marker_color='#4472C4'))
    fig_comp.add_trace(go.Bar(name='Matriculados', x=df_seg_total['Segmento'], y=df_seg_total['matriculados'], marker_color='#70AD47'))
    fig_comp.update_layout(barmode='group', title='Vagas x Matriculados', height=400)
    st.plotly_chart(fig_comp, use_container_width=True)

with col_y:
    fig_ocup = px.bar(df_seg_total, x='Segmento', y='Ocupação', text='Ocupação',
                      color='Ocupação', color_continuous_scale=['#ff6b6b', '#ffd93d', '#6bcb77'])
    fig_ocup.update_traces(texttemplate='%{text}%', textposition='outside')
    fig_ocup.update_layout(title='Taxa de Ocupação por Segmento', height=400, showlegend=False)
    st.plotly_chart(fig_ocup, use_container_width=True)

# Rodapé
st.divider()
st.caption(f"🔄 Dados atualizados diariamente às 6h | Total de extrações: {num_extracoes}")
