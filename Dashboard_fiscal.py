import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import json
import numpy as np
from scipy import stats
import re

# Configuração da página
st.set_page_config(layout="wide", page_title="Análise CFOP e Mapa do Brasil ICMS")
)

# Estilo CSS personalizado
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
        font-weight: bold;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1.5rem;
        border-radius: 10px;
        border-left: 4px solid #1f77b4;
    }
    .section-header {
        font-size: 1.5rem;
        color: #1f77b4;
        margin: 2rem 0 1rem 0;
        border-bottom: 2px solid #1f77b4;
        padding-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Título principal
st.markdown('<div class="main-header">📊 Dashboard Fiscal - ICMS Bonificação e Devolução</div>', unsafe_allow_html=True)


# Função para carregar dados da pasta específica
@st.cache_data
def load_data_from_folder(folder_path):
    """
    Carrega todos os arquivos Excel da pasta especificada e combina em um único DataFrame
    """
    # Verifica se a pasta existe
    if not os.path.exists(folder_path):
        st.error(f"❌ Pasta não encontrada: {folder_path}")
        return pd.DataFrame()

    # Encontra todos os arquivos Excel na pasta
    excel_files = glob.glob(os.path.join(folder_path, "*.xlsx")) + glob.glob(os.path.join(folder_path, "*.xls"))

    if not excel_files:
        st.error(f"❌ Nenhum arquivo Excel encontrado na pasta: {folder_path}")
        return pd.DataFrame()

    st.sidebar.success(f"✅ {len(excel_files)} arquivo(s) encontrado(s)")

    all_data = []

    for file_path in excel_files:
        try:
            # Lê o arquivo Excel
            df = pd.read_excel(file_path)

            # Adiciona coluna com o nome do arquivo de origem
            df['Arquivo_Origem'] = os.path.basename(file_path)

            all_data.append(df)
            st.sidebar.info(f"📂 {os.path.basename(file_path)} - {len(df)} registros")

        except Exception as e:
            st.sidebar.error(f"❌ Erro ao ler {os.path.basename(file_path)}: {str(e)}")

    if not all_data:
        st.error("❌ Nenhum dado foi carregado com sucesso")
        return pd.DataFrame()

    # Combina todos os DataFrames
    combined_df = pd.concat(all_data, ignore_index=True)

    # Processamento dos dados (baseado na estrutura da sua planilha)
    # Converte coluna de data se existir
    date_columns = ['Data de emissão atualizada', 'Data', 'Emissão', 'DATA EMISSÃO']
    for col in date_columns:
        if col in combined_df.columns:
            combined_df[col] = pd.to_datetime(combined_df[col], errors='coerce')
            # Usa a primeira coluna de data encontrada
            combined_df['Data de emissão atualizada'] = combined_df[col]
            break

    # Se não encontrou coluna de data, cria uma fictícia baseada no nome do arquivo
    if 'Data de emissão atualizada' not in combined_df.columns:
        combined_df['Data de emissão atualizada'] = pd.to_datetime('2025-01-01')
        st.warning("⚠️ Coluna de data não encontrada. Usando data padrão.")

    # Processa mês e ano
    combined_df['Mês'] = combined_df['Data de emissão atualizada'].dt.month
    combined_df['Ano'] = combined_df['Data de emissão atualizada'].dt.year
    combined_df['Mês/Ano'] = combined_df['Data de emissão atualizada'].dt.to_period('M').astype(str)

    # Mapeamento de CFOP para descrições
    cfop_map = {
        '5102': 'Venda de mercadorias',
        '6102': 'Venda de mercadorias',
        '5905': 'Remessa para depósito fechado',
        '5910': 'Bonificação - Dentro do estado',
        '6910': 'Bonificação - Fora do estado',
        '6905': 'Remessa para depósito',
        '6108': 'Venda de mercadorias',
        '5101': 'Venda de mercadorias',
        '6101': 'Venda de mercadorias'
    }

    # Identifica a coluna CFOP
    cfop_columns = ['CFOP', 'Código CFOP', 'CFOP Code']
    cfop_col = None
    for col in cfop_columns:
        if col in combined_df.columns:
            cfop_col = col
            break

    if cfop_col:
        combined_df['CFOP'] = combined_df[cfop_col].astype(str)
        combined_df['Descrição CFOP'] = combined_df['CFOP'].map(cfop_map)
        combined_df['Descrição CFOP'] = combined_df['Descrição CFOP'].fillna('Outros')
    else:
        combined_df['CFOP'] = 'Não identificado'
        combined_df['Descrição CFOP'] = 'CFOP não encontrado'
        st.warning("⚠️ Coluna CFOP não encontrada nos arquivos")

    # Identifica coluna de situação
    situacao_columns = ['Situação', 'Status', 'SITUAÇÃO', 'STATUS']
    situacao_col = None
    for col in situacao_columns:
        if col in combined_df.columns:
            situacao_col = col
            break

    if situacao_col:
        combined_df['Situação'] = combined_df[situacao_col]
    else:
        combined_df['Situação'] = 'Emitida DANFE'
        st.warning("⚠️ Coluna Situação não encontrada nos arquivos")

    # Identifica coluna de valor ICMS
    icms_columns = ['Valor de ICMS', 'ICMS', 'VALOR ICMS', 'ICMS Valor']
    icms_col = None
    for col in icms_columns:
        if col in combined_df.columns:
            icms_col = col
            break

    if icms_col:
        combined_df['Valor de ICMS'] = pd.to_numeric(combined_df[icms_col], errors='coerce').fillna(0)
    else:
        combined_df['Valor de ICMS'] = 0
        st.warning("⚠️ Coluna Valor de ICMS não encontrada nos arquivos")

    return combined_df


# Sidebar - Configuração do caminho da pasta
st.sidebar.header("📁 Configuração dos Dados")

# Caminho padrão
default_path = r"G:\Drives compartilhados\Moon Ventures - Admin Fin\Minimal Club\Dados\Notas fiscais de saída\2025"

folder_path = st.sidebar.text_input(
    "Caminho da pasta com os arquivos Excel:",
    value=default_path,
    help="Cole o caminho completo da pasta que contém os arquivos Excel"
)

# Botão para carregar dados
if st.sidebar.button("🔄 Carregar Dados", type="primary"):
    st.rerun()

# Carregar dados
if folder_path:
    with st.spinner("📂 Carregando dados da pasta..."):
        df = load_data_from_folder(folder_path)
else:
    st.info("👆 Por favor, insira o caminho da pasta no menu lateral")
    st.stop()

# Verifica se há dados carregados
if df.empty:
    st.error("❌ Nenhum dado foi carregado. Verifique o caminho da pasta e os arquivos.")
    st.stop()

# Sidebar com filtros
st.sidebar.header("🔍 Filtros")

# Filtro de data
if 'Data de emissão atualizada' in df.columns:
    min_date = df['Data de emissão atualizada'].min()
    max_date = df['Data de emissão atualizada'].max()
    date_range = st.sidebar.date_input(
        "Período",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )
else:
    date_range = None

# Filtro de CFOP
cfop_options = ['Todos'] + sorted(df['CFOP'].unique().tolist())
selected_cfop = st.sidebar.selectbox("CFOP", cfop_options)

# Filtro de situação
situacao_options = ['Todos'] + sorted(df['Situação'].unique().tolist())
selected_situacao = st.sidebar.selectbox("Situação", situacao_options)

# Aplicar filtros
filtered_df = df.copy()

if date_range and len(date_range) == 2:
    filtered_df = filtered_df[
        (filtered_df['Data de emissão atualizada'] >= pd.to_datetime(date_range[0])) &
        (filtered_df['Data de emissão atualizada'] <= pd.to_datetime(date_range[1]))
        ]

if selected_cfop != 'Todos':
    filtered_df = filtered_df[filtered_df['CFOP'] == selected_cfop]

if selected_situacao != 'Todos':
    filtered_df = filtered_df[filtered_df['Situação'] == selected_situacao]

# Informações do carregamento
st.sidebar.markdown("---")
st.sidebar.markdown("### 📊 Estatísticas do Carregamento")
st.sidebar.metric("Total de Registros", f"{len(df):,}")
st.sidebar.metric("Registros Filtrados", f"{len(filtered_df):,}")
st.sidebar.metric("Período dos Dados", f"{df['Mês/Ano'].min()} a {df['Mês/Ano'].max()}")

# Métricas principais
st.markdown('<div class="section-header">📈 Métricas Principais</div>', unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)

with col1:
    total_notas = len(filtered_df)
    st.metric("Total de Notas", f"{total_notas:,}")

with col2:
    total_icms = filtered_df['Valor de ICMS'].sum()
    st.metric("Valor Total ICMS", f"R$ {total_icms:,.2f}")

with col3:
    valor_medio = filtered_df['Valor de ICMS'].mean() if len(filtered_df) > 0 else 0
    st.metric("Valor Médio por NF", f"R$ {valor_medio:,.2f}")

with col4:
    taxa_cancelamento = (len(filtered_df[filtered_df['Situação'] == 'Cancelada']) / len(filtered_df) * 100) if len(
        filtered_df) > 0 else 0
    st.metric("Taxa de Cancelamento", f"{taxa_cancelamento:.1f}%")

# Gráficos e visualizações
st.markdown('<div class="section-header">📊 Visualizações</div>', unsafe_allow_html=True)

# Primeira linha de gráficos
col1, col2 = st.columns(2)

with col1:
    # Evolução mensal do ICMS
    if len(filtered_df) > 0:
        monthly_data = filtered_df.groupby('Mês/Ano').agg({
            'Valor de ICMS': 'sum',
            'CFOP': 'count'
        }).reset_index()

        fig_evolution = px.line(
            monthly_data,
            x='Mês/Ano',
            y='Valor de ICMS',
            title='Evolução Mensal do ICMS',
            markers=True
        )
        fig_evolution.update_layout(height=400)
        st.plotly_chart(fig_evolution, use_container_width=True)
    else:
        st.info("📈 Não há dados para exibir o gráfico de evolução")

with col2:
    # Distribuição por situação
    if len(filtered_df) > 0:
        situacao_count = filtered_df['Situação'].value_counts()
        fig_situacao = px.pie(
            values=situacao_count.values,
            names=situacao_count.index,
            title='Distribuição por Situação da NF'
        )
        fig_situacao.update_layout(height=400)
        st.plotly_chart(fig_situacao, use_container_width=True)
    else:
        st.info("📊 Não há dados para exibir o gráfico de situações")

# Segunda linha de gráficos
col1, col2 = st.columns(2)

with col1:
    # Top CFOPs por valor
    if len(filtered_df) > 0:
        cfop_value = filtered_df.groupby(['CFOP', 'Descrição CFOP'])['Valor de ICMS'].sum().reset_index()
        cfop_value = cfop_value.sort_values('Valor de ICMS', ascending=False).head(10)

        fig_cfop = px.bar(
            cfop_value,
            x='Valor de ICMS',
            y='CFOP',
            orientation='h',
            title='Top CFOPs por Valor de ICMS',
            hover_data=['Descrição CFOP']
        )
        fig_cfop.update_layout(height=400)
        st.plotly_chart(fig_cfop, use_container_width=True)
    else:
        st.info("📋 Não há dados para exibir o gráfico de CFOPs")

with col2:
    # Heatmap de situações por mês
    if len(filtered_df) > 0:
        heatmap_data = filtered_df.groupby(['Mês/Ano', 'Situação']).size().unstack(fill_value=0)
        fig_heatmap = px.imshow(
            heatmap_data.T,
            title='Heatmap - Situações por Mês',
            aspect="auto",
            color_continuous_scale='Blues'
        )
        fig_heatmap.update_layout(height=400)
        st.plotly_chart(fig_heatmap, use_container_width=True)
    else:
        st.info("🔥 Não há dados para exibir o heatmap")

# Tabela detalhada
st.markdown('<div class="section-header">📋 Dados Detalhados</div>', unsafe_allow_html=True)

if len(filtered_df) > 0:
    # Resumo por CFOP e Situação
    summary_table = filtered_df.pivot_table(
        index=['CFOP', 'Descrição CFOP'],
        columns='Situação',
        values='Valor de ICMS',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    st.dataframe(
        summary_table,
        use_container_width=True,
        height=400
    )

    # Download dos dados filtrados
    csv = filtered_df.to_csv(index=False)
    st.download_button(
        label="📥 Baixar Dados Filtrados (CSV)",
        data=csv,
        file_name=f"dados_fiscais_filtrados_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv"
    )
else:
    st.info("📄 Não há dados para exibir na tabela")

# Informações adicionais na sidebar
st.sidebar.markdown("---")
st.sidebar.markdown("### ℹ️ Legenda CFOP")
st.sidebar.markdown("""
- **5102/6102/6108**: Venda de mercadorias
- **5905/6905**: Remessa para depósito
- **5910**: Bonificação - Dentro do estado  
- **6910**: Bonificação - Fora do estado
""")

# Rodapé
st.markdown("---")

st.markdown(f"*Dashboard desenvolvido com Streamlit - Dados carregados de: {folder_path}*")
