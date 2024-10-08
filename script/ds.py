# Bibliotecas
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import plotly.express as px
import os

# Definir o diretório base como o caminho do próprio script
# Definir o diretório base como o caminho do próprio script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Defina os caminhos dos arquivos usando caminhos relativos
DADOS_POINTING_PATH = os.path.join(BASE_DIR, '.database', 'DATABASE.xlsx')
DADOS_MONITORING_PATH = os.path.join(BASE_DIR, '.database', 'ACOMPANHAMENTO DE PRODUÇÃO ATUAL-.xlsx')  # Ajustado para incluir a barra invertida correta
DADOS_DEMAND_PATH = os.path.join(BASE_DIR, '.database', 'ACOMPANHAMENTO DE PRODUÇÃO ATUAL-.xlsx')  # Ajustado para incluir a barra invertida correta

@st.cache_data
def carregar_todos_os_dados():
    with st.spinner():
        dados_monitoring = carregar_dados_monitoring()
        dados_pointing = carregar_dados_pointing()
        dados_demand = carregar_dados_demand()
    return dados_monitoring, dados_pointing, dados_demand
@st.cache_data
def carregar_dados_monitoring():
    try:
        wb = load_workbook(DADOS_MONITORING_PATH, data_only=True)
        sheet = wb.active
        data = sheet.values
        columns = next(data)  # Pega a primeira linha como cabeçalho
        return pd.DataFrame(data, columns=columns)
    except FileNotFoundError:
        st.error(f"Arquivo '{DADOS_MONITORING_PATH}' não encontrado.")
        return None

@st.cache_data 
def carregar_dados_pointing():
    try:
        wb = load_workbook(DADOS_POINTING_PATH, data_only=True)
        dados = pd.DataFrame()
        for sheet in wb.sheetnames:
            if '-' in sheet:  # Verifica se a aba tem um nome que indica mês e ano
                mes_ano = sheet.split('-')
                if len(mes_ano) == 2:
                    mes, ano = mes_ano[0].strip(), mes_ano[1].strip()
                    meses_validos = {
                        'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                        'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
                    }
                    if mes in meses_validos and ano.isdigit():
                        df = pd.read_excel(DADOS_POINTING_PATH, sheet_name=sheet)
                        df['Mês'] = mes
                        df['Ano'] = ano  # Adiciona a coluna 'Ano'
                        dados = pd.concat([dados, df], ignore_index=True)
        return dados
    except FileNotFoundError:
        st.error(f"Arquivo '{DADOS_POINTING_PATH}' não encontrado.")
        return None

@st.cache_data
def carregar_dados_demand():
    try:
        wb = load_workbook(DADOS_DEMAND_PATH, data_only=True)
        sheet = wb.active
        data = sheet.values
        columns = next(data)  # Pega a primeira linha como cabeçalho
        return pd.DataFrame(data, columns=columns)
    except FileNotFoundError:
        st.error(f"Arquivo '{DADOS_DEMAND_PATH}' não encontrado.")
        return None

# Funções que representam o conteúdo de cada página
def pagina1(dados_monitoring):
    st.write('## Monitoring')
    st.write('#### Programação')
    if dados_monitoring is not None:
        st.dataframe(dados_monitoring)

def pagina2(dados_pointing):
    st.write('## Pointing')
    st.write('#### Acompanhamento de produção')

    # Carregar os dados
    dados = carregar_dados_pointing()

    if dados is not None:
        # Filtros
        anos = dados['Ano'].unique()
        meses = dados['Mês'].unique()

        # Utilizando session_state para manter a seleção dos filtros
        if 'anos_selecionados' not in st.session_state:
            st.session_state.anos_selecionados = anos.tolist()  # Selecionar todos por padrão
        if 'meses_selecionados' not in st.session_state:
            st.session_state.meses_selecionados = meses.tolist()  # Selecionar todos por padrão

        ano_selecionado = st.multiselect('Selecione o(s) Ano(s)', anos, default=st.session_state.anos_selecionados)
        mes_selecionado = st.multiselect('Selecione o(s) Mês(es)', meses, default=st.session_state.meses_selecionados)

        # Atualizar o estado com a seleção atual
        st.session_state.anos_selecionados = ano_selecionado
        st.session_state.meses_selecionados = mes_selecionado

        # Filtrar dados por ano e mês
        if ano_selecionado and mes_selecionado:
            dados_filtrados = dados[(dados['Ano'].isin(ano_selecionado)) & (dados['Mês'].isin(mes_selecionado))]

            if not dados_filtrados.empty:
                # Exibir dados filtrados em uma lista expansível (toggle list)
                with st.expander(f'Exibir Dados Filtrados para os Anos {ano_selecionado} e Meses {mes_selecionado}'):
                    st.dataframe(dados_filtrados)

                # Verificar se as colunas 'Produção Cobre Realizado' e 'Produção Alumínio Realizado' existem
                if 'Produção Cobre Realizado' in dados_filtrados.columns and 'Produção Alumínio Realizado' in dados_filtrados.columns:
                    # Somar as colunas 'Produção Cobre Realizado' e 'Produção Alumínio Realizado'
                    total_cobre = dados_filtrados['Produção Cobre Realizado'].sum()
                    total_aluminio = dados_filtrados['Produção Alumínio Realizado'].sum()

                    # Exibir métricas de Produção Cobre e Produção Alumínio
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric(label="Cobre", value=f"{total_cobre:.2f}")
                    with col2:
                        st.metric(label="Alumínio", value=f"{total_aluminio:.2f}")

                    # Agrupar por 'Ano' e 'Mês' e somar as colunas de produção
                    dados_filtrados_grouped = dados_filtrados.groupby(['Ano', 'Mês'])[['Produção Cobre Realizado', 'Produção Alumínio Realizado']].sum()

                    # Gráfico de Linhas
                    st.write("### Gráfico de Linhas")
                    fig_line = px.line(dados_filtrados_grouped.reset_index(), x='Mês', y=['Produção Cobre Realizado', 'Produção Alumínio Realizado'], 
                                       labels={'value': 'Produção', 'index': 'Ano-Mês'}, 
                                       title="Evolução da Produção Realizada (Cobre vs Alumínio)")
                    st.plotly_chart(fig_line)

                    # Gráfico de Barras
                    st.write("### Gráfico de Barras")
                    fig_bar = px.bar(dados_filtrados_grouped.reset_index(), x='Mês', 
                                     y=['Produção Cobre Realizado', 'Produção Alumínio Realizado'], 
                                     labels={'value': 'Produção', 'index': 'Ano-Mês'}, 
                                     title="Produção Realizada Agregada (Cobre vs Alumínio)")
                    st.plotly_chart(fig_bar)

                    # Gráfico de Pizza (Proporção)
                    st.write("### Gráfico de Setores")
                    proporcoes = pd.DataFrame({
                        'Material': ['Cobre', 'Alumínio'],
                        'Produção': [total_cobre, total_aluminio]
                    })
                    fig_pie = px.pie(proporcoes, names='Material', values='Produção', 
                                     title="Proporção da Produção Realizada (Cobre vs Alumínio)")
                    st.plotly_chart(fig_pie)
                else:
                    st.write("Colunas de produção realizadas não encontradas nos dados.")
            else:
                st.write("Nenhum dado encontrado para os filtros selecionados.")
        else:
            st.write("Selecione pelo menos um ano e um mês para comparar os dados.")
    else:
        st.write("Erro ao carregar os dados.")
def pagina3(dados_demand):
    st.write('## Demand')
    st.write('#### Relevância por composto')
    if dados_demand is not None:
        st.dataframe(dados_demand)

# Interface do sistema
st.set_page_config(layout="wide")

# Logo
imagem_caminho = os.path.join(BASE_DIR, '.uploads', 'Logo.png')
if os.path.exists(imagem_caminho):
    st.sidebar.image(imagem_caminho, use_column_width=True)
else:
    st.sidebar.error(f"Imagem no caminho '{imagem_caminho}' não encontrada.")

# Menu lateral com botões para navegação entre as páginas
st.sidebar.markdown("<br><br>", unsafe_allow_html=True)

# Criação dos botões com espaçamento entre eles
if 'pagina_atual' not in st.session_state:
    st.session_state.pagina_atual = 'pagina1'  # Página inicial

botao_pagina1 = st.sidebar.button('Página 1 (ICON)', on_click=lambda: st.session_state.update({'pagina_atual': 'pagina1'}))
st.sidebar.markdown("<br><br><br><br><br>", unsafe_allow_html=True)

botao_pagina2 = st.sidebar.button('Página 2 (ICON)', on_click=lambda: st.session_state.update({'pagina_atual': 'pagina2'}))
st.sidebar.markdown("<br><br><br><br><br>", unsafe_allow_html=True)

botao_pagina3 = st.sidebar.button('Página 3 (ICON)', on_click=lambda: st.session_state.update({'pagina_atual': 'pagina3'}))
# Carregar os dados
dados_monitoring, dados_pointing, dados_demand = carregar_todos_os_dados()

# Exibição da página atual
pagina_atual = st.session_state.pagina_atual
if pagina_atual == 'pagina1':
    pagina1(dados_monitoring)
elif pagina_atual == 'pagina2':
    pagina2(dados_pointing)
elif pagina_atual == 'pagina3':
    pagina3(dados_demand)