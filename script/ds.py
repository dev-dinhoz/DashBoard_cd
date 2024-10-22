# Bibliotecas
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import plotly

# Definir o diretório base como o caminho do próprio script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Defina os caminhos dos arquivos usando caminhos relativos
DADOS_POINTING_PATH = os.path.join(BASE_DIR, '.database', 'ACOMPANHAMENTO.xlsx')
DADOS_MONITORING_PATH = os.path.join(BASE_DIR, '.database', 'DATABASE.xlsx')
DADOS_DEMAND_PATH = os.path.join(BASE_DIR, '.database', 'DATABASE.xlsx')

@st.cache_data
def carregar_dados_monitoring():
    try:
        wb1 = load_workbook(DADOS_MONITORING_PATH, data_only=True)
        sheet = wb1.active
        data = sheet.values
        columns = next(data)  # Pega a primeira linha como cabeçalho
        
        # Corrigir a leitura dos dados
        dados = pd.DataFrame(data, columns=columns)
        dados = dados[['Produção Cobre Realizado', 'Produção Alumínio Realizado']]
        return dados
    except FileNotFoundError:
        st.error(f"Arquivo '{DADOS_MONITORING_PATH}' não encontrado.")
        return None
@st.cache_data
def carregar_dados_pointing(arquivo=DADOS_POINTING_PATH):
    try:
        # Inicializa uma lista para armazenar os DataFrames
        dados_list = []
        
        # Lê o arquivo Excel e verifica as abas
        xls = pd.ExcelFile(arquivo)  # Abre o arquivo Excel
        print("Abas encontradas:")
        for sheet in xls.sheet_names:
            print(sheet)  # Imprime o nome da aba encontrada
            # Filtra as abas que contém mês e ano
            if '-' in sheet:
                mes_ano = sheet.split('-')
                if len(mes_ano) == 2:
                    mes, ano = mes_ano[0].strip(), mes_ano[1].strip()
                    meses_validos = {
                        'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                        'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
                    }
                    if mes in meses_validos and ano.isdigit():
                        df = pd.read_excel(arquivo, sheet_name=sheet)
                        df['Mês'] = mes  # Adiciona a coluna 'Mês'
                        df['Ano'] = int(ano)  # Adiciona a coluna 'Ano'
                        dados_list.append(df)  # Adiciona o DataFrame à lista
        # Concatena todos os DataFrames da lista em um único DataFrame
        if dados_list:
            return pd.concat(dados_list, ignore_index=True)
        else:
            st.error("Nenhuma aba válida encontrada.")
            return None
    except FileNotFoundError:
        st.error(f"Arquivo '{arquivo}' não encontrado.")
        return None
@st.cache_data
def carregar_dados_demand():
    try:
        wb3 = load_workbook(DADOS_DEMAND_PATH, data_only=True)
        sheet = wb3.active
        data = sheet.values
        columns = next(data)
        
        dados = pd.DataFrame(data, columns=columns)  # Corrigir a leitura dos dados
        dados = dados[['Data', 'Produção Cobre Realizado', 'Produção Alumínio Realizado']]
        return dados
    except FileNotFoundError:
        st.error(f"Arquivo '{DADOS_DEMAND_PATH}' não encontrado.")
        return None
    
# Funções para cada página
def pagina1():
    st.write('## Monitoring')
    st.write('#### Programação')
def pagina2():
    st.write('## Pointing')
    st.write('#### Acompanhamento de produção')
    # Carregar os dados
    dados = carregar_dados_pointing()

    if dados is not None:
        # Extraindo anos e meses diretamente das colunas
        anos = dados['Ano'].unique() if 'Ano' in dados.columns else []
        meses = dados['Mês'].unique() if 'Mês' in dados.columns else []

        # Seleção dos anos e meses
        ano_selecionado = st.multiselect('Selecione o(s) Ano(s)', anos)
        mes_selecionado = st.multiselect('Selecione o(s) Mês(es)', meses)

        # Filtrar dados por ano e mês
        if ano_selecionado and mes_selecionado:
            dados_filtrados = dados[
                (dados['Mês'].isin(mes_selecionado)) &
                (dados['Ano'].isin(ano_selecionado))
            ]

            if not dados_filtrados.empty:
                # Exibir dados filtrados em uma lista expansível (toggle list)
                with st.expander(f'Exibir Dados Filtrados para os Anos {ano_selecionado} e Meses {mes_selecionado}'):
                    st.dataframe(dados_filtrados)

                # Somar as colunas 'Produção Cobre' e 'Produção Alumínio'
                if 'Produção Cobre Realizado' in dados_filtrados.columns and 'Produção Alumínio Realizado' in dados_filtrados.columns:
                    total_cobre = dados_filtrados['Produção Cobre Realizado'].sum()
                    total_aluminio = dados_filtrados['Produção Alumínio Realizado'].sum()

                    # Exibir métricas de Produção Cobre e Produção Alumínio
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric(label="Cobre", value=f"{total_cobre:.2f}")
                    with col2:
                        st.metric(label="Alumínio", value=f"{total_aluminio:.2f}")

                    # Gráfico de Linhas
                    dados_filtrados_grouped = dados_filtrados.groupby(['Ano', 'Mês'])[['Produção Cobre Realizado', 'Produção Alumínio Realizado']].sum()

                    st.write("### Gráfico de Linhas")
                    st.plotly_chart(plotly.graph_objs.Figure(data=[plotly.graph_objs.Scatter(x=dados_filtrados_grouped.index, y=dados_filtrados_grouped['Produção Cobre Realizado'])]))

                    # Gráfico de Barras
                    st.write("### Gráfico de Barras")
                    st.plotly_chart(plotly.graph_objs.Figure(data=[plotly.graph_objs.Bar(x=dados_filtrados_grouped.index, y=dados_filtrados_grouped['Produção Alumínio Realizado'])]))

                    # Gráfico de Pizza (Proporção)
                    st.write("### Gráfico de Setores")
                    st.plotly_chart(plotly.graph_objs.Figure(data=[plotly.graph_objs.Pie(labels=dados_filtrados_grouped.index, values=dados_filtrados_grouped['Produção Cobre Realizado'])]))
                else:
                    st.write("Colunas de produção não encontradas nos dados.")
            else:
                st.write("Nenhum dado encontrado para os filtros selecionados.")
        else:
            st.write("Selecione pelo menos um ano e um mês para comparar os dados.")
    else:
        st.write("Erro ao carregar os dados.")
def pagina3():
    st.write('## Demand')
    st.write('#### Relevância por composto')

# Interface do sistema
st.set_page_config(page_title="Dashboard", page_icon="💡", layout="wide")

imagem_caminho = os.path.join(BASE_DIR, '.uploads', 'Logo.png')
if os.path.exists(imagem_caminho):
    st.sidebar.image(imagem_caminho, use_column_width=True)
else:
    st.sidebar.error(f"Imagem no caminho '{imagem_caminho}' não encontrada.")

if 'pagina_atual' not in st.session_state:
    st.session_state.pagina_atual = 'pagina1'

st.sidebar.markdown("<br><br>", unsafe_allow_html=True)
botao_pagina1 = st.sidebar.button('(ICON1)', on_click=lambda: st.session_state.update({'pagina_atual': 'pagina1'}))
st.sidebar.markdown("<br><br><br><br><br>", unsafe_allow_html=True)
botao_pagina2 = st.sidebar.button('(ICON2)', on_click=lambda: st.session_state.update({'pagina_atual': 'pagina2'}))
st.sidebar.markdown("<br><br><br><br><br>", unsafe_allow_html=True)
botao_pagina3 = st.sidebar.button('(ICON3)', on_click=lambda: st.session_state.update({'pagina_atual': 'pagina3'}))

if st.session_state.pagina_atual == 'pagina1':
    pagina1()
elif st.session_state.pagina_atual == 'pagina2':
    pagina2()
elif st.session_state.pagina_atual == 'pagina3':
    pagina3()
