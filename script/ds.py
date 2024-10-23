# Bibliotecas
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import plotly.express as px

# Definir o diretório base como o caminho do próprio script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Defina os caminhos dos arquivos usando caminhos relativos
DADOS_POINTING_PATH = os.path.join(BASE_DIR, '.database', 'ACOMPANHAMENTO DE PRODUÇÃO ATUAL-.xlsx')
DADOS_MONITORING_PATH = os.path.join(BASE_DIR, '.database', 'DATABASE.xlsx')
DADOS_DEMAND_PATH = os.path.join(BASE_DIR, '.database', 'DATABASE.xlsx')

# Função para carregar e limpar dados da aba "Pointing"
@st.cache_data
def carregar_dados_pointing_ajustado(arquivo, sheet_name):
    try:
        df = pd.read_excel(arquivo, sheet_name=sheet_name)

        # Limpar e organizar os dados
        df_cleaned = df.dropna(how='all').iloc[2:, [1, 2, 5, 8]]
        df_cleaned.columns = ['Data', 'Produção Cobre Realizado', 'Meta/Dia Cobre', 'Produção Alumínio Realizado']
        
        df_cleaned['Data'] = pd.to_datetime(df_cleaned['Data'], errors='coerce')
        df_cleaned.dropna(subset=['Data'], inplace=True)

        df_cleaned['Produção Cobre Realizado'] = pd.to_numeric(df_cleaned['Produção Cobre Realizado'], errors='coerce').fillna(0)
        df_cleaned['Produção Alumínio Realizado'] = pd.to_numeric(df_cleaned['Produção Alumínio Realizado'], errors='coerce').fillna(0)

        return df_cleaned
    except Exception as e:
        st.error(f"Erro ao carregar a aba {sheet_name}: {e}")
        return None
# Função para carregar todas as abas válidas e processar os dados de pointing
@st.cache_data
def carregar_todas_abas_ajustado(arquivo):
    xls = pd.ExcelFile(arquivo)
    dados_list = []
    meses_validos = {
        'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
        'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
    }

    for sheet in xls.sheet_names:
        if '-' in sheet:
            mes_ano = sheet.split('-')
            if len(mes_ano) == 2:
                mes, ano = mes_ano[0].strip(), mes_ano[1].strip()
                if mes in meses_validos and ano.isdigit():
                    df_cleaned = carregar_dados_pointing_ajustado(arquivo, sheet)
                    if df_cleaned is not None:
                        df_cleaned['Mês'] = mes
                        df_cleaned['Ano'] = int(ano)
                        dados_list.append(df_cleaned)
    
    return pd.concat(dados_list, ignore_index=True) if dados_list else None
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
    dados = carregar_todas_abas_ajustado(DADOS_POINTING_PATH)

    if dados is not None:
        # Seleção entre Cobre e Alumínio
        producao_tipo = st.radio("Escolha o tipo de produção", ('Cobre', 'Alumínio'))

        # Seleção entre Comparação por Anos ou Meses
        comparacao_tipo = st.radio("Escolha como deseja comparar os dados", ('Comparação por Anos', 'Comparação por Meses'))

        if comparacao_tipo == 'Comparação por Anos':
            # Usuário seleciona um ou mais anos para análise
            anos = dados['Ano'].unique() if 'Ano' in dados.columns else []
            anos_selecionados = st.multiselect('Selecione o(s) Ano(s)', anos)

            if anos_selecionados:
                # Filtrar os dados pelos anos selecionados
                dados_filtrados = dados[dados['Ano'].isin(anos_selecionados)]
                
                # Exibir uma tabela com a produção total e a expectativa por ano
                producao_total_ano = []
                expectativa_total_ano = []
                
                for ano in anos_selecionados:
                    dados_ano = dados_filtrados[dados_filtrados['Ano'] == ano]
                    if producao_tipo == 'Cobre':
                        total_producao_ano = dados_ano['Produção Cobre Realizado'].sum()
                        expectativa_ano = dados_ano['Meta/Dia Cobre'].sum()
                    else:
                        total_producao_ano = dados_ano['Produção Alumínio Realizado'].sum()
                        expectativa_ano = dados_ano['Meta/Dia Cobre'].sum()

                    producao_total_ano.append(total_producao_ano)
                    expectativa_total_ano.append(expectativa_ano)
                
                # Exibir os dados em formato de tabela
                st.write("### Relação entre Anos")
                df_anos = pd.DataFrame({
                    'Ano': anos_selecionados,
                    'Quantidade Total Produzida': producao_total_ano,
                    'Expectativa de Produção': expectativa_total_ano
                })
                st.dataframe(df_anos)

                # Gráfico de setores para a relação entre os anos
                fig_anos = px.pie(df_anos, names='Ano', values='Quantidade Total Produzida',
                                  title=f"Distribuição da Produção nos Anos Selecionados - {producao_tipo}")
                st.plotly_chart(fig_anos)

        elif comparacao_tipo == 'Comparação por Meses':
            # O usuário seleciona um ano
            anos = dados['Ano'].unique() if 'Ano' in dados.columns else []
            ano_selecionado = st.selectbox('Selecione o Ano', anos)

            if ano_selecionado:
                # Filtrar os dados pelo ano selecionado
                dados_filtrados = dados[dados['Ano'] == ano_selecionado]
                
                # Exibir uma tabela com a produção de cada mês do ano selecionado
                meses = dados_filtrados['Mês'].unique()

                producao_total_mes = []
                expectativa_total_mes = []
                
                for mes in meses:
                    dados_mes = dados_filtrados[dados_filtrados['Mês'] == mes]

                    if producao_tipo == 'Cobre':
                        total_mes = dados_mes['Produção Cobre Realizado'].sum()
                        expectativa_mes = dados_mes['Meta/Dia Cobre'].sum()
                    else:
                        total_mes = dados_mes['Produção Alumínio Realizado'].sum()
                        expectativa_mes = dados_mes['Meta/Dia Cobre'].sum()

                    producao_total_mes.append(total_mes)
                    expectativa_total_mes.append(expectativa_mes)

                # Exibir os dados em formato de tabela
                st.write(f"### Produção Mensal no Ano de {ano_selecionado}")
                df_meses = pd.DataFrame({
                    'Meses': meses,
                    'Quantidade Total Produzida': producao_total_mes,
                    'Expectativa de Produção': expectativa_total_mes
                })
                st.dataframe(df_meses)

                # Gráfico de setores para a relação entre os meses do ano selecionado
                fig_meses = px.pie(df_meses, names='Meses', values='Quantidade Total Produzida',
                                   title=f"Distribuição da Produção nos Meses de {ano_selecionado} - {producao_tipo}")
                st.plotly_chart(fig_meses)

                # Exibir total produzido no ano, expectativa de produção e média anual
                total_anual = df_meses['Quantidade Total Produzida'].sum()
                expectativa_anual = df_meses['Expectativa de Produção'].sum()
                media_anual = df_meses['Quantidade Total Produzida'].mean()

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.write(f"**Total Produzido no Ano**: {total_anual:.2f}")
                with col2:
                    st.write(f"**Expectativa de Produção no Ano**: {expectativa_anual:.2f}")
                with col3:
                    st.write(f"**Média de Produção Anual**: {media_anual:.2f}")

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
