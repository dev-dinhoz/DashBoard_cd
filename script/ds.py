import locale
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import plotly.express as px

# Definir o diretório base como o caminho do próprio script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Definir o formato de números como pt-BR
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
locale.atof = lambda x: float(x.replace('.', '').replace(',', '.'))  # Ignorar separadores de milhar e considerar apenas 2 casas decimais

def formatar_valores(valor):
    """ Formatar valores numéricos no formato 10.000,00 """
    return locale.format_string('%.2f', valor / 1000, grouping=True)  # Dividimos por 1000 para mostrar em milhares

def formatar_data_brasileira(data):
    """ Formatar data no formato brasileiro dd/mm/yyyy """
    return data.strftime('%d/%  m/%Y')

# Defina os caminhos dos arquivos usando caminhos relativos
DADOS_POINTING_PATH = os.path.join(BASE_DIR, '.database', 'ACOMPANHAMENTO DE PRODUÇÃO ATUAL-.xlsx')
DADOS_MONITORING_PATH = os.path.join(BASE_DIR, '.database', 'DATABASE.xlsx')
DADOS_DEMAND_PATH = os.path.join(BASE_DIR, '.database', 'DATABASE.xlsx')

# Função para carregar e limpar dados da aba "Pointing"
                        # ADICIONAR FORMATAÇÃO DOS VALORES AQUI 
@st.cache_data
def carregar_dados_pointing_ajustado(arquivo, sheet_name):
    
    try:
        df = pd.read_excel(arquivo, sheet_name=sheet_name, header=0)  # Ignora fórmulas e apenas retorna os valores

        # Limpar e organizar os dados
        df_cleaned = df.dropna(how='all').iloc[1:, [1, 2, 5, 8, 11]]  # Seleciona as colunas relevantes
        df_cleaned.columns = ['Data', 'Produção Cobre Realizado', 'Meta/Dia Cobre', 'Produção Alumínio Realizado', 'Meta/Dia Alumínio']
        
        df_cleaned['Data'] = pd.to_datetime(df_cleaned['Data'], errors='coerce')
        df_cleaned.dropna(subset=['Data'], inplace=True)

        # Adicionar coluna "Dia" com contagem que reinicia a cada mês
        df_cleaned[''] = df_cleaned.groupby(df_cleaned['Data'].dt.to_period("M")).cumcount() + 1 # Testar passar df vazio como argumento

        # Aplicar a formatação brasileira na data
        df_cleaned['Data'] = df_cleaned['Data'].dt.strftime('%d/%m/%Y')

        # Convertendo as colunas para numérico
        df_cleaned['Produção Cobre Realizado'] = pd.to_numeric(df_cleaned['Produção Cobre Realizado'], errors='coerce').fillna(0)
        df_cleaned['Meta/Dia Cobre'] = pd.to_numeric(df_cleaned['Meta/Dia Cobre'], errors='coerce').fillna(0) # ajustar método de calculo
        df_cleaned['Produção Alumínio Realizado'] = pd.to_numeric(df_cleaned['Produção Alumínio Realizado'], errors='coerce').fillna(0)
        df_cleaned['Meta/Dia Alumínio'] = pd.to_numeric(df_cleaned['Meta/Dia Alumínio'], errors='coerce').fillna(0) # ajustar método de calculo

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
 df = pd.read_excel(DADOS_MONITORING_PATH, header=2, sheet = ('ProgramaExtrusão'))
 return df
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
    st.header('_Status Máquina_', divider='gray')                     #WEBSCRAPING MALDITO

# ajustar os o carregamento de dados de alumínio
# ajustar formatação dos valores do DataFrame META/DIA para apenas 2 casas decimais sejam exibidas depois da vírgula

def pagina2():
    st.header('_Acompanhamento de Produção_', divider='gray')

    # Carregar os dados
    dados = carregar_todas_abas_ajustado(DADOS_POINTING_PATH)

    if dados is not None:
        col1, col2 = st.columns(2)
        with col1:
        # Seleção entre Cobre e Alumínio
            producao_tipo = st.radio("Escolha o tipo de produção", ('Cobre', 'Alumínio'))
        with col2:
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
                        expectativa_ano = dados_ano['Meta/Dia Alumínio'].sum()

                    producao_total_ano.append(total_producao_ano)
                    expectativa_total_ano.append(expectativa_ano)
                
                # Exibir os dados em formato de tabela
                st.write("### Relação entre Anos")
                df_anos = pd.DataFrame({
                    'Ano': anos_selecionados,
                    'Quantidade Total Produzida': [formatar_valores(val) for val in producao_total_ano],
                    'Expectativa de Produção': [formatar_valores(val) for val in expectativa_total_ano]
                })
                df_anos.index = df_anos.index + 1  # Ajuste de índice para iniciar do 1
                st.dataframe(df_anos, use_container_width=True)

                # Gráfico de setores para a relação entre os anos
                fig_anos = px.pie(pd.DataFrame({
                    'Ano': anos_selecionados,
                    'Quantidade Total Produzida': producao_total_ano
                }), names='Ano', values='Quantidade Total Produzida',
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
                for mes in meses:
                    # Filtrar dados do mês
                    dados_mes = dados_filtrados[dados_filtrados['Mês'] == mes].copy()

                    if producao_tipo == 'Cobre':
                        total_mes = dados_mes['Produção Cobre Realizado'].sum()
                        expectativa_mes = dados_mes['Meta/Dia Cobre'].iloc[0] * len(dados_mes.index)
                    else:
                        total_mes = dados_mes['Produção Alumínio Realizado'].sum()
                        expectativa_mes = dados_mes['Meta/Dia Alumínio'].iloc[0] * len(dados_mes.index) # ajustar método de calculo
                        
                    # Exibir a produção e expectativa do mês fora do toggle
                    col1, col2 = st.columns([1, 1])
                    with col1:
                        st.write(f"**Produção Total em {mes}:** {formatar_valores(total_mes)}")
                    with col2:
                        st.write(f"**Expectativa Total em {mes}:** {formatar_valores(expectativa_mes)}")

                    # Toggle list para mostrar os detalhes do mês
                    with st.expander(f"Exibir detalhes de {mes}"):
                        st.write(f"### Produção Diária - {mes}/{ano_selecionado}")
                        # Exibir DataFrame com a coluna "Dia" como índice
                        dados_mes.set_index([''], inplace=True) # testar passar argumento vazio
                        st.dataframe(dados_mes[['Data', f'Produção {producao_tipo} Realizado', f'Meta/Dia {producao_tipo}']], use_container_width=True)
                        
                # Gráfico de setores para a relação entre os meses do ano selecionado
                df_meses = pd.DataFrame({
                    'Meses': meses,
                    'Quantidade Total Produzida': [dados_filtrados[dados_filtrados['Mês'] == m][f'Produção {producao_tipo} Realizado'].sum() for m in meses],
                    'Expectativa de Produção': [dados_filtrados[dados_filtrados['Mês'] == m][f'Meta/Dia {producao_tipo}'].sum() for m in meses]
                })

                st.write(f"### Distribuição da Produção Mensal - {ano_selecionado}")
                fig_meses = px.pie(df_meses, names='Meses', values='Quantidade Total Produzida',
                                   title=f"Distribuição da Produção nos Meses de {ano_selecionado} - {producao_tipo}")
                st.plotly_chart(fig_meses)

                # Exibir total produzido no ano, expectativa de produção e média anual
                total_anual = df_meses['Quantidade Total Produzida'].sum()
                expectativa_anual = df_meses['Expectativa de Produção'].sum()
                media_anual = df_meses['Quantidade Total Produzida'].mean()

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.write(f"**Total Produzido no Ano**: {formatar_valores(total_anual)}")
                with col2:
                    st.write(f"**Expectativa de Produção no Ano**: {formatar_valores(expectativa_anual)}")
                with col3:
                    st.write(f"**Média de Produção Mensal**: {formatar_valores(media_anual)}")

    else:
        st.write("Erro ao carregar os dados.")

def pagina3():
     st.header('_Demanda por Composto_', divider='gray')
     dados = carregar_dados_demand(DADOS_DEMAND_PATH)
     st.dataframe(dados.iloc[:, 17:])
     
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
st.sidebar.markdown("<br><br><br><br><br><br>", unsafe_allow_html=True)
botao_pagina2 = st.sidebar.button('(ICON2)', on_click=lambda: st.session_state.update({'pagina_atual': 'pagina2'}))
st.sidebar.markdown("<br><br><br><br><br><br>", unsafe_allow_html=True)
botao_pagina3 = st.sidebar.button('(ICON3)', on_click=lambda: st.session_state.update({'pagina_atual': 'pagina3'}))

if st.session_state.pagina_atual == 'pagina1':
    pagina1()
elif st.session_state.pagina_atual == 'pagina2':
    pagina2()
elif st.session_state.pagina_atual == 'pagina3':
    pagina3()