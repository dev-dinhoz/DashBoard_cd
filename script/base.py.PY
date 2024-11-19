import locale
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import plotly.express as px

# Definir o diretório base como o caminho do próprio script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DADOS_MONITORING_PATH = os.path.join(BASE_DIR, '.database', 'DATABASE.xlsx')
DADOS_ALMMOXERIFADO_PATH = os.path.join(BASE_DIR, '.database', 'Novembro-2024.xlsx')

# Definir o formato de números como pt-BR
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
locale.atof = lambda x: float(x.replace('.', '').replace(',', '.'))  # Ignorar separadores de milhar e considerar apenas 2 casas decimais

def formatar_valores(valor):
    """ Formatar valores numéricos no formato 10.000,00 """
    return locale.format_string('%.2f', valor / 1000, grouping=True)  # Dividimos por 1000 para mostrar em milhares

def carregar_dados_monitoring():
    # Carregar a aba "Programa Extrusão"
    df = pd.read_excel(DADOS_MONITORING_PATH, header=4, sheet_name='ProgramaExtrusão')
    return df


def pagina3():

    # Facilita a organização das compras de polimeros para a empresa, com base na quantidade total em peso.
    # Balanceamento da demanda de utilização de polimeros para proução menos a quantiade de estoque de cada material.
    # Futuramente adicionar balanceamento das cores (preciso de outa base de dados para isso), por enquanto exibir um data frame no priprio site para indicar esse controle

    st.header('_Demanda de Polimeros_', divider='gray')
    
    # Carregar os dados
    dados = carregar_dados_monitoring()
    
    if dados is not None:
        # Dividir colunas descritivas e colunas de compostos (considerando que compostos iniciam a partir da coluna S)
        colunas_descritivas = dados.columns[:18]
        colunas_compostos = dados.columns[18:]
        
        # Separar dados
        dados_descritivos = dados[colunas_descritivas]
        dados_compostos = dados[colunas_compostos]

        # Exibir informações iniciais
        st.subheader("Informações Gerais dos Produtos")
        st.dataframe(dados_descritivos)
        
        # Calcular o total de horas de produção
        if "Tot Hrs" in dados_descritivos.columns:
            total_horas = dados_descritivos["Tot Hrs"].sum()
            st.write(f"**Total de Horas de Produção:** {formatar_valores(total_horas)} mil horas")

        # Calcular a distribuição dos compostos
        st.subheader("Distribuição dos Compostos Utilizados")
        total_compostos = dados_compostos.sum().sort_values(ascending=False)
        st.dataframe(total_compostos.rename("Quantidade Total"), use_container_width=True)

        # Ajuste no gráfico de pizza
        df_compostos = total_compostos.reset_index()
        df_compostos.columns = ['Composto', 'Quantidade Total']
        
        fig_compostos = px.pie(df_compostos, 
                               names='Composto', values='Quantidade Total',
                               title="Distribuição dos Compostos na Produção")
        st.plotly_chart(fig_compostos)
        
    else:
        st.error("Erro ao carregar os dados da aba 'Programa Extrusão'.")

# Interface do sistema
st.set_page_config(page_title="Dashboard", page_icon="💡", layout="wide")

imagem_caminho = os.path.join(BASE_DIR, '.uploads', 'Logo.png')
if os.path.exists(imagem_caminho):
    st.sidebar.image(imagem_caminho, use_column_width=True)
else:
    st.sidebar.error(f"Imagem no caminho '{imagem_caminho}' não encontrada.")

if 'pagina_atual' not in st.session_state:
    st.session_state.pagina_atual = 'pagina1'

st.sidebar.markdown("<br><br><br>", unsafe_allow_html=True)
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