import locale
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import plotly.express as px

# Configurações gerais
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DADOS_MONITORING_PATH = os.path.join(BASE_DIR, '.database', 'DATABASE.xlsx')
DADOS_ALMMOXERIFADO_PATH = os.path.join(BASE_DIR, '.database', 'Novembro-2024.xlsx')

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
locale.atof = lambda x: float(x.replace('.', '').replace(',', '.'))

def carregar_dados_monitoring():
    return pd.read_excel(DADOS_MONITORING_PATH, header=4, sheet_name='ProgramaExtrusão')

def carregar_dados_almoxarifado():
    df = pd.read_excel(DADOS_ALMMOXERIFADO_PATH, sheet_name=0, header=1)
    df.columns = df.columns.str.strip()
    return df

def comparar_demanda_estoque(demanda, estoque):
    saldo = demanda - estoque
    return pd.DataFrame({
        "Composto": demanda.index,
        "Demanda (kg)": demanda.values,
        "Estoque Atual (kg)": estoque.values,
        "Saldo (kg)": saldo
    })

def pagina3():
    st.header('_Demanda de Polímeros_', divider='gray')
    dados_producao = carregar_dados_monitoring()
    dados_almoxarifado = carregar_dados_almoxarifado()

    if dados_producao is not None and dados_almoxarifado is not None:
        colunas_descritivas = dados_producao.columns[:18]
        colunas_compostos = dados_producao.columns[18:]
        demanda_compostos = dados_producao[colunas_compostos].sum()

        try:
            # Preparar filtro no almoxarifado
            dados_almoxarifado['Produto'] = dados_almoxarifado['Produto'].str.strip()
            colunas_compostos = [col.strip() for col in colunas_compostos]
            estoque_filtrado = dados_almoxarifado.set_index("Produto")
            estoque_atual = estoque_filtrado.loc[colunas_compostos, '19/nov'].fillna(0)
        except KeyError as e:
            st.error(f"Erro ao filtrar os dados de estoque: {e}")
            st.write("Produtos disponíveis no almoxarifado:", estoque_filtrado.index.tolist())
            st.write("Colunas disponíveis no almoxarifado:", estoque_filtrado.columns.tolist())
            return

        resultado_comparacao = comparar_demanda_estoque(demanda_compostos, estoque_atual)

        st.subheader("Comparativo de Demanda e Estoque")
        st.dataframe(resultado_comparacao)

        compostos_deficit = resultado_comparacao[resultado_comparacao["Saldo (kg)"] < 0]
        if not compostos_deficit.empty:
            st.subheader("Compostos com Necessidade de Reposição")
            st.dataframe(compostos_deficit)

        fig_compostos = px.pie(
            resultado_comparacao,
            names='Composto',
            values='Demanda (kg)',
            title="Distribuição dos Compostos na Produção"
        )
        st.plotly_chart(
            resultado_comparacao,
            names='Composto',
            values='Demanda (kg)',
            title="Distribuição dos Compostos na Produção"
        )
        st.plotly_chart(fig_compostos)
    else:
        st.error("Erro ao carregar os dados da aba 'Programa Extrusão' ou 'Almoxarifado'.")

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