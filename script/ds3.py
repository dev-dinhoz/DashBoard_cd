import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Configurações gerais
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DADOS_MONITORING_PATH = os.path.join(BASE_DIR, '.database', 'DATABASE.xlsx')
DADOS_ALMOXARIFADO_PATH = os.path.join(BASE_DIR, '.database', 'Novembro-2024.xlsx')


@st.cache_data
def carregar_dados_almoxarifado(path):
    """Carregar dados do almoxarifado e retornar DataFrame e última coluna."""
    df = pd.read_excel(path, sheet_name=0, header=1)
    df.columns = df.columns.str.strip()
    ultima_data = df.columns[-1]  # Última coluna com dados
    return df, ultima_data


@st.cache_data
def carregar_dados_monitoring(path):
    """Carregar dados do programa de extrusão e retornar DataFrame."""
    dados = pd.read_excel(path, sheet_name='ProgramaExtrusão', header=4)
    return dados


def calcular_demanda_por_composto(dados):
    """Calcular a demanda total por composto (somar colunas específicas)."""
    colunas_compostos = dados.columns[18:]  # Ajustar o índice conforme a planilha
    return dados[colunas_compostos].sum()


def comparar_demanda_estoque(demanda, estoque, ultima_data):
    """
    Comparar demanda com estoque:
    - demanda: DataFrame de demanda por composto.
    - estoque: DataFrame de estoque.
    - ultima_data: Coluna com o estoque mais recente.
    """
    estoque_atual = estoque.set_index("Produto")[ultima_data].fillna(0)
    saldo = demanda - estoque_atual
    resultado = pd.DataFrame({
        "Composto": demanda.index,
        "Demanda (kg)": demanda.values,
        "Estoque Atual (kg)": estoque_atual.reindex(demanda.index, fill_value=0).values,
        "Saldo (kg)": saldo.values
    })
    return resultado


def exibir_detalhes_composto(resultado, dados_monitoring):
    """Expandir compostos com informações específicas, detalhando por cor."""
    for composto in resultado["Composto"].unique():
        with st.expander(f"Detalhes do Composto: {composto}"):
            dados_filtro = dados_monitoring[dados_monitoring[composto] > 0]
            if not dados_filtro.empty:
                st.dataframe(dados_filtro)


# Página principal do Streamlit
def pagina3():
    st.title("Demanda de Polímeros")
    st.write("### Comparativo entre demanda de produção e estoque de polímeros.")

    # Carregar dados das planilhas
    dados_producao = carregar_dados_monitoring(DADOS_MONITORING_PATH)
    dados_almoxarifado, ultima_data = carregar_dados_almoxarifado(DADOS_ALMOXARIFADO_PATH)

    if dados_producao is not None and dados_almoxarifado is not None:
        # Processar demanda e estoque
        demanda_compostos = calcular_demanda_por_composto(dados_producao)
        resultado_comparacao = comparar_demanda_estoque(demanda_compostos, dados_almoxarifado, ultima_data)

        # Exibir saldo geral
        st.subheader("Resumo Geral")
        st.dataframe(resultado_comparacao.style.applymap(
            lambda x: 'background-color: red' if x < 0 else 'background-color: green',
            subset=["Saldo (kg)"]
        ))

        # Gráfico comparativo
        st.subheader("Distribuição de Demanda por Composto")
        fig = px.bar(
            resultado_comparacao,
            x="Composto",
            y=["Demanda (kg)", "Estoque Atual (kg)"],
            barmode="group",
            title="Comparativo de Demanda e Estoque"
        )
        st.plotly_chart(fig)

        # Detalhamento por composto
        exibir_detalhes_composto(resultado_comparacao, dados_producao)

    else:
        st.error("Erro ao carregar os dados das planilhas.")


# Interface do sistema
st.set_page_config(page_title="Dashboard de Polímeros", layout="wide")

# Navegação entre páginas
if st.sidebar.button("Demanda de Polímeros"):
    pagina3()
