import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Configurações gerais
BASE_DIR = os.getcwd()
DADOS_PRODUCAO_PATH = os.path.join(BASE_DIR,'.database', 'DATABASE.xlsx')  # Programa de Extrusão
DADOS_ESTOQUE_PATH = os.path.join(BASE_DIR,'.database', 'Novembro-2024.xlsx')  # Almoxarifado


# Funções utilitárias
@st.cache_data
def carregar_dados_producao(caminho):
    """Carrega dados do programa de extrusão."""
    dados = pd.read_excel(caminho, sheet_name="ProgramaExtrusão", header=4)
    return dados


@st.cache_data
def carregar_dados_estoque(caminho):
    """Carrega dados do almoxarifado e retorna a última coluna de estoque."""
    dados = pd.read_excel(caminho, sheet_name="Folha1", header=1)
    ultima_data = dados.columns[-1]  # Última coluna com dados do estoque
    return dados, ultima_data


def processar_demanda(dados_producao):
    """Processa as demandas por composto e organiza por cor."""
    # Colunas descritivas e colunas de compostos
    colunas_descritivas = dados_producao.columns[:18]
    colunas_compostos = dados_producao.columns[18:]

    # Extrair a cor do material da descrição
    def extrair_cor(descricao):
        if pd.isna(descricao):
            return "Indefinido"
        if '-' in descricao:
            return descricao.split('-')[-1].strip()
        return "Indefinido"

    dados_producao["Cor"] = dados_producao["DESCRIÇÃO"].apply(extrair_cor)

    # Somar demandas por composto
    demanda_total = dados_producao[colunas_compostos].sum()
    demanda_por_cor = (
        dados_producao.groupby("Cor")[colunas_compostos].sum().reset_index()
    )
    return demanda_total, demanda_por_cor


def comparar_demanda_estoque(demanda_total, dados_estoque, ultima_data):
    """Compara a demanda total de compostos com o estoque disponível."""
    estoque_atual = (
        dados_estoque.set_index("Produto")[ultima_data]
        .reindex(demanda_total.index, fill_value=0)
        .fillna(0)
    )
    saldo = demanda_total - estoque_atual

    resultado = pd.DataFrame({
        "Composto": demanda_total.index,
        "Demanda (kg)": demanda_total.values,
        "Estoque Atual (kg)": estoque_atual.values,
        "Saldo (kg)": saldo.values,
    })
    return resultado


def exibir_detalhes_composto(resultado, demanda_por_cor):
    """Cria expansores com detalhes de compostos e informações por cor."""
    for composto in resultado["Composto"]:
        with st.expander(f"Detalhes do Composto: {composto}"):
            if composto in demanda_por_cor.columns:
                dados_detalhados = demanda_por_cor[["Cor", composto]]
                st.dataframe(dados_detalhados)


# Página principal  
def pagina_demanda_polimeros():
    st.title("Demanda de Polímeros")
    st.write("### Comparação entre demanda de produção e estoque de polímeros")

    # Carregar dados
    dados_producao = carregar_dados_producao(DADOS_PRODUCAO_PATH)
    dados_estoque, ultima_data = carregar_dados_estoque(DADOS_ESTOQUE_PATH)

    if dados_producao is not None and dados_estoque is not None:
        # Processar dados de demanda
        demanda_total, demanda_por_cor = processar_demanda(dados_producao)

        # Comparar demanda e estoque
        resultado_comparacao = comparar_demanda_estoque(demanda_total, dados_estoque, ultima_data)

        # Exibir resultados gerais
        st.subheader("Resumo Geral")
        st.dataframe(
            resultado_comparacao.style.applymap(
                lambda x: "background-color: red" if x < 0 else "background-color: green",
                subset=["Saldo (kg)"]
            )
        )

        # Gráfico de distribuição
        st.subheader("Distribuição de Demanda por Composto")
        fig = px.bar(
            resultado_comparacao,
            x="Composto",
            y=["Demanda (kg)", "Estoque Atual (kg)"],
            barmode="group",
            title="Comparativo de Demanda e Estoque"
        )
        st.plotly_chart(fig)

        # Exibir detalhes por composto
        exibir_detalhes_composto(resultado_comparacao, demanda_por_cor)
    else:
        st.error("Erro ao carregar os dados das planilhas.")


# Interface do sistema
st.set_page_config(page_title="Dashboard Operacional", layout="wide")

# Exibir a página
pagina_demanda_polimeros()
