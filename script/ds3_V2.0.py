import streamlit as st
import pandas as pd
import plotly.express as px
import os
import locale

# Configurações Gerais
BASE_DIR = os.getcwd()
DADOS_PRODUCAO_PATH = os.path.join(BASE_DIR, ".database", "DATABASE.xlsx")  # Programa de Extrusão
DADOS_ESTOQUE_PATH = os.path.join(BASE_DIR, ".database", "Novembro-2024.xlsx")  # Almoxarifado

# Configuração Regional para Números
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
locale.atof = lambda x: float(x.replace('.', '').replace(',', '.'))


def formatar_valores(valor):
    """Formata valores numéricos no estilo 10.000,00."""
    return locale.format_string('%.2f', valor, grouping=True)


# Classe para Mapear e Ajustar Compostos e Colorações
class Polimero:
    MAPA_POLIMEROS = {
        "PVC ANTICHAMA": ["PRETO", "AZUL", "VERMELHO", "VERDE", "AMARELO", "BRANCO"],
        "PVC CRISTAL": ["PRETO", "AZUL", "VERMELHO", "VERDE", "AMARELO", "BRANCO"],
        "XLPE": ["PRETO", "NATURAL"],
        "ST2": ["PRETO", "NATURAL", "UV"],
        "ST1": ["PRETO"]
    }

    @staticmethod
    def ajustar_nome(nome):
        """Ajusta nomes para mapear compostos conhecidos."""
        nome = str(nome).upper() if pd.notnull(nome) else "INDEFINIDO"
        for chave in Polimero.MAPA_POLIMEROS:
            if chave in nome:
                return chave
        return nome

    @staticmethod
    def ajustar_cor(nome):
        """Identifica a cor com base no nome."""
        nome = str(nome).upper() if pd.notnull(nome) else "INDEFINIDO"
        for chave, cores in Polimero.MAPA_POLIMEROS.items():
            if chave in nome:
                for cor in cores:
                    if cor in nome:
                        return cor
        return "Indefinido"


# Funções de Carregamento de Dados
@st.cache_data
def carregar_dados_producao():
    """Carrega os dados do programa de extrusão."""
    try:
        dados = pd.read_excel(DADOS_PRODUCAO_PATH, sheet_name="ProgramaExtrusão", header=4, usecols="B:AI")

        # Ajustar os dados de produção
        dados["Composto"] = dados["DESCRIÇÃO"].apply(Polimero.ajustar_nome)
        dados["Cor"] = dados["DESCRIÇÃO"].apply(Polimero.ajustar_cor)

        return dados
    except Exception as e:
        st.error(f"Erro ao carregar os dados de produção: {e}")
        return None


@st.cache_data
def carregar_dados_estoque():
    """Carrega os dados do almoxarifado, localizando a última coluna de atualização."""
    try:
        dados = pd.read_excel(DADOS_ESTOQUE_PATH, sheet_name="Folha1", header=1)

        # Identificar a última coluna com dados
        colunas_datas = dados.columns[6:]  # A partir da 7ª coluna
        ultima_coluna = next((col for col in reversed(colunas_datas) if dados[col].notnull().any()), None)

        if not ultima_coluna:
            raise ValueError("Nenhuma coluna válida encontrada no estoque.")

        # Ajustar os dados de estoque
        dados_estoque = dados[["Produto", ultima_coluna]].dropna(subset=["Produto", ultima_coluna])
        dados_estoque = dados_estoque.rename(columns={ultima_coluna: "Estoque (kg)"})
        dados_estoque["Produto"] = dados_estoque["Produto"].apply(Polimero.ajustar_nome)
        dados_estoque["Cor"] = dados_estoque["Produto"].apply(Polimero.ajustar_cor)

        return dados_estoque
    except Exception as e:
        st.error(f"Erro ao carregar os dados de estoque: {e}")
        return None


# Processamento de Demanda
def processar_demanda(dados_producao):
    """
    Processa os dados de produção para calcular a demanda total e a demanda por cor.
    """
    try:
        if "Composto" not in dados_producao.columns or "Cor" not in dados_producao.columns:
            raise KeyError("Colunas 'Composto' ou 'Cor' ausentes nos dados de produção.")

        # Agrupar os dados por Composto e Cor
        demanda_por_cor = (
            dados_producao.groupby(["Composto", "Cor"])["PESO"].sum().reset_index()
        )
        demanda_total = demanda_por_cor.groupby("Composto")["PESO"].sum()

        return demanda_total, demanda_por_cor
    except Exception as e:
        st.error(f"Erro ao processar a demanda: {e}")
        return None, None


# Comparação de Demanda e Estoque
def comparar_demanda_estoque(demanda_total, dados_estoque):
    """Compara a demanda total com o estoque disponível."""
    try:
        estoque_atual = dados_estoque.set_index("Produto")["Estoque (kg)"].reindex(demanda_total.index, fill_value=0)
        saldo = demanda_total - estoque_atual

        return pd.DataFrame({
            "Composto": demanda_total.index,
            "Demanda (kg)": demanda_total,
            "Estoque Atual (kg)": estoque_atual,
            "Saldo (kg)": saldo
        }).reset_index(drop=True)
    except Exception as e:
        st.error(f"Erro ao comparar demanda e estoque: {e}")
        return None


# Interface do Dashboard
def pagina3():
    st.header("_Demanda de Polímeros_", divider="gray")

    dados_producao = carregar_dados_producao()
    dados_estoque = carregar_dados_estoque()

    if dados_producao is not None and dados_estoque is not None:
        # Processar dados
        demanda_total, demanda_por_cor = processar_demanda(dados_producao)
        resultado = comparar_demanda_estoque(demanda_total, dados_estoque)

        # Exibição de Dados
        st.subheader("Resumo Geral")
        if resultado is not None:
            st.dataframe(
                resultado.style.applymap(
                    lambda x: "background-color: red" if x < 0 else "background-color: green",
                    subset=["Saldo (kg)"]
                )
            )

        # Gráfico Comparativo
        st.subheader("Distribuição de Demanda por Composto")
        fig = px.bar(
            resultado,
            x="Composto",
            y=["Demanda (kg)", "Estoque Atual (kg)"],
            barmode="group",
            title="Comparativo de Demanda e Estoque"
        )
        st.plotly_chart(fig)

        # Detalhes por Composto
        for composto, grupo in demanda_por_cor.groupby("Composto"):
            with st.expander(f"Detalhes do Composto: {composto}"):
                grupo = grupo.sort_values("PESO", ascending=False)
                st.dataframe(grupo)

    else:
        st.error("Erro ao carregar os dados das planilhas.")


# Configuração da Interface
st.set_page_config(page_title="Dashboard de Polímeros", layout="wide")

if st.sidebar.button("Demanda de Polímeros"):
    pagina3()
