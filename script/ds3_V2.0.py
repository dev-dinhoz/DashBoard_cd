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
        "PVC ANTICHAMA": ["PRETO", "AZUL", "VERMELHO"],
        "PVC CRISTAL": ["VERDE", "AMARELO", "BRANCO", "VERMELHO", "PRETO", "NATURAL"],
        "XLPE": ["PRETO", "NATURAL"],
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
    def extrair_cor(descricao):
        """
        Extrai a cor da descrição do produto.
        A cor é localizada após o primeiro '-' na descrição.
        """
        if pd.notnull(descricao):
            partes = descricao.split("-")
            if len(partes) > 1:
                return partes[1].strip().split(" ")[0].upper()
        return "INDEFINIDO"

# Funções de Carregamento de Dados
@st.cache_data
def carregar_dados_producao():
    """
    Carrega os dados do programa de extrusão.
    As colunas começam na coluna B e a partir da quinta linha da aba 'ProgramaExtrusão'.
    """
    try:
        dados = pd.read_excel(DADOS_PRODUCAO_PATH, sheet_name="ProgramaExtrusão", header=4, usecols="B:AZ")

        # Ajustar os nomes e cores
        dados["Composto"] = dados["DESCRIÇÃO"].apply(Polimero.ajustar_nome)
        dados["Cor"] = dados["DESCRIÇÃO"].apply(Polimero.extrair_cor)

        return dados
    except Exception as e:
        st.error(f"Erro ao carregar os dados de produção: {e}")
        return None
@st.cache_data
def carregar_dados_estoque():
    """
    Carrega os dados do almoxarifado.
    O peso é extraído da última coluna preenchida.
    """
    try:
        dados = pd.read_excel(DADOS_ESTOQUE_PATH, sheet_name="Folha1", header=1)

        # Identificar a última coluna preenchida
        ultima_coluna = dados.iloc[:, 6:].columns[dados.iloc[:, 6:].notnull().any()[-1]]

        # Selecionar colunas relevantes
        dados_estoque = dados[["Produto", ultima_coluna]].dropna()
        dados_estoque = dados_estoque.rename(columns={ultima_coluna: "Estoque (kg)"})

        # Ajustar os nomes dos produtos
        dados_estoque["Produto"] = dados_estoque["Produto"].apply(Polimero.ajustar_nome)

        return dados_estoque
    except Exception as e:
        st.error(f"Erro ao carregar os dados de estoque: {e}")
        return None
# Processamento de Demanda
def processar_demanda(dados_producao):
    """
    Processa os dados de produção para calcular a demanda total e por cor.
    """
    if "Composto" not in dados_producao.columns or "Cor" not in dados_producao.columns:
        st.error("Colunas 'Composto' ou 'Cor' ausentes nos dados de produção.")
        return None, None

    # Agrupar os dados por Composto e Cor
    demanda_por_cor = (
        dados_producao.groupby(["Composto", "Cor"]).sum().reset_index()
    )
    demanda_total = demanda_por_cor.groupby("Composto").sum()["PESO"]
    return demanda_total, demanda_por_cor
# Comparação de Demanda e Estoque
def comparar_demanda_estoque(demanda_total, dados_estoque):
    """Compara a demanda total com o estoque disponível."""
    estoque_atual = dados_estoque.set_index("Produto")["Estoque (kg)"].reindex(demanda_total.index, fill_value=0)
    saldo = demanda_total.sum(axis=1) - estoque_atual

    return pd.DataFrame({
        "Composto": demanda_total.index,
        "Demanda (kg)": demanda_total.sum(axis=1),
        "Estoque Atual (kg)": estoque_atual,
        "Saldo (kg)": saldo
    })

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
                grupo = grupo.sort_values("Demanda (kg)", ascending=False)
                st.dataframe(grupo)

    else:
        st.error("Erro ao carregar os dados das planilhas.")

# Interface do sistema
st.set_page_config(page_title="Dashboard de Polímeros", layout="wide")

# Navegação entre páginas
if st.sidebar.button("Demanda de Polímeros"):
    pagina3()
