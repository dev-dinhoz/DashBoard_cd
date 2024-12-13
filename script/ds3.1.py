import streamlit as st
import pandas as pd
import plotly.express as px
import os
import locale 

# Configurações gerais
BASE_DIR = os.getcwd()
DADOS_PRODUCAO_PATH = os.path.join(BASE_DIR, ".database", "DATABASE.xlsx")  # Programa de Extrusão
DADOS_ESTOQUE_PATH = os.path.join(BASE_DIR, ".database", "Novembro-2024 - Copia.xlsx")  # Almoxarifado

st.set_page_config(page_title="Dashboard Operacional", layout="wide")

# Definir o formato de números como pt-BR
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
locale.atof = lambda x: float(x.replace('.', '').replace(',', '.'))  
# (talvez seja inútilKKK, mas deixa ai, não vamos mexer no que está quieto) Ignorar separadores de milhar e considerar apenas 2 casas decimais

def formatar_valores(valor):
    """ Formatar valores numéricos no formato 10.000,00 """
    return locale.format_string('%.2f', valor)

# Funções utilitárias
@st.cache_data
def carregar_dados(caminho, aba, header):
    """Carrega dados de uma planilha Excel."""
    try:
        dados = pd.read_excel(caminho, sheet_name=aba, header=header)
        return dados
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return None

def carregar_dados_estoque(caminho):
    """Carrega dados do almoxarifado, soma os estoques por produto e retorna o saldo."""
    try:
        st.write("Iniciando carregamento de dados de estoque...")
        st.write(f"Lendo a planilha do caminho: {caminho}")
        
        # Carregar planilha
        dados = pd.read_excel(caminho, sheet_name="Folha1", header=1)
        st.write("Dados carregados com sucesso! Primeiras linhas:")
        st.dataframe(dados.head())
        
        # Identificar colunas de datas (última com dados válidos)
        dados = pd.read_exel (caminho, sheet_name="Folha1", header=1 )
        
        colunas_datas = dados.columns[6:]  # As datas começam na coluna 6
        ultima_coluna_valida = next(
            (col for col in reversed(colunas_datas) if dados[col].notnull().any()), None
        )
        if not ultima_coluna_valida:
            raise ValueError("Nenhuma coluna com valores atualizados foi encontrada.")
        
        st.write(f"Última coluna válida encontrada: {ultima_coluna_valida}")
        
        # Selecionar a última coluna de estoque
        dados_estoque = dados[["Unnamed: 2", ultima_coluna_valida]].copy()
        dados_estoque.columns = ["Produto", "Estoque (kg)"]  # Renomear colunas
        dados_estoque = dados_estoque.dropna(subset=["Produto", "Estoque (kg)"])
        
        # Agrupar por produto e somar o estoque
        dados_estoque = dados_estoque.groupby("Produto", as_index=False).sum()
        
        st.write("Dados de estoque processados com sucesso:")
        st.dataframe(dados_estoque.head())
        return dados_estoque
    except Exception as e:
        st.error(f"Erro ao carregar os dados de estoque: {e}")
        return None

def extrair_cor(descricao):
    """Extrai a cor de um material a partir da descrição."""
    if pd.isna(descricao):
        return "Indefinido"
    if '-' in descricao:
        return descricao.split('-')[-1].strip()
    return "Indefinido"

def processar_demanda(dados_producao):
    """Processa demandas e organiza por composto e cor."""
    # Identificar colunas descritivas e compostos
    colunas_descritivas = dados_producao.columns[:18]
    colunas_compostos = dados_producao.columns[18:]

    # Adicionar coluna de cor
    dados_producao["Cor"] = dados_producao["DESCRIÇÃO"].apply(extrair_cor)

    # Soma total por composto
    demanda_total = dados_producao[colunas_compostos].sum()

    # Soma por cor
    demanda_por_cor = (
        dados_producao.groupby("Cor")[colunas_compostos].sum().reset_index()
    )
    return demanda_total, demanda_por_cor

# Comparar demanda e estoque no dataframe principal
def comparar_demanda_estoque(demanda_total, dados_estoque):
    """Compara a demanda total de compostos com o estoque disponível."""
    # Ajustar índices do estoque para alinhar com a demanda
    estoque_atual = dados_estoque.set_index("Produto")["Estoque (kg)"]

    # Reindexar para garantir alinhamento
    estoque_atual = estoque_atual.reindex(demanda_total.index, fill_value=0)

    # Calcular o saldo
    saldo = demanda_total - estoque_atual

    # Criar dataframe de resultado
    resultado = pd.DataFrame({
        "Composto": demanda_total.index,
        "Demanda (kg)": demanda_total.values,
        "Estoque Atual (kg)": estoque_atual.values,
        "Saldo (kg)": saldo.values,
    })
    return resultado

def exibir_detalhes_composto(resultado, demanda_por_cor):
    """Organiza os expansores e as informações internas em colunas, ordenando por valor."""
    colunas_por_linha_expansores = 3  # Número de expansores por linha
    colunas_por_linha_info = 2  # Número de colunas dentro de cada expansor

    num_compostos = len(resultado)

    # Criar linhas de expansores
    linhas_expansores = [
        resultado.iloc[i : i + colunas_por_linha_expansores]
        for i in range(0, num_compostos, colunas_por_linha_expansores)
    ]

    # Renderizar expansores organizados em colunas
    for linha in linhas_expansores:
        cols = st.columns(colunas_por_linha_expansores)
        for col, composto_data in zip(cols, linha.itertuples()):
            with col:
                with st.expander(f"Detalhes do Composto: {composto_data.Composto}"):
                    # Exibir informações detalhadas por cor em grade
                    if composto_data.Composto in demanda_por_cor.columns:
                        dados_detalhados = demanda_por_cor[["Cor", composto_data.Composto]]
                        dados_detalhados = dados_detalhados.rename(
                            columns={composto_data.Composto: "Demanda (kg)"}
                        )

                        # Ordenar os dados por valor de demanda (decrescente)
                        dados_detalhados = dados_detalhados.sort_values(
                            by="Demanda (kg)", ascending=False
                        )

                        # Organizar os detalhes em colunas
                        num_infos = len(dados_detalhados)
                        linhas_info = [
                            dados_detalhados.iloc[i : i + colunas_por_linha_info]
                            for i in range(0, num_infos, colunas_por_linha_info)
                        ]

                        for linha_info in linhas_info:
                            cols_info = st.columns(colunas_por_linha_info)
                            for col_info, info in zip(cols_info, linha_info.itertuples()):
                                with col_info:
                                    st.metric(
                                        label=f"Cor: {info.Cor}",
                                        value=f"{info._2} kg",
                                    )

# Página principal
def pagina_demanda_polimeros():
    st.title("Demanda de Polímeros")
    st.write("### Comparação entre demanda de produção e estoque de polímeros")
    
    # Carregar dados de produção
    dados_producao = carregar_dados(DADOS_PRODUCAO_PATH, "ProgramaExtrusão", 4)
    
    # Carregar dados de estoque
    dados_estoque = carregar_dados_estoque(DADOS_ESTOQUE_PATH)
    
    if dados_producao is not None and dados_estoque is not None:
        # Processar demanda
        demanda_total, demanda_por_cor = processar_demanda(dados_producao)
        
        # Comparar demanda e estoque
        resultado_comparacao = comparar_demanda_estoque(demanda_total, dados_estoque)
        
        # Exibir resultados gerais
        st.subheader("Resumo Geral")
        st.dataframe(
            resultado_comparacao.style.applymap(
                lambda x: "background-color: red" if x < 0 else "background-color: green",
                subset=["Saldo (kg)"]
            )
        )
        
        # Gráfico de barras
        st.subheader("Distribuição de Demanda por Composto")
        fig = px.bar(
            resultado_comparacao,
            x="Composto",
            y=["Demanda (kg)", "Estoque Atual (kg)"],
            barmode="group",
            title="Comparativo de Demanda e Estoque"
        )
        st.plotly_chart(fig)
        
        # Detalhes por composto
        exibir_detalhes_composto(resultado_comparacao, demanda_por_cor)
    else:
        st.error("Erro ao carregar os dados das planilhas.")

# Configuração e exibição
pagina_demanda_polimeros()