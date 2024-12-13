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

# Configuração para formato de números em pt-BR
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
locale.atof = lambda x: float(x.replace('.', '').replace(',', '.'))  # Ajustar separadores de milhar

def formatar_valores(valor):
    """ Formatar valores numéricos no formato 10.000,00 """
    return locale.format_string('%.2f', valor)

# Dicionário para exceções manuais de grupos
EXCECOES_GRUPOS = {
    "GrupoInvalido1": "GrupoCorreto1",
    "GrupoInvalido2": "GrupoCorreto2",
    # Adicione mais exceções conforme necessário
}

# Funções utilitárias
def carregar_dados(caminho, aba, header):
    """Carrega dados de uma planilha Excel."""
    try:
        dados = pd.read_excel(caminho, sheet_name=aba, header=header)
        return dados
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return None

def carregar_dados_estoque(caminho):
    """Carrega dados do almoxarifado, soma os estoques por grupo e retorna o saldo."""
    try:
        dados = pd.read_excel(caminho, sheet_name="Folha1", header=1)

        # Identificar as colunas de datas
        colunas_datas = dados.columns[6:]
        ultima_coluna_valida = next(
            (col for col in reversed(colunas_datas) if dados[col].notnull().any()), None
        )

        if not ultima_coluna_valida:
            raise ValueError("Nenhuma coluna com valores atualizados foi encontrada.")

        # Extrair dados relevantes
        dados_estoque = dados[["Produto", ultima_coluna_valida]].copy()
        dados_estoque = dados_estoque.rename(columns={ultima_coluna_valida: "Estoque (kg)"})
        dados_estoque = dados_estoque.dropna(subset=["Produto", "Estoque (kg)"])
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

def identificar_grupo_dinamico(composto, grupos_demanda):
    """Identifica o grupo de um composto com base na demanda e nas exceções."""
    for grupo in grupos_demanda:
        if grupo in composto:
            return grupo
    for grupo_invalido, grupo_correto in EXCECOES_GRUPOS.items():
        if grupo_invalido in composto:
            return grupo_correto
    return "OUTROS"

def extrair_grupos_dinamicos(demanda_total):
    """Extrai os grupos a partir da demanda."""
    return [composto for composto in demanda_total.index if not composto.startswith("Cor")]

def processar_demanda(dados_producao):
    """Processa demandas e organiza por composto e cor."""
    colunas_descritivas = dados_producao.columns[:18]
    colunas_compostos = dados_producao.columns[18:]
    dados_producao["Cor"] = dados_producao["DESCRIÇÃO"].apply(extrair_cor)
    demanda_total = dados_producao[colunas_compostos].sum()
    demanda_por_cor = dados_producao.groupby("Cor")[colunas_compostos].sum().reset_index()
    return demanda_total, demanda_por_cor

def agregar_dados_por_grupo(demanda_total, dados_estoque):
    """Consolida a demanda e o estoque por grupo."""
    grupos_demanda = extrair_grupos_dinamicos(demanda_total)
    demanda_total = demanda_total.reset_index()
    demanda_total["Grupo"] = demanda_total["index"].apply(
        lambda composto: identificar_grupo_dinamico(composto, grupos_demanda)
    )
    demanda_agrupada = demanda_total.groupby("Grupo")[0].sum().reset_index()
    demanda_agrupada = demanda_agrupada.rename(columns={0: "Demanda (kg)"})

    dados_estoque["Grupo"] = dados_estoque["Produto"].apply(
        lambda produto: identificar_grupo_dinamico(produto, grupos_demanda)
    )
    estoque_agrupado = dados_estoque.groupby("Grupo")["Estoque (kg)"].sum().reset_index()
    estoque_agrupado = estoque_agrupado.rename(columns={"Estoque (kg)": "Estoque Total (kg)"})

    df_consolidado = pd.merge(
        demanda_agrupada, estoque_agrupado, on="Grupo", how="outer"
    ).fillna(0)
    df_consolidado["Saldo (kg)"] = df_consolidado["Estoque Total (kg)"] - df_consolidado["Demanda (kg)"]
    return df_consolidado

def pagina_demanda_polimeros():
    st.title("Demanda de Polímeros")
    st.write("### Comparação entre demanda de produção e estoque de polímeros")

    dados_producao = carregar_dados(DADOS_PRODUCAO_PATH, "ProgramaExtrusão", 4)
    dados_estoque = carregar_dados_estoque(DADOS_ESTOQUE_PATH)

    if dados_producao is not None and dados_estoque is not None:
        demanda_total, demanda_por_cor = processar_demanda(dados_producao)
        df_consolidado = agregar_dados_por_grupo(demanda_total, dados_estoque)

        st.subheader("Resumo por Grupo")
        st.dataframe(
            df_consolidado.style.applymap(
                lambda x: "background-color: red" if x < 0 else "background-color: green",
                subset=["Saldo (kg)"]
            )
        )

        st.subheader("Distribuição por Grupo")
        fig = px.bar(
            df_consolidado,
            x="Grupo",
            y=["Demanda (kg)", "Estoque Total (kg)"],
            barmode="group",
            title="Comparativo de Demanda e Estoque por Grupo"
        )
        st.plotly_chart(fig)

        st.subheader("Detalhes por Composto")
        for composto in demanda_por_cor.columns[1:]:
            with st.expander(f"Composto: {composto}"):
                detalhamento = demanda_por_cor[["Cor", composto]].rename(
                    columns={composto: "Demanda (kg)"}
                )
                detalhamento = detalhamento.sort_values(by="Demanda (kg)", ascending=False)
                
                st.dataframe(detalhamento)
    else:
        st.error("Erro ao carregar os dados das planilhas.")

# Execução
pagina_demanda_polimeros()
