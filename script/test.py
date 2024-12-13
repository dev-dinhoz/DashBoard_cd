import pandas as pd
import os

# Configurações gerais
BASE_DIR = os.getcwd()
DADOS_PRODUCAO_PATH = os.path.join(BASE_DIR, ".database", "DATABASE.xlsx")  # Programa de Extrusão
DADOS_ESTOQUE_PATH = os.path.join(BASE_DIR, ".database", "Novembro-2024 - Copia.xlsx")  # Almoxarifado

def carregar_dados_estoque(caminho):
    """Carrega dados do almoxarifado e retorna os estoques na última data válida."""
    try:
        # Carregar a planilha completa
        print("Lendo a planilha do caminho:", caminho)
        dados = pd.read_excel(caminho, sheet_name="Folha1", header=1)
        print("Dados carregados com sucesso! Primeiras linhas:")
        print(dados.head())  # Mostra as primeiras linhas do DataFrame

        # Identificar as colunas de datas (a partir da 7ª coluna)
        colunas_datas = dados.columns[6:]  # Colunas de datas começam na coluna 7
        print("Colunas de datas identificadas:", colunas_datas)

        # Identificar a última coluna com valores válidos
        ultima_coluna_valida = next(
            (col for col in reversed(colunas_datas) if dados[col].notnull().any()), None
        )
        print("Última coluna válida encontrada:", ultima_coluna_valida)

        if not ultima_coluna_valida:
            raise ValueError("Nenhuma coluna com valores atualizados foi encontrada.")

        # Extrair as informações relevantes (Produto e última coluna válida)
        dados_estoque = dados[["Produto", ultima_coluna_valida]].copy()
        dados_estoque = dados_estoque.rename(columns={ultima_coluna_valida: "Estoque (kg)"})
        print("Dados de estoque extraídos:")
        print(dados_estoque.head())  # Mostra as primeiras linhas do DataFrame filtrado

        # Remover linhas sem produto ou estoque válido
        dados_estoque = dados_estoque.dropna(subset=["Produto", "Estoque (kg)"])
        dados_estoque = dados_estoque.groupby("Produto", as_index=False).sum()
        
        print("Dados de estoque após remoção de valores nulos:")
        print(dados_estoque.head())  # Mostra as primeiras linhas após remoção de NaN

        return dados_estoque
    except Exception as e:
        print(f"Erro ao carregar os dados de estoque: {e}")
        return None

# Chamando a função diretamente no terminal para debug
if __name__ == "__main__":
    caminho_estoque = DADOS_ESTOQUE_PATH
    print("Iniciando carregamento de dados de estoque...")
    resultado = carregar_dados_estoque(caminho_estoque)
    if resultado is not None:
        print("Resultado final do carregamento:")
        print(resultado)
    else:
        print("Nenhum dado foi carregado.")
