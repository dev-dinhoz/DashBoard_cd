
import pandas as pd
import os
import locale
from openpyxl import load_workbook
import plotly.express as px

# Definir o diretório base como o caminho do próprio script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Defina os caminhos dos arquivos usando caminhos relativos
DADOS_POINTING_PATH = os.path.join(BASE_DIR, '.database', 'ACOMPANHAMENTO DE PRODUÇÃO ATUAL-.xlsx')
DADOS_MONITORING_PATH = os.path.join(BASE_DIR, '.database', 'DATABASE.xlsx')
DADOS_DEMAND_PATH = os.path.join(BASE_DIR, '.database', 'DATABASE.xlsx')

# Definir o formato de números como pt-BR
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
locale.atof = lambda x: float(x.replace('.', '').replace(',', '.'))  
# (talvez seja inútilKKK, mas deixa ai, não vamos mexer no que está quieto) Ignorar separadores de milhar e considerar apenas 2 casas decimais

def formatar_valores(valor):
    """ Formatar valores numéricos no formato 10.000,00 """
    return locale.format_string('%.2f', valor / 1000, grouping=True)  # Divide por 1000 para mostrar em milhares

def formatar_data_brasileira(data):
    """ Formatar data no formato brasileiro dd/mm/yyyy """
    return data.strftime('%d/%  m/%Y')

def carregar_dados_pointing_ajustado(arquivo, sheet_name):
        
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

def carregar_todas_abas_ajustado_pointing(arquivo):
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

