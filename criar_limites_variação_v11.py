'''
SISTEMA PARA GERAÇÃO DE ÍNDICES DE VARIAÇÃO DE ESTOQUE COM BASE NO HISTÓRICO E REGRAS DE NEGÓCIO

Autor: Fabrício Pinheiro Souza
Analista Sênior - Vibra Energia S.A.
Rev 0
Data revisão: 03/02/2024
'''

#===============================================================

# Bloco 1: Importação das Bibliotecas e Configuração Inicial

# Importação das bibliotecas necessárias
import pandas as pd
import numpy as np
import openpyxl
from datetime import datetime, timedelta
import seaborn as sns
import os
import xlrd

print("\nINÍCIO DA EXECUÇÃO DO SCRIPT PARA GERAR LIMITES DE VARIAÇÃO - VIBRA ENERGIA S.A.")

# Inicialização e Verificação das Versões das Bibliotecas
print("\n")
print(85 * "*")
print("Iniciando o ambiente de desenvolvimento")
print("Verificando versões das bibliotecas instaladas:")
print(f"Versão do pandas: {pd.__version__}")
print(f"Versão do numpy: {np.__version__}")
print(f"Versão do openpyxl: {openpyxl.__version__}")
print(f"Versão do datetime: {datetime.now().strftime('%Y-%m-%d')}")
print(f"Versão do seaborn: {sns.__version__}")
print(f"Versão do xlrd: {xlrd.__version__}")
print(85 * "*")

# Alerta para primeira execução
print("\n")
print(85 * "*")
print(35*" "+" "+"IMPORTANTE"+" "+35*" ")
print(85 * "*")

print("\nSE ESTA É A PRIMEIRA VEZ QUE EXECUTA ESTE CÓDIGO, É NECESSÁRIO PERSONALIZAR A EXECUÇÃO DO CÓDIGO, OBSERVANDO O EXPOSTO ABAIXO")

print("\n# Na primeira execução, garanta estar acessando localmente o OneDrive da Vibra, e ter o acesso necessário às pastas do NCMV - Indicador.")
print(" # Os dados encontram-se em: C:\\Users\\CHAVE\\OneDrive\\NCMV - Indicador\\Dados (verificar caminho no seu PC.)")
print(" # A pasta para os dados de saída encontram-se em: C:\\Users\\CHAVE\\OneDrive\\NCMV - Indicador\\BI-StageArea\\AIVI\\PATHSAIDA \n(verificar caminho no seu PC.)")
print(" # As tabelas auxiliares encontram-se em: C:\\Users\\fpsou\\OneDrive - VIBRA\\NCMV - Indicador\\BI-StageArea\\Tabelas Auxiliares (verificar caminho no seu PC.)")
print(' # Verifique se os arquivos "ANO batentes.xlsx" está em uso por alguma pessoa ou processo e na pasta indicada em BATENTESPATH/ Tabelas Auxiliares.')
print(' # Verifique se existe arquivo "Limites Novos.xlsx" aprovado pela OPER e se o arquivo "ANO batentes.xlsx" está em uso por alguma pessoa ou processo \ne na pasta indicada em BATENTESPATH/ Tabelas Auxiliares.')
print(' # Se algum dos arquivos estiver em uso por outra pessoa, aparecerá uma mensagem similar a:\nPermissionError: [Errno 13] Permission denied: \n"C:\\Users\\fpsou\\OneDrive - VIBRA\\NCMV - Indicador\\BI-StageArea\\AIVI\\PATHSAIDA/Limites Ajustados.xlsx"')
print("\n")
print(35*" " + " " + "BLOCO 1 CONCLUÍDO"+ " "+ 35*" ")
print("\n")
print(85 * "*")

#===============================================================

# Bloco 2: Definição de Caminhos e Modo de Execução

# Definição padrão dos caminhos dos arquivos
PATHDADOS = r'C:\Users\fpsou\OneDrive - VIBRA\NCMV - Indicador\Dados'
PATHSAIDA = r'C:\Users\fpsou\OneDrive - VIBRA\NCMV - Indicador\BI-StageArea\AIVI\PATHSAIDA'
PATHTABELAS = r'C:\Users\fpsou\OneDrive - VIBRA\NCMV - Indicador\BI-StageArea\Tabelas Auxiliares'

# Verificação do Modo de Execução (Padrão ou Personalizado)
modo_execucao = input("\nDeseja executar o código no modo padrão (a) ou personalizar (b) a execução? [a/b]: ").lower()

def obter_ano_para_analise():
    """
    Solicita ao usuário o ano para análise.
    :return: Ano para análise.
    """
    ano = int(input("Digite o ano para análise (ex: 2023): "))
    return ano

print("\n")
print(35*" " + " " + "BLOCO 2 CONCLUÍDO"+ " "+ 35*" ")
print("\n")
print(85 * "*")
print("\n")

#===============================================================

# Bloco 3: Carregamento e Limpeza dos Dados

# Carregamento e Concatenação de Vários Arquivos de Dados e Limpeza do DataFrame
print(85 * "*")
print("\nCarregando e concatenando dados dos arquivos...\n")
print(85 * "*")

dataframes = []
for arquivo in os.listdir(PATHDADOS):
    if arquivo.startswith('Dado') and arquivo.endswith('.xlsx'):
        caminho_completo = os.path.join(PATHDADOS, arquivo)
        df = pd.read_excel(caminho_completo, header=None)
        df = df.drop(index=0)  # Eliminar a primeira linha de cada arquivo
        df = df.dropna(how='all').dropna(axis=1, how='all')
        df.columns = df.iloc[0]
        df = df[1:]
        df = df[df.iloc[:, 0] != df.columns[0]]
        dataframes.append(df)

# Concatenando todos os dataframes
df = pd.concat(dataframes)

# Tratamento Inicial de Valores Nulos
df = df.dropna(how='all').dropna(axis=1, how='all')
df.fillna(0, inplace=True)  # Substituir nulos por zero (ajustar conforme necessário)

# Atualização da seleção do ano para análise na execução padrão
if modo_execucao == 'a':
    anos_disponiveis = df['Ano do documento do material'].unique()
    anos_completos = []

    for ano in anos_disponiveis:
        meses = df[df['Ano do documento do material'] == ano]['Mês do exercício'].unique()
        if len(meses) == 12:
            anos_completos.append(ano)

    # Seleciona o ano seguinte ao mais recente ano completo para análise
    ano_para_analise = max(anos_completos) + 1 if anos_completos else None
    print(f"Ano selecionado para análise (automático): {ano_para_analise}")
else:
    # Em modo personalizado, o usuário define o ano para análise
    ano_para_analise = obter_ano_para_analise()

print('\nValores "Null" dos dados concatenados tratados\n')
print(85 * "*")

# Ajustando os tipos de dados das colunas
cols_to_float = ['Expedição c/ Veí', 'Variação Interna', 'Variação Manual', 'VarInt + VarMan', 'Quant. Exceden', 'Custo Unitário', 'Valor Excede', 'Imposto (R$)', 'Valor Exced. da', 'Valor da VI (R$)', 'Valor da VI +']
for col in cols_to_float:
    df[col] = df[col].astype(float).round(2)

cols_to_decimal = ['Percentual de V', 'Limite Inferior', 'Limite Su', 'Histórico', 'Percentual Excedente']
for col in cols_to_decimal:
    df[col] = df[col].astype(str).str.rstrip('%').astype('float') / 10000

df.fillna(0, inplace=True)

print("\nColunas convertidas para FLOAT C/ 2 CASAS: ['Expedição c/ Veí', 'Variação Interna', 'Variação Manual', 'VarInt + VarMan', 'Quant. Exceden',\n 'Custo Unitário', 'Valor Excede', 'Imposto (R$)', 'Valor Exced. da', 'Valor da VI (R$)', 'Valor da VI +']\n")
print("\nColunas convertidas para FLOAT c/ 7 CASAS: ['Percentual de V', 'Limite Inferior', 'Limite Su', 'Histórico', 'Percentual Excedente']\n")
print(85 * "*")

# Crie uma nova coluna 'data' combinando as colunas 'Ano do documento do material' e 'Mês do exercício'
print('\nCriando uma coluna de data...\n')
df['data'] = pd.to_datetime(df['Ano do documento do material'].astype(str) + '-' + df['Mês do exercício'].astype(str))
print('Coluna "Data" criada.\n')
print(85 * "*")

'''
# 'df' - DataFrame
for col in df.columns:
    print(f"Valores únicos para {col}:")
    print(df[col].unique())
'''

#===============================================================

# Bloco 4: Funções de Configuração Personalizada

if modo_execucao == 'b':
    print(85 * "*")
    PATHDADOS = input("Digite o caminho para os dados de entrada: ")
    print(85 * "*")
    PATHSAIDA = input("Digite o caminho para o arquivo de saída: ")
    print(85 * "*")
    PATHTABELAS = input('Digite o caminho para a tabela de batentes e "Limites Novos.xlsx" (se existir limite aprovado): ')
    print(85 * "*")
    # Funções para interação com o usuário em modo personalizado
    ano_para_analise = obter_ano_para_analise

    def escolher_periodo_analise(ano):
        """
        Pergunta ao usuário se deseja gerar limites para o ano inteiro ou um período específico.
        :param ano: Ano para análise.
        :return: Mês inicial e final para o período de análise.
        """
        escolha = input("Deseja gerar limites para o ano inteiro ou um período específico? [Ano/Período]: ").lower()
        if escolha == 'ano':
            return 1, 12
        else:
            mes_inicial = int(input("Digite o primeiro mês do período (1-12): "))
            return mes_inicial, 12

    def escolher_filtragem():
        """
        Pede ao usuário para escolher o tipo de filtragem desejada.
        :return: Parâmetros de filtragem escolhidos pelo usuário.
        """
        opcao_todos_centros = input("Deseja gerar limites para todos os centros? [S/N]: ").upper()
        if opcao_todos_centros == 'S':
            return None, None, None

        opcao_set = input("Deseja gerar limites para uma OPx (Set) inteira? [S/N]: ").upper()
        set_especifico = input("Digite o Set específico (OPx): ") if opcao_set == 'S' else None

        opcao_centro = input("Deseja gerar limites para um Centro específico? [S/N]: ").upper()
        centro_especifico = input("Digite o código do Centro específico: ") if opcao_centro == 'S' else None

        opcao_produto = input("Deseja gerar limites para um Cód Grupo de produto específico? [S/N]: ").upper()
        produto_especifico = input("Digite o Cód Grupo de produto específico: ") if opcao_produto == 'S' else None

        return set_especifico, centro_especifico, produto_especifico

if modo_execucao == 'a':
    # Em modo padrão, o ano para análise é definido automaticamente como o ano seguinte ao mais recente nos dados
    print("\nExecução no modo padrão. Determinando automaticamente o ano para análise...\n")
    anos_arquivos = [int(arquivo.split(" ")[1].split(".")[0]) for arquivo in os.listdir(PATHDADOS) if arquivo.startswith('Dado') and arquivo.endswith('.xlsx')]
    ano_para_analise = max(anos_arquivos) + 1 if anos_arquivos else datetime.datetime.now().year
    print(f"Ano para análise definido como: {ano_para_analise}")

    # No modo padrão, considera-se todo o ano para análise
    mes_inicial, mes_final = 1, 12
    print(f"Período para análise: Janeiro a Dezembro de {ano_para_analise}")

    # No modo padrão, os limites são gerados para todos os centros e produtos
    set_especifico, centro_especifico, produto_especifico = None, None, None
    print("Limites sendo gerados para todos os centros e produtos.")

print("\nConfiguração inicial completa. Prosseguindo para o carregamento e processamento dos dados.\n")
print(35*" " + " " + "BLOCO 3 CONCLUÍDO"+ " "+ 35*" ")
print("\n")
print(85 * "*")

#===============================================================

# Bloco 5: Filtragem de Dados

# Filtragem de Dados
print(85 * "*")
print("\nIniciando a filtragem de dados históricos...\n")

# 'df' é o DataFrame 'data' é a coluna com as datas
# Converta a coluna de data para o tipo de data correto
df['data'] = pd.to_datetime(df['data'])

# Obtenha a data do último registro no DataFrame
ultima_data = df['data'].max()

# Calcule as datas de início para cada período
inicio_12_meses = ultima_data - timedelta(days=365)
inicio_6_meses = ultima_data - timedelta(days=180)
inicio_24_meses = ultima_data - timedelta(days=730)

# Dentro do Bloco 5, antes da filtragem específica de dados
if modo_execucao == 'a':
    # Utiliza dados do último ano completo disponível
    df_filtrado = df[df['Ano do documento do material'] == (ano_para_analise - 1)]
else:
    # Modo personalizado, filtragem baseada na entrada do usuário
    # Aqui você pode inserir lógica personalizada baseada nas entradas do usuário
    # Exemplo: df_filtrado = df[(df['Ano do documento do material'] == ano_customizado) & (condições adicionais)]
    pass

# Selecione os dados com base nas regras de negócio
if modo_execucao == 'b':
    set_especifico, centro_especifico, produto_especifico = escolher_filtragem()
    if set_especifico is not None:
        df = df[df['Nome do set'] == set_especifico]
    if centro_especifico is not None:
        df = df[df['Centro'] == centro_especifico]
    if produto_especifico is not None:
        df = df[df['Cód Grupo de produto'] == produto_especifico]

if len(df[df['data'] >= inicio_12_meses]) >= 12:
    df_filtrado = df[df['data'] >= inicio_12_meses]
elif len(df[df['data'] >= inicio_6_meses]) >= 6:
    df_filtrado = df[df['data'] >= inicio_6_meses]
else:
    df_filtrado = df[df['data'] >= inicio_24_meses]

imprimir_configuracoes = True
imprimir_dados_concatenados = False
imprimir_filtragem = False
imprimir_analises = True

if imprimir_configuracoes:
    print('\n')
    print(f"Modo de execução: {'Padrão' if modo_execucao == 'a' else 'Personalizado'}")
    print(f"Caminho dos dados: {PATHDADOS}")
    print(f"Caminho de saída: {PATHSAIDA}")
    print(f"Caminho da tabela de batentes: {PATHTABELAS}")
    print('\n')
print(85 * "*")

if modo_execucao == 'b':
    imprimir_dados_concatenados = True
    imprimir_filtragem = True
print(85 * "*")

# Imprimir dados concatenados
if imprimir_dados_concatenados:
    print("Dados carregados e concatenados com sucesso!")
    print(df.head())
print(85 * "*")

# Imprimir processo de filtragem
if imprimir_filtragem:
    print("Filtragem de dados históricos:")
    print(df_filtrado.head())
print(85 * "*")

# Verificar e remover as colunas 'Limite Inferior' e 'Limite Superior', se existirem
if 'Limite Inferior' in df.columns:
    df = df.drop(columns='Limite Inferior')
if 'Limite Su' in df.columns:
    df = df.drop(columns='Limite Su')

if 'Limite Inferior' in df_filtrado.columns:
    df_filtrado = df_filtrado.drop(columns='Limite Inferior')
if 'Limite Su' in df_filtrado.columns:
    df_filtrado = df_filtrado.drop(columns='Limite Su')

print("\nColunas 'Limite Inferior' e 'Limite SU' originais removidas de df e df_filtrado\n")
print(85 * "*")

# Criar as novas colunas no DataFrame original
df['Variação Total'] = df['Variação Interna'] + df['Variação Manual']
df['% de Variação Total'] = (df['Variação Total'] / df['Expedição c/ Veí']).round(7)
df['% de Variação Total'] = df['% de Variação Total'].replace([np.inf, -np.inf], np.nan)

print("\nColunas 'Variação Total' e '% de Variação Total' calculadas e criadas.\n")
print(85 * "*")

# Agora, criar a cópia df_filtrado com as novas colunas
df_filtrado = df.copy()

# Verificar e tratar nulos nas colunas críticas antes de cálculos
colunas_criticas = ['Variação Total', '% de Variação Total']

for col in colunas_criticas:
    df_filtrado[col] = df_filtrado[col].fillna(0)

'''
print("\nDADOS ÚNICOS DF FILTRADO")
for coluna in df_filtrado.columns:
    print(f"Valores únicos para a coluna {coluna}:")
    print(df_filtrado[coluna].unique())
    print("\n")
print(85 * "*")
'''

print('\nInformações de df_filtrado\n')
print(df_filtrado.info())

print("\nDados nulos das colunas 'Variação Total', '% de Variação Total' substituídos por zeros.\n")
print(35*" " + " " + "BLOCO 5 CONCLUÍDO"+ " "+ 35*" ")
print('\n')
print(85 * "*")

df_pregroupby = df_filtrado

#===============================================================

# Bloco 6: Definindo e filtrando Outliers

# Função para calcular estatísticas iniciais no DataFrame filtrado
def calcular_estatisticas_iniciais(df_filtrado):
    # Substituir nulos por zero na coluna '% de Variação Total'
    df_filtrado['% de Variação Total'] = df_filtrado['% de Variação Total'].fillna(0)

    df_grouped = df_filtrado.groupby(['Nome do set', 'Centro', 'Cód Grupo de produto', 'Nome'])
    df_estatisticas = df_grouped['% de Variação Total'].agg(['mean', 'std']).reset_index()
    
    # Calcular os quantis separadamente
    q1 = df_grouped['% de Variação Total'].quantile(0.25).reset_index(name='quantile_25')
    q3 = df_grouped['% de Variação Total'].quantile(0.75).reset_index(name='quantile_75')
    
    # Unir os quantis com as estatísticas
    df_estatisticas = pd.merge(df_estatisticas, q1, on=['Nome do set', 'Centro', 'Cód Grupo de produto', 'Nome'])
    df_estatisticas = pd.merge(df_estatisticas, q3, on=['Nome do set', 'Centro', 'Cód Grupo de produto', 'Nome'])

    # Calcular IQR e limites para outliers
    df_estatisticas['IQR'] = df_estatisticas['quantile_75'] - df_estatisticas['quantile_25']
    df_estatisticas['limite_inferior_iqr'] = df_estatisticas['quantile_25'] - 1.5 * df_estatisticas['IQR']
    df_estatisticas['limite_superior_iqr'] = df_estatisticas['quantile_75'] + 1.5 * df_estatisticas['IQR']
    return df_estatisticas

# Calculando estatísticas iniciais para df_filtrado
df_estatisticas_iniciais = calcular_estatisticas_iniciais(df_filtrado)

# Função para identificar e marcar outliers no DataFrame filtrado
def marcar_outliers(df_filtrado, df_estatisticas):
    df_filtrado = df_filtrado.merge(df_estatisticas, on=['Nome do set', 'Centro', 'Cód Grupo de produto', 'Nome'], how='left')
    df_filtrado['É Outlier?'] = (df_filtrado['% de Variação Total'] < df_filtrado['limite_inferior_iqr']) | (df_filtrado['% de Variação Total'] > df_filtrado['limite_superior_iqr'])
    return df_filtrado

# Identificando e marcando outliers em df_filtrado
df_com_outliers_marcados = marcar_outliers(df_filtrado, df_estatisticas_iniciais)

# Removendo outliers de df_filtrado e criando uma cópia independente
df_sem_outliers = df_com_outliers_marcados[~df_com_outliers_marcados['É Outlier?']].copy()

# Função para recalcular estatísticas sem outliers
def recalcular_estatisticas(df_sem_outliers):
    # Substituir nulos por zero na coluna '% de Variação Total'
    df_sem_outliers['% de Variação Total'] = df_sem_outliers['% de Variação Total'].fillna(0)

    return calcular_estatisticas_iniciais(df_sem_outliers)

# Recalculando estatísticas sem outliers para df_filtrado
df_estatisticas_finais = recalcular_estatisticas(df_sem_outliers)

# Calcular limites de variação com base nas estatísticas finais
df_estatisticas_finais['Limite Inferior'] = df_estatisticas_finais['mean'] - df_estatisticas_finais['std']
df_estatisticas_finais['Limite Superior'] = df_estatisticas_finais['mean'] + df_estatisticas_finais['std']

df_estatisticas_finais = df_estatisticas_finais.fillna(0)

print("\nOutliers (IQR+-1,5 intervalo interquartil) identificados e removidos em df_estatístisticas_finais.\n")
print('Estatísticas finais calculadas sem considerar outliers\n')
print(85 * "*")
print("\nInformações de df_estatístisticas_finais\n")
print(df_estatisticas_finais.head)
print('\n')
print(df_estatisticas_finais.info())
print(35*" " + " " + "BLOCO 6 CONCLUÍDO"+ " "+ 35*" ")
print('\n')
print(85 * "*")

'''
for col in df_estatisticas_finais.columns:
    print(f"\nValores únicos na coluna {col}: {df_estatisticas_finais[col].unique()}\n")
'''

#===============================================================

# Bloco 7: Criando dataframe final

# Verifique os nomes das colunas atuais
print("\nColunas atuais em df_estatisticas_finais:")
print(df_estatisticas_finais.columns)

# Criando uma cópia independente para trabalhar
df_final = df_estatisticas_finais[['Nome do set', 'Centro', 'Cód Grupo de produto', 'Nome', 'mean', 'std', 'limite_inferior_iqr', 'limite_superior_iqr']].copy()

# Renomeando as colunas
df_final.columns = ['Nome do set', 'Centro', 'Cód Grupo de produto', 'Nome', 'Média % Variação Total', 'Std % Variação Total', 'Limite Inferior', 'Limite Superior']

# Convertendo tipos e arredondando
df_final['Média % Variação Total'] = df_final['Média % Variação Total'].astype(float).round(7)
df_final['Std % Variação Total'] = df_final['Std % Variação Total'].astype(float).round(7)
df_final['Limite Inferior'] = df_final['Limite Inferior'].astype(float).round(7)
df_final['Limite Superior'] = df_final['Limite Superior'].astype(float).round(7)

# Calculando o intervalo e verificando se é adequado
df_final['Intervalo'] = df_final['Limite Superior'] - df_final['Limite Inferior']
df_final['Intervalo Adequado'] = df_final['Intervalo'] >= 0.001  # 0.10%

# Convertendo as colunas para os tipos desejados
df_final['Nome do set'] = df_final['Nome do set'].astype(str)
df_final['Centro'] = df_final['Centro'].astype(str)
df_final['Cód Grupo de produto'] = df_final['Cód Grupo de produto'].astype(str)
df_final['Média % Variação Total'] = df_final['Média % Variação Total'].astype(float).round(7)
df_final['Std % Variação Total'] = df_final['Std % Variação Total'].astype(float).round(7)
df_final['Limite Inferior'] = df_final['Limite Inferior'].astype(float).round(7)
df_final['Limite Superior'] = df_final['Limite Superior'].astype(float).round(7)

# Verificando o DataFrame final
print("\nDataframe final criado com as colunas ['Nome do set', 'Centro', 'Cód Grupo de produto', 'Nome', 'Média % Variação Total', 'Std % Variação Total', 'Limite Inferior', 'Limite Superior']\n")
print(df_final.head())
print('\n')
print(35*" " + " " + "BLOCO 7 CONCLUÍDO"+ " "+ 35*" ")
print('\n')
print(85 * "*")

#===============================================================

# Bloco 8: Verificando Regras de Negócio

#1 Regra de Negócio: Intervalo entre limite superior e inferior de pelo menos 0,10%

print('REGRAS DE NEGÓCIO CONSIDERADAS NO CÓDIGO\n'
'Checando aderência às regras de Negócio:\n'
'-Base histórica de 2 anos para movimentação\n'
'-Base histórica de 1 ano para cálculo das estatísticas\n'
'-Para cálculo da média e desvios, são eliminados outliers que extrapolem IQR+-1,5 quartil\n'
'-Calculo dos ranges: Média +- 1 DesvPad da variação por Centro/ Código Produto, obedecendo às seguintes restrições:\n'
'    # Piso e teto % do produto\n'
'    # Intervalo entre limite superior e inferior de pelo menos 0,1%\n'
'    # Limite inferior máximo: - 0,03%\n'
'    # Limite superior mínimo: + 0,03%\n'
'    # Índices são revisados anualmente ou em caso de mudança drástica relevante no processo da unidade operacional.\n'
'Aqui serão criadas colunas booleanas (True or False) para checar se os limites calculados com a média e desvio\n'
'padrão se adequam às regras de negócio.')

'''
Intervalo entre limite superior e inferior de pelo menos 0,10%
'''
# Informando status
print("\nCalculando intervalos entre Limite Superior e Limite Inferior...")

# Adicionando coluna "Intervalo" para calcular a diferença percentual entre os limites
df_final['Intervalo'] = df_final['Limite Superior'] - df_final['Limite Inferior']

# Verificando nulos
print("\nNulos em 'Limite Inferior':", df_final['Limite Inferior'].isna().sum())
print("Nulos em 'Limite Superior':", df_final['Limite Superior'].isna().sum())
print("Nulos em 'Intervalo':", df_final['Intervalo'].isna().sum())

# Informando status
print("\nVerificando quais intervalos estão adequados ao range mínimo de 0,1%...\n")

# Definindo o intervalo mínimo
intervalo_minimo = 0.001  # 0,10%

# Verificando a adequação ao intervalo mínimo entre limites
df_final['Intervalo Adequado'] = df_final['Intervalo'] >= intervalo_minimo

# Definindo os limites inferior máximo e superior mínimo conforme as regras de negócio
limite_inferior_maximo = -0.0003  # -0,03%
limite_superior_minimo = 0.0003  # 0,03%

# Verificando a adequação dos limites
df_final['Limite Inferior Adequado'] = df_final['Limite Inferior'] >= limite_inferior_maximo
df_final['Limite Superior Adequado'] = df_final['Limite Superior'] >= limite_superior_minimo

# Calculando Meio do Range do Intervalo
df_final['Meio do Range do Intervalo'] = (df_final['Limite Superior'] + df_final['Limite Inferior']) / 2

# Ajustes conforme as regras
def ajustar_limites(row):
    if row['Limite Inferior'] > limite_inferior_maximo:  # Correção aplicada aqui
        row['Limite Inferior'] = limite_inferior_maximo
    if row['Limite Superior'] < limite_superior_minimo:
        row['Limite Superior'] = limite_superior_minimo
    
    # Recalculando o intervalo após ajustes iniciais
    row['Intervalo'] = row['Limite Superior'] - row['Limite Inferior']
    
    # Ajustando para garantir o intervalo mínimo
    if row['Intervalo'] < intervalo_minimo:
        # Ajustar ambos os limites para atender o intervalo mínimo
        diferenca_para_ajuste = (intervalo_minimo - row['Intervalo']) / 2
        row['Limite Inferior'] -= diferenca_para_ajuste
        row['Limite Superior'] += diferenca_para_ajuste
    
    return row

# Aplicando a função de ajuste
df_final = df_final.apply(ajustar_limites, axis=1)

# Verificando novamente a adequação ao intervalo mínimo e aos limites adequados após ajustes
df_final['Intervalo Adequado Pos'] = df_final['Intervalo'] >= intervalo_minimo
df_final['Limite Inferior Adequado Pos'] = df_final['Limite Inferior'] <= limite_inferior_maximo  # Correção aqui
df_final['Limite Superior Adequado Pos'] = df_final['Limite Superior'] >= limite_superior_minimo

# Imprimindo verificações
print("Contagem de valores únicos True or False para cada coluna booleana antes de ajustar:\n")
print(df_final['Intervalo Adequado'].value_counts())
print('\n')
print(df_final['Limite Inferior Adequado'].value_counts())
print('\n')
print(df_final['Limite Superior Adequado'].value_counts())

print('\n Adequando Limite Inferior Máximo e Limite Superior Mínimo...')

print("\nApós ajustar:")
print('\n')
print(df_final['Intervalo Adequado Pos'].value_counts())
print('\n')
print(df_final['Limite Inferior Adequado Pos'].value_counts())
print('\n')
print(df_final['Limite Superior Adequado Pos'].value_counts())

print("\nLimite Inferior Máximo e Limite Superior Mínimo após ajustes (devem ser -0,3% e 0,3% respectivamente):")
print(f"O Limite Inferior Máximo encontrado foi {limite_inferior_maximo}")
print(f"O Limite Superior Mínimo encontrado foi {limite_superior_minimo}")

# Verificando nulos após ajuste dos ranges mínimos
print("\nNulos em 'Limite Inferior':", df_final['Limite Inferior'].isna().sum())
print("Nulos em 'Limite Superior':", df_final['Limite Superior'].isna().sum())
print("Nulos em 'Intervalo':", df_final['Intervalo'].isna().sum())

# Imprimindo as primeiras e últimas linhas de df_final para verificação
print("\nDataFrame final (cabeçalho):\n")
print(df_final.head())

print("\nDataFrame final (rodapé):\n")
print(df_final.tail())

print("\nInformações do DataFrame final:\n")
print(df_final.info())
print(limite_inferior_maximo)

'''
Piso e teto % do produto
'''
print('Carregando arquivo de batentes. Valores em % precisam ser convertidos para decimal.')
# Carregar o arquivo de batentes
df_batentes = pd.read_excel(PATHTABELAS + '\\2024 batentes.xlsx')

# Converter valores de porcentagem para decimal
df_batentes['Limite Inferior'] = df_batentes['Limite Inferior'] / 100
df_batentes['Limite Superior'] = df_batentes['Limite Superior'] / 100

# Verificar colunas de df_batentes antes da renomeação
print("\nColunas de df_batentes antes da renomeação:", df_batentes.columns.tolist())
print("\n")

# Renomear as colunas de df_batentes
df_batentes.rename(columns={'Limite Inferior': 'Batente Inferior', 'Limite Superior': 'Batente Superior'}, inplace=True)

# Verificar colunas de df_batentes após a renomeação
print("Colunas de df_batentes após a renomeação:", df_batentes.columns.tolist())
print("\n")

# Mesclar df_final com df_batentes
df_final = df_final.merge(df_batentes[['Cód Grupo de produto', 'Batente Inferior', 'Batente Superior']], on='Cód Grupo de produto', how='left')

# Definindo valores padrão
limite_inferior_padrao = -0.0003
limite_superior_padrao = 0.0003

print("Ajustando valores eventuais valores nulos (não possuem batentes) para valores padrão (-0.0003 e 0.0003).\n")
# Preencher valores nulos para 'Batente Inferior' e 'Batente Superior' após a mesclagem
df_final['Batente Inferior'].fillna(limite_inferior_padrao, inplace=True)  # Ajustando para valores padrão
df_final['Batente Superior'].fillna(limite_superior_padrao, inplace=True)   # Ajustando para valores padrão

# Verificar colunas de df_final após a mesclagem
print("Colunas de df_final após a mesclagem:", df_final.columns.tolist())
print("\n")

# Adicionar as colunas de verificação em relação aos batentes
df_final['Inferior no Batente'] = df_final['Limite Inferior'] >= df_final['Batente Inferior']
df_final['Superior no Batente'] = df_final['Limite Superior'] <= df_final['Batente Superior']

# Ajustando limites conforme batentes
df_final.loc[~df_final['Inferior no Batente'], 'Limite Inferior'] = df_final['Batente Inferior']
df_final.loc[~df_final['Superior no Batente'], 'Limite Superior'] = df_final['Batente Superior']

# Reavaliando as colunas de verificação após ajustes
df_final['Inferior no Batente'] = df_final.apply(lambda row: row['Limite Inferior'] >= row['Batente Inferior'], axis=1)
df_final['Superior no Batente'] = df_final.apply(lambda row: row['Limite Superior'] <= row['Batente Superior'], axis=1)

# Imprimir informações antes e depois dos ajustes de batentes
print("\nInformações do DataFrame final após ajustes:")
print(df_final.info())

# Recalcular as colunas de verificação após os ajustes
df_final['Inferior no Batente'] = df_final['Limite Inferior'] >= df_final['Batente Inferior']
df_final['Superior no Batente'] = df_final['Limite Superior'] <= df_final['Batente Superior']

# Contagem dos valores True e False após os ajustes
print("\n")
print('Contagem total de registros em Limite Inferior:')
print(f'{df_final["Limite Inferior"].count()} ')
print("\n")
print(df_final['Inferior no Batente'].value_counts())
print("\nContagem de valores únicos em Limite Inferior após ajustes às Regras de Negócio:")
print(df_final["Limite Inferior"].value_counts())

print("\n")
print('Contagem total de registros em Limite Superior:')
print(f'{df_final["Limite Superior"].count()} ')
print("\n")
print(df_final['Superior no Batente'].value_counts())
print("\nContagem de valores únicos em Limite Inferior após ajustes às Regras de Negócio:")
print(df_final["Limite Superior"].value_counts())

print('\n')
print("Batente Inferior:", df_final['Batente Inferior'].min())
print("Limite Inferior Máximo:", limite_inferior_maximo)
print("Limite Superior Mínimo:", limite_superior_minimo)
print("Batente Superior:", df_final['Batente Superior'].max())

print('\n')
print("Limite Inferior Mínimo após ajustes:", df_final['Limite Inferior'].min())
print("Limite Inferior Máximo após ajustes:", df_final['Limite Inferior'].max())
print("Limite Superior Mínimo após ajustes:", df_final['Limite Superior'].min())
print("Limite Superior Máximo após ajustes:", df_final['Limite Superior'].max())

# Verificando o DataFrame final
print('\nDataframe final criado\n')
print(35*" " + " " + "BLOCO 8 CONCLUÍDO"+ " "+ 35*" ")
print("\n")
print(85 * "*")

#===============================================================

# Bloco 9: Verificando se existe tabela de limites aprovada anteriormente a este cálculo e adequando os valores dos limites

print('\nVerificando se existe tabela de limites aprovada anteriormente a este cálculo e adequando os valores dos limites...')

# Caminho para o arquivo "Limites Novos.xls"
caminho_arquivo_limites_novos = os.path.join(PATHTABELAS, "Limites Novos.xls")

# Verificando se o arquivo existe
arquivo_existe = os.path.exists(caminho_arquivo_limites_novos)

# Processamento condicional baseado na existência do arquivo
if arquivo_existe:
    print('\n')
    print(f'Existe arquivo de Limites Aprovados no momento desta execução: Limites Novos.xlsx) em {PATHTABELAS}.')
    print('\n')
    
    # O arquivo existe, proceder com o carregamento e ajustes
    df_limites_novos = pd.read_excel(caminho_arquivo_limites_novos)

    # Convertendo vírgula para ponto nos valores decimais
    #df_limites_novos['LmInferior'] = df_limites_novos['LmInferior'].astype(str).str.replace(',', '.').astype(float)
    #df_limites_novos['LmSuperior'] = df_limites_novos['LmSuperior'].astype(str).str.replace(',', '.').astype(float)

    # Renomeando colunas
    df_limites_novos.rename(columns={
        'Set': 'Nome do set',
        'Cen.': 'Centro',
        'Nome 1': 'Nome',
        # 'LmInferior' permanece
        # 'Histórico' permanece
        # 'LmSuperior' permanece
    }, inplace=True)

    # Certificando que a coluna 'Centro' é do mesmo tipo nos dois DataFrames
    df_limites_novos['Centro'] = df_limites_novos['Centro'].astype(str)
    df_final['Centro'] = df_final['Centro'].astype(str)

    # Mesclagem com df_final_mergido
    df_final_mergido = df_final.merge(df_limites_novos[['Nome do set', 'Centro', 'Nome', 'Cód Grupo de produto', 'LmInferior', 'LmSuperior']], on=['Nome do set', 'Centro', 'Cód Grupo de produto', 'Nome'], how='left')

    # Substituição dos valores de limites por valores de 'Limites Novos', se disponíveis
    df_final_mergido['Limite Inferior'] = df_final_mergido['LmInferior'].combine_first(df_final_mergido['Limite Inferior'])
    df_final_mergido['Limite Superior'] = df_final_mergido['LmSuperior'].combine_first(df_final_mergido['Limite Superior'])

    # Descartando as colunas de 'LmInferior' e 'LmSuperior' após a substituição
    #df_final_mergido.drop(['LmInferior', 'LmSuperior'], axis=1, inplace=True)

    # Impressões antes e após as mesclagens para as colunas LmInferior e LmSuperior adaptadas
    # Estas impressões foram solicitadas antes da mesclagem, mas agora adaptaremos para verificar após a mesclagem
    print("\nDataFrame final com limites novos mesclados:")

    # Realizando as impressões para verificação
    print("\nContagem de valores de LmInferior e LmSuperior:")
    print("\n")
    print(f'Contagem total de registros em LmInferior: {df_final_mergido["LmInferior"].count()} ')
    print(f'Contagem total de registros em LmSuperior: {df_final_mergido["LmSuperior"].count()} ')
    print("\n")

else:
    # O arquivo não existe, prosseguir com o processo sem ajustes
    print("\nArquivo 'Limites Novos.xlsx' não existe, prosseguir com o processo sem ajustes.\n")

    # Impressões antes e após as mesclagens para as colunas LmInferior e LmSuperior adaptadas
    # Estas impressões foram solicitadas antes da mesclagem, mas agora adaptaremos para verificar após a mesclagem
    print("\nDataFrame final com limites novos mesclados.")

    # Realizando as impressões dos Limites para verificação
    print("\n")
    print('Contando registros totais (esta contagem deve ser igual para ambos os limites):')
    print(f'Contagem total de registros em Limite Inferior: {df_final["Limite Inferior"].count()} ')
    print(f'Contagem total de registros em Limite Superior: {df_final["Limite Superior"].count()} ')
    print("\n")

# Atualizando Status e bloco
print("\n")
print(35*" " + " " + "BLOCO 9 CONCLUÍDO"+ " "+ 35*" ")
print("\n")
print(85 * "*")

#===============================================================

# Bloco 10: Ajustes Finais e Verificação das Regras de Negócio

if arquivo_existe:
    df_final_ajustado = df_final_mergido[['Nome do set', 'Centro', 'Cód Grupo de produto', 'Nome', 'Limite Inferior', 'LmInferior', 'Limite Superior', 'LmSuperior', 'Batente Inferior', 'Batente Superior']].copy()
    # Aplicar fallback para Limites Inferiores e Superiores, se LmInferior e LmSuperior estiverem disponíveis
    df_final_ajustado['Limite Inferior'] = df_final_ajustado.apply(
        lambda row: row['LmInferior'] if pd.notnull(row['LmInferior']) else row['Limite Inferior'], axis=1)
    df_final_ajustado['Limite Superior'] = df_final_ajustado.apply(
        lambda row: row['LmSuperior'] if pd.notnull(row['LmSuperior']) else row['Limite Superior'], axis=1)

    # Remover as colunas 'LmInferior' e 'LmSuperior' após a aplicação dos fallbacks
    df_final_ajustado.drop(['LmInferior', 'LmSuperior'], axis=1, inplace=True)

else:
    # Se o arquivo "Limites Novos.xls" não existiu, manter a estrutura original de df_final
    df_final_ajustado = df_final[['Nome do set', 'Centro', 'Cód Grupo de produto', 'Nome', 'Limite Inferior', 'Limite Superior', 'Batente Inferior', 'Batente Superior']].copy()

# Ajustar limites considerando também nulos, NaN e zero
def ajustar_limites_conforme_regras_e_valores(row):
    # Trata Limite Inferior nulo, NaN ou zero
    if pd.isnull(row['Limite Inferior']) or row['Limite Inferior'] == 0:
        row['Limite Inferior'] = limite_inferior_maximo  # Define como batente inferior se nulo, NaN ou zero
    
    # Trata Limite Superior nulo, NaN ou zero
    if pd.isnull(row['Limite Superior']) or row['Limite Superior'] == 0:
        row['Limite Superior'] = limite_superior_minimo  # Define como batente superior se nulo, NaN ou zero

    # Ajusta o limite inferior somente se fora do batente ou fora do limite máximo permitido
    if row['Limite Inferior'] < row['Batente Inferior']:
        row['Limite Inferior'] = row['Batente Inferior']
    if row['Limite Inferior'] > limite_inferior_maximo:
        row['Limite Inferior'] = limite_inferior_maximo

    # Ajusta o limite superior somente se fora do batente ou fora do limite mínimo permitido
    if row['Limite Superior'] > row['Batente Superior']:
        row['Limite Superior'] = row['Batente Superior']
    if row['Limite Superior'] < limite_superior_minimo:
        row['Limite Superior'] = limite_superior_minimo

    return row

# Aplica os ajustes finais conforme as regras de negócio e valores nulos, NaN ou zero
df_final_ajustado = df_final_ajustado.apply(ajustar_limites_conforme_regras_e_valores, axis=1)

# Após ajustar, verifica se os limites estão conforme as regras de negócio
df_final_ajustado['Dentro dos Batentes e Limites'] = df_final_ajustado.apply(
    lambda row: row['Batente Inferior'] <= row['Limite Inferior'] <= limite_inferior_maximo < limite_superior_minimo <= row['Limite Superior'] <= row['Batente Superior'],
    axis=1
)

# Imprime as verificações finais
print("\nReverificando aderência às regras de negócio após eventual mesclagem com Limites Novos.xlsx:")
print(df_final_ajustado[['Limite Inferior', 'Limite Superior', 'Batente Inferior', 'Batente Superior', 'Dentro dos Batentes e Limites']].head())

print('\n')
print('Imprimindo batentes para conferência:')
print("Batente Inferior:", df_final['Batente Inferior'].min())
print("Limite Inferior Máximo:", limite_inferior_maximo)
print("Limite Superior Mínimo:", limite_superior_minimo)
print("Batente Superior:", df_final['Batente Superior'].max())

print("\nEstatísticas Finais dos Limites (devem estar dentro do range dos batentes acima):")
print("Limite Inferior Mínimo:", df_final_ajustado['Limite Inferior'].min())
print("Limite Inferior Máximo:", df_final_ajustado['Limite Inferior'].max())
print("Limite Superior Mínimo:", df_final_ajustado['Limite Superior'].min())
print("Limite Superior Máximo:", df_final_ajustado['Limite Superior'].max())

# Resumo da aderência
print("\nResumo da aderência aos limites e batentes:")
print(df_final_ajustado['Dentro dos Batentes e Limites'].value_counts())

# Filtrar registros fora dos batentes e limites
registros_fora = df_final_ajustado[~df_final_ajustado['Dentro dos Batentes e Limites']]

# Exibir esses registros
print("\nRegistros fora dos batentes e limites:")
print(registros_fora[['Nome do set', 'Centro', 'Cód Grupo de produto', 'Limite Inferior', 'Limite Superior', 'Batente Inferior', 'Batente Superior']])
print("\n")

# Atualizando Status e bloco
print(35*" " + " " + "BLOCO 10 CONCLUÍDO"+ " "+ 35*" ")
print("\n")
print(85 * "*")

#===============================================================

# Bloco 11: Reajustando intervalo mínimo se necessário

print("\n")
print('Reajustando intervalo mínimo se necessário...')
print("\n")

print("\nReverificação do Intervalo Mínimo de 0,1% (0,001) após eventual mesclagem com 'Limites Novos.xlsx' antes do ajuste:")
print(f"Intervalo Mínimo: {(df_final_ajustado['Limite Superior']-df_final_ajustado['Limite Inferior']).min()}")

print("Ajustando intervalos mínimos entre os limites...")

# Ajustar limites para garantir um intervalo mínimo de 0,001
def ajustar_intervalo_minimo(row):
    intervalo = row['Limite Superior'] - row['Limite Inferior']
    if intervalo < 0.001:
        centro_do_limite = (row['Limite Superior'] + row['Limite Inferior']) / 2
        row['Limite Inferior'] = centro_do_limite - 0.0005  # Ajusta para metade do intervalo mínimo desejado
        row['Limite Superior'] = centro_do_limite + 0.0005  # Ajusta para metade do intervalo mínimo desejado
    return row

df_final_ajustado = df_final_ajustado.apply(ajustar_intervalo_minimo, axis=1)

print("\nVerificação se os Limites foram adequados ao Intervalo Mínimo de 0,1% (0,001) após ajustes:")
print(f"Intervalo Mínimo: {(df_final_ajustado['Limite Superior']-df_final_ajustado['Limite Inferior']).min()}")

# Atualizando Status e bloco
print("\n")
print(35*" " + " " + "BLOCO 11 CONCLUÍDO"+ " "+ 35*" ")
print("\n")
print(85 * "*")

#===============================================================

# Bloco 12: Salvando arquivos

# Verifique se todos os registros na coluna 'Dentro dos Batentes e Limites' são True
todos_aderem = df_final_ajustado['Dentro dos Batentes e Limites'].all()

# Calcule o intervalo mínimo
intervalo_minimo = (df_final_ajustado['Limite Superior'] - df_final_ajustado['Limite Inferior']).min()

# Verifique se o intervalo mínimo é pelo menos 0,001
intervalo_suficiente = intervalo_minimo >= 0.001

# Imprime os nomes das colunas de df_final_ajustado antes do agrupamento
print("\n")
print('Nomes das colunas de df_final_ajustado')
print(df_final_ajustado.columns)

# Se ambas as condições forem atendidas, imprima e salve os arquivos
if todos_aderem and intervalo_suficiente:
    print("\nResumo da aderência aos limites e batentes:")
    print(df_final_ajustado['Dentro dos Batentes e Limites'].value_counts())
    print("\n")
    print("Informações do dataframe final")
    print(df_final_ajustado.info())
    print("\n")

    # Calcular 'Histórico' para cada registro
    df_final_ajustado['Histórico'] = (df_final_ajustado['Limite Superior'] - df_final_ajustado['Limite Inferior']) / 2

    # Renomear colunas
    df_final_ajustado_renomeado = df_final_ajustado.rename(columns={
    'Nome do set': 'Set',
    'Centro': 'Cen.',
    'Nome' :  'Nome 1',
    'Limite Inferior': 'LmInferior',
    'Limite Superior': 'LmSuperior'
    })

    # Selecionar apenas as colunas desejadas
    df_final_ajustado_renomeado = df_final_ajustado_renomeado[['Set', 'Cód Grupo de produto', 'Cen.', 'Nome 1', 'LmInferior', 'Histórico', 'LmSuperior']]

    # Arredondar colunas numéricas para no máximo 4 casas decimais
    df_final_ajustado_renomeado['LmInferior'] = df_final_ajustado_renomeado['LmInferior'].round(4)
    df_final_ajustado_renomeado['Histórico'] = df_final_ajustado_renomeado['Histórico'].round(4)
    df_final_ajustado_renomeado['LmSuperior'] = df_final_ajustado_renomeado['LmSuperior'].round(4)

    # Verificar o DataFrame antes de salvar
    print(df_final_ajustado_renomeado.head())

    # Salvar em um arquivo Excel (XLSX)
    df_final_ajustado_renomeado.to_excel(f"{PATHSAIDA}/Limites Ajustados.xlsx", index=False)

    # Salvar em um arquivo CSV
    df_final_ajustado_renomeado.to_csv(f"{PATHSAIDA}/Limites Ajustados.csv", sep=";", encoding='utf-8', index=False)

    # Salvar em um arquivo TXT
    df_final_ajustado_renomeado.to_csv(f"{PATHSAIDA}/Limites Ajustados.txt", sep=";", encoding='utf-8', index=False)

    # Atualizando Status e bloco
    print("\n")
    print("TODOS OS LIMITES CALCULADOS E ARQUIVOS DE SAÍDA GERADOS")
    print("\n")
    print(35*" " + " " + "BLOCO 12 CONCLUÍDO"+ " "+ 35*" ")
    print("\n")
    print(85 * "*")
else:
    print(85 * "*")
    print("\nAs condições não foram atendidas ou houve erro durante o input dos dados.\n")
    print(85 * "*")
