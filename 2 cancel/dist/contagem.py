import pandas as pd

# Carregar o arquivo Excel
arquivo_excel = 'C:/Users/victor.cals/Downloads/conthhc.xlsx'  # Usando barra normal
df = pd.read_excel(arquivo_excel)

# Verifique se as colunas estão corretas
print(df.columns)

# Certifique-se de que as colunas estão nomeadas corretamente
df.columns = ['Data', 'Acessos']

# Converter a coluna de data para datetime
df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y')

# Adicionar colunas para o mês e ano
df['AnoMes'] = df['Data'].dt.to_period('M')

# Calcular a quantidade de acessos por mês
acessos_por_mes = df.groupby('AnoMes')['Acessos'].sum()

# Exibir os resultados
print(acessos_por_mes)

# Se desejar salvar os resultados em um novo arquivo Excel
acessos_por_mes.to_excel('acessos_por_mes.xlsx')
