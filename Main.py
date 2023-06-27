import pandas as pd

caminho_arquivo = r'C:\Users\Iago Piai\Desktop\resultado.xlsx'
planilha = pd.read_excel(r'C:\Users\Iago Piai\Desktop\1.xlsx')


with open(r'C:\Users\Iago Piai\Desktop\Elétrico.txt', 'r') as arquivo:
    lista_elétricos = [int(numero) for numero in arquivo.readlines()]

with open(r'C:\Users\Iago Piai\Desktop\Spectra.txt', 'r') as arquivo:
    lista_spectra = [int(numero) for numero in arquivo.readlines()]

with open(r'C:\Users\Iago Piai\Desktop\Modular.txt', 'r') as arquivo:
    lista_modular = [int(numero) for numero in arquivo.readlines()]

with open(r'C:\Users\Iago Piai\Desktop\Pressão.txt', 'r') as arquivo:
    lista_pressão = [int(numero) for numero in arquivo.readlines()]

with open(r'C:\Users\Iago Piai\Desktop\Mod Bus.txt', 'r') as arquivo:
    lista_mod_bus = [int(numero) for numero in arquivo.readlines()]


agrupado = planilha.groupby('positionId')

ids = []
maiores_datas = []
menores_datas = []
valores_nome_menor = []
valores_nome_maior = []
valores_boards = []

for id_grupo, grupo in agrupado:
    maior_data = grupo['date'].max()
    menor_data = grupo['date'].min()

    valores_board = grupo.loc[grupo['date'] == menor_data, 'boardId'].values
    valores_nome_maiorr = grupo.loc[grupo['date'] == maior_data, 'lastCollectBatteryVoltage'].values
    valores_nome_menorr = grupo.loc[grupo['date'] == menor_data, 'lastCollectBatteryVoltage'].values

    ids.append(id_grupo)
    maiores_datas.append(maior_data)
    menores_datas.append(menor_data)
    valores_nome_maior.append(valores_nome_maiorr[0] if len(valores_nome_maior) > 0 else None)
    valores_nome_menor.append(valores_nome_menorr[0] if len(valores_nome_menor) > 0 else None)
    valores_boards.append(valores_board[0] if len(valores_board) > 0 else None)

resultado = pd.DataFrame({
    'ID': ids,
    'Maior Data': maiores_datas,
    'Menor Data': menores_datas,
    'Valor Nome Maior': valores_nome_maior,
    'Valor Nome Menor': valores_nome_menor,
    'Valor Board': valores_boards
})

resultado = resultado.rename(columns={'Valor Nome Maior': 'Valor Bateria 2', 'Valor Nome Menor': 'Valor Bateria 1', 'ID': 'Position', 'Valor Board': 'BoardID'})

df = resultado[['Position', 'Menor Data', 'Valor Bateria 1', 'Maior Data', 'Valor Bateria 2', 'BoardID']]

df['Maior Data'] = pd.to_datetime(df['Maior Data'])
df['Menor Data'] = pd.to_datetime(df['Menor Data'])

df['Valor Bateria 1'] = df['Valor Bateria 1'].round(1)
df['Valor Bateria 2'] = df['Valor Bateria 2'].round(1)

df['Dias'] = (df['Maior Data'] - df['Menor Data']).dt.days

df['Maior Data'] = df['Maior Data'].dt.strftime('%d-%m-%Y')
df['Menor Data'] = df['Menor Data'].dt.strftime('%d-%m-%Y')

df['SensorType'] = ''

df.loc[df['Position'].isin(lista_elétricos), 'SensorType'] = 'Connect Elétrico'

df.loc[df['Position'].isin(lista_spectra), 'SensorType'] = 'Spectra'

df.loc[df['Position'].isin(lista_modular), 'SensorType'] = 'Analógico Modular'

df.loc[df['Position'].isin(lista_pressão), 'SensorType'] = 'Pressão'

df.loc[df['Position'].isin(lista_mod_bus), 'SensorType'] = 'Mod Bus'

df.to_excel(caminho_arquivo, index=False)

print("O arquivo Excel foi salvo com sucesso!")
