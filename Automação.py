
import pandas as pd

# Lê as duas planilhas em objetos DataFrame
df1 = pd.read_excel('MOCITEC2023.xlsx', sheet_name='inscritos')
df2 = pd.read_excel('MOCITEC2023.xlsx', sheet_name='AtividadesRealizadas')

# Cria um dicionário para armazenar a soma das horas para cada nome
horas = {}


# Percorre as linhas da primeira planilha
for i, row in df1.iterrows():
    # Obtém o nome e as horas da linha atual
    nome = row['Nome dos inscritos'].upper()
    horas_nome = 0

    if nome in df2['Nome Completo:'].values:
        horas_nome = df2.loc[df2['Nome Completo:']== nome, 'Hora:'].sum()

    # Armazena a soma das horas no dicionário
    horas[nome] = horas_nome

# Cria um objeto DataFrame a partir do dicionário horas
df = pd.DataFrame(horas.items(), columns=['Nome', 'Horas'])

# Escreve o DataFrame em uma planilha do Excel
df.to_excel('horas.xlsx', index=False)
