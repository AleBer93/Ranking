import pandas as pd
df = pd.read_csv('Export_2022-08-08_46156_Complet.csv', sep=";", decimal=',', engine='python', encoding='unicode_escape')
print(df)
lista_micro = []
for micro in df['Categoria Quantalys'].unique():
    lista_micro.append(micro)
print(lista_micro)
df2 = pd.DataFrame(lista_micro)
print(df2)
df2.to_excel('lol.xlsx')