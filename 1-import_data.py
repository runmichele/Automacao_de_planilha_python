import pandas as pd #para evirar usar  pandas usa-se pd e o as é um apelido

#Importando Dados
data = pd.read_excel('data/VendaCarros.xlsx') # Excel, pois quero importar uma planilha
print(data)

# 2-Lista os primeiros registros
print(data.head(10)) # o head lista tais registros que eu escolher

# 3-Lista os últimos registros
print(data.tail(10))

# Vou trabalhar com o Valor de venda, fabricante e o ano
print(data['Fabricante'].value_counts())#Colchetes para usar o pandas


