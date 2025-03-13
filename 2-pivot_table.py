import pandas as pd 

#Importando Dados
data = pd.read_excel('data/VendaCarros.xlsx') # Excel, pois quero importar uma planilha
nome = 'Michele'
#print(type(nome)) #type diz qual o tipo de dado de uma variável, estrutura...
# print(type(data)) #ele vai mostrar que data é um dataframe

# Como entender esse DataFrame?
# 2-Selecionando colunas
# print(data['Estado'])
# print(data) #retorna TODOS os dados da planilha

# 2-Selecionando colunas específicas do Dataframe
df = data[['Fabricante', 'ValorVenda', 'Ano']]  # dentro de colchetes quero saber quantas colunas ter dentro do dataframe #vai obter apenas as tranformações que eu preciso
print(df)
#Isos agora é um dataframe

# 3- Criando a tabela pivô
pivot_table = df.pivot_table(
    index='Ano',
    columns='Fabricante',
    values='ValorVenda',
    aggfunc='sum'
    # esses são parametros
)

print(pivot_table)

#4- Exportando tabela pivô em arquivo excel
pivot_table.to_excel('data/pivot_table.xlsx', 'Relatório') # aqui foi criada uma planilha excel chamada Relatório
