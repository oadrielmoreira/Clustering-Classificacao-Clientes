#Importando bibliotecas
import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler
from scipy.cluster.hierarchy import dendrogram, linkage
import matplotlib.pyplot as plt
from scipy.cluster.hierarchy import fcluster


#Importando bases
db = pd.read_excel('C:/Users/a.moreira/Downloads/Relatório - 003 - DB_CONSOLIDADO.xlsx')
clientes = pd.read_excel('C:/Users/a.moreira/Downloads/Clientes.xlsx')


#Limpando para ficar apenas CNPJ
filtro = db['Cpf cnpj'].str.len() < 16
db = db.drop(db[filtro].index)


#Limpando para ficar apenas os últimos 6 meses considerando o dia de hoje
db['Data'] = pd.to_datetime(db['Data'])
hoje = pd.to_datetime('today')
db['Ultimos 6 meses'] = (db['Data'] < hoje) & (db['Data'] >= hoje - pd.DateOffset(months=6))
db = db[db['Ultimos 6 meses']]


#Criando a coluna Receita = Faturamento Total
receita = db.groupby('Cpf cnpj')['Receita'].sum().reset_index()
receita['Cpf cnpj e Receita'] = receita['Cpf cnpj'] + ' - ' + df['Receita'].astype(str)
receita = receita.drop('Cpf cnpj e Receita', axis=1)
receita = receita.rename(columns={'Cpf cnpj': 'CNPJ'})


#Criando a coluna Frequêncio
frequencia_clientes = db.groupby(['Cpf cnpj', 'Data']).size().reset_index(name='Frequencia')
frequencia = frequencia_clientes.groupby('Cpf cnpj')['Frequencia'].mean().reset_index(name='Frequencia')
frequencia = frequencia.rename(columns={'Cpf cnpj': 'CNPJ'})


#Junção de Data Frames
receita = pd.merge(receita, frequencia, on='CNPJ', how='left')


#Criando a coluna que manterá o Valor Médio de Pedido na "Valor"
receita_clientes = db.groupby(['Cpf cnpj', 'Data'])['Receita'].sum().reset_index()
receita_total_clientes_por_dia = receita_clientes.groupby(['Cpf cnpj', 'Data'])['Receita'].sum().reset_index()
valor = receita_total_clientes_por_dia.groupby('Cpf cnpj')['Receita'].mean().reset_index(name='Valor')
valor = valor.rename(columns={'Cpf cnpj': 'CNPJ'})


#Junção de Data Frames
df = pd.merge(receita, valor, on='CNPJ', how='left')


#Selecionando Colunas de Interesse
X = df[['Receita', 'Frequencia', 'Valor']].values


#Padronizando os dados
#Todos os dados ficam com média 0 e desvio padrão = 1
scaler = StandardScaler()
X = scaler.fit_transform(X)


#Executando a Hierarquia de Cluster (Hierarchical Clustering)
Z = linkage(X, 'ward')


#Criação do dendograma
plt.figure(figsize=(10, 5))
plt.title('Dendrograma')
plt.xlabel('Clientes')
plt.ylabel('Distância')
dendrogram(Z, leaf_rotation=90, leaf_font_size=8)
plt.show()


#Criação dos Clusters
k = 3 #Adicione o número de clusters desejados aqui
clusters = fcluster(Z, k, criterion='maxclust')
df['cluster'] = clusters

plt.scatter(X[clusters == 1, 0], X[clusters == 1, 1], s = X[clusters == 1, 2]*10, c = '#CD7F32', label = 'Bronze')
plt.scatter(X[clusters == 2, 0], X[clusters == 2, 1], s = X[clusters == 2, 2]*10, c = '#C0C0C0', label = 'Prata')
plt.scatter(X[clusters == 3, 0], X[clusters == 3, 1], s = X[clusters == 3, 2]*10, c = '#FFD700', label = 'Ouro')
plt.title('Clusters de clientes')
plt.xlabel('Faturamento')
plt.ylabel('Frequência')
plt.legend()
plt.show()


#Tratando base para exportação e análise
df['cluster'] = clusters
df = pd.merge(df, clientes[['CNPJ', 'NOME']], on='CNPJ', how='left')
df = df.reindex(columns=['CNPJ', 'NOME', 'Receita', 'Frequencia', 'Valor', 'cluster'])
df['Classificação'] = df['cluster'].apply(lambda x: 'Bronze' if x == 1 else ('Prata' if x == 2 else 'Ouro'))
df.to_excel('Ranking Clientes.xlsx', index=False)
