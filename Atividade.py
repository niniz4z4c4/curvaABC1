import pandas as pd
import plotly as px

a = input("Digite o Nome do Arquivo: ")  #importa a tabela de dados
a = a + ".xlsx"
dados = pd.read_excel(a)

cortea = int(input("Digite o valor do Corte A: "))
corteb = int(input("Digite o valor do Corte B: "))
cortec = int(input("Digite o valor do Corte C: "))

dados = dados.groupby("Material", as_index=False).sum() #remove duplicatas levando em cosideração o nome
dados["Valor Total"] = dados["Quantidade"] * dados["Preço"] #Gera os valores Totais de cada Produto

total = sum(dados["Valor Total"])

for i, infos in dados.iterrows(): #Criação da coluna de Porcentagem individual
    ic2 = dados.iloc[i, 3]
    ic3 = ((ic2 / total) * 100)
    dados.loc[i, "Porcentagem"] = ic3

dados = dados.sort_values(by='Valor Total', ascending=False) #Coloca em ordem (final muda para decrescente)
dados["Porcentagem Acumulada"] = dados["Porcentagem"].cumsum() #Criação coluna da Porcentagem acumulada

#Criação da coluna Classificação
for i, infos in dados.iterrows():
    if dados.loc[i, "Porcentagem Acumulada"] <= cortea:
        dados.loc[i, "Classificação"] = "A"
    elif dados.loc[i, "Porcentagem Acumulada"] <= corteb:
        dados.loc[i, "Classificação"] = "B"
    else:
        dados.loc[i, "Classificação"] = "C"

proportion_sku = (dados.groupby('Classificação')['Material'].nunique() / dados['Material'].nunique())*100
proportion_value = (dados.groupby('Classificação')['Valor Total'].sum() / dados['Valor Total'].sum())*100

corte_sku = pd.Series([cortea, corteb, cortec], name='Corte')
corte_sku = corte_sku.rename_axis('Classificação')
corte_sku = corte_sku.rename({0: 'A', 1: 'B', 2: 'C'})

with pd.ExcelWriter('Curva_ABC.xlsx') as writer:
    dados.to_excel(writer, sheet_name='Tabela Principal', index=False)
    corte_sku.to_frame('Corte').to_excel(writer, sheet_name='Tabela Principal', index=True, startcol= 9)
    proportion_sku.to_frame('Proporção de SKU').to_excel(writer, sheet_name='Tabela Principal', index=False, startcol= 11)
    proportion_value.to_frame('Proporção de Valor').to_excel(writer, sheet_name='Tabela Principal', index=False, startcol= 12)

print("Arquivo com a Curva ABC gerado com sucesso!")