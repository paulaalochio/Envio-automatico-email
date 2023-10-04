import pandas as pd
import win32com.client as win32

# importando base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# faturamento por loja

fatporloja = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(fatporloja)

# quantidade de produtos vendidos por loja

qtdprodutos = tabela_vendas[['ID Loja', 'Produto', 'Quantidade']].groupby('ID Loja').sum()
print(qtdprodutos)

print('-' * 50)

# ticket médio por produto em cada loja

ticket_medio = (fatporloja['Valor Final'] / qtdprodutos['Quantidade']).to_frame()
print(ticket_medio)  

# enviar um e-mail com o relatório
outlook = win32.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)
mail.To = 'paula.alochio@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de vendas por cada loja.</p> 

<p>Faturamento:</p>
{fatporloja.to_html()}

<p>Quantidade vendida</p>
{qtdprodutos.to_html()}

<p>Ticket Médio dos produtos de cada loja:</p>
{ticket_medio.to_html()}

<p>Qualquer dúvida estou à disposição</p>

<p>Att</p>
<p>Paula Alochio</p>
'''
mail.Send()