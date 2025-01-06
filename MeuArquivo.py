import pandas as pd
import win32com.client as win32

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# facturamento por loja
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# quantidade de productos vendidos por loja
quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
# ticket medio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'ussene.c.matato@gmail.com'
mail.Subject = 'Relatorio de Vendas por Loja'
mail.HTMLBody = f'''
<html>
    <body>
        <p>Prezados,</p>
        <p>Segue o Relatório de Vendas por cada Loja.</p>
        <h3>Faturamento:</h3>
        <p>{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}</p>
        <h3>Quantidade Vendida:</h3>
        <p>{quantidade.to_html}</p>
        <h3>Ticket Médio dos Produtos em cada Loja:</h3>
        <p>{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}</p>
        <p>Qualquer dúvida estou à disposição.</p>
        <p>Att...</p>
        <p>Matato</p>
    </body>
</html>
'''
mail.Send()
print('Email enviado')
