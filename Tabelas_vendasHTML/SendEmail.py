import pandas as pd
import win32com.client as win32
import smtplib
import email.message

#1° importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')  #vai armazenar o arq do excel

# visualizar a base de dados
pd.set_option('display.max_columns', None)  #demonstar todas as colunas
print(tabela_vendas)                        #aproveita e verifica se ta tudo ok

#Faturamento por loja / filtrar por coluna ex:( tabela_vendas [['ID loja','Valor Final']]) [[]] tabela + lista / soma o que quer e mostra o faturamento tabela_vendas.grouby('ID Loja').sum()
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum() #sum: somar, media . . . 
print(faturamento)

#quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
#ticket médio por produto de cada loja  / faturamento // quantidade
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame() #.to_frame (transforma em tabela)
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'}) #mudar o nome de 0 para o que vc quer
print(ticket_medio)


#email com relatóri o
def enviar_email():
    corpo_email = f"""
    <p>Olá, <b>Jovem Aprendiz</b></p>
    <p>Treine aqui</p>
    <img src="PÍER.png" />
    <p>Prezados,</p>

    <p>Segue o relatório de vendas de cada Loja.</p>

    <p>Faturamento:</p>
    {faturamento.to_html(formatters={'Valor Final':'R${:.2f}'.format})}

    <p>Quantidade Vendida:</p>
    {quantidade.to_html()}

    <p>Ticket Médio dos Produtos de cada loja:</p>
    {ticket_medio.to_html()}
 
    <p>Qualquer dúvida, estou a disposição!<p>

    <p>att</p>
    <p>Rcrislane</P>
    """
    
    msg = email.message.Message()
    msg['Subject'] = "Relatório de Cada Produto da Loja"
    msg['From'] = 'rcavalcanti.ads@gmail.com'
    msg['To'] = 'rcavalcanti.ads@gmail.com'
    password = 'qfkabbapkalmppxq'
    msg.add_header('content-Type','text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()

    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

enviar_email()

# (formatters={'Valor Final':'R${:.2f}'.format})} coluna que quer mudar -> formatar em money