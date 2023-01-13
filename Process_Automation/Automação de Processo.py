#!/usr/bin/env python
# coding: utf-8

# In[1]:


#blibiotecas
import pandas as pd
import win32com.client as win32
import pathlib

#Importação das bases de dados
lojas_df = pd.read_csv(r'C:\Users\T-Gamer\Desktop\python\Projeto 1\Projeto AutomacaoIndicadores\Bases de Dados\Lojas.csv',encoding='latin-1',sep=';')
emails_df = pd.read_excel(r'C:\Users\T-Gamer\Desktop\python\Projeto 1\Projeto AutomacaoIndicadores\Bases de Dados\Emails.xlsx')
vendas_df = pd.read_excel(r'C:\Users\T-Gamer\Desktop\python\Projeto 1\Projeto AutomacaoIndicadores\Bases de Dados\Vendas.xlsx')

#juntor base de dados de lojas com a de vendas para facilitar a análise
vendas_df = vendas_df.merge(lojas_df,on='ID Loja')
vendas_df.head()

#criando o DF de cada loja
dicionario_lojas = {} #aqui foi usando um dicionário pela sua praticidade, porque ficava mais fácil usar ele dentro de um for, pois se eu colocar uma chave que ele não tem, essa chave é criada.
                       
for loja in lojas_df['Loja']:
    dicionario_lojas[loja] = vendas_df.loc[vendas_df['Loja']==loja,:]

# O indicar tem que rodar no último dia disponível na base de dados, com isso, iremos usar o max(), pois a data é tratada como um número

dia_indicador = vendas_df['Data'].max()

#identificar se pasta já existe
caminho_backup = pathlib.Path(r'C:\Users\T-Gamer\Desktop\python\Projeto 1\Projeto AutomacaoIndicadores\Backup Arquivos Lojas')
arquivos_pasta_backup = caminho_backup.iterdir()#mostra todos os aquivos dentro de uma pasta
lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]


#criando pastas
for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja #consigo usar uma barra(/) para concatenar por causa do pathlib.Path
        nova_pasta.mkdir()
        #salvar dentro da pasta
        nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month,dia_indicador.day,loja)#aqui o nome do arquivo é feito
        local_arquivo = caminho_backup/loja/nome_arquivo#esse local é a junção do caminho mais o nome da loja que é o nome da pasta e mais o nome do arquivo
        dicionario_lojas[loja].to_excel(local_arquivo)#tranformando o DF do pandas em um excel e salvando dentro do local estipulado

#Variáveis 
meta_faturamento_ano = 1650000
meta_faturamento_dia = 1000
qtd_meta_ano = 120
qtd_meta_dia = 4
ticket_medio_meta_ano = 500
ticket_medio_meta_dia = 500

#Cálculo de indicadores para cada loja e envio de email para seus respectivos gerentes
for loja in dicionario_lojas:
    vendas_lojas = dicionario_lojas[loja]
    vendas_dia = vendas_lojas.loc[vendas_lojas['Data']==dia_indicador,:]
    #faturamento
    faturamento_ano = vendas_lojas['Valor Final'].sum()
    faturamento_dia = vendas_dia['Valor Final'].sum()
    #diversidade de produtos
    qtd_produto_ano = len(vendas_lojas['Produto'].unique())
    qtd_produto_dia = len(vendas_dia['Produto'].unique())
    #ticket médio
    ticket_medio_ano = vendas_lojas.groupby('ID Loja').sum()
    ticket_medio_ano = ticket_medio_ano['Valor Final'].mean()
    ticket_medio_dia = vendas_dia.groupby('ID Loja').sum()
    ticket_medio_dia = ticket_medio_dia['Valor Final'].mean()
    #enviar email
    outlook = win32.Dispatch('outlook.application')
    nome = emails_df.loc[emails_df['Loja']==loja,'Gerente'].values[0]#como é um DF só com uma informação, podemos usar o .values[0] para pegar essa informação
    mail = outlook.CreateItem(0)
    mail.To = emails_df.loc[emails_df['E-mail']==loja,''].values[0]
    mail.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
    # lógica para mudar a cor do ◙
    if faturamento_dia >= meta_faturamento_dia:
        cor_faturamento_dia = 'green'
    else:
        cor_faturamento_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_faturamento_ano = 'green'
    else:
        cor_faturamento_ano = 'red'
    if ticket_medio_dia >= ticket_medio_meta_dia:
        cor_ticket_medio_dia = 'green'
    else:
        cor_ticket_medio_dia = 'red'
    if ticket_medio_ano >= ticket_medio_meta_ano:
        cor_ticket_medio_ano = 'green'
    else:
        cor_ticket_medio_ano = 'red'
    if qtd_produto_dia >= qtd_meta_dia:
        cor_qtd_dia= 'green'
    else: 
        cor_qtd_dia = 'red'
    if qtd_produto_ano >= qtd_meta_ano:
        cor_qtd_ano = 'green'
    else:
        cor_qtd_ano = 'red'
    #Corpo do email com formatações em HTML
    mail.HTMLBody = f'''
    <p> Bom, dia {nome}</p>
    <p> O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi: </p>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font = color="{cor_faturamento_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtd_produto_dia:.2f}</td>
        <td style="text-align: center">{qtd_meta_dia:.2f}</td>
        <td style="text-align: center"><font = color="{cor_qtd_dia}">◙</font></td>                
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${ticket_medio_meta_dia:.2f}</td>
        <td style="text-align: center"><font = color="{cor_ticket_medio_dia}">◙</font></td>
      </tr>


    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font = color="{cor_faturamento_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtd_produto_ano:.2f}</td>
        <td style="text-align: center">{qtd_meta_ano:.2f}</td>
        <td style="text-align: center"><font = color="{cor_qtd_ano}">◙</font></td>                
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${ticket_medio_meta_ano:.2f}</td>
        <td style="text-align: center"><font = color="{cor_ticket_medio_ano}">◙</font></td>
      </tr>


    </table>


    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>
    <p> Qualquer dúvida estou à disposição.</p>
    <p> Att., Daniel</p>


    '''
    # Enviando arquivo de cada loja no email
    attachment  = caminho_backup/loja/f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))

    mail.Send()
    

#Criação do Ranking para Diretoria
faturamento_loja_ano = vendas_df.groupby('Loja')[['Valor Final','Loja']].sum()
faturamento_loja_ano = faturamento_loja_ano.sort_values(by = 'Valor Final', ascending=False)
nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month,dia_indicador.day)
faturamento_loja_ano.to_excel(r'C:\Users\T-Gamer\Desktop\python\Projeto 1\Projeto AutomacaoIndicadores\Backup Arquivos Lojas\{}'.format(nome_arquivo))

vendas_loja_dia = vendas_df.loc[vendas_df['Data']==dia_indicador,:]
faturamento_loja_dia = vendas_loja_dia.groupby('Loja')[['Valor Final','Loja']].sum()
faturamento_loja_dia=faturamento_loja_dia.sort_values(by = 'Valor Final',ascending=False)
nome_arquivo = '{}_{}_Ranking Dia.xlsx'.format(dia_indicador.month,dia_indicador.day)
faturamento_loja_dia.to_excel(r'C:\Users\T-Gamer\Desktop\python\Projeto 1\Projeto AutomacaoIndicadores\Backup Arquivos Lojas\{}'.format(nome_arquivo))

#Envio de email para Diretoria
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = emails_df.loc[emails_df['Loja']=='Diretoria','E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
mail.Body = f'''
Prezados, bom dia!

Melhor loja do dia em faturamento:Loja {faturamento_loja_dia.index[0]} com faturamento:R${faturamento_loja_dia.iloc[0,0]:.2f}
Pior loja do dia em faturamento:Loja {faturamento_loja_dia.index[-1]} com faturamento:R${faturamento_loja_dia.iloc[-1,0]:.2f}

Melhor loja do ano em faturamento:Loja {faturamento_loja_ano.index[0]} com faturamento:R${faturamento_loja_ano.iloc[0,0]:.2f}
Pior loja do ano em faturamento:Loja {faturamento_loja_ano.index[-1]} com faturamento:R${faturamento_loja_ano.iloc[-1,0]:.2f}



Segue em anexos os rankings do ano e do dia de todas as lojas.
Qualquer dúvida estou à disposição.
Att.,
Daniel


'''
attachment  = caminho_backup/f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment  = caminho_backup/f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))

mail.Send()

