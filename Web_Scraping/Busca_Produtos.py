#!/usr/bin/env python
# coding: utf-8

# In[8]:


#Importar blibliotecas
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import time
import win32com.client as win32

#criar navegador
navegador = webdriver.Chrome()

#Importar base de dados
buscas_df = pd.read_excel('buscas2.xlsx')

def busca_buscape(produto,preco_maximo,preco_minimo,termos_banidos,navegador):    
    navegador.get('https://www.buscape.com.br/?og=19220&og=19220&gclid=EAIaIQobChMIotX-lsKo-gIVJOVcCh3P2wrAEAAYASAAEgK9X_D_BwE')
    #tratando os termos banidos, preço mínimo e preço máximo
    termos_banidos = termos_banidos.lower()
    produto = produto.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_produto = produto.split(" ")
    preco_maximo = float(preco_maximo)
    preco_minimo = float(preco_minimo)
    time.sleep(2)
    #buscando item na barra de busca do buscapé
    navegador.find_element(By.CLASS_NAME,'AutoCompleteStyle_input__HG105').send_keys(produto,Keys.ENTER)
    time.sleep(2)
    #lista de elemtos que corresponde a uma lista com todos os itens encontrados na busca
    lista_elementos_buscape = navegador.find_elements(By.CLASS_NAME,'Paper_Paper__HIHv0')
    lista_ofertas=[] # lista vazia para ser preenchida caso as ofertas dentro dos critérios sejam encontradas
    for resultado in lista_elementos_buscape:
        nome = resultado.find_element(By.CLASS_NAME,'Text_Text__h_AF6').text #nome do produto
        nome = nome.lower()#tratar nome
        #lógica para tratar se tem termos banidos e se tem o nome completo que estamos pesquisando
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra in nome:
                tem_termos_banidos = True
        tem_todos_termos = True
        for palavra in lista_termos_produto:
            if palavra not in nome:
                tem_todos_termos = False
        #se não houver termos banidos e tiver todos os termos do produto, iremos pegar o preço e o link e colocar dentro da lista de ofertas
        if not tem_termos_banidos and tem_todos_termos:
            preco = resultado.find_element(By.CLASS_NAME,'Text_MobileHeadingS__Zxam2').text #pegando o preço
            preco = preco.replace(" ", "").replace(".", "").replace(",", ".").replace("R$","") #tratando o preço
            preco=float(preco)#tratando o preço
            if preco_minimo<= preco <= preco_maximo:# só pega o link do produto se o preço estiver dentro da faixa estipulada
                link = resultado.find_element(By.CLASS_NAME,'SearchCard_ProductCard_Inner__7JhKb').get_attribute('href')#pegnado link
                lista_ofertas.append((nome,preco,link))#adicionando nome,preço e link na lista de ofertas 
    return lista_ofertas #retorna lista de ofertas


def busca_google(produto,preco_maximo,preco_minimo,termos_banidos,navegador):
    navegador.get('https://www.google.com.br/')
    #tratando os termos banidos, preço mínimo e preço máximo
    termos_banidos = termos_banidos.lower()
    produto = produto.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_produto = produto.split(" ")
    preco_maximo = float(preco_maximo)
    preco_minimo = float(preco_minimo)
    #buscando item na barra de busca do google
    navegador.find_element(By.XPATH,'/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(produto,Keys.ENTER)
    #Lógica para clicar na aba Shopping do google
    elementos_google = navegador.find_elements(By.CLASS_NAME,'hdtb-mitem')
    for item in elementos_google:
        if 'Shopping' in item.text:
            item.click()
            break
    #lista buscas que corresponde a uma lista com todos os itens encontrados na busca
    lista_buscas = navegador.find_elements(By.CLASS_NAME,'KZmu8e')
    lista_ofertas = []
    for resultado in lista_buscas:
        nome = resultado.find_element(By.CLASS_NAME,'translate-content').text #pegando o nome
        nome = nome.lower()#tratando o nome
        #lógica para tratar se tem termos banidos e se tem o nome completo que estamos pesquisando
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra in nome:
                tem_termos_banidos = True
        tem_todos_termos = True
        for palavra in lista_termos_produto:
            if not palavra in nome:
                        tem_todos_termos = False
        #se não houver termos banidos e tiver todos os termos do produto, iremos pegar o preço e o link e colocar dentro da lista de ofertas
        if not tem_termos_banidos and tem_todos_termos:
                preco = resultado.find_element(By.CLASS_NAME,'hn9kf').text #pegando o preço
                #tratando o preço
                preco = preco.replace(" ", "").replace(".", "").replace(",", ".") 
                lista_preco = preco.split('R$')
                preco = lista_preco[1]
                preco= float(preco)
                if preco_minimo<= preco <= preco_maximo:# só pega o link do produto se o preço estiver dentro da faixa estipulada
                    link = resultado.find_element(By.CLASS_NAME,'shntl').get_attribute('href')#pegando o link
                    lista_ofertas.append((nome,preco,link))#adicionando nome,preço e link na lista de ofertas 
    return lista_ofertas#retorna lista de ofertas




    
tabela_resultado_buscas = pd.DataFrame()
for linha in buscas_df.index:
    #definindo argumentos das funções
    produto = buscas_df.loc[linha,'Nome']
    termos_banidos = buscas_df.loc[linha,'Termos banidos']
    preco_minimo = buscas_df.loc[linha,'Preço mínimo']
    preco_maximo = buscas_df.loc[linha,'Preço máximo']
    lista_ofertas_google = busca_google(produto,preco_maximo,preco_minimo,termos_banidos,navegador)#rodando a função de buscas no google
    if lista_ofertas_google:
        tabela_google = pd.DataFrame(lista_ofertas_google,columns=['Produto','Preço','Link'])#criando DF com as buscas do Google
        tabela_resultado_buscas = pd.concat([tabela_resultado_buscas, tabela_google], ignore_index=True)# Juntando DF das buscas do Google com o DF final
    else:
        tabela_google = None
    lista_ofertas_buscape = busca_buscape(produto,preco_maximo,preco_minimo,termos_banidos,navegador)#rodando a função de buscas no buscapé
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape,columns=['Produto','Preço','Link'])#criando DF com as buscas do buscapé
        tabela_resultado_buscas = pd.concat([tabela_resultado_buscas, tabela_buscape], ignore_index=True)# Juntando DF das buscas do buscapés com o DF final
    else:
        tabela_buscape = None                                                   
                                
            

#exportar por excel
tabela_resultado_buscas = tabela_resultado_buscas.reset_index(drop = True)
tabela_resultado_buscas.to_excel('Ofertas1.xlsx',index = False)


#enviar email
if len(tabela_resultado_buscas.index) > 0: #garantir que há alguma busca na tabela
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'danielcardosomds@gmail.com'
    mail.Subject = 'Produto(s) Encontrado(s) na faixa de preço desejada'
    mail.HTMLBody = f"""
    <p>Prezados,</p>
    <p>Encontramos alguns produtos em oferta dentro da faixa de preço desejada. Segue tabela com detalhes</p>
    {tabela_resultado_buscas.to_html(index=False)}
    <p>Qualquer dúvida estou à disposição</p>
    <p>Att.,</p>
    """
    
    mail.Send()

navegador.quit()  


# In[ ]:




