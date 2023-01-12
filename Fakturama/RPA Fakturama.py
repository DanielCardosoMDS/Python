#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pyautogui
import pyperclip
import subprocess
import time
import pandas as pd


# In[4]:


#abrir o fakturama(ERP)
subprocess.Popen([r"C:\Program Files\Fakturama2\Fakturama.exe"])#exectuar o programa fakturama
pyautogui.FAILSAFE = True

pyautogui.PAUSE= 1
# função para pedir para o programa esperar até encontrar o elemento na tela e me retornar o elemento
def encontrar_imagem(imagem):
    while not pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.9):# reconhecimento de imagem
        time.sleep(1)
    encontrou = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.9)
    return encontrou

#função para clicar no lado direito a imagem. retorna o ponto no lado direito para eu clicar
def direita(posicoes_imagem):
    return posicoes_imagem[0] + posicoes_imagem[2], posicoes_imagem[1] + posicoes_imagem[3]/2

#função para copiar e colar o texto. o pyautogui da erro quando encontra as pontuações do portugês
def escrever_texto(texto):
    pyperclip.copy(texto)
    pyautogui.hotkey("ctrl", "v")

encontrou = encontrar_imagem('logo.PNG') # garantia que o programa abriu completamente para começar a rodar a automação

df_produtos= pd.read_excel('Produtos.xlsx')
for item in df_produtos.index:
    id_produto = df_produtos.loc[item, 'ID']# o .loc foi usado para pegar um item específico de uma coluna específica 
    nome = df_produtos.loc[item, 'Nome']
    categoria = df_produtos.loc[item, 'Categoria']
    gtin = df_produtos.loc[item, 'GTIN']
    supplier = df_produtos.loc[item, 'Supplier']
    descricao = df_produtos.loc[item, 'Descrição']
    preco = df_produtos.loc[item, 'Preço']
    custo = df_produtos.loc[item, 'Custo']
    estoque = df_produtos.loc[item, 'Estoque']
    imagem = df_produtos.loc[item, 'Imagem']
    
    encontrou = encontrar_imagem('new.PNG')#procurar a imagem
    pyautogui.click(pyautogui.center(encontrou))#clicar na imagem procurada
    encontrou = encontrar_imagem('new product.png')
    pyautogui.click(pyautogui.center(encontrou))
    
    encontrou = encontrar_imagem('item number.PNG')
    pyautogui.click(direita(encontrou))
    escrever_texto(str(id_produto))#foi usado o o str() porque irei preencher um texto. 
    
    encontrou = encontrar_imagem('name.PNG')
    pyautogui.click(direita(encontrou))
    escrever_texto(str(nome))
    
    encontrou = encontrar_imagem('category.PNG')
    pyautogui.click(direita(encontrou))
    escrever_texto(str(categoria))
    
    encontrou = encontrar_imagem('gtin.PNG')
    pyautogui.click(direita(encontrou))
    escrever_texto(str(gtin))
    
    encontrou = encontrar_imagem('supplier.PNG')
    pyautogui.click(direita(encontrou))
    escrever_texto(str(supplier))
    
    encontrou = encontrar_imagem('description.PNG')
    pyautogui.click(direita(encontrou))
    escrever_texto(str(descricao))
    
    encontrou = encontrar_imagem('price.PNG')
    pyautogui.click(direita(encontrou))
    preco_texto = f"{preco:.2f}".replace(".", ",")# tive que criar essa nova variável porque o python troca ponto por vírgula e o fakturama só aceita vírgula
    escrever_texto(str(preco_texto))
    
    encontrou = encontrar_imagem('cost price.PNG')
    pyautogui.click(direita(encontrou))
    custo_texto = f'{custo:.2f}'.replace('.',',')# tive que criar essa nova variável porque o python troca ponto por vírgula e o fakturama só aceita vírgula
    escrever_texto(str(custo_texto))
    
    encontrou = encontrar_imagem('stock.PNG')
    pyautogui.click(direita(encontrou))
    estoque_texto = f'{estoque:.2f}'.replace('.',',')# tive que criar essa nova variável porque o python troca ponto por vírgula e o fakturama só aceita vírgula
    escrever_texto(str(estoque_texto))
    
    encontrou = encontrar_imagem('picture.PNG')
    pyautogui.click(pyautogui.center(encontrou))
    
    
    encontrou = encontrar_imagem('selecionar arquivo.PNG')
    pyautogui.click(direita(encontrou))
    escrever_texto(rf'C:\Users\T-Gamer\Desktop\python\RPA\Fakturama\{str(imagem)}')
    pyautogui.press('enter')
    
    encontrou = encontrar_imagem('save.PNG')
    pyautogui.click(pyautogui.center(encontrou))
    
    
    
    
    
    
    
    
    
 


        
    
    







# In[ ]:




