# 📖Sobre  Projeto

**Process_Automation** é um projeto que tem por objetivo automatizar o processo de gestão de indicadores de uma rede de lojas. É enviado para cada gerente um e-mail contendo um OnePage e uma planilha com as vendas do dia e o acumulado do ano. E para a diretoria é enviado um e-mail com o ranking das lojas do dia e do ano.

## 📊O que é um OnePage?
**OnePage** é um resumo dos principais indicadores em uma só pagina. Por isso o nome OnePage e esses indicadores podem variar de acordo com a empresa. 


**Exemplo de OnePage**

![exempli_onepage](https://github.com/DanielCardosoMDS/Python/blob/main/Process_Automation/Imagens/OnePage.PNG)

## ⚙Funcionamento
* Pegar a base de dados global e criar uma planilha para cada loja
* Verificar se na pasta **Backup Arquivos Lojas** há uma pasta para cada loja, se não, iremos criar uma.
* Salvar a planilha de cada loja dentro de sua respectiva pasta
* Calcular os indicadores para cada loja
* Enviar o e-mail para cada gerente
* Criar ranking das lojas para diretoria
* Enviar e-mail para diretoria

### ✉E-mail para os gerentes
![email_gerente](https://github.com/DanielCardosoMDS/Python/blob/main/Process_Automation/Imagens/e-mail_final_gerente.PNG)

### ✉E-mail para a diretoria
![email_gerente](https://github.com/DanielCardosoMDS/Python/blob/main/Process_Automation/Imagens/e-mail_final_diretoria.PNG)

## 🔧Ferramentas
- Python.

