##  Automação de Sistemas e Processos com Python -- Código criado no Jupyter

##  Desafio:
##  Todos os dias, o nosso sistema atualiza as vendas do dia anterior. O seu trabalho diário, como analista, é enviar um e-mail para a diretoria,
#   assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior.

##E-mail da diretoria: seugmail+diretoria@gmail.com
##Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
##Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado.

import pyautogui as pa
import pyperclip as pc
import time

pa.hotkey('ctrl', 't') (#Abrindo nova aba no chrome)
time.sleep(2)
pc.copy(" https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pa.hotkey("ctrl", "v")
pa.press("enter")
time.sleep(15)
pa.doubleClick(x=390, y=362)
time.sleep(10)
pa.click(x=424, y=497)

pa.click(x=1112, y=221)
time.sleep(5)
pa.click(x=923, y=706)
time.sleep(10)

#Calcular faturamento e quantidade de vendas

import pandas as pd
tabela = pd.read_excel(r"C:\Users\Raia Del Rey\Downloads\Vendas - Dez.xlsx")
display(tabela)
faturamento = tabela["Valor Final"].sum()
print("Faturamento: R$ {}".format(faturamento))
quantidade = tabela["Quantidade"].sum()
print("Quantidade{}:".format(quantidade))


#Envio do e-mail pelo gmail
pa.hotkey('ctrl', 't')
time.sleep(2)
pc.copy("https://mail.google.com/mail/u/0/#inbox?compose=new")
pa.hotkey("ctrl", "v")
pa.press("enter")
time.sleep(30)
pc.copy("seugmail+diretoria@gmail.com")
pa.hotkey("ctrl", "v")

pa.press("tab")
time.sleep(2)
pc.copy("Relatório Vendas")
pa.hotkey('ctrl','v')
pa.press("tab")
texto = f"""
Prezados,
bom dia!
 Segue o relatório de vendas:
 Faturamento:{faturamento}
 Quantidade:{quantidade}
 """
pa.write(texto)



pa.click(x=955, y=698)
time.sleep(15)
pa.click(x=511, y=49)
pc.copy(r"C:\Users\nomepasta\nomepasta")
pa.hotkey("ctrl", "v")
time.sleep(10)
pa.click(x=513, y=667)
pa.write("Vendas")
pa.press('enter')

#enviar email

pa.hotkey("ctrl", "enter")
#pa.click(x=841, y=699) --> Opção alternativa

