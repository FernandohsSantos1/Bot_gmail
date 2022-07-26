#Importar biblitecas

import pandas as pd
import pyautogui as pg
import pyperclip as pc
import time

pg.PAUSE = (0.5)

#Passo 1: Pedir ao usuário nome e email do destinatario
print('        Bem vindo ao MailBot\n')
print('Dados do destinatario:')
nome = input('Digite o nome do destinatário: ')
email = input('Digite o email do destinatário: ')
print('\nBase de dados: ')
link = input('Digite o link do google drive: ')
arq = input('Digite o nome do arquivo: ')
col1 = input('Digite o nome da coluna para calcular a quantidade de itens vendidos: ')
col2 = input('Digite o nome da coluna para calcular o faturamento: ')
time.sleep(3)
print('\nInicando bot...')

# #Passo 2: Abrir google chrome
pg.hotkey('win')
pg.write('google')
pg.press('Enter')

# #Passo 3: Abrir local base de dados (planilha)
pc.copy(link)
pg.hotkey('ctrl', 'v')
pg.press('Enter')
time.sleep(5)

#Passo 4: Baixar planilha
pg.rightClick(x=292, y=263)
time.sleep(1)
pg.click(x=420, y=630)
time.sleep(3)

#Passo 5: Abrir a planilha com pandas e atribur a variaveis os valores
tabela = pd.read_excel(rf"C:\Users\Fernando\Downloads\{arq}")
quantidade = tabela[col1].sum()
faturamento = tabela[col2].sum()
itens = tabela[col1].count()

#Passo 6: Abrir gmail
pg.hotkey('ctrl', 't')
pg.write('https://mail.google.com/mail/u/0/#inbox')
pg.press('Enter')
time.sleep(2)

#Passo 7: enviar email
pg.click(x=75, y=178)
time.sleep(1)
pg.write(email)
pg.press('tab')
pg.press('tab')
pg.write('Relatorio de vendas', interval= 0.05)
pg.press('tab')
texto = f'''
Prezado, {nome}!

Segue o relatorio,
Total de vendas realizadas: {itens:,}
Total de itens vendidos: {quantidade:,}
Total de faturamento: R${faturamento:,.2f} 

Grato,
Bot rs
'''
pg.write(texto, interval=0.05)
pg.hotkey('ctrl', 'Enter')

print('\nEmail enviado com sucesso\n\n'
      'Encerrando programa... ' )
