# projeto-automacao-web
# Projeto criado a partir do curso do canal "Hashtag Programação"
# O que faz? Entra no sistema da empresa, baixa e lê o arquivo com os dados que a empresa dispomibiliza, calcula o que necessita e manda um email para a empresa com os dados calculados(tudo automatico) 

# Passo 1: Entrar no sistema da empresa(link do drive)
import pyautogui as pg
import time
import pyperclip
pg.PAUSE = 1
pg.press('win')
pg.write('Navegador Opera Gx')
pg.press('enter')
pg.write('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pg.press('enter') 
# Passo 2: Navegar até o local do relatorio(entrar na pasta exportar)
time.sleep(5)
pg.click(484,319,  clicks=2)
# Passo 3: Exportar o relatório(Fazer o dowload)
pg.click(513,418)
pg.click(513,418, button='right')
pg.click(619,911)
# Passo 4: Cálcular os indicadores(faturamento e quantidade de produtos)
import pandas

tabela = pandas.read_excel(r'D:\windows\Downloads\Vendas - Dez.xlsx')
print(tabela)

faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

print(faturamento)
print(quantidade)
# Passo 5: Enviar email para diretoria
time.sleep(1)
pg.click(1004,507)
pg.hotkey('ctrl', 't')
pg.write('https://mail.google.com/mail/u/4/#inbox')
pg.press('enter')
pg.click(179,216)
pg.click(1196,427)
pg.write('luismotafernandes@hotmail.com')
pg.click(1292,483)
pg.click(1160,504)

time.sleep(2)

pyperclip.copy('Relatório de Vendas')
pg.hotkey('ctrl','v')
pg.press('tab')

texto = f'''
Prezados, Boa tarde

O faturamento de hoje foi de: R${faturamento:,.2f}
A quantidade de produtos foi de:{quantidade:,}

Abs
Raphael Belfort'''

pyperclip.copy(texto)
pg.hotkey('ctrl','v')

time.sleep(3)

pg.click(1330,986)
pg.click(1051,898)
pg.click(361,523)
pg.write('Vendas - Dez (9)')
pg.press('enter')
pg.click(1183,981)


