# Descrever os passos manuais e transformar em código
# Ler a planilha e guardar informações sobre nome e telefone
import openpyxl
import webbrowser
import pyautogui
from urllib.parse import quote
from time import sleep

webbrowser.open('https://web.whatsapp.com/')
sleep(25)

workbook = openpyxl.load_workbook('servos.xlsx')
pagina = workbook['Planilha1']

for linha in pagina.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    celula = linha[2].value
    email = linha[3].value
    escala = linha[4].value
    
    # print(nome)
    # print(telefone)
    # print(celula)
    # https://web.whatsapp.com/send?phone=&text

    if (celula == 'Vx' or celula == 'Outras'):
        linha[4].value = '1'
    elif (celula == 'Bt' or celula == 'Ms' or celula == 'Kahal' or celula == 'One way' or celula == 'Fc'):
        linha[4].value = '2'
    elif (celula == 'Wake'):
        linha[4].value = '3'
    elif (celula == 'Marvel'):
        linha[4].value = '4'
    else:
        linha[4].value = 'Sem escala definida'
# Salvar as alterações na planilha
workbook.save('servos_atualizado.xlsx')