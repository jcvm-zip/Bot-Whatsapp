import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep

import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(20)
pyautogui.hotkey('ctrl', 'w')

workbook_alunos = openpyxl.load_workbook('Planilha.xlsx')
sheet_clientes = workbook_alunos['Planilha1']

for indice, line in enumerate(sheet_clientes.iter_rows(min_row=2)):
    nome_cliente = line[0].value
    telefone = line[1].value
    data_vencimento = line[2].value
    data_vencimento_str = data_vencimento.strftime('%d/%m/%Y')

    mensagem = f'Olá, {nome_cliente} Essa é uma mensagem para informar o vencimento da sua conta: {data_vencimento_str}'

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        pyautogui.hotkey('enter')
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Erro ao enviar a pensagem para {nome_cliente}')
        with open('erro.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome_cliente}, {telefone}')
