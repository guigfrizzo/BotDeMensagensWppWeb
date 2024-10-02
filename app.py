import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os
import datetime

#Abrir WhatsApp Web
webbrowser.open('https://web.whatsapp.com/')
sleep(20)   

# Ler planilha e guardar informações sobre nome, telefone e data
workbook = openpyxl.load_workbook('pacientes.xlsx')
pagina_clientes = workbook['Página1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, data
    nome = linha[0].value
    telefone = linha[1].value
    data = linha[2].value

    # Formatando a data
    if isinstance(data, datetime.date):
        data_formatada = data.strftime("%d/%m/%Y")
    else:
        data_formatada = data

    telefone = str(telefone).replace('.0', '')  # Remover o '.0' do final
    telefone = telefone.replace("+", "").replace(" ", "").replace("-", "")  # Limpar o formato do telefone

    mensagem = f'Olá {nome}, sua consulta é hoje, dia {data_formatada}. Não se esqueça! :)'

    try:
        telefone = telefone.replace("+", "").replace(" ", "").replace("-", "")
        link_mensagem_wpp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_wpp)
        sleep(10)

        # Localizar e clicar na seta
        seta = pyautogui.locateCenterOnScreen('seta.png')
        if seta:
            pyautogui.click(seta[0], seta[1])
        else:
            print(f"Não foi possível localizar a seta para {nome}")
            continue  # Pular para o próximo cliente

        sleep(5)
        pyautogui.hotkey('ctrl', 'w')  # Fechar a aba
        sleep(5)

    except Exception as e:
        print(f'Não foi possível enviar mensagem para {nome}: {e}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
