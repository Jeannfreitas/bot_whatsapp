# Importando bibliotecas/módulos necessários
import openpyxl  # Para trabalhar com arquivos do Excel
from urllib.parse import quote  # Para codificar URLs
import webbrowser  # Para abrir o WhatsApp Web
from time import sleep  # Para criar atrasos/delays
import pyautogui  # Para automatizar ações de teclado e mouse
import os  # Para interagir com o sistema operacional

#
webbrowser.open('https://web.whatsapp.com/')  #-Abre o WhatsApp Web em uma janela do navegador
sleep(15)  

# Carrega o arquivo do Excel e acessa uma planilha específica 
clientes = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = clientes['Sheet1']

# Itera por cada linha na planilha do Excel para extrair informações dos clientes
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value 
    telefone = linha[1].value 
    vencimento = linha[2].value  
    
    # Cria uma mensagem para cada cliente
    mensagem = f'Olá {nome}, seu boleto vence no dia {vencimento.strftime("%d/%m/%Y")}. Favor pagar no link https://www.link_do_pagamento.com'

    try:
        # Gera o URL do WhatsApp com a mensagem codificada e o número de telefone do cliente
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        
        # Abre a janela de chat do WhatsApp para o respectivo cliente em uma nova janela do navegador
        webbrowser.open(link_mensagem_whatsapp)
        sleep(20)  
        
        sleep(2)  
        
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)  
        
    except Exception as e:
        # Em caso de erro, imprime a mensagem de erro e registra o cliente com erro em um arquivo 'erros.csv'
        print(f'Não foi possível enviar mensagem para {nome}. Erro: {str(e)}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
