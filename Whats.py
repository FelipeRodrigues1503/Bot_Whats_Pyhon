#Processo Automomatização
#Ler planinha com nome, telefone etc
#Criar links personalizados no whats e enviar mensagens com base na planilha


#importando openpyxl para usar planilha excell
import openpyxl
#importando formatação para formartar lins para links especiais
from urllib.parse import quote
#importando webbrowser para abrir o navegador
import webbrowser
#comando para carregar planilha
#usando pyautogui para o processo de automatização
import pyautogui
import time



workbook = openpyxl.load_workbook('Planilha.xlsx')
#pagina da planilha
pagina_clientes = workbook['Plan1']
#lendo com for cada linha da planilha, começando pela linha 2 da planilha


for linha in pagina_clientes.iter_rows(min_row=2):
    #capturando dados planilha
    #primeira coluna dos dados começa lendo com [0], e assim por diante..
    nome = linha[0].value
    telefone = linha[1].value
    
    mensagem = f'Olá {nome} esta é uma mensagem automatica'
    #criando link personalizado direto na conversa do whats
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    #abrindo a pagina usando webbrowesr.open
    webbrowser.open(link_mensagem_whatsapp)
    time.sleep(30)
    #usando try para aviso de erro ao enviar a mensagem
    try:
        pyautogui.press('enter')
        time.sleep(5)
        pyautogui.hotkey('ctrl','w')
        time.sleep(5)
    except:
        print(f'Não foi possivel enviar essa mensagem{nome}')
        with open ('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}{telefone},')

