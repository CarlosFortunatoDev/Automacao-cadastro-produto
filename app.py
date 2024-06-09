import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']


for linha in sheet_produtos.iter_rows(min_row=2):

    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(-1523,348, duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(-1556,462, duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(-1553,575, duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(-1560,652, duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(-1576,747, duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(-1577,832, duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(-1568,888, duration=1)
    sleep(3)

    # preco = linha[6].value

    # quantidade_em_estoque = linha[7].value

    # data_de_validade = linha[8].value

    # cor = linha[9].value

    # tamanho = linha[10].value

    # material = linha[11].value

    # fabricante = linha[12].value

    # pais_origen = linha[13].value

    # observacoes = linha[14].value

    # codigo_de_barras = linha[15].value

    # localizacao_armazem = linha[16].value
