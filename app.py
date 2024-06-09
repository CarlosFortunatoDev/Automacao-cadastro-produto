import openpyxl
import pyperclip
import pyautogui
from time import sleep

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']


for linha in sheet_produtos.iter_rows(min_row=2):

# Página (1)
# Nome do Produto
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(-1523,348, duration=1)
    pyautogui.hotkey('ctrl','v')

# Descrição
    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(-1556,462, duration=1)
    pyautogui.hotkey('ctrl','v')

# Categoria
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(-1553,575, duration=1)
    pyautogui.hotkey('ctrl','v')

# Código do produto
    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(-1560,652, duration=1)
    pyautogui.hotkey('ctrl','v')

# Peso
    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(-1576,747, duration=1)
    pyautogui.hotkey('ctrl','v')

# Dimensões
    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(-1577,832, duration=1)
    pyautogui.hotkey('ctrl','v')

# Próxima página (2)
    pyautogui.click(-1568,888, duration=1)
    sleep(3)

# Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(-1547,372, duration=1)
    pyautogui.hotkey('ctrl','v')

# Quantidade em estoque
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(-1552,461, duration=1)
    pyautogui.hotkey('ctrl','v')

# Data de Validade
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(-1556,542, duration=1)
    pyautogui.hotkey('ctrl','v')

# Cor
    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(-1553,631, duration=1)
    pyautogui.hotkey('ctrl','v')

# Tamanho
    tamanho = linha[10].value
    pyautogui.click(-1569,720,duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(-1563,747, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(-1560,773, duration=1)
    else:
        pyautogui.click(-1572,794, duration=1)

# Material
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(-1575,802, duration=1)
    pyautogui.hotkey('ctrl','v')

# Próxima página (3)
    pyautogui.click(-1566,865, duration=1)
    sleep(3)

    # fabricante = linha[12].value

    # pais_origen = linha[13].value

    # observacoes = linha[14].value

    # codigo_de_barras = linha[15].value

    # localizacao_armazem = linha[16].value
