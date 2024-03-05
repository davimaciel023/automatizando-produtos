import openpyxl
import pyperclip
import pyautogui
from time import sleep

#Entrar na planilha
workbook = openpyxl.load_workbook ('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Copiar informações de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(196,300, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(264,405, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(192,577, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    codigo_do_produto = linha[3].value
    pyperclip.copy(codigo_do_produto)
    pyautogui.click(249,678, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(265,798, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    dimensao = linha[5].value
    pyperclip.copy(dimensao)
    pyautogui.click(238,896, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    pyautogui.click(177,970, duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(268,330, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    quantidade_estoque = linha[7].value
    pyperclip.copy(quantidade_estoque)
    pyautogui.click(230,430, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(239,541, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(211,651, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    tamanho = linha[10].value
    pyautogui.click(231,755, duration=1)
    if tamanho == 'Pequeno':
        pyautogui.click(247,790, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(238,821, duration=1)
    else:
        pyautogui.click(239,846, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(225,868, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    pyautogui.click(182,937, duration=1)
    sleep(3) 

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(244,354, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(222,458, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    observacao = linha[14].value
    pyperclip.copy(observacao)
    pyautogui.click(211,569, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    codigo_barra = linha[15].value
    pyperclip.copy(codigo_barra)
    pyautogui.click(228,736, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    loc_armazem = linha[16].value
    pyperclip.copy(loc_armazem)
    pyautogui.click(247,835, duration=1)
    pyautogui.hotkey('Ctrl', 'V')

    pyautogui.click(177,914, duration=1)
    sleep(1)
    pyautogui.click(670,238, duration=1)
    sleep(1)
    pyautogui.click(670,238, duration=1)
    sleep(3)
    pyautogui.click(453,633, duration=1)
    sleep(2)