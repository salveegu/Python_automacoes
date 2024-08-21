import openpyxl
import pyautogui

# Carregar a planilha
workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

# Iterar sobre as linhas da planilha, começando pela segunda linha
for linha in vendas_sheet.iter_rows(min_row=2):
    pyautogui.click(2897, 178, duration=1.0)
    pyautogui.write(str(linha[0].value))  # produto
    
    pyautogui.click(2898, 202, duration=1.0)
    pyautogui.write(str(linha[1].value))  # quantidade
    
    pyautogui.click(2906, 232, duration=1.0)
    pyautogui.write(str(linha[2].value))  # categoria
    
    pyautogui.click(2966, 256, duration=1.0)
    pyautogui.write(str(linha[3].value))  # salvar
    
    pyautogui.click(2835, 283, duration=1.0)  # botão salvar
    pyautogui.click(822, 573, duration=1.0)  # botão ok
