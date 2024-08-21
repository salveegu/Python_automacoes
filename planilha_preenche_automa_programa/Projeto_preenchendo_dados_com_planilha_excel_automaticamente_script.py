#Pacote que permite trabalhar com planilhas excel com Py
import openpyxl

#Pacote que permite pegar a interação do mouse na tela
import pyautogui

#Abrindo a planilha que desejo trabalhar
planilha = openpyxl.load_workbook('vendas_de_produtos.xlsx')

#Abrindo a pagina da planilha que desejo operar
pagina_vendas = planilha['vendas']


#iterando cada linha da planilha

#especificando que parte da coluna eu quero que ele pegue os dados
for linha in pagina_vendas.iter_rows(min_row=2):
    #alt+shift+setaBaixo para duplicar a linha anterior
    #nome
    pyautogui.click(2897,178,duration=1.0)
    pyautogui.write(linha[0])
    #produto
    pyautogui.click(2898,202,duration=1.0)
    pyautogui.write(linha[1])
    #quantidade
    pyautogui.click(2906,232,duration=1.0)
    pyautogui.write(str(linha[2]))
    #categoria
    pyautogui.click(2966,256,duration=1.0)
    pyautogui.write(linha[3])
    #salvar
    pyautogui.click(2835,283,duration=1.0)
    #ok
    pyautogui.click(822,573,duration=1.0)
    # print(linha[0].value)
    # print(linha[1].value)
    # print(linha[2].value)
    # print(linha[3].value)




