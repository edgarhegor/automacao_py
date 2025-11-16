import pyautogui
import pandas
import pyperclip
import time
import locale
locale.setlocale(locale.LC_ALL, "pt_BR.UFT-8") ##currency BR

#Baixando Arquivo
pyautogui.press("win")
pyautogui.write("Edge")
pyautogui.press("enter")
time.sleep(1)
pyautogui.hotkey("crtl, l")
gdrive = "https://drive.google.com/drive/my-drive"
pyautogui.write(gdrive)
pyautogui.press("enter")
time.sleep(4)
pyautogui.click(x=1132, y=337, clicks=2)
time.sleep(2) 
pyautogui.click(x=423, y=548, button="right")
time.sleep(2)
pyautogui.click(x=529, y=626)

time.sleep(4)
#pyautogui.press("enter")
#pyautogui.click(x=684, y=254)
#time.sleep(3)
#print(pyautogui.position())
#pyautogui.click(x=233, y=301, click= "right") # parametro do clicks = 2 da dois clicks 


# EXCEL
caminho = r"C:\Users\edgar\Downloads\Vendas - Dez.xlsx"

tabela = pandas.read_excel(caminho)

print(tabela)

faturamento = tabela["Valor Final"].sum()

qntd_produtos = tabela["Quantidade"].sum()

faturamento_fmt = locale.format_string("%.2f", faturamento, grouping=True )

qntd_fmt = locale.format_string("%.2f", qntd_produtos)# grouping=True)
print(f" Faturamento = {faturamento}")
print(f" Quantidade de vendas = {qntd_produtos}")

#Abrir email e enviar

time.sleep(3)
pyautogui.hotkey("ctrl","l")

time.sleep(2)
pyautogui.write("https://mail.google.com/mail/u/0/#inbox")
pyautogui.press("enter")
time.sleep(4)
pyautogui.click(x=100, y=203)
time.sleep(2)

pyautogui.click(x=1379, y=472)
email = "edgarhegor@gmail.com"
pyautogui.write(email)
pyautogui.press("tab")
time.sleep(2)
pyautogui.press("tab")
time.sleep(1)
assunto = "Relatório atualizado"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl","v")
pyautogui.press("tab")
time.sleep(1)

texto = f"""
Prezados,

Segue o relatório atualizado de hoje.

Faturamento = R${faturamento_fmt}
Quantidade de Vendas =  {qntd_produtos}

Qualquer dúvida fico à disposição
Abs, Edgar Hegor.
"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

pyautogui.press("tab")
pyautogui.press("enter")


