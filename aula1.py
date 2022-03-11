from sys import displayhook
import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

strings = [
   r"https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing\n",
   "https://www.gmail.com",   
]


def abrirNavegador(url):
   pyautogui.press("winleft")
   pyautogui.write("firefox")
   pyautogui.press("down")
   pyautogui.press("enter")
   pyautogui.write(url)
   pyautogui.press("enter")
   time.sleep(2)


# entrar no sistema
#abrir navegador
# navegar no sistema até encontrar a base de dados
abrirNavegador(strings[0])
pyautogui.click(2269,328, clicks=2)
time.sleep(5)
pyautogui.click(2284,326, button='right')
time.sleep(3)
pyautogui.click(2435,780)
time.sleep(10)

# Importando base de dados
tabela = pd.read_excel("/home/tiago/Downloads/Vendas - Dez.xlsx")

# calcular os indicadores ( fatureamento e quantidade de produtos vendidos)
faturamento = tabela['Valor Final'].sum()
qtd_produtos = tabela['Quantidade'].sum()


# enviar o relatório dos indicadores por e-mail
pyautogui.hotkey("ctrl", "t")
time.sleep(3)
pyautogui.write(strings[1])
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(2017,245)
time.sleep(5)
pyautogui.write("tiago.test.developer@gmail.com")
pyautogui.press("enter")
pyautogui.press("tab")
titulo = "Relatório de Vendas"
pyperclip.copy(titulo)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
texto = f"""Prezados, bom dia!
            
            O faturamento de ontem foi de: R${faturamento:,.2f}
            A quantidade de produtos vendidos foi de: {qtd_produtos:,}

            Abs
            TiagoAragão"""
pyperclip.copy(texto)
time.sleep(2)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")




