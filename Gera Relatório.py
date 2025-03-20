import pyautogui
import time

# Passo 1: Abre o menu do Windows
pyautogui.press("win")
time.sleep(1)

# Passo 2: Digita "SiatWEB"
pyautogui.write("SiatWEB", interval=0.1)
time.sleep(1)

# Passo 3: Abre o aplicativo (pressionando Enter)
pyautogui.press("enter")

# Passo 4: Aguarda 2 segundos para o aplicativo abrir
time.sleep(2)

# Passo 5: Digita "4008"
pyautogui.write("4008", interval=0.1)

# Passo 6: Pressiona "Tab" 4 vezes
for _ in range(4):
    pyautogui.press("tab")
    time.sleep(0.2)

# Passo 7: Aperta Ctrl + A (Seleciona tudo)
pyautogui.hotkey("ctrl", "a")
time.sleep(0.5)

# Passo 8: Digita "ARTUR"
pyautogui.write("ARTUR", interval=0.1)

# Passo 9: Aperta Enter
pyautogui.press("enter")

# **Passo intermediário: Aguarda 3 segundos**
time.sleep(3)

# Passo 10: Posiciona o mouse em X=93, Y=537
pyautogui.moveTo(93, 537, duration=0.5)

# Passo 11: Clica com o botão esquerdo
pyautogui.click()

# Passo 12: Posiciona o mouse em X=80, Y=118
pyautogui.moveTo(80, 118, duration=0.5)

# Passo 13: Clica com o botão esquerdo
pyautogui.click()

# Passo 14: Posiciona o mouse em X=131, Y=145
pyautogui.moveTo(131, 145, duration=0.5)

# Passo 15: Clica com o botão esquerdo
pyautogui.click()

# **Passo 16: Aguarda 4 segundos**
time.sleep(4)

# **Passo 17: Aperta Tab 17 vezes**
for _ in range(17):
    pyautogui.press("tab")
    # time.sleep(0.1)

# **Passo 18: Digita "01/03/2025"**
pyautogui.write("01/03/2025", interval=0.1)

# **Passo 19: Aperta Tab**
pyautogui.press("tab")
time.sleep(0.2)

# **Passo 20: Digita "31/03/2025"**
pyautogui.write("31/03/2025", interval=0.1)

# **Passo 21: Posiciona o mouse em X=428, Y=198**
pyautogui.moveTo(428, 198, duration=0.5)

# **Passo 22: Clica com o botão esquerdo**
pyautogui.click()

# **Passo 23: Posiciona o mouse em X=465, Y=429**
pyautogui.moveTo(465, 429, duration=0.5)

# **Passo 24: Clica com o botão esquerdo**
pyautogui.click()

# **Passo 25: Posiciona o mouse em X=1191, Y=125**
pyautogui.moveTo(1191, 125, duration=0.5)

# **Passo 26: Clique com o botão direito**
pyautogui.click(button="right")

# **Passo 27: Posiciona o mouse em X=1106, Y=203**
pyautogui.moveTo(1106, 203, duration=0.5)

# **Passo 28: Clique com o botão esquerdo**
pyautogui.click()

# **Passo 29: Posiciona o mouse em X=1334, Y=131**
pyautogui.moveTo(1334, 131, duration=0.5)

# **Passo 30: Clique com o botão esquerdo**
pyautogui.click()

# **Passo 31: Aguarda 15 segundos**
time.sleep(15)

# **Passo 32: Posiciona o mouse em X=667, Y=571**
pyautogui.moveTo(667, 571, duration=0.5)

# **Passo 33: Clique com o botão direito**
pyautogui.click(button="right")

# **Passo 34: Posicione o mouse em X=746, Y=611**
pyautogui.moveTo(746, 611, duration=0.5)

# **Passo 35: Clique com o botão esquerdo**
pyautogui.click()

# Intervalo
time.sleep(3)

# **Passo 36: Digita "lancamentos-mes"**
pyautogui.write("lancamentos-mes", interval=0.1)

# **Passo 37: Posiciona o mouse em X=1089, Y=46**
pyautogui.moveTo(1089, 46, duration=0.5)

# **Passo 38: Clique com o botão esquerdo**
pyautogui.click()

# **Passo 39: Digita "Área de Trabalho"**
pyautogui.write(r"C:\Users\Cliente\Desktop", interval=0.1)

# **Passo 40: Aperta Enter**
pyautogui.press("enter")
time.sleep(1)

# **Passo 41: Alt + L**
pyautogui.hotkey("alt", "l")
time.sleep(1)

# **Passo 42: Enter 2 vezes**
pyautogui.press("enter")
time.sleep(0.5)
pyautogui.press("enter")

# **Passo 43: Abre o menu iniciar**
pyautogui.press("win")
time.sleep(1)

# **Passo 44: Digita "lancamentos mes.csv"**
pyautogui.write("lancamentos mes.csv", interval=0.1)
time.sleep(1)

# **Passo 45: Aperta Enter para abrir o arquivo**
pyautogui.press("enter")

# **Passo 46: Aguarda 2 segundos para o arquivo abrir**
time.sleep(2)

# **Passo 47: Aperta Enter (caso haja algum aviso na abertura)**
pyautogui.press("enter")

# **Passo 48: Aguarda 2 segundos**
time.sleep(2)

# **Passo 49: Pressiona Ctrl + Shift + S (Salvar Como)**
pyautogui.hotkey("ctrl", "shift", "s")
time.sleep(1)

# **Passo 50: Aperta a tecla de seta para a direita para acessar o formato do arquivo**
pyautogui.press("right")

# **Passo 51: Pressiona Backspace 3 vezes para apagar a extensão existente**
for _ in range(3):
    pyautogui.press("backspace")
    time.sleep(0.1)

# **Passo 52: Digita "xlsx" para alterar o formato**
pyautogui.write("xlsx", interval=0.1)

# **Passo 53: Aperta Enter para confirmar a alteração**
pyautogui.press("enter")
