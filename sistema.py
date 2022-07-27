import pyautogui
import time
import pyperclip
import pandas

pyautogui.PAUSE = 1


# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
# Passo 1: Entrar no sistema (no nosso caso entrar no link)
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

pyautogui.alert("Para iniciar o programa clique em OK e não mexa em nada!")
pyautogui.press("win")
pyautogui.write("chrome")
time.sleep(2)
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
# site está carregando 
time.sleep(5)


# =-=-=-=-=-=-=-=-=-=-s=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
# Passo 2: Navegar no sistema e encontrar a base de dados (entrar na pasta Exportar) 
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-= -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

pyautogui.click(x=479, y=402, clicks=2)
time.sleep(1)


# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
# Passo 3: Fazer o Download da base de dados
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

pyautogui.click(x=522, y=507) # clicou no arquivo
pyautogui.click(x=1639, y=249) # clico nos 3 pontos
pyautogui.click(x=1296, y=845) # fazer download
time.sleep(5)


# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
# Passo 4: Calcular os indicadores (faturamento, quantidade de produtos)
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

tabela = pandas.read_excel(r"C:\Users\Camilli\Downloads\Vendas - Dez.xlsx")

# soma da coluna valor final 
faturamento = tabela["Valor Final"].sum()
#soma da coluna quantidade
quantidade = tabela["Quantidade"].sum()

print(quantidade)
print(faturamento)


# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
# Passo 5: Entrar no Email
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

pyautogui.hotkey("ctrl", "t") #abrindo a nova aba
pyperclip.copy("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox") #copiando o link
pyautogui.hotkey("ctrl", "v") #colando o link
pyautogui.press("enter")
time.sleep(5)


# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
# Passo 6: Mandar um email para a diretoria com os indicadores
# =-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

pyautogui.click(x=44, y=268) # clicar no +
time.sleep(2)
pyautogui.click(x=1760, y=326)
time.sleep(4)
#escrevendo o destinatário
pyperclip.copy("caahjoannes@gmail.com") #copiando o email
pyautogui.hotkey("ctrl", "v") #colando o email
pyautogui.press("tab") # passando para o assunto

#escrevendo o assunto
pyperclip.copy("Relatório Diário") #copiando o assunto
pyautogui.hotkey("ctrl", "v") #colando o assunto
pyautogui.press("tab") #passando para o texto

#escrevendo o texto
texto = f"""
Prezados, Bom dia.

O faturamento de ontem foi de: R$ {faturamento:,.2f}
A quantidade de produtos foi de {quantidade:,}

Ass
Camilli Joannes - Gerente de Vendas
"""
pyperclip.copy(texto) #copiando o texto
pyautogui.hotkey("ctrl", "v") #colando o texto

#enciar o email
pyautogui.hotkey("ctrl", "enter")

