import pandas as pd
import pyautogui as pt
import os
import time

pt.FAILSAFE = True  # Mover para canto superior esquerdo
df = pd.read_excel("caminho do arquivo") 

# Conta o número de linhas
num_linhas = df.shape[0]

os.system('start notepad')  # Apenas para testes
time.sleep(0.5)

i = 0
while i < num_linhas:  
    valorREF = df.loc[i, "REFERENCIA"] #REFENCIA DA NOTA    
    valorVAL = df.loc[i, "VALOR"] #VALOR DA NOTA
    valorCOMPT_data = pd.to_datetime(df.loc[i, "DATA"]).strftime("%d/%m/%Y") #data inicial
    valorCOMPT_comp = pd.to_datetime(df.loc[i, "DATA"]).strftime("%m/%Y") #valor de competencia
    valorCNT = df.loc[i, "CONTA"] #CONTA 501/301/101
    valorAPROP = df.loc[i, "APROPRIACAO"] #APROPRIAÇÃO DA NOTA 

    pt.write(f'{valorCOMPT_data}') 
    pt.sleep(1)
    pt.press("tab")
    pt.write(f'{valorCOMPT_comp}')
    pt.sleep(1)
    pt.press("tab")
    pt.write(f'{valorREF}') 
    pt.sleep(1) 

    for _ in range(6):
        pt.press("tab") 

    pt.write(f'{valorAPROP}')
    pt.press("tab")
    pt.sleep(1)
    pt.write(f'{valorCNT}')
    pt.press("tab")
    pt.sleep(1)
    pt.write(f'{valorVAL}')
    pt.sleep(1)
    
    for _ in range(4):
        pt.press("enter")

    pt.press("f5")
    time.sleep(2.5)

    i += 1 