import pyautogui as p
import pandas as pd
import os
import traceback
import logging
from datetime import date,timedelta
import datetime
## DATA HOJE
data_atual = datetime.datetime.now()
horas =  data_atual.hour

if int(horas) < 18:
   data_inicial = (data_atual - timedelta(1)).strftime("%d/%m/%Y")
   data_final = data_inicial
else:
    data_inicial = data_atual.strftime("%d/%m/%Y")
    data_final = data_inicial







def bi_estoque():
    print("ACESSANDO MODULO DE BI ESTOQUE")

    try:

        novo_pesquisar = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\novo_pesquisar.png', confidence=0.95)
        while novo_pesquisar == None:
            novo_pesquisar = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\novo_pesquisar.png', confidence=0.95)
        if novo_pesquisar != None:
            p.click(novo_pesquisar)
        p.typewrite('BI Estoque')
        p.sleep(1)

        novo_bi_estoque = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\novo_bi_estoque.png', confidence=0.95)
        if novo_bi_estoque != None:
            p.click(novo_bi_estoque)
        p.sleep(2)
       
        recuperar_cenario = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\recuperar_cenario.png', confidence=0.95)
        while recuperar_cenario == None:
            recuperar_cenario = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\recuperar_cenario.png', confidence=0.95)
        if recuperar_cenario != None:
            p.click(recuperar_cenario)
        p.sleep(1)
        p.typewrite('Bot - Verificar se a nota esta no sitram',interval=0.06)
        p.sleep(1)

        cenario_verificar_se_a_nota_esta_sitram = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\cenario_verificar_se_a_nota_esta_sitram.png',confidence=.95)
        while cenario_verificar_se_a_nota_esta_sitram == None:
            cenario_verificar_se_a_nota_esta_sitram = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\cenario_verificar_se_a_nota_esta_sitram.png',confidence=.95)
        if cenario_verificar_se_a_nota_esta_sitram != None:
            p.doubleClick(cenario_verificar_se_a_nota_esta_sitram)
        p.sleep(1)

        p.press('\t',presses=12)
        p.sleep(1)
        p.typewrite(data_inicial)
        p.sleep(1)
        p.press('tab')
        p.typewrite(data_final)
        p.sleep(2)
        filtrar = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\filtrar.png',confidence=.95)
        if filtrar != None:
            p.click(filtrar)
        p.sleep(1)
        aguarde = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\aguarde.png',confidence=.95)
        while aguarde != None:
            aguarde = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\aguarde.png',confidence=.95)
        

        p.sleep(2)
        exportar = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\exportar4.png',confidence=.95)
        while exportar == None:
            exportar = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\exportar4.png',confidence=.95)
        
        if exportar != None:
            p.click(exportar)
        #p.alert('clicoou exportar')
        p.sleep(1)
        p.typewrite(r'C:\CODIGOS\SILICON_60_VERIIFICAR_NF_SITRAM\notas_para_consultar_sitram.xlsx',interval=0.06)
        p.sleep(3)
        p.press('tab',presses=2)
        p.sleep(1)
        p.press('enter')
        p.sleep(2)
        salvando_estatistica = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\salvando_estatistica.png', confidence=0.95)
        while salvando_estatistica != None:
            salvando_estatistica = p.locateCenterOnScreen(r'C:\RPA\arquivos\images\salvando_estatistica.png', confidence=0.95)
   


    except Exception:
        msg_error = traceback.format_exc()
        print(f'{msg_error}')
        logging.info(f'{msg_error}')

        print('Erro!')
        logging.info('Erro!')

