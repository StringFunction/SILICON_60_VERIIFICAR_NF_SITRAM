import logging
import pyautogui as p
import os
import traceback
from src.model.login_alterdata import login_alterdata
from src.model.bi_estoque import bi_estoque
from src.model.tratar_planilha import tratar_planilha
from src.model.enviar_email import enviar_email

logging.basicConfig(filename=r'C:\CODIGOS\SILICON_60_VERIIFICAR_NF_SITRAM\log.log',level=logging.INFO, format=' \
                    %(asctime)s - %(levelname)s - %(message)s', datefmt='%d/%m/%Y - %I:%M:%S %p', filemode='w')

print('RPA SILICONTECH -VERIFICAR NOTAS NO SITRAM\n')
logging.info('RPA SILICONTECH -VERIFICAR NOTAS NO SITRAM\n')
print('ESSE BOTVERIFICAR NOTAS NO SITRAM')

print('Encerrando instancias anteriores')
logging.info('Encerrando instancias anteriores')

os.system("taskkill /f /im Bimer.exe")


try:
    caminho = r'C:\RPA\arquivos\SILICON_26_RPA_ATUALIZA_ENTRADA_NF_SIGET'
    extensao_excel = '.xlsx'
    for arquivos in os.listdir(caminho):
        if arquivos.endswith(extensao_excel):
            os.remove(os.path.join(caminho,arquivos))
    print('removendo arquivos inicialmente')


    login_alterdata()
    bi_estoque()
    
    #tratar_planilha()
    #enviar_email()

except Exception:
    msg_error = traceback.format_exc()
    print(f'{msg_error}')
    logging.info(f'{msg_error}')

    print('Enviando Email de Erro!')
    logging.info('Enviando Email de Erro!')

finally:
    os.system("taskkill /f /im Bimer.exe")
    #caminho = r'C:\RPA\arquivos\ATUALIZA_ENTRADA_NF_SIGET'
    extensao_excel = '.xlsx'
    for arquivos in os.listdir(caminho):
        if arquivos.endswith(extensao_excel):
            os.remove(os.path.join(caminho,arquivos))
    print('\nTRABALHO COMPLETO!')
    logging.info('\nTRABALHO COMPLETO!')
