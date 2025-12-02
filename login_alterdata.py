import pyautogui as p
import os
import logging
import ctypes
from pynput.keyboard import Controller, Key


def login_alterdata():
    #ler o arquivo com as credenciais de acesso
    credenciais = open('C:\\RPA\\arquivos\\credenciais\\login_alterdata.txt','r')
    chaves = credenciais.readlines()
    credenciais.close()
    
    # login = chaves[0][:-1]
    password = chaves[1]
    #VERIFICAR SE O CAPSLOCK ESTA ATIVO
    keyboard = Controller()
    if ctypes.windll.user32.GetKeyState(0x14) == 1:
        keyboard.press(Key.caps_lock)
        keyboard.release(Key.caps_lock)
        print("Caps Lock foi desativado.")
        
    else:
        print("Caps Lock já está desativado.")
   

    #chamar o executavel do Bimer
    os.startfile('C:\Program Files\Alterdata\ERP\Bimer.exe')
    p.sleep(3)
    
    
    bimmer = p.locateCenterOnScreen('C:/RPA/arquivos/images/novo_bimmer.png', confidence = 0.9)
    while bimmer == None:
        bimmer = p.locateCenterOnScreen('C:/RPA/arquivos/images/novo_bimmer.png', confidence = 0.9)
    
    if bimmer  != None:
        p.sleep(2)
        p.typewrite(password)
        
    
    #procura a imgagem do entrar no alterdata
    novo_entrar = p.locateCenterOnScreen('C:/RPA/arquivos/images/novo_entrar.png', confidence = 0.95)
    if novo_entrar != None:
        p.click(novo_entrar)
       
    
    

    print('Login Concluido')
    logging.info('Login Concluido')