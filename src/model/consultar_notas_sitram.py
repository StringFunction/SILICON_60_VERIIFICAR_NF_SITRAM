from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
from selenium  import webdriver
from traceback import format_exc
import logging
import pyautogui as p




def Consulta_Chave_Acesso(chave, navegador):
    print("PESQUISANDO CHAVE DE ACESSO SINTRAM")
    logging.info("PESQUISANDO CHAVE DE ACESSO SINTRAM")
    p.sleep(1)
    WebDriverWait(navegador,20).until(ec.visibility_of_element_located((By.ID, "chaveAcesso")))


    try:
        navegador.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/section/div[1]/form/div[2]/div/input").click()
        p.sleep(1)
        navegador.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/section/div[1]/form/div[2]/div/input").send_keys(chave, Keys.RETURN)
    except:
        navegador.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/section/div[1]/form/div[3]/div/input").click()
        p.sleep(1)
        navegador.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/section/div[1]/form/div[3]/div/input").send_keys(chave, Keys.RETURN)
    
    p.sleep(1)
    WebDriverWait(navegador, 15).until(ec.visibility_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/div/section/form/div[3]/table")))
    p.sleep(1)
    listas =  navegador.find_elements(By.XPATH, "/html/body/div[1]/div[2]/div/section/form/div[3]/table/tbody/tr")
    
    if len(listas) > 0:
       dados = listas[0].text.split(" ")
       print(dados[1])
       p.sleep(5)
       navegador.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/section/form/div[2]/div/button").click()
       return True
    else:

        return False