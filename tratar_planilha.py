import pandas as pd
import pyautogui as p
import gspread
import gspread_dataframe as gd
import openpyxl
from xls2xlsx import XLS2XLSX


def tratar_planilha():
   
    # LER A PLANILHA NO DIRETORIO E O NOME DA ABA QUE DESEJO TRABALHAR
    df = pd.read_excel(r'C:\RPA\arquivos\SILICON_26_RPA_ATUALIZA_ENTRADA_NF_SIGET\entrada_nf_siget.xlsx',sheet_name='entrada_nf_siget')
    
    #  GERO UM NOVO ARQUIVO NO DIRETORIO , NOMEIO A ABA,RETIRO O INDEX E O CABECALHO
    nova_planilha = df.to_excel(r'C:\RPA\arquivos\SILICON_26_RPA_ATUALIZA_ENTRADA_NF_SIGET\nova.xlsx',sheet_name='base',index=False)
    
    # LEIO O NOVO ARQUIVO, E A ABA ESPECIFICA
    df = pd.read_excel(r'C:\RPA\arquivos\SILICON_26_RPA_ATUALIZA_ENTRADA_NF_SIGET\nova.xlsx',sheet_name ='base')

    
    # SELECIONO AS LINHAS E COLUNAS QUE SERAO GERADAS NO DATA FRAME
    # df.iloc[seleciona linhas, seleciona colunas]
    df2 = df.iloc[0:-1,0:-1]

    

#####################################  CONFIGURACAO DE ACESSO PLANILHA GOOGLE  #####################################
    
    # INICIANDO O SERVICO DA PLANILHA COM A CREDENCIAL DO JSON
    gc = gspread.service_account(r'C:\RPA\arquivos\credenciais\novoprojetojander.json')

    #ABRINDO A PLANILHA PELO NOME
    sh = gc.open('2025 - Operação Interestadual')

    #ESCOLHO QUAL SHEET
    worksheet = sh.worksheet("ENTRADA")
    #worksheet.batch_clear(["A2:O2"])
    worksheet.clear()

    # Atualizando a planilha do Google Sheets
    gd.set_with_dataframe(worksheet,df2,1,1)#(3,2 copiar no indice linha 3 e indice coluna 2)
    p.sleep(3)

    ###################### ATUALIZAR PLANILHA INTERNA  ##############################################

    #p.alert('vai atualizar planilha interna')
    # INICIANDO O SERVICO DA PLANILHA COM A CREDENCIAL DO JSON
    gc = gspread.service_account(r'C:\RPA\arquivos\credenciais\novoprojetojander.json')

    #ABRINDO A PLANILHA PELO NOME
    sh = gc.open('2025 - Operação Interna')

    #ESCOLHO QUAL SHEET
    worksheet = sh.worksheet("ENTRADA")
    #worksheet.batch_clear(["A2:O2"])
    worksheet.clear()

    # Atualizando a planilha do Google Sheets
    gd.set_with_dataframe(worksheet,df2,1,1)#(3,2 copiar no indice linha 3 e indice coluna 2)
    p.sleep(3)



    
