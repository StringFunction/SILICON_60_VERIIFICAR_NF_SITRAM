import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage 
from datetime import date
import os
from email import encoders # Codifica para base 64
import PyPDF2 # Manipula PDF's
import logging



credenciais = open('C:\\RPA\\arquivos\\credenciais\\email_pybot.txt','r')
chaves = credenciais.readlines()
credenciais.close()

#sender = 'janderamancio@gmail.com'
sender = chaves[0][:-1]
password = chaves[1] 


data_atual = date.today()
data_em_texto = data_atual.strftime('%d_%m_%Y')
data_em_texto_barr = data_atual.strftime('%d/%m/%Y')
inicio_mes = data_em_texto_barr[2:]


def enviar_email():

    receiver = 'janderamancio@gmail.com'

    message = MIMEMultipart('related')
    message['From'] = sender
    message['To'] = receiver
    message['Subject'] = 'RPA-26---ATUALIZA ENTRADA NF SIGET ' + data_em_texto_barr

    msgAlternative = MIMEMultipart('alternative')
    message.attach(msgAlternative) 

    msgText = MIMEText('RPA-26---ATUALIZA ENTRADA NF SIGET')
    msgAlternative.attach(msgText) 

    msgText = MIMEText(f'<br>RPA-26---ATUALIZA ENTRADA NF SIGET .\nESSE É UM EMAIL AUTOMATICO, POR FAVOR NAO RESPONDER.</br>', 'html')
    msgAlternative.attach(msgText)

    session = smtplib.SMTP('smtp.gmail.com', 587)

    session.starttls()

    session.login(sender, password)

    text = message.as_string()
    session.sendmail(sender, receiver.split(","), text)
    session.quit()

    print('Email Enviado ')




def enviar_email_erro():
   
    receiver = 'janderamancio@gmail.com'

    message = MIMEMultipart('related')
    message['From'] = sender
    message['To'] = receiver
    message['Subject'] = 'RPA-26---ATUALIZA ENTRADA NF SIGET  ' + data_em_texto_barr

    msgAlternative = MIMEMultipart('alternative')
    message.attach(msgAlternative) 

    msgText = MIMEText('RPA-26---ATUALIZA ENTRADA NF SIGET ')
    msgAlternative.attach(msgText) 

    msgText = MIMEText(f'<br>AVISO: ERRO INESPERADO FOI ENCONTRADO IMPEDINDO O ENVIO DA NF À SEFAZ.\nSEGUE IMAGEM DO ERRO EM ANEXO.\nESSE É UM EMAIL AUTOMATICO, POR FAVOR NAO RESPONDER.</br>', 'html')
    msgAlternative.attach(msgText)

    fp = open(r'C:\RPA\arquivos\erro_fatbot.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()

    msgImage.add_header('Content-ID', '<image>')
    message.attach(msgImage)

    session = smtplib.SMTP('smtp.gmail.com', 587)

    session.starttls()

    session.login(sender, password)

    text = message.as_string()
    session.sendmail(sender, receiver.split(","), text)
    session.quit()

    print('Enviado e-mail com imagem de erro.')
    os.remove(r'C:\RPA\arquivos\erro_fatbot.png')
