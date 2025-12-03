import pandas as pd

from estilo_email import estilo
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage 
from datetime import date
import os
from email import encoders # Codifica para base 64
import logging




#* Importacoes




def enviar_email():
    # credenciais = open('C:\\RPA\\arquivos\\credenciais\\email_pybot.txt','r')

    # chaves = credenciais.readlines()
    # credenciais.close()

    # #sender = 'janderamancio@gmail.com'
    # sender = chaves[0][:-1]
    # password = chaves[1] 

    credenciais = open(r'D:\RPA\credenciais\email_clecio.txt')
    chaves = credenciais.readlines()
    credenciais.close()

    sender = chaves[0][:-1]
    password = chaves[1]

    #* Data de hoje com formatacao
    data_atual = date.today()
    data_em_texto = data_atual.strftime('%d_%m_%Y')
    data_em_barra = data_atual.strftime('%d/%m/%Y')

    receiver = ("franciscoclecioti@gmail.com")

  
    print(f'-- enviando email para {receiver}')
    logging.info(f'-- enviando relatorio por email para {receiver}')

    # Configurando o objeto MIME
    message = MIMEMultipart('related')
    message['From'] = sender
    message['To'] = receiver
    message['Subject'] = f'RPA - PASSAGEM DE NOTAS PELO SITRAM - {data_em_texto}'

    # Encapsulando a mensagem alternativa
    msgAlternative = MIMEMultipart('alternative')
    message.attach(msgAlternative)

    #* Caminho do Arquivo
    nome_arquivo = r'D:\RPA\SILICON_60_VERIIFICAR_NF_SITRAM\planilha_para_consultar_sitram.xlsx'

    #* Carregando o DataFrame para Exibir no Corpo do receiver
    df = pd.read_excel(nome_arquivo)
    df = df[["Nr. Documento","Chave de Acesso","Fornecedor","Status"]]
    tabela = df.to_html(index=False)

    #* Corpo do receiver
    body = f"""
        <h3 style="color: #000000; font-weight: bold;">
        RPA - Passagem de notas pelo sitram - {data_em_barra}:
        </h3>
        {estilo(tabela)}
    """
    msgText = MIMEText(body, 'html')
    msgAlternative.attach(msgText)

    # #* Adicionando Arquivo como Anexo
    with open(nome_arquivo, 'rb') as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={nome_arquivo.split("/")[-1]}')
        message.attach(part)

    #* Abrindo uma nova sessao e se conectando no servidor SMTP do Gmail
    session = smtplib.SMTP('smtp.gmail.com', 587)
    session.starttls()

    #* Login utilizando usuario e password
    session.login(sender, password)
    
    #* Enviando e-mail
    text = message.as_string()
    session.sendmail(sender, receiver.split(','), text)

    #* Fechando conexao com Gmail
    session.quit()

    print('-- email enviado com sucesso')
    logging.info('-- email enviado com sucesso')





# def enviar_email_erro():
    # data_atual = date.today()
    # data_em_texto = data_atual.strftime('%d_%m_%Y')
    # data_em_barra = data_atual.strftime('%d/%m/%Y')
    # receiver = 'janderamancio@gmail.com'

    # message = MIMEMultipart('related')
    # message['From'] = sender
    # message['To'] = receiver
    # message['Subject'] = 'RPA-26---ATUALIZA ENTRADA NF SIGET  ' + data_em_texto

    # msgAlternative = MIMEMultipart('alternative')
    # message.attach(msgAlternative) 

    # msgText = MIMEText('RPA-26---ATUALIZA ENTRADA NF SIGET ')
    # msgAlternative.attach(msgText) 

    # msgText = MIMEText(f'<br>AVISO: ERRO INESPERADO FOI ENCONTRADO IMPEDINDO O ENVIO DA NF À SEFAZ.\nSEGUE IMAGEM DO ERRO EM ANEXO.\nESSE É UM EMAIL AUTOMATICO, POR FAVOR NAO RESPONDER.</br>', 'html')
    # msgAlternative.attach(msgText)

    # fp = open(r'C:\RPA\arquivos\erro_fatbot.png', 'rb')
    # msgImage = MIMEImage(fp.read())
    # fp.close()

    # msgImage.add_header('Content-ID', '<image>')
    # message.attach(msgImage)

    # session = smtplib.SMTP('smtp.gmail.com', 587)

    # session.starttls()

    # session.login(sender, password)

    # text = message.as_string()
    # session.sendmail(sender, receiver.split(","), text)
    # session.quit()

    # print('Enviado e-mail com imagem de erro.')
    # os.remove(r'C:\RPA\arquivos\erro_fatbot.png')
if __name__ == '__main__':
    enviar_email()
