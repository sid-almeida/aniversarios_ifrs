import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import pandas as pd
from datetime import date

# Configurações do Gmail
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "sidnei.alves@caxias.ifrs.edu.br"
sender_password = "apuculacata"
context = ssl.create_default_context()

# Ler a tabela de dados com os aniversariantes
data = pd.read_csv("dados/tabela_aniversariantes.csv")

# Pegar a data atual
data_atual = date.today()

# Iterar sobre os aniversariantes
for index, row in data.iterrows():
    nome = row['nome']
    email_destinatario = row['e-mail']
    data_aniversario = row['data_aniversario']

    # Converter a data de aniversário que está no formato dd/mm/yyyy para um objeto date
    data_aniversario = date(int(data_aniversario[6:]), int(data_aniversario[3:5]), int(data_aniversario[0:2]))

    if data_atual == data_aniversario:
        # Criar uma mensagem de e-mail
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = email_destinatario
        msg['Subject'] = 'O IFRS te deseja um Feliz Aniversário e muitos anos de vida!'

        # Adicionar o corpo do e-mail em HTML
        mensagem = f"""<html>
            <body>
                <p>Olá {nome},</p>
                <p>O IFRS Campus Caxias do Sul está te desejando um Feliz Aniversário!\u2764</p>
                <p>Aqui está uma imagem para celebrar o seu dia:</p>
                <img src="cid:image"><br>
                <p>Tenha um ótimo dia!</p>
            </body>
            </html>
            """
        msg.attach(MIMEText(mensagem, 'html'))

        # Adicionar a imagem .png ao corpo do e-mail
        with open('cartao_aniver.png', 'rb') as image_file:
            image = MIMEImage(image_file.read())
            image.add_header('Content-ID', '<image>')
            msg.attach(image)

        # Enviar o e-mail
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls(context=context)
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, email_destinatario, msg.as_string())

print("E-mails de aniversário enviados com sucesso!")