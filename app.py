import streamlit as st
import pandas as pd
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import pandas as pd
from datetime import date
import os
import codecs
import openpyxl

# configurei o streamlit para abrir em wide mode
st.set_page_config(layout="wide")

# Configurações do Gmail
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "sidnei.alves@caxias.ifrs.edu.br"
sender_password = "apuculacata"
context = ssl.create_default_context()

with st.sidebar:
    st.image("https://github.com/sid-almeida/aniversarios_ifrs/blob/main/logo_ifrs.png?raw=true", width=250)
    st.write('---')
    choice = st.radio("**Navegação:**", ("Atualizar dados", "Enviar e-mail"))
    st.info("Esta aplicação permite o envio de e-mails em massa     para os aniversariantes do dia.")
    st.write('---')

if os.path.exists("dados/tabela_aniversariantes.csv"):
    dataframe = pd.read_csv("dados/tabela_aniversariantes.csv")

# Se a opção for atualizar os dados
if choice == "Atualizar dados":
    # Upload do arquivo xlsx
    st.write("## Atualizar dados")
    uploaded_file = st.file_uploader("Escolha o arquivo .xlsx com os dados dos aniversariantes", type="xlsx")
    # Transformar o arquivo em dataframe
    if uploaded_file is not None:
        dataframe = pd.read_excel(uploaded_file)
        # Salvar o dataframe em um arquivo csv na pasta do Github
        dataframe.to_csv("dados/tabela_aniversariantes.csv", index=False)
        st.success("Dados atualizados com sucesso!")
        st.dataframe(dataframe)
    else:
        st.info("Faça o upload do arquivo .xlsx com os dados dos aniversariantes.")


# Se a opção for enviar os e-mails
if choice == "Enviar e-mail":
    data = pd.read_csv("dados/tabela_aniversariantes.csv")
    # Pegar a data atual
    data_atual = date.today()

    # Mostrei a data atual com o mês por extenso e em português
    st.write("## Enviar e-mail")
    st.write("Hoje é: ", data_atual.strftime("%d de %B de %Y"))
    st.write('---')

    # Mostrei um resumo dos dados
    st.write("### Resumo dos dados:")
    st.write("Total de aniversariantes: ", data.shape[0])
    st.write("Total de aniversariantes hoje: ", data[data['ANIVERSARIO'] == data_atual.strftime("%d/%m/%Y")].shape[0])
    st.dataframe(data[data['ANIVERSARIO'] == data_atual.strftime("%d/%m/%Y")])
    st.write('---')

    # Salvei o dataframe filtrado com os aniversariantes de hoje em um nono dataframe
    data_hoje = data[data['ANIVERSARIO'] == data_atual.strftime("%d/%m/%Y")]


    # Se o botão for pressionado
    if st.button("Enviar e-mails"):
        for index, row in data.iterrows():
            nome = row['SERVIDOR']
            email_destinatario = row['E-MAIL']
            data_aniversario = row['ANIVERSARIO']

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
                        <img src="cid:image"><br>
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

                # Excluí os dados da coluna data que já foram enviados por e-mail
                data = data[data['ANIVERSARIO'] != data_atual.strftime("%d/%m/%Y")]

                # Modifiquei o dataframe data_hoje adicionando um ano à coluna ANIVERSARIO
                data_hoje['ANIVERSARIO'] = data_hoje['ANIVERSARIO'].apply(lambda x: x[0:6] + str(int(x[6:])+1))

                # Concatenei o dataframe data_hoje com o dataframe data
                data = pd.concat([data, data_hoje], ignore_index=True)

                # Salvei o dataframe data em um arquivo csv
                data.to_csv("dados/tabela_aniversariantes.csv", index=False)


        st.success("E-mails de aniversário enviados com sucesso!")
    else:
        st.info("Pressione o botão para parabenizar os aniversariantes.")

    st.write('Made with \u2764 by [Sidnei Almeida](https://www.linkedin.com/in/saaelmeida93/)')

