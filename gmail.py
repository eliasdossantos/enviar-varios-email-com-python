import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib

# Devemos instalar a biblioteca ( pip install secure-smtplib ) do python pelo cmd.


# Aqui é o comando que vai pegar a planilha com qual vamos utilizar para mandas as informações para cada e-mail.
clientes = pd.read_excel('./clientes.xlsx')

# Aqui é um loop que vai de linha em linha da planilha pegando o nome e o e-mail para envia as informações.
for index, cliente in clientes.iterrows():
    msg = MIMEMultipart()

    # Aqui vai ser o título do email que deverá ser enviado.
    msg['Subject'] = "Asssunto"

    # Email de onde vai ser enviado as informações.
    msg['From'] = "e-mail de quem vai manda as informações"

    # Todos os email que devem receber as informações.
    msg['To'] = cliente['E-mail']

    # Aqui vai o corpo do email com toda a mensagem ou o conteúdo.
    # Deixar as  3 """ para envia texto com mais de 20 linhas.
    message = """ Olá, programadores...
                  Olá, programadores...
                  Olá, programadores...
                  Olá, programadores...
    """

    # Aqui não mexer pois nesse comando ele vai colocar todo os dados no corpo do email e os anexos necessários.
    msg.attach(MIMEText(message, 'plain'))

    # ======================== Começo da configuração do servidor ========================
    
    # Servidor do email é um serviço de hospedagem de e-mail no qual eles são armazenados.
    server =smtplib.SMTP('smtp.gmail.com', port=587)

    # Aqui é o Start para se conectar com o servidor. 
    server.starttls()

    # Aqui é para fazer login na conta do email para envia as informações.
    # Conta do Gmail juntamente com a senha de aplicativo que já deve ter sido criada depois de ativar a verificação em duas etapas.
    server.login("e-mail de quem vai manda as informações", "senha de aplicativo do email que vai manda as informações")

    # Aqui vai enviar nosso email padrão juntamente com todas as informações que colocamos.
    server.sendmail(msg['From'], msg['To'], msg.as_string().encode('utf-8'))

    # Aqui fecha a conexão com o servidor para não ficar aberto depois de enviar toda a menagem.
    server.quit()

    print("Email Enviado:")
