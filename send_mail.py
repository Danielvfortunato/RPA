from email.message import EmailMessage
import smtplib

def enviar_email(id_solicitacao, numero_nota):
    with open('senha_email.txt') as f:
        senha_do_email = f.readline().strip()
    
    email = "filipeafort@gmail.com"
    destinatario = "tiago@geracao.com.br"
    assunto = "Lançamento bem sucedido!"
    corpo_email = f"""
    Lançamento bem sucedido!
    Número da solicitação: {id_solicitacao}
    Número da nota: {numero_nota}
    """

    msg = EmailMessage()
    msg['Subject'] = assunto
    msg['From'] = email
    msg['To'] = destinatario
    msg.set_content(corpo_email)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email, senha_do_email)
        smtp.send_message(msg)

