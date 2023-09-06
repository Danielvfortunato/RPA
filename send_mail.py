from email.message import EmailMessage
import smtplib

def enviar_email(id_solicitacao, numero_nota):
    senha_do_email = "F39659122332645EAE5F781FFE818E74EE1D"

    email = "alertawise@wisemanager.com.br"
    destinatarios = ["thiago@geracao-motor.com.br", "daniel@geracao-motor.com.br"]
    assunto = f"Lançamento da SG: {id_solicitacao} efetuado com sucesso pelo RPA"
    corpo_email = f"""
    Número da solicitação: {id_solicitacao}
    Número da nota: {numero_nota}
    """

    with smtplib.SMTP_SSL('smtp.elasticemail.com', 465) as smtp:
        smtp.login(email, senha_do_email)
        
        for destinatario in destinatarios:
            msg = EmailMessage()
            msg['Subject'] = assunto
            msg['From'] = email
            msg['To'] = destinatario
            msg.set_content(corpo_email)
            
            smtp.send_message(msg)

# Exemplo de uso
# enviar_email("123123", "123123")
