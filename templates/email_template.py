import smtplib, ssl
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import os


class Email():
    def __init__(self, smtp_server, sender, password):
        self.smtp_server = smtp_server
        self.port = 587
        self.sender = sender
        self.password = password

    #context = ssl.create_default_context()
    def format_list_emails_to_string(self, list_emails):
        string_emails = list_emails.__repr__()[1:-1].replace("'", '').replace('"', '')
        return string_emails

    def attach(self, message, path_file):
        with open(path_file, 'rb') as file:
            part = MIMEApplication(file.read(), Name=os.path.basename(path_file))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(path_file)}"'
        message.attach(part)
        return

    def create_message(self, receivers, subject, body, path_files=[], assinatura=None, html=True):
        message = MIMEMultipart()
        # HEADER
        message["From"] = self.sender
        receivers.append(self.sender)
        message["To"] = ', '.join(receivers)
        message["Subject"] = subject
        # message["Bcc"] = self.receiver #Com cópia oculta
        # BODY
        if html == True:
            message.attach(MIMEText(body, "html"))
        else:
            message.attach(MIMEText(body, "plain"))
        for path_file in path_files:
            self.attach(message, path_file)

        if assinatura !=None:
            with open(assinatura, 'rb') as img:
                assinatura = MIMEImage(img.read())
            message.attach(assinatura)
        else:
            pass
        return message

    def send_email(self, message):
        context = ssl._create_unverified_context()
        with smtplib.SMTP(self.smtp_server, self.port) as server:
            server.ehlo('Thiago Martins')
            server.starttls(context=context)
            server.login(self.sender, self.password)
            server.sendmail(message['From'], message['To'].split(), message.as_string())
        return

def TEST():
    sender = input('Sender e-mail: ')
    password = input('Password: ')
    smtp_server = input('SMTP-Server: ')
    receivers = ['thiago.martins@solarar.com.br', 'thiago.martins@solarar.com.br']
    subject = 'E-mail Subject test'
    body = """
<html>
<body>
<h2>Prezado cliente,</h2><br>
<br>
<p>TESTE</p><br>
<br>
Atenciosamente,<br>
</body>
</html>
"""
    path_files = ['./email_template.py',]
    e_mail = Email(smtp_server=smtp_server, sender=sender, password=password)
    message = e_mail.create_message(receivers, subject, body, path_files)
    e_mail.send_email(message)
