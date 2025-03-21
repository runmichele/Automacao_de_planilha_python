import smtplib
import ssl
import mimetypes
from email.message import EmailMessage 

# 1- Dados do E-mail
password = open('data/senha.txt', 'r').read() # Lendo um arquivo chamado "senha"
print(password)  # Exibe a senha no terminal

from_email='michelemmono@gmail.com'
to_email='michelemmono@gmail.com'
subject='Automação Planilha'
body = """   
Olá. Segue em anexo a automação da planilha
para a empresa XYZ Automação.

Qualquer dúvida estou a disposição!
"""

#2- Montando a estrutura do e-mail
message = EmailMessage()
message['From'] = from_email
message['To'] = to_email
message['Subject'] = subject

message.set_content(body)
safe = ssl.create_default_context()

# 3- Adcionar Anexo 
anexo = 'data/test.xlsx'

# print(mimetypes.guess_type(anexo)[0].split('/'))
mime_type, mime_subtype = mimetypes.guess_type(anexo)[0].split('/')
with open(anexo, 'rb') as a:
    message.add_attachment(
        a.read(),
        maintype=mime_type,
        subtype=mime_subtype,
        filename=anexo
)
    
# 4- Envio do E-mail
with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=safe) as smtp:
    smtp.login(from_email, password)
    smtp.sendmail(
        from_email,
        to_email,
        message.as_string()
    )
