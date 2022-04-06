import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from platform import python_version
from config import password, login, path, your_email
from message_html import html
from create_set_of_email_from_excel import create_list_excel

x = create_list_excel(path)
server = 'smtp.mail.ru'
user = login         # create config.py with login
password = password  # create config.py with password
recipients = x
sender = login
to_addr = your_email
subject = 'Поздравление от Группы Компаний "Стройбат!"'

filepath = 'static/happy_NY2022.jpg'
basename = os.path.basename(filepath)
filesize = os.path.getsize(filepath)

msg = MIMEMultipart('alternative')
msg['Subject'] = subject
msg['From'] = 'stroybat <' + sender + '>'
msg['To'] = to_addr  # ','.join(recipients)
msg['Reply-To'] = sender
msg['Return-Path'] = sender
msg['X-Mailer'] = 'Python/' + (python_version())

part_html = MIMEText(html, 'html')
part_file = MIMEBase('application', 'octet-stream; name="{}"'.format(basename))
part_file.set_payload(open(filepath, 'rb').read())
part_file.add_header('Content-Description', basename)
part_file.add_header('Content-Description', 'attachment; filename="{}"; size={}'.format(basename, filesize))
encoders.encode_base64(part_file)

msg.attach(part_html)
msg.attach(part_file)

mail = smtplib.SMTP_SSL(server)
mail.login(user, password)
mail.sendmail(sender, [to_addr] + recipients, msg.as_string())
mail.quit()
print('Все письма отправлены, открываем шампанское!')
