#!/usr/bin/env python3
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import yaml
import os

with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f)

files = [
    'output/20.06.2026_07/zakupki_20.06.2026_01.xlsx',
    'plans_parser/output/20.06.2026_10/Plans_20.06.2026_01.xlsx',
    'output/20.06.2026_08/zakupki_extended_20.06.2026_08.xlsx'
]

msg = MIMEMultipart()
msg['From'] = config['email']['sender_email']
msg['To'] = config['email']['recipient_email']
msg['Subject'] = f"Все отчеты за 20.06.2026"

body = """
Все три файла за 20.06.2026:
1. zakupki_20.06.2026_01.xlsx - Закупки (225 шт.)
2. Plans_20.06.2026_01.xlsx - Планы-графики (690 позиций)
3. zakupki_extended_20.06.2026_08.xlsx - Детали лотов (225 лотов)
"""
msg.attach(MIMEText(body, 'plain', 'utf-8'))

for filepath in files:
    if os.path.exists(filepath):
        with open(filepath, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            filename = os.path.basename(filepath)
            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
            msg.attach(part)
            print(f"📎 Добавлен: {filename}")

context = ssl.create_default_context()
with smtplib.SMTP_SSL(config['email']['smtp_server'], config['email']['smtp_port'], context=context) as server:
    server.login(config['email']['sender_email'], config['email']['sender_password'])
    server.send_message(msg)

print("✅ Отправлено!")
