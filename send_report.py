#!/usr/bin/env python3
import smtplib
import ssl
import subprocess
from datetime import datetime
import os
import yaml
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

with open('config.yaml', 'r') as f:
    config = yaml.safe_load(f)

EMAIL = config['email']['sender_email']
PASSWORD = config['email']['sender_password']
TO = config['email']['recipient_email']

print(f"{datetime.now()}: Запуск парсера...")

result = subprocess.run(
    ['python3', 'parser.py'],
    cwd='/opt/zakupki-service',
    capture_output=True,
    text=True
)

print(result.stdout)

output_dir = config['excel']['output_dir']
files = [f for f in os.listdir(output_dir) if f.endswith('.xlsx')]
if files:
    latest_file = max(files, key=lambda x: os.path.getctime(os.path.join(output_dir, x)))
    filepath = os.path.join(output_dir, latest_file)
    
    total = 0
    for line in result.stdout.split('\n'):
        if 'Всего закупок:' in line:
            try:
                total = int(line.split(':')[1].strip())
            except:
                pass
    
    msg = MIMEMultipart()
    msg['From'] = EMAIL
    msg['To'] = TO
    msg['Subject'] = f"Отчет по закупкам от {datetime.now().strftime('%d.%m.%Y')}"
    
    body = f"""
Отчет сформирован: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}
Всего закупок найдено: {total}
Файл приложен к письму.
"""
    msg.attach(MIMEText(body, 'plain', 'utf-8'))
    
    with open(filepath, 'rb') as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{latest_file}"')
        msg.attach(part)
    
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(config['email']['smtp_server'], config['email']['smtp_port'], context=context) as server:
        server.login(EMAIL, PASSWORD)
        server.send_message(msg)
    
    print(f"✅ Отчет отправлен на {TO}")
else:
    print("❌ Файлы не найдены")
