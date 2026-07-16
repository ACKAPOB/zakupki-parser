#!/usr/bin/env python3
"""
Общий запуск всех парсеров и отправка отчётов
Создаёт общую папку для всех результатов и передаёт путь парсерам
через переменную окружения ZAKUPKI_OUTPUT_DIR
"""

import subprocess
import os
import sys
import re
import glob
from datetime import datetime
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import yaml
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/run_all.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

try:
    with open('config.yaml', 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)
    logger.info("✅ Конфиг загружен")
except Exception as e:
    logger.error(f"❌ Ошибка загрузки конфига: {e}")
    sys.exit(1)

OUTPUT_DIR = config.get('excel', {}).get('output_dir', 'output')


def create_session_folder():
    """Создаёт папку для текущей сессии запуска"""
    date_str = datetime.now().strftime("%d.%m.%Y")
    pattern = f"{OUTPUT_DIR}/{date_str}_*"
    existing_folders = glob.glob(pattern)
    
    if not existing_folders:
        folder_number = 1
    else:
        numbers = []
        for folder in existing_folders:
            match = re.search(r'(\d{2}\.\d{2}\.\d{4})_(\d+)', folder)
            if match:
                numbers.append(int(match.group(2)))
        folder_number = max(numbers) + 1 if numbers else 1
    
    folder_name = f"{date_str}_{folder_number:02d}"
    folder_path = os.path.abspath(os.path.join(OUTPUT_DIR, folder_name))
    os.makedirs(folder_path, exist_ok=True)
    
    logger.info(f"📁 Создана папка сессии: {folder_path}")
    return folder_path


def send_email(subject, body, attachments):
    msg = MIMEMultipart()
    msg['From'] = config['email']['sender_email']
    msg['To'] = config['email']['recipient_email']
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain', 'utf-8'))
    
    for filepath in attachments:
        if os.path.exists(filepath):
            with open(filepath, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                filename = os.path.basename(filepath)
                part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                msg.attach(part)
                logger.info(f"📎 Добавлен файл: {filename}")
    
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(config['email']['smtp_server'], config['email']['smtp_port'], context=context) as server:
        server.login(config['email']['sender_email'], config['email']['sender_password'])
        server.send_message(msg)
    logger.info("✅ Письмо отправлено")


def run_parser_realtime(name, command, env, timeout=7200):
    """Запуск парсера с выводом в реальном времени"""
    logger.info(f"\n{'='*70}")
    logger.info(f"🚀 Запуск: {name}")
    logger.info(f"{'='*70}\n")
    
    try:
        process = subprocess.Popen(
            command,
            cwd='/opt/zakupki-service',
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            universal_newlines=True,
            env=env  # ← Передаём окружение с путём к папке
        )
        
        start_time = datetime.now()
        for line in process.stdout:
            line = line.rstrip()
            if line:
                logger.info(line)
            
            elapsed = (datetime.now() - start_time).total_seconds()
            if elapsed > timeout:
                logger.error(f"❌ Превышен таймаут ({timeout}с)")
                process.kill()
                return False
        
        process.wait()
        
        if process.returncode == 0:
            logger.info(f"\n✅ {name} завершён успешно")
            return True
        else:
            logger.error(f"\n❌ {name} завершён с кодом {process.returncode}")
            return False
            
    except Exception as e:
        logger.error(f"❌ Ошибка запуска {name}: {e}")
        return False


def main():
    start_time = datetime.now()
    
    logger.info("=" * 70)
    logger.info("ЗАПУСК ВСЕХ ПАРСЕРОВ")
    logger.info(f"Время запуска: {start_time.strftime('%d.%m.%Y %H:%M:%S')}")
    logger.info("=" * 70)
    
    # === ШАГ 0: Создаём общую папку для сессии ===
    session_folder = create_session_folder()
    
    # === Готовим окружение для парсеров ===
    env = os.environ.copy()
    env['ZAKUPKI_OUTPUT_DIR'] = session_folder
    env['REQUESTS_CA_BUNDLE'] = '/etc/ssl/certs/ca-certificates.crt'
    env['SSL_CERT_FILE'] = '/etc/ssl/certs/ca-certificates.crt'
    logger.info(f"🔧 Переменная ZAKUPKI_OUTPUT_DIR={session_folder}")
    logger.info(f"🔧 SSL сертификаты: {env['REQUESTS_CA_BUNDLE']}")
    
    results = {
        
        'plans': False,
        'extended': False
    }
    
    # 1. Парсер планов-графиков
    logger.info(f"\n[1/3] Запуск парсера планов-графиков...")
    results['plans'] = run_parser_realtime(
        "Парсер планов-графиков",
        [sys.executable, 'plans_parser/final_parser_v6.py', '--year', str(config['plans_parser']['default_year'])],
        env=env,
        timeout=7200
    )
    
    # 3. Расширенный парсер закупок
    logger.info(f"\n[2/3] Запуск расширенного парсера закупок с деталями лотов...")
    results['extended'] = run_parser_realtime(
        "Расширенный парсер закупок",
        [sys.executable, 'parser_extended.py', '-y', str(config['parser_extended']['default_year'])],
        env=env,
        timeout=5400
    )
    
    # 4. Сбор файлов из папки сессии
    logger.info(f"\n[3/3] Сбор файлов из папки сессии...")
    logger.info(f"📁 Папка: {session_folder}")
    
    attachments = sorted(glob.glob(f"{session_folder}/*.xlsx"))
    
    if not attachments:
        logger.error("❌ Файлы не найдены в папке сессии!")
        return
    
    logger.info(f"📊 Найдено файлов: {len(attachments)}")
    for f in attachments:
        logger.info(f"   - {os.path.basename(f)}")
    
    # 4.1. Сравнение файлов планов-графиков
    logger.info(f"\n[3.1/3] Сравнение файлов планов-графиков...")
    today_str = datetime.now().strftime("%d.%m.%Y")
    changes_txt_path = os.path.join(session_folder, f"changes_{today_str}.txt")
    changes_body_text = ""
    
    try:
        res = subprocess.run(
            [sys.executable, 'compare_plans.py', '--output-file', changes_txt_path],
            cwd='/opt/zakupki-service',
            capture_output=True, text=True, timeout=120
        )
        
        if res.returncode == 0 and os.path.exists(changes_txt_path):
            attachments.append(changes_txt_path)
            logger.info(f"✅ Файл изменений создан: {os.path.basename(changes_txt_path)}")
            
            # Читаем содержимое для вставки в тело письма
            with open(changes_txt_path, 'r', encoding='utf-8') as cf:
                changes_body_text = "\n" + cf.read()
        else:
            changes_body_text = "\n⚠️ Не удалось сформировать отчёт об изменениях."
            logger.warning(f"⚠️ Ошибка сравнения планов: {res.stderr}")
    except Exception as e:
        changes_body_text = f"\n⚠️ Ошибка запуска сравнения: {e}"
        logger.error(f"❌ Ошибка: {e}")
    
    # 5. Отправка письма
    logger.info(f"\n[4/3] Отправка {len(attachments)} файлов на почту...")
    
    body_lines = [
        f"Отчеты сформированы: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}",
        f"Время выполнения: {datetime.now() - start_time}",
        "",
        "Результаты парсеров:",
        f"  - Парсер планов: {'✅ Успешно' if results['plans'] else '❌ Ошибка'}",
        f"  - Расширенный парсер: {'✅ Успешно' if results['extended'] else '❌ Ошибка'}",
        "",
        f"Всего файлов: {len(attachments)}",
        "",
        "Приложены файлы:"
    ]
    
    for f in attachments:
        body_lines.append(f"  - {os.path.basename(f)}")
    
    # Добавляем блок изменений в конец письма
    if changes_body_text:
        body_lines.append("\n" + "="*40)
        body_lines.append("ОТЧЁТ ОБ ИЗМЕНЕНИЯХ В ПЛАНАХ")
        body_lines.append("="*40)
        body_lines.append(changes_body_text)
    
    body = "\n".join(body_lines)
    subject = f"Отчет по закупкам от {datetime.now().strftime('%d.%m.%Y')}"
    
    try:
        send_email(subject, body, attachments)
        logger.info("✅ Отправлено!")
    except Exception as e:
        logger.error(f"❌ Ошибка отправки: {e}")
    
    # Итоговая статистика
    logger.info(f"\n{'='*70}")
    logger.info("📊 ИТОГОВАЯ СТАТИСТИКА")
    logger.info(f"{'='*70}")
    logger.info(f"   Начало: {start_time.strftime('%d.%m.%Y %H:%M:%S')}")
    logger.info(f"   Конец:  {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}")
    logger.info(f"   Длительность: {datetime.now() - start_time}")
    logger.info(f"   Парсер планов: {'✅' if results['plans'] else '❌'}")
    logger.info(f"   Расширенный парсер: {'✅' if results['extended'] else '❌'}")
    logger.info(f"   Файлов отправлено: {len(attachments)}")
    logger.info(f"   Папка сессии: {session_folder}")
    logger.info(f"{'='*70}")


if __name__ == "__main__":
    main()
