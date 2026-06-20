#!/usr/bin/env python3
"""
Общий запуск всех парсеров и отправка отчётов
Исправленная версия с поддержкой parser_extended.py
"""

import subprocess
import os
import sys
from datetime import datetime
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import glob
import yaml
import logging

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/run_all.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Загрузка конфига
try:
    with open('config.yaml', 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)
    logger.info("✅ Конфиг загружен")
except Exception as e:
    logger.error(f"❌ Ошибка загрузки конфига: {e}")
    sys.exit(1)


def get_latest_file(pattern):
    """Найти последний файл по шаблону"""
    files = glob.glob(pattern)
    if not files:
        return None
    return max(files, key=os.path.getctime)


def send_email(subject, body, attachments):
    """Отправка письма с вложениями"""
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
        else:
            logger.warning(f"⚠️  Файл не найден: {filepath}")
    
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(config['email']['smtp_server'], config['email']['smtp_port'], context=context) as server:
            server.login(config['email']['sender_email'], config['email']['sender_password'])
            server.send_message(msg)
        logger.info("✅ Письмо отправлено")
    except Exception as e:
        logger.error(f"❌ Ошибка отправки письма: {e}")
        raise


def run_parser(name, command, timeout=3600):
    """Запуск парсера с обработкой ошибок"""
    logger.info(f"\n{'='*70}")
    logger.info(f"🚀 Запуск: {name}")
    logger.info(f"{'='*70}")
    
    try:
        result = subprocess.run(
            command,
            cwd='/opt/zakupki-service',
            capture_output=True,
            text=True,
            timeout=timeout
        )
        
        if result.stdout:
            logger.info(result.stdout)
        
        if result.stderr:
            logger.warning(f"⚠️  Ошибки/предупреждения:\n{result.stderr}")
        
        if result.returncode == 0:
            logger.info(f"✅ {name} завершён успешно")
            return True
        else:
            logger.error(f"❌ {name} завершён с кодом {result.returncode}")
            return False
            
    except subprocess.TimeoutExpired:
        logger.error(f"❌ {name} превысил таймаут ({timeout}с)")
        return False
    except Exception as e:
        logger.error(f"❌ Ошибка запуска {name}: {e}")
        return False


def find_all_output_files():
    """Найти файлы результатов через манифест"""
    try:
        sys.path.insert(0, '/opt/zakupki-service')
        from folder_registry import get_today_folders, clear_today
        
        folders = get_today_folders()
        
        if not folders:
            logger.warning("️  Манифест пуст — ищем файлы по старому методу")
            return find_files_by_glob()
        
        logger.info(f"📋 Найдено записей в манифесте: {len(folders)}")
        
        attachments = []
        for parser_name, folder_path in folders.items():
            logger.info(f"   📁 {parser_name}: {folder_path}")
            
            if not os.path.exists(folder_path):
                logger.warning(f"      ️  Папка не существует!")
                continue
            
            for f in glob.glob(f"{folder_path}/*.xlsx"):
                if os.path.exists(f):
                    attachments.append(f)
                    logger.info(f"       Найден: {os.path.basename(f)}")
        
        # Очищаем манифест после сбора файлов
        clear_today()
        
        return attachments
        
    except Exception as e:
        logger.error(f"❌ Ошибка чтения манифеста: {e}")
        logger.warning("⚠️  Используем резервный метод поиска")
        return find_files_by_glob()


def find_files_by_glob():
    """Резервный метод: поиск файлов по glob"""
    date_str = datetime.now().strftime("%d.%m.%Y")
    pattern = f"output/{date_str}_*"
    folders = glob.glob(pattern)
    
    if not folders:
        return []
    
    folders.sort(key=os.path.getctime, reverse=True)
    
    attachments = []
    for folder in folders:
        for f in glob.glob(f"{folder}/*.xlsx"):
            if os.path.exists(f):
                attachments.append(f)
    
    return attachments

def main():
    start_time = datetime.now()
    
    logger.info("=" * 70)
    logger.info("ЗАПУСК ВСЕХ ПАРСЕРОВ")
    logger.info(f"Время запуска: {start_time.strftime('%d.%m.%Y %H:%M:%S')}")
    logger.info("=" * 70)
    
    results = {
        'parser': False,
        'plans': False,
        'extended': False
    }
    
    # 1. Запуск парсера закупок
    logger.info(f"\n[1/4] Запуск парсера закупок...")
    results['parser'] = run_parser(
        "Парсер закупок",
        [sys.executable, 'parser.py', '-y', str(config['parser']['default_year'])],
        timeout=1800  # 30 минут
    )
    
    # 2. Запуск парсера планов-графиков
    logger.info(f"\n[2/4] Запуск парсера планов-графиков...")
    results['plans'] = run_parser(
        "Парсер планов-графиков",
        [sys.executable, 'plans_parser/final_parser_v6.py', '--year', str(config['parser']['default_year'])],
        timeout=3600  # 60 минут (сбор деталей долгий)
    )
    
    # 3. Запуск расширенного парсера закупок
    logger.info(f"\n[3/4] Запуск расширенного парсера закупок с деталями лотов...")
    results['extended'] = run_parser(
        "Расширенный парсер закупок",
        [sys.executable, 'parser_extended.py', '-y', str(config['parser']['default_year'])],
        timeout=5400  # 90 минут (сбор деталей лотов очень долгий)
    )
    
    # 4. Поиск файлов
    logger.info(f"\n[4/4] Поиск файлов для отправки...")
    attachments = find_all_output_files()
    
    if not attachments:
        logger.error(" Файлы не найдены!")
        logger.error("Проверь логи парсеров:")
        logger.error("  - parser.py")
        logger.error("  - plans_parser/final_parser_v6.py")
        logger.error("  - parser_extended.py")
        return
    
    # 5. Отправка письма
    logger.info(f"\n Отправка {len(attachments)} файлов на почту...")
    
    # Формируем тело письма
    body_lines = [
        f"Отчеты сформированы: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}",
        f"Время выполнения: {datetime.now() - start_time}",
        "",
        "Результаты парсеров:",
        f"  - Парсер закупок: {'✅ Успешно' if results['parser'] else '❌ Ошибка'}",
        f"  - Парсер планов: {'✅ Успешно' if results['plans'] else '❌ Ошибка'}",
        f"  - Расширенный парсер: {'✅ Успешно' if results['extended'] else '❌ Ошибка'}",
        "",
        "Приложены файлы:"
    ]
    
    for f in attachments:
        body_lines.append(f"  - {os.path.basename(f)}")
    
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
    logger.info(f"   Парсер закупок: {'✅' if results['parser'] else ''}")
    logger.info(f"   Парсер планов: {'✅' if results['plans'] else '❌'}")
    logger.info(f"   Расширенный парсер: {'✅' if results['extended'] else '❌'}")
    logger.info(f"   Файлов отправлено: {len(attachments)}")
    logger.info(f"{'='*70}")
    
    # Проверяем, все ли парсеры отработали успешно
    if all(results.values()):
        logger.info("\n🎉 Все парсеры отработали успешно!")
    else:
        logger.warning("\n⚠️  Некоторые парсеры завершились с ошибками")
        logger.warning("Проверь логи:")
        if not results['parser']:
            logger.warning("  - parser.py")
        if not results['plans']:
            logger.warning("  - plans_parser/final_parser_v6.py")
        if not results['extended']:
            logger.warning("  - parser_extended.py")


if __name__ == "__main__":
    main()
