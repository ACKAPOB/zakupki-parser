#!/usr/bin/env python3
"""
Конфигурация парсера планов-графиков
Читает настройки из общего config.yaml в корне проекта
"""

import os
import sys
from pathlib import Path
import yaml

# ===== ПУТИ =====
# Базовая директория проекта (корень проекта, не plans_parser)
BASE_DIR = Path(__file__).parent.parent

# Загружаем общий конфиг из корня проекта
try:
    config_path = BASE_DIR / 'config.yaml'
    with open(config_path, 'r', encoding='utf-8') as f:
        general_config = yaml.safe_load(f)
except Exception as e:
    print(f"⚠️  Не удалось загрузить config.yaml: {e}")
    general_config = {}

# Директория для выходных файлов (из общего конфига)
OUTPUT_DIR = BASE_DIR / general_config.get('excel', {}).get('output_dir', 'output')

# Файл со списком ИНН (в корне проекта)
INN_LIST_FILE = BASE_DIR / "inn_list.txt"

# Файл лога (будет создаваться в папке с результатом)
LOG_FILE_NAME = "parser.log"

# ===== НАСТРОЙКИ ПОИСКА =====
# Год плана по умолчанию (из конфига)
plans_config = general_config.get('plans_parser', {})
DEFAULT_YEAR = plans_config.get('default_year', 2026)

# Тип ФЗ: "44", "223", "all"
DEFAULT_FZ = "all"

# Максимум страниц поиска планов
MAX_SEARCH_PAGES = 5

# Размер страницы поиска
SEARCH_PAGE_SIZE = 50

# ===== НАСТРОЙКИ ПАРСИНГА =====
# Максимум страниц пагинации внутри плана (из конфига)
MAX_PAGINATION_PAGES = plans_config.get('max_pagination_pages', 50)

# Собирать детальные данные по позициям (из конфига)
COLLECT_DETAILS = plans_config.get('collect_details', True)

# Запускать браузер с интерфейсом (из конфига)
HEADLESS = plans_config.get('headless', True)

# ===== ЗАДЕРЖКИ (в секундах) - из конфига =====
DELAY_BETWEEN_INN = plans_config.get('delay_between_inn', 5)
DELAY_BETWEEN_PLANS = plans_config.get('delay_between_plans', 3)
DELAY_BETWEEN_PAGES = plans_config.get('delay_between_pages', 2)
DELAY_AFTER_POSITION = plans_config.get('delay_after_position', 1.5)
DELAY_BETWEEN_SEARCH_PAGES = 1
DELAY_RETRY = 5
MAX_RETRIES = plans_config.get('max_retries', 3)
PAGE_LOAD_TIMEOUT = plans_config.get('page_load_timeout', 45000)

# ===== РАЗМЕР БРАУЗЕРА =====
BROWSER_WIDTH = 1920
BROWSER_HEIGHT = 1080

# ===== HTTP ЗАГОЛОВКИ =====
HEADERS = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:97.0) Gecko/20100101 Firefox/97.0",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3",
}

# ===== URL =====
BASE_URL = "https://zakupki.gov.ru"
SEARCH_URL = "https://zakupki.gov.ru/epz/orderplan/search/results.html"

# ===== МЕСЯЦЫ =====
MONTHS = {
    'Январь': '01', 'Февраль': '02', 'Март': '03', 'Апрель': '04',
    'Май': '05', 'Июнь': '06', 'Июль': '07', 'Август': '08',
    'Сентябрь': '09', 'Октябрь': '10', 'Ноябрь': '11', 'Декабрь': '12'
}
