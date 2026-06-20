#!/usr/bin/env python3
"""
Конфигурация парсера планов-графиков
Все настройки, которые могут меняться при запуске на разных серверах
"""

import os
from pathlib import Path

# ===== ПУТИ =====
# Базовая директория проекта (где лежит этот конфиг)
BASE_DIR = Path(__file__).parent

# Директория для выходных файлов
OUTPUT_DIR = BASE_DIR / "output"

# Файл со списком ИНН
#INN_LIST_FILE = BASE_DIR / "inn_list.txt"
INN_LIST_FILE = BASE_DIR.parent / "inn_list.txt"

# Файл лога (будет создаваться в папке с результатом)
LOG_FILE_NAME = "parser.log"

# ===== НАСТРОЙКИ ПОИСКА =====
# Год плана по умолчанию
DEFAULT_YEAR = 2026

# Тип ФЗ: "44", "223", "all"
DEFAULT_FZ = "all"

# Максимум страниц поиска планов
MAX_SEARCH_PAGES = 5

# Размер страницы поиска
SEARCH_PAGE_SIZE = 50

# ===== НАСТРОЙКИ ПАРСИНГА =====
# Максимум страниц пагинации внутри плана
MAX_PAGINATION_PAGES = 50

# Собирать детальные данные по позициям
COLLECT_DETAILS = True

# Запускать браузер с интерфейсом (False = headless)
HEADLESS = True

# ===== ЗАДЕРЖКИ (в секундах) =====
# Между организациями
DELAY_BETWEEN_INN = 5

# Между планами одной организации
DELAY_BETWEEN_PLANS = 3

# Между страницами пагинации
DELAY_BETWEEN_PAGES = 2

# После сбора деталей позиции
DELAY_AFTER_POSITION = 1.5

# Между страницами поиска планов
DELAY_BETWEEN_SEARCH_PAGES = 1

# Перед повторной попыткой загрузки
DELAY_RETRY = 5

# ===== ПОВТОРНЫЕ ПОПЫТКИ =====
# Максимум попыток загрузки плана
MAX_RETRIES = 3

# Таймаут загрузки страницы (мс)
PAGE_LOAD_TIMEOUT = 45000

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
