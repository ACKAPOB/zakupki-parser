#!/usr/bin/env python3
"""
Сравнение файлов планов-графиков за два дня
Находит новые позиции и сохраняет их в txt файл
"""

import os
import sys
import glob
import argparse
from datetime import datetime, timedelta
import openpyxl
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)


def find_plans_file(output_dir, date_str):
    """Найти файл Plans за указанную дату"""
    pattern = f"{output_dir}/{date_str}_*/Plans_{date_str}_*.xlsx"
    files = glob.glob(pattern)
    if files:
        return sorted(files)[-1]
    return None


def compare_plans(current_file, previous_file):
    """Сравнить два файла планов и найти новые позиции"""
    if not previous_file:
        return [], "Файл за предыдущий день не найден"
    
    logger.info(f"📊 Сравнение файлов планов...")
    logger.info(f"   Текущий: {current_file}")
    logger.info(f"   Предыдущий: {previous_file}")
    
    # Загружаем файлы
    wb_current = openpyxl.load_workbook(current_file)
    ws_current = wb_current['Позиции планов']
    
    wb_previous = openpyxl.load_workbook(previous_file)
    ws_previous = wb_previous['Позиции планов']
    
    # Собираем позиции из предыдущего файла
    previous_positions = set()
    for row in ws_previous.iter_rows(min_row=2, values_only=True):
        if row[2] and row[4]:
            key = (str(row[2]).strip(), str(row[4]).strip())
            previous_positions.add(key)
    
    logger.info(f"   Позиций в предыдущем файле: {len(previous_positions)}")
    
    # Находим новые позиции
    new_positions = []
    for row in ws_current.iter_rows(min_row=2, values_only=True):
        if row[0] and row[2] and row[4]:
            key = (str(row[2]).strip(), str(row[4]).strip())
            if key not in previous_positions:
                new_positions.append({
                    'inn': str(row[0]).strip(),
                    'org': row[1] if row[1] else '',
                    'plan_number': str(row[2]).strip(),
                    'position_number': str(row[4]).strip(),
                    'name': row[5] if row[5] else '',
                    'price': row[6] if row[6] else 0
                })
    
    logger.info(f"   Новых позиций: {len(new_positions)}")
    
    return new_positions, None


def save_changes_to_txt(new_positions, filename, current_file, previous_file, error_message=None):
    """Сохранить изменения в txt файл с двумя блоками"""
    with open(filename, 'w', encoding='utf-8') as f:
        # ===== БЛОК 1: ТЕХНИЧЕСКАЯ ИНФОРМАЦИЯ =====
        f.write("=" * 80 + "\n")
        f.write("ТЕХНИЧЕСКАЯ ИНФОРМАЦИЯ\n")
        f.write("=" * 80 + "\n\n")
        
        f.write(f"Дата формирования отчёта: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
        f.write(f"Сравниваемые файлы:\n")
        f.write(f"  - Текущий:    {os.path.basename(current_file) if current_file else 'не найден'}\n")
        f.write(f"  - Предыдущий: {os.path.basename(previous_file) if previous_file else 'не найден'}\n")
        
        if previous_file:
            wb_prev = openpyxl.load_workbook(previous_file)
            ws_prev = wb_prev['Позиции планов']
            prev_count = sum(1 for _ in ws_prev.iter_rows(min_row=2))
            f.write(f"  - Позиций в предыдущем файле: {prev_count}\n")
        
        if current_file:
            wb_curr = openpyxl.load_workbook(current_file)
            ws_curr = wb_curr['Позиции планов']
            curr_count = sum(1 for _ in ws_curr.iter_rows(min_row=2))
            f.write(f"  - Позиций в текущем файле:    {curr_count}\n")
        
        f.write(f"  - Новых позиций обнаружено:     {len(new_positions)}\n")
        
        if error_message:
            f.write(f"\n️  Ошибка: {error_message}\n")
        
        f.write("\n")
        
        # ===== БЛОК 2: ДАННЫЕ ДЛЯ КОЛЛЕГ =====
        f.write("=" * 80 + "\n")
        f.write("ИЗМЕНЕНИЯ В ПЛАНАХ-ГРАФИКАХ\n")
        f.write("=" * 80 + "\n\n")
        
        if error_message:
            f.write(f"⚠️  {error_message}\n")
            f.write("Сравнение не проводилось.\n")
        elif not new_positions:
            f.write("✅ Новых позиций не обнаружено.\n")
            f.write("\nВсе позиции планов-графиков остались без изменений.\n")
        else:
            f.write(f"Найдено новых позиций: {len(new_positions)}\n\n")
            
            # Группируем по организациям
            by_org = {}
            for pos in new_positions:
                org = pos['org']
                if org not in by_org:
                    by_org[org] = []
                by_org[org].append(pos)
            
            for org, positions in by_org.items():
                f.write(f"📋 {org} ({len(positions)} новых позиций):\n")
                f.write("-" * 80 + "\n")
                
                for idx, pos in enumerate(positions, 1):
                    f.write(f"{idx}. План №{pos['plan_number']}, позиция №{pos['position_number']}\n")
                    f.write(f"   Наименование: {pos['name']}\n")
                    
                    # Форматируем цену
                    if pos['price'] and pos['price'] != 0:
                        try:
                            price = float(pos['price'])
                            f.write(f"   Цена: {price:,.2f} руб.\n")
                        except:
                            f.write(f"   Цена: {pos['price']}\n")
                    else:
                        f.write(f"   Цена: не указана\n")
                    
                    f.write("\n")
                
                f.write("\n")
            
            f.write("=" * 80 + "\n")
            f.write("КОНЕЦ ОТЧЁТА\n")
            f.write("=" * 80 + "\n")
    
    logger.info(f"💾 Изменения сохранены в: {filename}")


def main():
    parser = argparse.ArgumentParser(description="Сравнение файлов планов-графиков")
    parser.add_argument("--output-dir", default="output", help="Директория с файлами")
    parser.add_argument("--date", help="Дата для сравнения (ДД.ММ.ГГГГ), по умолчанию сегодня")
    parser.add_argument("--days-back", type=int, default=1, help="Сколько дней назад брать для сравнения")
    parser.add_argument("--output-file", help="Путь к выходному txt файлу")
    
    args = parser.parse_args()
    
    # Определяем даты
    if args.date:
        today = datetime.strptime(args.date, "%d.%m.%Y")
    else:
        today = datetime.now()
    
    today_str = today.strftime("%d.%m.%Y")
    previous = today - timedelta(days=args.days_back)
    previous_str = previous.strftime("%d.%m.%Y")
    
    logger.info("=" * 70)
    logger.info("СРАВНЕНИЕ ПЛАНОВ-ГРАФИКОВ")
    logger.info(f"Текущая дата: {today_str}")
    logger.info(f"Предыдущая дата: {previous_str}")
    logger.info("=" * 70)
    
    # Ищем файлы
    current_file = find_plans_file(args.output_dir, today_str)
    previous_file = find_plans_file(args.output_dir, previous_str)
    
    if not current_file:
        logger.error(f"❌ Файл планов за {today_str} не найден!")
        sys.exit(1)
    
    # Сравниваем
    new_positions, error_message = compare_plans(current_file, previous_file)
    
    # Определяем путь к выходному файлу
    if args.output_file:
        output_file = args.output_file
    else:
        # Создаём в папке сессии
        output_file = f"{args.output_dir}/{today_str}_01/changes_{today_str}.txt"
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    # Сохраняем
    save_changes_to_txt(new_positions, output_file, current_file, previous_file, error_message)
    
    # Выводим путь к файлу (для использования в run_all.py)
    print(f"\nOUTPUT_FILE={output_file}")
    print(f"NEW_POSITIONS_COUNT={len(new_positions)}")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
