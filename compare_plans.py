#!/usr/bin/env python3
"""
Сравнение файлов планов-графиков за два дня.
Простое построчное сравнение: если строки (план + позиция + наименование) 
нет в предыдущем файле — она новая.
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
    """Найти файл Plans за указанную дату (берём последний, если их несколько)"""
    pattern = f"{output_dir}/{date_str}_*/Plans_{date_str}_*.xlsx"
    files = glob.glob(pattern)
    if files:
        return sorted(files)[-1]
    return None


def load_rows(filepath):
    """Загрузить строки из файла. Возвращает список словарей"""
    wb = openpyxl.load_workbook(filepath, read_only=True)
    ws = wb['Позиции планов']
    
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        # Правильные индексы колонок из Excel:
        # [0]=ИНН, [1]=Организация, [2]=№ плана, [3]=Тип ФЗ, 
        # [4]=Тип позиции, [5]=№ позиции, [6]=Наименование, [7]=Цена
        if row[2] and row[5]:  # Должны быть номер плана и номер позиции
            plan = str(row[2]).strip()
            position = str(row[5]).strip()
            name = str(row[6]).strip() if row[6] else ''
            price = row[7] if row[7] else ''
            inn = str(row[0]).strip() if row[0] else ''
            org = row[1] if row[1] else ''
            rows.append({
                'plan': plan,
                'position': position,
                'name': name,
                'price': price,
                'inn': inn,
                'org': org
            })
    
    wb.close()
    return rows


def compare_plans(current_file, previous_file):
    """Сравнить два файла планов построчно"""
    if not previous_file:
        return [], "Файл за предыдущий день не найден"
    
    logger.info(f"📊 Сравнение файлов планов...")
    logger.info(f"   Текущий: {current_file}")
    logger.info(f"   Предыдущий: {previous_file}")
    
    prev_rows = load_rows(previous_file)
    curr_rows = load_rows(current_file)
    
    logger.info(f"   Строк в предыдущем файле: {len(prev_rows)}")
    logger.info(f"   Строк в текущем файле: {len(curr_rows)}")
    
    # Создаём множество ключей из предыдущего файла
    # Ключ = (план, позиция, наименование)
    prev_keys = set()
    for row in prev_rows:
        key = (row['plan'], row['position'], row['name'])
        prev_keys.add(key)
    
    # Ищем новые строки в текущем файле
    new_rows = []
    seen_keys = set()  # Чтобы не добавлять дубликаты в результат
    for row in curr_rows:
        key = (row['plan'], row['position'], row['name'])
        if key not in prev_keys and key not in seen_keys:
            new_rows.append(row)
            seen_keys.add(key)
    
    logger.info(f"   Новых строк: {len(new_rows)}")
    
    return new_rows, None


def save_changes_to_txt(new_rows, filename, current_file, previous_file, error_message=None):
    """Сохранить изменения в txt файл"""
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
            prev_rows = load_rows(previous_file)
            f.write(f"  - Строк в предыдущем файле: {len(prev_rows)}\n")
        
        if current_file:
            curr_rows = load_rows(current_file)
            f.write(f"  - Строк в текущем файле:    {len(curr_rows)}\n")
        
        f.write(f"  - Новых строк обнаружено:   {len(new_rows)}\n")
        
        if error_message:
            f.write(f"\n⚠️  Ошибка: {error_message}\n")
        
        f.write("\n")
        
        # ===== БЛОК 2: ДАННЫЕ ДЛЯ КОЛЛЕГ =====
        f.write("=" * 80 + "\n")
        f.write("ИЗМЕНЕНИЯ В ПЛАНАХ-ГРАФИКАХ\n")
        f.write("=" * 80 + "\n\n")
        
        if error_message:
            f.write(f"⚠️  {error_message}\n")
            f.write("Сравнение не проводилось.\n")
        elif not new_rows:
            f.write("✅ Новых позиций не обнаружено.\n")
            f.write("\nВсе позиции планов-графиков остались без изменений.\n")
        else:
            f.write(f"Найдено новых позиций: {len(new_rows)}\n\n")
            
            # Группируем по организациям
            by_org = {}
            for row in new_rows:
                org = row['org']
                if org not in by_org:
                    by_org[org] = []
                by_org[org].append(row)
            
            for org, rows in by_org.items():
                f.write(f"📋 {org} ({len(rows)} новых позиций):\n")
                f.write("-" * 80 + "\n")
                
                for idx, row in enumerate(rows, 1):
                    f.write(f"{idx}. План №{row['plan']}, позиция №{row['position']}\n")
                    f.write(f"   Наименование: {row['name']}\n")
                    
                    if row['price'] and row['price'] != '':
                        try:
                            price = float(row['price'])
                            f.write(f"   Цена: {price:,.2f} руб.\n")
                        except:
                            f.write(f"   Цена: {row['price']}\n")
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
    
    current_file = find_plans_file(args.output_dir, today_str)
    previous_file = find_plans_file(args.output_dir, previous_str)
    
    if not current_file:
        logger.error(f"❌ Файл планов за {today_str} не найден!")
        sys.exit(1)
    
    new_rows, error_message = compare_plans(current_file, previous_file)
    
    if args.output_file:
        output_file = args.output_file
    else:
        output_file = f"{args.output_dir}/{today_str}_01/changes_{today_str}.txt"
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    save_changes_to_txt(new_rows, output_file, current_file, previous_file, error_message)
    
    print(f"\nOUTPUT_FILE={output_file}")
    print(f"NEW_ROWS_COUNT={len(new_rows)}")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
