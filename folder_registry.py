#!/usr/bin/env python3
"""
Реестр папок результатов.
Парсеры регистрируют свои папки здесь, run_all.py читает отсюда.
"""

import json
import os
from datetime import datetime
from pathlib import Path

MANIFEST_FILE = Path("output/manifest.json")


def register_folder(parser_name: str, folder_path: str):
    """Зарегистрировать папку результата от парсера"""
    manifest = {}
    
    # Читаем существующий манифест
    if MANIFEST_FILE.exists():
        try:
            with open(MANIFEST_FILE, 'r', encoding='utf-8') as f:
                manifest = json.load(f)
        except Exception:
            manifest = {}
    
    # today_key = "20.06.2026"
    today_key = datetime.now().strftime("%d.%m.%Y")
    
    if today_key not in manifest:
        manifest[today_key] = {}
    
    manifest[today_key][parser_name] = folder_path
    
    # Сохраняем
    MANIFEST_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(MANIFEST_FILE, 'w', encoding='utf-8') as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    
    print(f"📝 Зарегистрирована папка: {parser_name} -> {folder_path}")


def get_today_folders() -> dict:
    """Получить все папки за сегодня"""
    if not MANIFEST_FILE.exists():
        return {}
    
    try:
        with open(MANIFEST_FILE, 'r', encoding='utf-8') as f:
            manifest = json.load(f)
    except Exception:
        return {}
    
    today_key = datetime.now().strftime("%d.%m.%Y")
    return manifest.get(today_key, {})


def clear_today():
    """Очистить записи за сегодня (после отправки письма)"""
    if not MANIFEST_FILE.exists():
        return
    
    try:
        with open(MANIFEST_FILE, 'r', encoding='utf-8') as f:
            manifest = json.load(f)
    except Exception:
        return
    
    today_key = datetime.now().strftime("%d.%m.%Y")
    if today_key in manifest:
        del manifest[today_key]
        with open(MANIFEST_FILE, 'w', encoding='utf-8') as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)
        print(f"️  Очищены записи за {today_key}")


def print_manifest():
    """Показать содержимое манифеста"""
    if not MANIFEST_FILE.exists():
        print("Манифест не существует")
        return
    
    with open(MANIFEST_FILE, 'r', encoding='utf-8') as f:
        print(json.dumps(json.load(f), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Использование:")
        print("  python3 folder_registry.py register <parser_name> <folder_path>")
        print("  python3 folder_registry.py get")
        print("  python3 folder_registry.py clear")
        print("  python3 folder_registry.py show")
        sys.exit(1)
    
    command = sys.argv[1]
    
    if command == "register" and len(sys.argv) >= 4:
        register_folder(sys.argv[2], sys.argv[3])
    elif command == "get":
        folders = get_today_folders()
        for name, path in folders.items():
            print(f"{name}: {path}")
    elif command == "clear":
        clear_today()
    elif command == "show":
        print_manifest()
    else:
        print(f"Неизвестная команда: {command}")
