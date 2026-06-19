# Парсер закупок ЕИС

Микросервис для автоматического сбора информации о закупках с портала zakupki.gov.ru по списку ИНН организаций.

## Функционал

- Сбор закупок по списку ИНН за текущий год
- Экспорт в Excel с полными данными (12 колонок)
- Отправка отчетов на электронную почту
- Гибкая настройка через конфигурационный файл
- Работа по расписанию через cron

## Структура проекта

- parser.py - основной парсер
- send_report.py - парсер + отправка на почту
- inn_list.txt - список организаций (ИНН|Название)
- config.yaml - конфигурация (не в репозитории)
- output/ - папка с отчетами
- logs/ - логи работы

## Быстрый старт

### Установка

git clone https://github.com/ACKAPOB/zakupki-parser.git
cd zakupki-parser
pip install requests beautifulsoup4 openpyxl pyyaml

### Настройка

Создайте config.yaml на основе примера:

parser:
  page_size: 50
  max_pages: 50
  delay: 0.5

email:
  smtp_server: "smtp.yandex.ru"
  smtp_port: 465
  sender_email: "your_email@yandex.ru"
  sender_password: "your_app_password"
  recipient_email: "recipient@mail.ru"

Добавьте ИНН в inn_list.txt:

2312054894|АО "АТЭК"
2311322773|ООО "КРАСНОДАРТЕПЛОЭНЕРГО"

## Использование

### Ручной запуск

python3 parser.py -y 2026
python3 parser.py -i 2312054894 -y 2026
python3 parser.py -y 2026 -m 6
python3 send_report.py

### Автоматический запуск (cron)

crontab -e
# Добавить строку для ежедневного запуска в 9:00
0 9 * * * cd /opt/zakupki-service && /usr/bin/python3 send_report.py >> logs/cron.log 2>&1

## Формат вывода

Excel-отчет содержит колонки:
- ID, Номер закупки, Название, Организация
- Цена (руб.), Стадия, ФЗ, Тип
- Дата размещения, Дата обновления, Дата окончания
- Ссылка

## Конфигурация

- page_size - записей на страницу (50)
- max_pages - максимум страниц (50)
- delay - задержка между запросами в секундах (0.5)
- smtp_server - SMTP сервер
- smtp_port - порт SMTP

## Безопасность

- config.yaml с паролями не хранится в репозитории
- Добавлен в .gitignore
- Для отправки почты используются пароли приложений

## Автор

Igor
Email: ackapob@yandex.ru
GitHub: https://github.com/ACKAPOB/zakupki-parser
