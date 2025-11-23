#!/usr/bin/env python3
"""
Упрощенная проверка дат в таблице
"""
import os
import json
import re
from datetime import datetime, timedelta
from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials

load_dotenv()

def connect_to_sheets():
    """Подключение к Google Sheets"""
    spreadsheet_url = os.getenv('GOOGLE_SHEET_URL')
    service_account_json = os.getenv('SERVICE_ACCOUNT_JSON')

    if not spreadsheet_url or not service_account_json:
        raise ValueError("GOOGLE_SHEET_URL or SERVICE_ACCOUNT_JSON not found")

    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]

    creds_dict = json.loads(service_account_json)
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open_by_url(spreadsheet_url)

    return sheet

def main():
    print("=== Проверка дат в Google Sheets ===\n")

    try:
        sheet = connect_to_sheets()
        worksheet = sheet.worksheet("ВЕЛ БЕГ")
        print("✓ Подключено к таблице\n")

        # Читаем строку 20
        row_20 = worksheet.row_values(20)

        print("Содержимое строки 20 (Лонг RUN вс):\n")
        print(f"{'Столбец':<10} {'Индекс':<10} {'Содержимое':<50}")
        print("-" * 70)

        for idx, cell in enumerate(row_20):
            col_letter = chr(65 + idx)
            if idx >= 2 and cell:  # Только с столбца C и дальше
                print(f"{col_letter:<10} {idx:<10} {cell:<50}")

        # Парсим даты
        print("\n" + "="*70)
        print("Парсинг дат из строки 20:")
        print("="*70 + "\n")

        sunday_columns = {}
        for idx, cell in enumerate(row_20):
            if idx < 2 or not cell or not isinstance(cell, str):
                continue

            col_letter = chr(65 + idx)

            # Ищем дату в формате dd.mm.yy
            date_match = re.search(r'\b(\d{2})\.(\d{2})\.(\d{2})\b', cell)
            if date_match:
                day, month, year = date_match.groups()
                year_full = '20' + year

                try:
                    sunday_date = datetime.strptime(f"{day}.{month}.{year_full}", "%d.%m.%Y").date()
                    saturday_date = sunday_date - timedelta(days=1)
                    friday_date = saturday_date + timedelta(days=6)

                    sunday_columns[col_letter] = sunday_date

                    print(f"Столбец {col_letter}:")
                    print(f"  Воскресенье: {sunday_date.strftime('%d.%m.%Y')}")
                    print(f"  Неделя (сб-пт): {saturday_date.strftime('%d.%m.%Y')} - {friday_date.strftime('%d.%m.%Y')}")
                    print()
                except Exception as e:
                    print(f"  ❌ Ошибка парсинга: {e}\n")

        # Проверка текущей даты
        print("="*70)
        print("Проверка попадания тренировок в недели:")
        print("="*70 + "\n")

        today = datetime.now().date()
        print(f"Сегодня: {today.strftime('%d.%m.%Y')} ({today.year} год)\n")

        # Проверяем несколько дат
        test_dates = [
            today,
            today - timedelta(days=7),  # Неделя назад
            today - timedelta(days=14), # 2 недели назад
        ]

        for test_date in test_dates:
            print(f"Дата: {test_date.strftime('%d.%m.%Y')}")
            found = False

            for col_letter, sunday_date in sunday_columns.items():
                saturday_date = sunday_date - timedelta(days=1)
                friday_date = saturday_date + timedelta(days=6)

                if saturday_date <= test_date <= friday_date:
                    print(f"  ✓ Попадает в столбец {col_letter} ({saturday_date.strftime('%d.%m')} - {friday_date.strftime('%d.%m')})")
                    found = True
                    break

            if not found:
                print(f"  ❌ НЕ попадает ни в один столбец!")

            print()

    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
