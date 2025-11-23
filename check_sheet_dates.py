#!/usr/bin/env python3
"""
Проверка дат в таблице Google Sheets
"""
import os
from datetime import datetime
from dotenv import load_dotenv
from main import connect_to_google_sheets, parse_week_dates_from_block_rows

load_dotenv()

def check_sheet_dates():
    """Проверяет даты в строке 20 таблицы"""
    print("=== Проверка дат в Google Sheets ===\n")

    try:
        # Подключаемся к таблице
        sheet = connect_to_google_sheets()
        worksheet = sheet.worksheet("ВЕЛ БЕГ")

        # Читаем строку 20
        print("Чтение строки 20...")
        row_20 = worksheet.row_values(20)

        print(f"\nВсего ячеек в строке 20: {len(row_20)}\n")

        # Показываем содержимое столбцов
        for idx, cell in enumerate(row_20):
            if idx < 2:  # Пропускаем A и B
                continue

            if not cell:
                continue

            col_letter = chr(65 + idx)
            print(f"Столбец {col_letter} (индекс {idx}): '{cell}'")

        print("\n" + "="*60)
        print("Парсинг дат функцией parse_week_dates_from_block_rows():")
        print("="*60 + "\n")

        # Парсим даты
        week_columns = parse_week_dates_from_block_rows(worksheet)

        if week_columns:
            print(f"Найдено {len(week_columns)} недель:\n")
            for col, date in sorted(week_columns.items()):
                # Вычисляем диапазон недели
                from datetime import timedelta
                saturday_date = date - timedelta(days=1)
                friday_date = saturday_date + timedelta(days=6)

                print(f"Столбец {col}:")
                print(f"  Воскресенье (из строки 20): {date.strftime('%d.%m.%Y')}")
                print(f"  Неделя (сб-пт): {saturday_date.strftime('%d.%m.%Y')} - {friday_date.strftime('%d.%m.%Y')}")
                print()
        else:
            print("❌ Не найдено ни одной недели!")

        # Проверяем текущую дату
        print("="*60)
        print("Проверка текущей даты:")
        print("="*60)
        today = datetime.now().date()
        print(f"Сегодня: {today.strftime('%d.%m.%Y')}\n")

        # Ищем столбец для сегодняшней даты
        from main import find_column_for_date
        column = find_column_for_date(today, week_columns)

        if column:
            print(f"✓ Сегодняшняя дата попадает в столбец: {column}")
            sunday = week_columns[column]
            saturday = sunday - timedelta(days=1)
            friday = saturday + timedelta(days=6)
            print(f"  Неделя: {saturday.strftime('%d.%m.%Y')} - {friday.strftime('%d.%m.%Y')}")
        else:
            print("❌ Сегодняшняя дата НЕ попадает ни в один столбец!")
            print("\nВозможные причины:")
            print("1. В таблице нет недели, покрывающей текущую дату")
            print("2. Даты в таблице указаны неправильно (например, 2025 вместо 2024)")

    except Exception as e:
        print(f"❌ Ошибка: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_sheet_dates()
