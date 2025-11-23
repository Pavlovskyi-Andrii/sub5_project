#!/usr/bin/env python3
"""Тест синхронизации с Google Sheets"""
import os
from datetime import datetime, timedelta
from dotenv import load_dotenv
from main import (
    connect_to_garmin,
    connect_to_google_sheets,
    parse_week_dates_from_block_rows,
    find_column_for_date
)

load_dotenv()

def test_google_sync():
    """Проверяем синхронизацию с Google Sheets"""
    print("="*60)
    print("ДИАГНОСТИКА СИНХРОНИЗАЦИИ С GOOGLE SHEETS")
    print("="*60)

    # Подключаемся к Google Sheets
    print("\n1. Подключаемся к Google Sheets...")
    sheet = connect_to_google_sheets()
    worksheet = sheet.worksheet("ВЕЛ БЕГ")
    print("✓ Подключено к листу 'ВЕЛ БЕГ'")

    # Парсим даты из строки 20
    print("\n2. Парсим даты недель из строки 20...")
    week_columns = parse_week_dates_from_block_rows(worksheet)
    print(f"✓ Найдено {len(week_columns)} недель:")
    for col, date in sorted(week_columns.items()):
        saturday = date - timedelta(days=1)
        friday = saturday + timedelta(days=6)
        print(f"   {col}: {saturday.strftime('%d.%m.%Y')} (Сб) - {friday.strftime('%d.%m.%Y')} (Пт), воскресенье: {date.strftime('%d.%m.%Y')}")

    # Проверяем, в какие столбцы попадают даты 19-23 ноября
    print("\n3. Проверяем столбцы для дат 19-23 ноября 2024:")
    test_dates = []
    for day in range(19, 24):
        test_date = datetime(2024, 11, day).date()
        test_dates.append(test_date)
        column = find_column_for_date(test_date, week_columns)
        status = "✓" if column else "✗"
        print(f"   {status} {test_date.strftime('%d.%m.%Y')}: столбец {column if column else 'НЕ НАЙДЕН'}")

    # Подключаемся к Garmin и проверяем активности
    print("\n4. Подключаемся к Garmin и проверяем активности...")
    garmin = connect_to_garmin()

    print("\n5. Получаем активности за последние 7 дней...")
    activities = garmin.get_activities(0, 50)

    print(f"\n6. Активности за период 19-23 ноября:")
    activities_in_range = []
    for activity in activities:
        start_time = activity.get('startTimeLocal', '')
        if start_time:
            activity_date = datetime.strptime(start_time[:10], '%Y-%m-%d').date()
            if activity_date >= datetime(2024, 11, 19).date() and activity_date <= datetime(2024, 11, 23).date():
                activities_in_range.append(activity)
                activity_name = activity.get('activityName', 'No name')
                activity_type = activity.get('activityType', {}).get('typeKey', 'unknown')
                column = find_column_for_date(activity_date, week_columns)
                status = "✓" if column else "✗"
                print(f"   {status} {activity_date.strftime('%d.%m.%Y')}: {activity_name[:40]} ({activity_type}) → {column if column else 'НЕТ СТОЛБЦА'}")

    print(f"\n{'='*60}")
    print(f"ИТОГО:")
    print(f"  - Найдено активностей за 19-23 ноября: {len(activities_in_range)}")
    print(f"  - Активностей без столбца: {sum(1 for a in activities_in_range if not find_column_for_date(datetime.strptime(a.get('startTimeLocal', '')[:10], '%Y-%m-%d').date(), week_columns))}")
    print(f"{'='*60}")

    # Проверяем, есть ли столбцы для этих дат
    missing_columns = []
    for test_date in test_dates:
        if not find_column_for_date(test_date, week_columns):
            missing_columns.append(test_date)

    if missing_columns:
        print(f"\n⚠️  ПРОБЛЕМА: В Google Sheets нет столбцов для следующих дат:")
        for date in missing_columns:
            print(f"   - {date.strftime('%d.%m.%Y')}")
        print(f"\n💡 РЕШЕНИЕ: Добавьте в строку 20 новые столбцы с датами воскресений для недель, содержащих эти даты")
        print(f"   Например, для 19-23 ноября нужно добавить столбец с воскресеньем 24.11.24")
    else:
        print(f"\n✓ Все даты 19-23 ноября имеют столбцы в таблице")

if __name__ == '__main__':
    test_google_sync()
