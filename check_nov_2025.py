#!/usr/bin/env python3
"""Проверка для ноября 2025"""
from datetime import datetime, timedelta
from dotenv import load_dotenv
from main import (
    connect_to_garmin,
    connect_to_google_sheets,
    parse_week_dates_from_block_rows,
    find_column_for_date
)

load_dotenv()

def check_nov_2025():
    """Проверяем ноябрь 2025"""
    print("="*60)
    print("ПРОВЕРКА НОЯБРЯ 2025")
    print("="*60)

    # Подключаемся к Google Sheets
    sheet = connect_to_google_sheets()
    worksheet = sheet.worksheet("ВЕЛ БЕГ")

    # Парсим даты
    week_columns = parse_week_dates_from_block_rows(worksheet)

    # Проверяем столбцы для ноября 2025
    print("\nСтолбцы для ноября 2025:")
    for col, date in sorted(week_columns.items()):
        if date.year == 2025 and date.month == 11:
            saturday = date - timedelta(days=1)
            friday = saturday + timedelta(days=6)
            print(f"  {col}: {saturday.strftime('%d.%m.%Y')} (Сб) - {friday.strftime('%d.%m.%Y')} (Пт)")

    # Проверяем активности
    print("\n" + "="*60)
    garmin = connect_to_garmin()
    activities = garmin.get_activities(0, 50)

    print("\nАктивности за ноябрь 2025:")
    nov_activities = []
    for activity in activities:
        start_time = activity.get('startTimeLocal', '')
        if start_time:
            activity_date = datetime.strptime(start_time[:10], '%Y-%m-%d').date()
            if activity_date.year == 2025 and activity_date.month == 11:
                nov_activities.append(activity)
                activity_name = activity.get('activityName', 'No name')
                column = find_column_for_date(activity_date, week_columns)
                status = "✓" if column else "✗"
                print(f"  {status} {activity_date.strftime('%d.%m.%Y')}: {activity_name[:40]} → {column if column else 'НЕТ СТОЛБЦА'}")

    print(f"\nВсего: {len(nov_activities)} активностей")
    without_column = sum(1 for a in nov_activities if not find_column_for_date(
        datetime.strptime(a.get('startTimeLocal', '')[:10], '%Y-%m-%d').date(),
        week_columns
    ))
    print(f"Без столбца: {without_column}")

if __name__ == '__main__':
    check_nov_2025()
