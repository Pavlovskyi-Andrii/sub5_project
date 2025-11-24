#!/usr/bin/env python3
"""
Тестовый скрипт для проверки полной синхронизации
"""
import os
import sys
from datetime import datetime
from dotenv import load_dotenv

# Загружаем переменные окружения
load_dotenv()

# Импортируем функции из main.py
from main import (
    connect_to_garmin,
    connect_to_google_sheets,
    parse_week_dates_from_block_rows
)

def test_full_sync():
    """Тест полной синхронизации"""
    print("="*60)
    print("ТЕСТ ПОЛНОЙ СИНХРОНИЗАЦИИ")
    print("="*60)

    # 1. Подключаемся к Garmin
    print("\n1. Подключение к Garmin...")
    try:
        garmin = connect_to_garmin()
        print("   ✓ Успешно подключено к Garmin")
    except Exception as e:
        print(f"   ✗ Ошибка: {e}")
        return

    # 2. Получаем активности
    start_date = datetime(2024, 9, 13).date()
    today = datetime.now().date()
    days_diff = (today - start_date).days

    print(f"\n2. Получение активностей с {start_date.strftime('%d.%m.%Y')} ({days_diff} дней)...")
    try:
        max_activities = max(800, days_diff * 3)
        print(f"   Запрашиваем до {max_activities} активностей...")
        activities = garmin.get_activities(0, max_activities)
        print(f"   ✓ Получено {len(activities)} активностей из Garmin")
    except Exception as e:
        print(f"   ✗ Ошибка: {e}")
        return

    # 3. Фильтруем активности
    print(f"\n3. Фильтрация активностей с {start_date.strftime('%d.%m.%Y')}...")
    filtered_activities = []
    for activity in activities:
        if not isinstance(activity, dict):
            continue

        start_time = activity.get('startTimeLocal', '')
        if start_time:
            activity_date = datetime.strptime(start_time[:10], '%Y-%m-%d').date()
            if activity_date >= start_date:
                filtered_activities.append(activity)

    print(f"   ✓ Отфильтровано {len(filtered_activities)} активностей")

    # 4. Показываем активности по датам
    print(f"\n4. Активности по датам:")
    activities_by_date = {}
    for activity in filtered_activities:
        start_time = activity.get('startTimeLocal', '')
        if start_time:
            activity_date = datetime.strptime(start_time[:10], '%Y-%m-%d').date()
            date_str = activity_date.strftime('%d.%m.%Y')
            if date_str not in activities_by_date:
                activities_by_date[date_str] = []
            activities_by_date[date_str].append(activity.get('activityName', 'No name'))

    for date_str in sorted(activities_by_date.keys()):
        count = len(activities_by_date[date_str])
        print(f"   {date_str}: {count} тренировок")

    # 5. Подключаемся к Google Sheets
    print(f"\n5. Подключение к Google Sheets...")
    try:
        sheet = connect_to_google_sheets()
        worksheet = sheet.worksheet("ВЕЛ БЕГ")
        print(f"   ✓ Подключено к листу: {worksheet.title}")
    except Exception as e:
        print(f"   ✗ Ошибка: {e}")
        return

    # 6. Парсим недели из таблицы
    print(f"\n6. Парсинг недель из таблицы...")
    try:
        week_columns = parse_week_dates_from_block_rows(worksheet)
        print(f"   ✓ Найдено {len(week_columns)} недель в таблице:")
        for col, date in sorted(week_columns.items()):
            print(f"      Столбец {col}: {date.strftime('%d.%m.%Y')}")
    except Exception as e:
        print(f"   ✗ Ошибка: {e}")
        return

    # 7. Проверяем сопоставление активностей и столбцов
    print(f"\n7. Проверка сопоставления активностей и столбцов...")
    from main import find_column_for_date

    activities_by_column = {}
    activities_without_column = []

    for activity in filtered_activities:
        start_time = activity.get('startTimeLocal', '')
        if start_time:
            activity_date = datetime.strptime(start_time[:10], '%Y-%m-%d').date()
            column = find_column_for_date(activity_date, week_columns)

            if column:
                if column not in activities_by_column:
                    activities_by_column[column] = []
                activities_by_column[column].append({
                    'date': activity_date,
                    'name': activity.get('activityName', 'No name')
                })
            else:
                activities_without_column.append({
                    'date': activity_date,
                    'name': activity.get('activityName', 'No name')
                })

    print(f"   ✓ Активностей распределено по столбцам: {len(activities_by_column)}")
    for col in sorted(activities_by_column.keys()):
        count = len(activities_by_column[col])
        print(f"      Столбец {col}: {count} активностей")

    if activities_without_column:
        print(f"\n   ⚠️ ВНИМАНИЕ: {len(activities_without_column)} активностей БЕЗ столбца:")
        for act in activities_without_column[:10]:
            print(f"      {act['date'].strftime('%d.%m.%Y')}: {act['name'][:50]}")

    print("\n" + "="*60)
    print("ТЕСТ ЗАВЕРШЕН")
    print("="*60)

    # Итоговая статистика
    print(f"\nИТОГО:")
    print(f"  Всего активностей из Garmin: {len(activities)}")
    print(f"  Активностей с {start_date.strftime('%d.%m.%Y')}: {len(filtered_activities)}")
    print(f"  Недель в таблице: {len(week_columns)}")
    print(f"  Столбцов для синхронизации: {len(activities_by_column)}")
    print(f"  Активностей без столбца: {len(activities_without_column)}")

if __name__ == "__main__":
    test_full_sync()
