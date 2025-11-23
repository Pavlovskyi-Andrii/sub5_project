#!/usr/bin/env python3
"""
Диагностический скрипт для проверки синхронизации
"""
import os
from datetime import datetime, timedelta
from dotenv import load_dotenv

load_dotenv()

print("="*80)
print("ДИАГНОСТИКА СИНХРОНИЗАЦИИ")
print("="*80)

# Шаг 1: Подключаемся к Garmin
print("\n[1] Подключение к Garmin...")
try:
    from main import connect_to_garmin
    garmin = connect_to_garmin()
    print("✓ Подключено к Garmin")
except Exception as e:
    print(f"✗ Ошибка подключения к Garmin: {e}")
    exit(1)

# Шаг 2: Получаем тренировки
print("\n[2] Получение тренировок из Garmin...")
try:
    activities = garmin.get_activities(0, 200)
    print(f"✓ Получено {len(activities)} тренировок")

    # Показываем последние 20 тренировок
    print("\nПоследние 20 тренировок:")
    print(f"{'Дата':<15} {'Название':<50} {'Тип':<20}")
    print("-" * 85)

    for i, activity in enumerate(activities[:20]):
        start_time = activity.get('startTimeLocal', 'Нет даты')
        date_str = start_time[:10] if start_time else 'Нет даты'
        name = activity.get('activityName', 'Без названия')[:47] + "..."
        activity_type = activity.get('activityType', {}).get('typeKey', 'unknown')
        print(f"{date_str:<15} {name:<50} {activity_type:<20}")

except Exception as e:
    print(f"✗ Ошибка получения тренировок: {e}")
    import traceback
    traceback.print_exc()
    exit(1)

# Шаг 3: Подключаемся к Google Sheets
print("\n[3] Подключение к Google Sheets...")
try:
    from main import connect_to_google_sheets
    sheet = connect_to_google_sheets()
    worksheet = sheet.worksheet("ВЕЛ БЕГ")
    print("✓ Подключено к таблице")
except Exception as e:
    print(f"✗ Ошибка подключения к Google Sheets: {e}")
    exit(1)

# Шаг 4: Парсим даты недель
print("\n[4] Парсинг дат недель из строки 20...")
try:
    from main import parse_week_dates_from_block_rows
    week_columns = parse_week_dates_from_block_rows(worksheet)
    print(f"✓ Найдено {len(week_columns)} недель")

    print("\nНедели в таблице:")
    for col, sunday_date in sorted(week_columns.items()):
        saturday_date = sunday_date - timedelta(days=1)
        friday_date = saturday_date + timedelta(days=6)
        print(f"  Столбец {col}: {saturday_date.strftime('%d.%m.%Y')} (сб) - {friday_date.strftime('%d.%m.%Y')} (пт)")

except Exception as e:
    print(f"✗ Ошибка парсинга дат: {e}")
    import traceback
    traceback.print_exc()
    exit(1)

# Шаг 5: Группируем тренировки по неделям
print("\n[5] Группировка тренировок по неделям...")
try:
    from main import find_column_for_date

    activities_by_week = {}
    activities_without_column = []

    for activity in activities:
        start_time = activity.get('startTimeLocal', '')
        if start_time:
            activity_date = datetime.strptime(start_time[:10], '%Y-%m-%d').date()
            activity_name = activity.get('activityName', 'Без названия')
            activity_type = activity.get('activityType', {}).get('typeKey', 'unknown')

            column = find_column_for_date(activity_date, week_columns)

            if column:
                if column not in activities_by_week:
                    activities_by_week[column] = []
                activities_by_week[column].append({
                    'date': activity_date,
                    'name': activity_name,
                    'type': activity_type
                })
            else:
                activities_without_column.append({
                    'date': activity_date,
                    'name': activity_name,
                    'type': activity_type
                })

    print(f"✓ Тренировки распределены по {len(activities_by_week)} неделям")

    # Показываем распределение
    for col in sorted(activities_by_week.keys()):
        week_activities = activities_by_week[col]
        sunday_date = week_columns[col]
        saturday_date = sunday_date - timedelta(days=1)
        friday_date = saturday_date + timedelta(days=6)

        print(f"\nСтолбец {col} ({saturday_date.strftime('%d.%m')} - {friday_date.strftime('%d.%m')}): {len(week_activities)} тренировок")
        for act in week_activities[:5]:  # Показываем первые 5
            print(f"  - {act['date'].strftime('%d.%m')}: {act['name'][:40]} ({act['type']})")
        if len(week_activities) > 5:
            print(f"  ... и еще {len(week_activities) - 5} тренировок")

    if activities_without_column:
        print(f"\n⚠️  {len(activities_without_column)} тренировок НЕ попали ни в один столбец:")
        for act in activities_without_column[:10]:
            print(f"  - {act['date'].strftime('%d.%m.%Y')}: {act['name'][:40]} ({act['type']})")
        if len(activities_without_column) > 10:
            print(f"  ... и еще {len(activities_without_column) - 10} тренировок")

except Exception as e:
    print(f"✗ Ошибка группировки: {e}")
    import traceback
    traceback.print_exc()
    exit(1)

# Шаг 6: Фокусируемся на столбцах K и L
print("\n[6] Проверка столбцов K и L:")
print("="*80)

for col in ['K', 'L']:
    if col in week_columns:
        sunday_date = week_columns[col]
        saturday_date = sunday_date - timedelta(days=1)
        friday_date = saturday_date + timedelta(days=6)

        print(f"\nСтолбец {col}:")
        print(f"  Неделя: {saturday_date.strftime('%d.%m.%Y')} (сб) - {friday_date.strftime('%d.%m.%Y')} (пт)")
        print(f"  Воскресенье из строки 20: {sunday_date.strftime('%d.%m.%Y')}")

        if col in activities_by_week:
            week_acts = activities_by_week[col]
            print(f"  ✓ Найдено {len(week_acts)} тренировок:")

            # Группируем по типам
            cycling = [a for a in week_acts if 'cycling' in a['type'].lower()]
            running = [a for a in week_acts if 'running' in a['type'].lower()]

            print(f"    🚴 Велосипед: {len(cycling)}")
            for act in cycling:
                print(f"      - {act['date'].strftime('%d.%m')}: {act['name'][:50]}")

            print(f"    🏃 Бег: {len(running)}")
            for act in running:
                print(f"      - {act['date'].strftime('%d.%m')}: {act['name'][:50]}")

            # Проверяем субботу отдельно
            saturday_acts = [a for a in week_acts if a['date'] == saturday_date]
            if saturday_acts:
                print(f"\n  Тренировки за субботу {saturday_date.strftime('%d.%m')}:")
                for act in saturday_acts:
                    print(f"    - {act['name'][:50]} ({act['type']})")
            else:
                print(f"\n  ⚠️  НЕТ тренировок за субботу {saturday_date.strftime('%d.%m')}")
        else:
            print(f"  ✗ НЕТ тренировок для этой недели")
    else:
        print(f"\nСтолбец {col}: НЕ НАЙДЕН в таблице")

print("\n" + "="*80)
print("ДИАГНОСТИКА ЗАВЕРШЕНА")
print("="*80)
