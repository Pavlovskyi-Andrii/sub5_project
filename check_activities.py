#!/usr/bin/env python3
"""Проверка активностей в Garmin"""
from datetime import datetime
from dotenv import load_dotenv
from main import connect_to_garmin

load_dotenv()

def check_activities():
    """Проверяем активности за последние дни"""
    print("="*60)
    print("ПРОВЕРКА АКТИВНОСТЕЙ В GARMIN")
    print("="*60)

    garmin = connect_to_garmin()

    print("\nПолучаем последние 50 активностей...")
    activities = garmin.get_activities(0, 50)

    print(f"\nВсего активностей: {len(activities)}")
    print("\nПоследние 20 активностей:")
    print("-" * 60)

    for i, activity in enumerate(activities[:20]):
        start_time = activity.get('startTimeLocal', '')
        if start_time:
            activity_date = datetime.strptime(start_time[:10], '%Y-%m-%d').date()
            activity_name = activity.get('activityName', 'No name')
            activity_type = activity.get('activityType', {}).get('typeKey', 'unknown')
            print(f"{i+1}. {activity_date.strftime('%d.%m.%Y')}: {activity_name[:40]} ({activity_type})")

    # Проверяем активности за ноябрь 2024
    print("\n" + "="*60)
    print("АКТИВНОСТИ ЗА НОЯБРЬ 2024:")
    print("="*60)
    nov_activities = []
    for activity in activities:
        start_time = activity.get('startTimeLocal', '')
        if start_time:
            activity_date = datetime.strptime(start_time[:10], '%Y-%m-%d').date()
            if activity_date.year == 2024 and activity_date.month == 11:
                nov_activities.append(activity)
                activity_name = activity.get('activityName', 'No name')
                activity_type = activity.get('activityType', {}).get('typeKey', 'unknown')
                print(f"{activity_date.strftime('%d.%m.%Y')}: {activity_name[:40]} ({activity_type})")

    print(f"\nВсего активностей в ноябре 2024: {len(nov_activities)}")

    # Проверяем диапазон дат
    if activities:
        dates = []
        for activity in activities:
            start_time = activity.get('startTimeLocal', '')
            if start_time:
                activity_date = datetime.strptime(start_time[:10], '%Y-%m-%d').date()
                dates.append(activity_date)

        if dates:
            print(f"\nДиапазон дат активностей:")
            print(f"  Самая ранняя: {min(dates).strftime('%d.%m.%Y')}")
            print(f"  Самая поздняя: {max(dates).strftime('%d.%m.%Y')}")

if __name__ == '__main__':
    check_activities()
