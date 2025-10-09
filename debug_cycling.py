import os
import json
from datetime import datetime
from garminconnect import Garmin
from dotenv import load_dotenv

load_dotenv()

email = os.getenv('GARMIN_EMAIL')
password = os.getenv('GARMIN_PASSWORD')
session_data = os.getenv('SESSION_SECRET')

client = Garmin(email, password)

client.login()

# Получаем тренировки за 07.10.2025
target_date = datetime(2025, 10, 7).date()
activities = client.get_activities(0, 50)

print(f"Ищем тренировки за {target_date}...\n")

for activity in activities:
    start_time_str = activity.get('startTimeLocal', '')
    if start_time_str:
        activity_date = datetime.strptime(start_time_str.split()[0], '%Y-%m-%d').date()
        if activity_date == target_date:
            activity_type = activity.get('activityType', {}).get('typeKey', '')
            print(f"="*60)
            print(f"Тренировка: {activity.get('activityName', 'Без названия')}")
            print(f"Тип: {activity_type}")
            print(f"ID: {activity['activityId']}")
            print(f"Время: {start_time_str}")
            
            # Получаем детали
            details = client.get_activity(activity['activityId'])
            
            print(f"\nДетали:")
            print(f"  avg_power: {details.get('avgPower', 'НЕТ')}")
            print(f"  normalized_power: {details.get('normalizedPower', 'НЕТ')}")
            print(f"  avg_speed: {details.get('averageSpeed', 'НЕТ')}")
            print(f"  avg_cadence: {details.get('averageBikingCadenceInRevPerMinute', 'НЕТ')}")
            print(f"  avg_hr: {details.get('averageHR', 'НЕТ')}")
            
            # Показываем все ключи в details
            print(f"\nВсе доступные поля:")
            for key in sorted(details.keys()):
                if 'power' in key.lower() or 'watt' in key.lower():
                    print(f"  {key}: {details[key]}")
