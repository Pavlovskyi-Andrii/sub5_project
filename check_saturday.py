import os
from datetime import datetime
from garminconnect import Garmin
from dotenv import load_dotenv

load_dotenv()

email = os.getenv('GARMIN_EMAIL')
password = os.getenv('GARMIN_PASSWORD')

client = Garmin(email, password)
client.login()

# Получаем тренировки за 04.10.2025 (суббота)
target_date = datetime(2025, 10, 4).date()
activities = client.get_activities(0, 50)

print(f"Тренировки за {target_date} (суббота - 2 вел + 1 бег):\n")

cycling_count = 0
running_count = 0

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
            
            if 'cycling' in activity_type.lower():
                cycling_count += 1
                print(f"\n🚴 ВЕЛОСИПЕД #{cycling_count}:")
                print(f"  avg_power: {details.get('avgPower', 'НЕТ')}")
                print(f"  normalized_power: {details.get('normalizedPower', 'НЕТ')}")
                print(f"  avg_speed: {details.get('averageSpeed', 'НЕТ') * 3.6 if details.get('averageSpeed') else 'НЕТ'} км/ч")
                print(f"  avg_cadence: {details.get('averageBikingCadenceInRevPerMinute', 'НЕТ')}")
                print(f"  avg_hr: {details.get('averageHR', 'НЕТ')}")
            elif 'running' in activity_type.lower():
                running_count += 1
                print(f"\n🏃 БЕГ #{running_count}:")
                duration = details.get('duration', 0)
                hours = int(duration // 3600)
                minutes = int((duration % 3600) // 60)
                secs = int(duration % 60)
                print(f"  Время: {hours}:{minutes:02d}:{secs:02d}")
                print(f"  Расстояние: {details.get('distance', 0) / 1000:.2f} км")
                print(f"  Средний темп: {details.get('averageSpeed', 'НЕТ')} м/с")
                print(f"  avg_hr: {details.get('averageHR', 'НЕТ')}")

print(f"\n{'='*60}")
print(f"Итого: {cycling_count} велосипедных, {running_count} беговых")
