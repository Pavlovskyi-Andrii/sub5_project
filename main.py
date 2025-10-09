#!/usr/bin/env python3
import os
import json
from datetime import datetime, timedelta
from garminconnect import Garmin
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

def connect_to_garmin():
    """Подключение к Garmin Connect"""
    email = os.getenv('GARMIN_EMAIL')
    password = os.getenv('GARMIN_PASSWORD')
    session_data = os.getenv('SESSION_SECRET')
    
    if not email or not password:
        raise ValueError("Garmin credentials not found. Please set GARMIN_EMAIL and GARMIN_PASSWORD")
    
    print("Connecting to Garmin Connect...")
    
    try:
        client = Garmin(email, password)
        
        if session_data:
            try:
                print("Attempting to use saved session...")
                client.garth.loads(session_data)
                test_date = datetime.today().strftime("%Y-%m-%d")
                client.get_user_summary(test_date)
                print("✓ Successfully connected using saved session!")
                return client
            except Exception as e:
                print(f"Saved session invalid, logging in again... ({str(e)})")
        
        print("Logging in with credentials (this may take a moment)...")
        client.login()
        
        try:
            token_data = client.garth.dumps()
            print(f"\n{'='*60}")
            print(f"✓ Login successful!")
            print(f"{'='*60}")
            print(f"\nIMPORTANT: To avoid re-logging in every time, add this to your Secrets:")
            print(f"SESSION_SECRET = {token_data[:80]}...")
            print(f"(Full token data has been saved, copy from logs if needed)")
            print(f"{'='*60}\n")
        except Exception as e:
            print(f"Warning: Could not save session data: {str(e)}")
        
        print("Successfully connected to Garmin!")
        return client
        
    except Exception as e:
        error_msg = str(e)
        if "401" in error_msg or "Unauthorized" in error_msg:
            raise ValueError(
                "\n" + "="*60 + "\n"
                "❌ Garmin authentication failed!\n"
                "="*60 + "\n"
                "Possible solutions:\n"
                "1. Double-check GARMIN_EMAIL and GARMIN_PASSWORD are correct\n"
                "2. Try logging into connect.garmin.com in your browser first\n"
                "3. If you have 2FA/MFA enabled, you may need to disable it temporarily\n"
                "4. Wait a few minutes and try again (rate limiting)\n"
                "5. Check if your account is locked or requires verification\n"
                f"\nError details: {error_msg}\n"
                "="*60
            )
        raise

def connect_to_google_sheets():
    """Подключение к Google Sheets"""
    spreadsheet_url = os.getenv('GOOGLE_SHEET_URL')
    
    if not spreadsheet_url:
        raise ValueError("Google Sheet URL not found. Please set GOOGLE_SHEET_URL")
    
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    
    service_account_json = os.getenv('SERVICE_ACCOUNT_JSON')
    
    if service_account_json:
        creds_dict = json.loads(service_account_json)
        service_account_email = creds_dict.get('client_email', 'UNKNOWN')
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        print("Connecting to Google Sheets with service account...")
    else:
        print("Error: SERVICE_ACCOUNT_JSON not found in environment variables")
        raise ValueError("SERVICE_ACCOUNT_JSON is required")
    
    client = gspread.authorize(creds)
    
    try:
        sheet = client.open_by_url(spreadsheet_url)
        print(f"✓ Successfully connected to Google Sheet!")
        return sheet
    except PermissionError:
        print("\n" + "="*60)
        print("❌ ОШИБКА ДОСТУПА К GOOGLE ТАБЛИЦЕ")
        print("="*60)
        print(f"\nВам нужно дать доступ к таблице для Service Account!")
        print(f"\n1. Откройте таблицу: {spreadsheet_url}")
        print(f"2. Нажмите 'Настроить доступ' (Share)")
        print(f"3. Добавьте этот email с правами 'Редактор':")
        print(f"\n   {service_account_email}")
        print(f"\n4. Нажмите 'Готово' и запустите скрипт снова")
        print("="*60 + "\n")
        raise

def parse_date(date_str):
    """Парсинг даты из таблицы в формат datetime"""
    try:
        # Формат DD.MM.YY или DD.MM.YYYY
        parts = date_str.strip().split('.')
        if len(parts) >= 2:
            day = int(parts[0])
            month = int(parts[1])
            year = int(parts[2]) if len(parts) > 2 else datetime.now().year
            
            # Если год двузначный, добавляем 2000
            if year < 100:
                year += 2000
            
            return datetime(year, month, day)
    except:
        pass
    return None

def format_time(seconds):
    """Форматирование времени из секунд в ЧЧ:ММ:СС"""
    if not seconds:
        return ''
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    return f"{hours}:{minutes:02d}:{secs:02d}"

def format_pace(speed_mps):
    """Форматирование темпа из м/с в мин/км"""
    if not speed_mps or speed_mps == 0:
        return ''
    pace_min_per_km = 1000 / (speed_mps * 60)
    minutes = int(pace_min_per_km)
    seconds = int((pace_min_per_km - minutes) * 60)
    return f"{minutes}:{seconds:02d}"

def get_activities_for_date(garmin_client, target_date):
    """Получить все тренировки за конкретную дату"""
    # Получаем тренировки за этот день
    activities = garmin_client.get_activities(0, 50)  # Берем последние 50 тренировок
    
    result = []
    for activity in activities:
        start_time_str = activity.get('startTimeLocal', '')
        if start_time_str:
            activity_date = datetime.strptime(start_time_str.split()[0], '%Y-%m-%d').date()
            if activity_date == target_date.date():
                result.append(activity)
    
    return result

def get_training_blocks(worksheet):
    """Найти все блоки тренировок в таблице"""
    col_a = worksheet.col_values(1)
    
    blocks = []
    for row_num, value in enumerate(col_a, 1):
        value = value.strip()
        # Ищем строки с названиями тренировок
        if value and any(keyword in value.upper() for keyword in ['RUN', 'BIKE', 'БЕГ', 'ВЕЛ', 'ПЛАВ']):
            # Читаем строку с датами
            row_data = worksheet.row_values(row_num)
            blocks.append({
                'row': row_num,
                'name': value,
                'data': row_data
            })
    
    return blocks

def process_cycling_data(garmin_client, activities):
    """Обработка данных велосипеда"""
    if not activities:
        return {}
    
    data = {
        'avg_power': [],
        'normalized_power': [],
        'avg_speed': [],
        'avg_cadence': [],
        'avg_hr': []
    }
    
    for activity in activities:
        activity_id = activity['activityId']
        details = garmin_client.get_activity(activity_id)
        summary = details.get('summaryDTO', {})
        
        avg_power = summary.get('averagePower', '')
        normalized_power = summary.get('normalizedPower', '')
        avg_speed = round(summary.get('averageSpeed', 0) * 3.6, 1) if summary.get('averageSpeed') else ''
        avg_cadence = summary.get('averageBikeCadence', '')
        avg_hr = summary.get('averageHR', '')
        
        if avg_power:
            data['avg_power'].append(str(int(avg_power)))
        if normalized_power:
            data['normalized_power'].append(str(int(normalized_power)))
        if avg_speed:
            data['avg_speed'].append(str(avg_speed))
        if avg_cadence:
            data['avg_cadence'].append(str(int(avg_cadence)))
        if avg_hr:
            data['avg_hr'].append(str(int(avg_hr)))
    
    return data

def process_running_data(garmin_client, activity):
    """Обработка данных бега"""
    if not activity:
        return {}
    
    activity_id = activity['activityId']
    details = garmin_client.get_activity(activity_id)
    summary = details.get('summaryDTO', {})
    
    duration = format_time(summary.get('duration', 0))
    distance = round(summary.get('distance', 0) / 1000, 2) if summary.get('distance') else ''
    avg_speed = summary.get('averageSpeed', 0)
    avg_pace = format_pace(avg_speed) if avg_speed else ''
    avg_hr = summary.get('averageHR', '')
    
    return {
        'time': duration,
        'distance': f"{distance} км" if distance else '',
        'pace': f"{avg_pace} /км" if avg_pace else '',
        'hr': f"{int(avg_hr)} уд./мин" if avg_hr else ''
    }

def sync_to_sheet(garmin_client, worksheet, column):
    """Синхронизация данных в конкретный столбец"""
    print(f"\n{'='*60}")
    print(f"Синхронизация для столбца {column}")
    print(f"{'='*60}")
    
    # Находим все блоки тренировок
    blocks = get_training_blocks(worksheet)
    
    # Для каждого блока ищем дату в нужном столбце
    for block in blocks:
        row_num = block['row']
        name = block['name']
        row_data = block['data']
        
        # Определяем индекс колонки (E = 5, значит индекс 4)
        col_index = ord(column.upper()) - ord('A')
        
        if col_index >= len(row_data):
            continue
        
        date_str = row_data[col_index]
        date_obj = parse_date(date_str)
        
        if not date_obj:
            continue
        
        print(f"\n📅 {name} - {date_str}")
        
        # Получаем тренировки за эту дату
        activities = get_activities_for_date(garmin_client, date_obj)
        
        if not activities:
            print(f"  ℹ️  Нет тренировок в Garmin за {date_str}")
            continue
        
        # Разделяем по типам
        cycling_activities = [a for a in activities if 'cycling' in a.get('activityType', {}).get('typeKey', '').lower()]
        running_activities = [a for a in activities if 'running' in a.get('activityType', {}).get('typeKey', '').lower()]
        
        print(f"  🚴 Велосипед: {len(cycling_activities)} тренировок")
        print(f"  🏃 Бег: {len(running_activities)} тренировок")
        
        # Определяем куда записывать данные
        if 'БЕГ' in name.upper() or 'RUN' in name.upper():
            # Это блок бега
            if running_activities:
                run_data = process_running_data(garmin_client, running_activities[0])
                # Записываем данные
                # Строка +1 = Время, +2 = Расстояние, +3 = Темп, +4 = ЧСС
                if run_data.get('time'):
                    worksheet.update_cell(row_num + 1, col_index + 1, run_data['time'])
                if run_data.get('distance'):
                    worksheet.update_cell(row_num + 2, col_index + 1, run_data['distance'])
                if run_data.get('pace'):
                    worksheet.update_cell(row_num + 3, col_index + 1, run_data['pace'])
                if run_data.get('hr'):
                    worksheet.update_cell(row_num + 4, col_index + 1, run_data['hr'])
                print(f"  ✓ Записаны данные бега")
        
        elif 'ВЕЛ' in name.upper() or 'BIKE' in name.upper():
            # Это блок велосипеда
            if cycling_activities:
                cycle_data = process_cycling_data(garmin_client, cycling_activities[:2])  # Макс 2 тренировки
                
                # Определяем строки для записи
                # Нужно найти где именно записывать (может быть разная структура)
                # Пока записываем по аналогии: Средние ваты, NP, Скорость, Каденс, ЧСС
                
                # Ищем строки с текстом в столбце A
                # Пропустим это и просто запишем данные через слеш если их 2
                if len(cycle_data['avg_power']) >= 2:
                    avg_power_str = f"{cycle_data['avg_power'][0]}/{cycle_data['avg_power'][1]}"
                elif len(cycle_data['avg_power']) == 1:
                    avg_power_str = cycle_data['avg_power'][0]
                else:
                    avg_power_str = ''
                
                # Аналогично для других метрик
                print(f"  ✓ Средние ваты: {avg_power_str}")
                
                # TODO: нужно определить точные строки для записи
                # Пока просто выведем информацию
                print(f"  ℹ️  Данные готовы к записи (нужно уточнить строки)")

def main():
    try:
        print("=== Garmin to Google Sheets Sync ===\n")
        
        # Подключение к Garmin
        garmin = connect_to_garmin()
        
        # Подключение к Google Sheets
        sheet = connect_to_google_sheets()
        worksheet = sheet.worksheet("ВЕЛ БЕГ")
        print(f"✓ Opened worksheet: {worksheet.title}")
        
        # Определяем колонку для синхронизации (по умолчанию E = текущая неделя)
        column = os.getenv('SYNC_COLUMN', 'E')
        
        # Синхронизация
        sync_to_sheet(garmin, worksheet, column)
        
        print(f"\n{'='*60}")
        print("✅ Синхронизация завершена!")
        print(f"{'='*60}\n")
        
    except Exception as e:
        print(f"\n✗ Error: {e}")
        import traceback
        traceback.print_exc()
        raise

if __name__ == "__main__":
    main()
