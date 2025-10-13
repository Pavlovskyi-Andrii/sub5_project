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
    """Парсинг даты из таблицы в формат datetime с поддержкой regex"""
    import re
    
    if not date_str:
        return None
    
    # Ищем первый паттерн даты вида DD.MM или DD.MM.YY или DD.MM.YYYY
    match = re.search(r'(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?', date_str)
    
    if match:
        try:
            day = int(match.group(1))
            month = int(match.group(2))
            year_str = match.group(3)
            
            if year_str:
                year = int(year_str)
                if year < 100:
                    year += 2000
            else:
                year = datetime.now().year
            
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
        
        # Если в строке блока нет даты, проверяем строку 1 (заголовки недель)
        if not date_obj:
            row1 = worksheet.row_values(1)
            if col_index < len(row1):
                date_str = row1[col_index]
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
        strength_activities = [a for a in activities if 'strength' in a.get('activityType', {}).get('typeKey', '').lower()]
        swimming_activities = [a for a in activities if 'swimming' in a.get('activityType', {}).get('typeKey', '').lower() or 'lap_swimming' in a.get('activityType', {}).get('typeKey', '').lower()]
        
        print(f"  🚴 Велосипед: {len(cycling_activities)} тренировок")
        print(f"  🏃 Бег: {len(running_activities)} тренировок")
        print(f"  💪 Силовая: {len(strength_activities)} тренировок")
        print(f"  🏊 Плавание: {len(swimming_activities)} тренировок")
        
        # Определяем куда записывать данные
        # Сначала проверяем комбинированные блоки (вел+бег, например суббота)
        if ('ВЕЛ' in name.upper() or 'BIKE' in name.upper()) and ('БЕГ' in name.upper() or 'RUN' in name.upper()):
            # Это комбинированный блок (суббота: 2 вел + 1 бег)
            # Сначала записываем велосипед
            if cycling_activities:
                cycle_data = process_cycling_data(garmin_client, cycling_activities[:2])
                col_a = worksheet.col_values(1)
                
                def format_values(values_list):
                    if len(values_list) >= 2:
                        return f"{values_list[0]}/{values_list[1]}"
                    elif len(values_list) == 1:
                        return values_list[0]
                    return ''
                
                avg_power_str = format_values(cycle_data['avg_power'])
                np_str = format_values(cycle_data['normalized_power'])
                speed_str = format_values(cycle_data['avg_speed'])
                cadence_str = format_values(cycle_data['avg_cadence'])
                hr_str = format_values(cycle_data['avg_hr'])
                
                print(f"  📊 Данные вел: power={avg_power_str}, NP={np_str}, speed={speed_str}, cadence={cadence_str}, HR={hr_str}")
                
                # Ищем строки в пределах блока
                block_end = len(col_a)
                for next_idx in range(row_num, len(col_a)):
                    next_text = col_a[next_idx].strip().upper()
                    if next_text and next_idx > row_num and any(kw in next_text for kw in ['RUN', 'BIKE', 'БЕГ', 'ВЕЛ', 'ПЛАВ', 'ЛОНГ', 'ИНТЕРВАЛ', 'КОРОТКИЕ']):
                        block_end = next_idx
                        break
                
                print(f"  🔍 Ищем вел данные с row {row_num} до {block_end}")
                
                for search_idx in range(row_num - 1, min(block_end, len(col_a))):
                    cell_text = col_a[search_idx].strip().lower()
                    actual_row = search_idx + 1
                    
                    if 'средн' in cell_text and 'ват' in cell_text:
                        if avg_power_str:
                            worksheet.update_cell(actual_row, col_index + 1, avg_power_str)
                            print(f"  ✓ Средние ваты: {avg_power_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'normalized' in cell_text or ('power' in cell_text and 'norm' in cell_text):
                        if np_str:
                            worksheet.update_cell(actual_row, col_index + 1, np_str)
                            print(f"  ✓ Normalized Power: {np_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'сред' in cell_text and 'скор' in cell_text:
                        if speed_str:
                            worksheet.update_cell(actual_row, col_index + 1, speed_str)
                            print(f"  ✓ Средняя скорость: {speed_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif ('средн' in cell_text or 'срадн' in cell_text) and 'чсс' in cell_text:
                        if hr_str:
                            worksheet.update_cell(actual_row, col_index + 1, hr_str)
                            print(f"  ✓ Средняя ЧСС: {hr_str} → {chr(64+col_index+1)}{actual_row}")
            
            # Потом записываем бег
            if running_activities:
                run_data = process_running_data(garmin_client, running_activities[0])
                # Ищем строки "Бег брик" и "ЧСС бег" для записи
                for search_idx in range(row_num - 1, min(row_num + 20, len(col_a))):
                    cell_text = col_a[search_idx].strip().lower()
                    actual_row = search_idx + 1
                    
                    if 'бег' in cell_text and 'брик' in cell_text:
                        # Записываем описание бега
                        desc = f"{run_data.get('distance', '')} {run_data.get('pace', '')}"
                        if desc.strip():
                            worksheet.update_cell(actual_row, col_index + 1, desc)
                            print(f"  ✓ Бег брик: {desc} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'чсс' in cell_text and 'бег' in cell_text:
                        if run_data.get('hr'):
                            hr_only = run_data['hr'].replace(' уд./мин', '')
                            worksheet.update_cell(actual_row, col_index + 1, hr_only)
                            print(f"  ✓ ЧСС бег: {hr_only} → {chr(64+col_index+1)}{actual_row}")
        
        elif 'БЕГ' in name.upper() or 'RUN' in name.upper():
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
                
                # Ищем строки для записи в столбце A
                col_a = worksheet.col_values(1)
                
                # Формируем данные (через слеш если 2 тренировки)
                def format_values(values_list):
                    if len(values_list) >= 2:
                        return f"{values_list[0]}/{values_list[1]}"
                    elif len(values_list) == 1:
                        return values_list[0]
                    return ''
                
                avg_power_str = format_values(cycle_data['avg_power'])
                np_str = format_values(cycle_data['normalized_power'])
                speed_str = format_values(cycle_data['avg_speed'])
                cadence_str = format_values(cycle_data['avg_cadence'])
                hr_str = format_values(cycle_data['avg_hr'])
                
                print(f"  📊 Данные вел: power={avg_power_str}, NP={np_str}, speed={speed_str}, cadence={cadence_str}, HR={hr_str}")
                
                # Ищем строки по тексту в колонке A (только в пределах блока)
                # row_num - это уже 1-based индекс из enumerate
                # col_a - это список со строками, индексы с 0
                
                # Находим конец блока (следующий заголовок блока или конец таблицы)
                block_end = len(col_a)
                for next_idx in range(row_num, len(col_a)):
                    next_text = col_a[next_idx].strip().upper()
                    if next_text and any(kw in next_text for kw in ['RUN', 'BIKE', 'БЕГ', 'ВЕЛ', 'ПЛАВ']):
                        # Это следующий блок
                        block_end = next_idx
                        break
                
                print(f"  🔍 Ищем с row {row_num} до {block_end}")
                
                for search_idx in range(row_num - 1, min(block_end, len(col_a))):
                    cell_text = col_a[search_idx].strip().lower()
                    actual_row = search_idx + 1  # Реальный номер строки в Google Sheets
                    
                    if 'средн' in cell_text and 'ват' in cell_text:
                        if avg_power_str:
                            worksheet.update_cell(actual_row, col_index + 1, avg_power_str)
                            print(f"  ✓ Средние ваты: {avg_power_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'normalized' in cell_text or ('power' in cell_text and 'norm' in cell_text):
                        if np_str:
                            worksheet.update_cell(actual_row, col_index + 1, np_str)
                            print(f"  ✓ Normalized Power: {np_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'сред' in cell_text and 'скор' in cell_text:
                        if speed_str:
                            worksheet.update_cell(actual_row, col_index + 1, speed_str)
                            print(f"  ✓ Средняя скорость: {speed_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'частот' in cell_text and 'вращ' in cell_text:
                        if cadence_str:
                            worksheet.update_cell(actual_row, col_index + 1, cadence_str)
                            print(f"  ✓ Частота вращения: {cadence_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    # Различаем ЧП (каденс) и ЧСС (пульс)
                    elif ('средн' in cell_text or 'срадн' in cell_text) and 'чп' in cell_text and 'чсс' not in cell_text:
                        # Это каденс (ЧП = частота педалирования)
                        if cadence_str:
                            worksheet.update_cell(actual_row, col_index + 1, cadence_str)
                            print(f"  ✓ Средняя ЧП (каденс): {cadence_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif ('средн' in cell_text or 'срадн' in cell_text) and 'чсс' in cell_text:
                        # Это пульс (ЧСС)
                        if hr_str:
                            worksheet.update_cell(actual_row, col_index + 1, hr_str)
                            print(f"  ✓ Средняя ЧСС: {hr_str} → {chr(64+col_index+1)}{actual_row}")
        
        elif ('СТАНОВ' in name.upper() or 'ПЛАВ' in name.upper()) and 'ПН' in name.upper():
            # Это понедельник - становая + плавание (строки 32-34)
            # Записываем длительность тренировок
            if strength_activities or swimming_activities:
                durations = []
                
                # Получаем длительность силовой
                if strength_activities:
                    activity_id = strength_activities[0]['activityId']
                    details = garmin_client.get_activity(activity_id)
                    summary = details.get('summaryDTO', {})
                    duration_sec = summary.get('duration', 0)
                    if duration_sec:
                        hours = int(duration_sec // 3600)
                        minutes = int((duration_sec % 3600) // 60)
                        seconds = int(duration_sec % 60)
                        duration_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                        durations.append(duration_str)
                
                # Получаем длительность плавания
                if swimming_activities:
                    activity_id = swimming_activities[0]['activityId']
                    details = garmin_client.get_activity(activity_id)
                    summary = details.get('summaryDTO', {})
                    duration_sec = summary.get('duration', 0)
                    if duration_sec:
                        hours = int(duration_sec // 3600)
                        minutes = int((duration_sec % 3600) // 60)
                        seconds = int(duration_sec % 60)
                        duration_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
                        durations.append(duration_str)
                
                # Записываем длительности в строки 33 и 34
                if len(durations) >= 1:
                    worksheet.update_cell(33, col_index + 1, durations[0])
                    print(f"  ✓ Длительность первой тренировки: {durations[0]} → {chr(64+col_index+1)}33")
                
                if len(durations) >= 2:
                    worksheet.update_cell(34, col_index + 1, durations[1])
                    print(f"  ✓ Длительность второй тренировки: {durations[1]} → {chr(64+col_index+1)}34")

def get_week_start(date_obj):
    """Получить субботу начала недели для данной даты"""
    # Неделя начинается с субботы (weekday 5)
    # Если сегодня суббота или позже - это текущая неделя
    # Если до субботы - берем предыдущую субботу
    days_since_saturday = (date_obj.weekday() + 2) % 7  # Суббота = 0
    week_start = date_obj - timedelta(days=days_since_saturday)
    return week_start

def parse_week_dates_from_row1(worksheet):
    """Парсит строку 1 и возвращает словарь {столбец: дата_начала_недели}"""
    row1 = worksheet.row_values(1)
    week_columns = {}
    
    for idx, cell in enumerate(row1):
        if not cell.strip():
            continue
        
        col_letter = chr(65 + idx)  # A, B, C, D, E, F...
        
        # Парсим дату
        cell_text = cell.strip()
        try:
            # Формат: "04.10.25" или "13.09"
            if '.' in cell_text:
                parts = cell_text.split('.')
                if len(parts) == 3:
                    day, month, year = parts
                    if len(year) == 2:
                        year = '20' + year
                    date_obj = datetime.strptime(f"{day}.{month}.{year}", "%d.%m.%Y")
                elif len(parts) == 2:
                    day, month = parts
                    year = datetime.now().year
                    date_obj = datetime.strptime(f"{day}.{month}.{year}", "%d.%m.%Y")
                else:
                    continue
                
                week_columns[col_letter] = date_obj.date()
        except:
            continue
    
    return week_columns

def find_column_for_date(activity_date, week_columns):
    """Находит столбец для записи данных на основе даты тренировки"""
    # Определяем начало недели для тренировки
    week_start = get_week_start(activity_date)
    
    # Ищем подходящий столбец
    for col_letter, col_week_start in sorted(week_columns.items()):
        if col_week_start == week_start:
            return col_letter
    
    # Если точное совпадение не найдено, ищем ближайшую неделю
    for col_letter, col_week_start in sorted(week_columns.items()):
        # Проверяем попадает ли тренировка в эту неделю (суббота + 6 дней)
        week_end = col_week_start + timedelta(days=6)
        if col_week_start <= activity_date <= week_end:
            return col_letter
    
    return None

def main():
    try:
        print("=== Garmin to Google Sheets Sync ===\n")
        
        # Подключение к Garmin
        garmin = connect_to_garmin()
        
        # Подключение к Google Sheets
        sheet = connect_to_google_sheets()
        worksheet = sheet.worksheet("ВЕЛ БЕГ")
        print(f"✓ Opened worksheet: {worksheet.title}")
        
        # Парсим даты недель из строки 1
        week_columns = parse_week_dates_from_row1(worksheet)
        print(f"✓ Найдено {len(week_columns)} недель в таблице")
        
        # Получаем тренировки за последние N дней
        days_to_sync = int(os.getenv('DAYS_TO_SYNC', '14'))  # По умолчанию 2 недели
        activities = garmin.get_activities(0, days_to_sync * 2)  # С запасом
        
        # Группируем тренировки по неделям
        activities_by_week = {}
        for activity in activities:
            activity_date = datetime.strptime(activity['startTimeLocal'][:10], '%Y-%m-%d').date()
            column = find_column_for_date(activity_date, week_columns)
            
            if column:
                if column not in activities_by_week:
                    activities_by_week[column] = []
                activities_by_week[column].append(activity)
        
        # Синхронизируем каждую неделю
        for column in sorted(activities_by_week.keys()):
            week_activities = activities_by_week[column]
            week_date = week_columns.get(column)
            print(f"\n{'='*60}")
            print(f"Синхронизация недели {column} (начало: {week_date.strftime('%d.%m.%Y')})")
            print(f"Найдено тренировок: {len(week_activities)}")
            print(f"{'='*60}")
            
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
