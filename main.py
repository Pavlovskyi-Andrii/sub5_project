#!/usr/bin/env python3
import os
import json
import re
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
    # Теперь столбец B (2) содержит названия блоков, столбец A - порядковые номера
    col_b = worksheet.col_values(2)
    
    blocks = []
    for row_num, value in enumerate(col_b, 1):
        if not value:
            continue
        value = str(value).strip()
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
        'pace': avg_pace if avg_pace else '',
        'hr': f"{int(avg_hr)} уд./мин" if avg_hr else ''
    }

class BatchUpdater:
    """Класс для накопления обновлений и отправки batch запросом"""
    def __init__(self, worksheet):
        self.worksheet = worksheet
        self.updates = []
    
    def add_update(self, row, col, value):
        """Добавить обновление в очередь"""
        self.updates.append({
            'row': row,
            'col': col,
            'value': str(value) if value else ''
        })
    
    def flush(self):
        """Отправить все накопленные обновления одним запросом"""
        if not self.updates:
            return
        
        # Формируем batch_update запрос
        cells_to_update = []
        for update in self.updates:
            cell = gspread.utils.rowcol_to_a1(update['row'], update['col'])
            cells_to_update.append({
                'range': cell,
                'values': [[update['value']]]
            })
        
        # Отправляем batch update
        if cells_to_update:
            self.worksheet.batch_update(cells_to_update, value_input_option='USER_ENTERED')
        
        self.updates = []

def sync_to_sheet(garmin_client, worksheet, column):
    """Синхронизация данных в конкретный столбец"""
    print(f"\n{'='*60}")
    print(f"Синхронизация для столбца {column}")
    print(f"{'='*60}")
    
    # Создаем batch updater
    batch = BatchUpdater(worksheet)
    
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
            # Фиксированные строки 7-15
            
            # Сначала записываем велосипед (строки 7-11)
            if cycling_activities:
                cycle_data = process_cycling_data(garmin_client, cycling_activities[:2])
                
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
                
                # Строка 7: Средние ваты
                if avg_power_str:
                    batch.add_update(7, col_index + 1, avg_power_str)
                    print(f"  ✓ Средние ваты: {avg_power_str} → {chr(64+col_index+1)}7")
                
                # Строка 8: Normalized Power
                if np_str:
                    batch.add_update(8, col_index + 1, np_str)
                    print(f"  ✓ Normalized Power: {np_str} → {chr(64+col_index+1)}8")
                
                # Строка 9: Сред.скорость
                if speed_str:
                    batch.add_update(9, col_index + 1, speed_str)
                    print(f"  ✓ Средняя скорость: {speed_str} → {chr(64+col_index+1)}9")
                
                # Строка 10: Частота вращения
                if cadence_str:
                    batch.add_update(10, col_index + 1, cadence_str)
                    print(f"  ✓ Частота вращения: {cadence_str} → {chr(64+col_index+1)}10")
                
                # Строка 11: Средняя ЧСС
                if hr_str:
                    batch.add_update(11, col_index + 1, hr_str)
                    print(f"  ✓ Средняя ЧСС: {hr_str} → {chr(64+col_index+1)}11")
            
            # Потом записываем бег брик (строки 13-15)
            if running_activities:
                run_data = process_running_data(garmin_client, running_activities[0])
                
                # Строка 13: Бег брик км
                if run_data.get('distance'):
                    distance_only = run_data['distance'].replace(' км', '')
                    batch.add_update(13, col_index + 1, distance_only)
                    print(f"  ✓ Бег брик км: {distance_only} → {chr(64+col_index+1)}13")
                
                # Строка 14: Бег брик темп
                if run_data.get('pace'):
                    batch.add_update(14, col_index + 1, run_data['pace'])
                    print(f"  ✓ Бег брик темп: {run_data['pace']} → {chr(64+col_index+1)}14")
                
                # Строка 15: Бег брик ЧСС
                if run_data.get('hr'):
                    hr_only = run_data['hr'].replace(' уд./мин', '')
                    batch.add_update(15, col_index + 1, hr_only)
                    print(f"  ✓ Бег брик ЧСС: {hr_only} → {chr(64+col_index+1)}15")
        
        elif 'БЕГ' in name.upper() or 'RUN' in name.upper():
            # Это блок бега
            if running_activities:
                run_data = process_running_data(garmin_client, running_activities[0])
                # Записываем данные
                # Строка +1 = Время, +2 = Расстояние, +3 = Темп, +4 = ЧСС
                if run_data.get('time'):
                    batch.add_update(row_num + 1, col_index + 1, run_data['time'])
                if run_data.get('distance'):
                    distance_only = run_data['distance'].replace(' км', '')
                    batch.add_update(row_num + 2, col_index + 1, distance_only)
                if run_data.get('pace'):
                    batch.add_update(row_num + 3, col_index + 1, run_data['pace'])
                if run_data.get('hr'):
                    hr_only = run_data['hr'].replace(' уд./мин', '')
                    batch.add_update(row_num + 4, col_index + 1, hr_only)
                print(f"  ✓ Записаны данные бега")
        
        elif 'ВЕЛ' in name.upper() or 'BIKE' in name.upper():
            # Это блок велосипеда
            if cycling_activities:
                cycle_data = process_cycling_data(garmin_client, cycling_activities[:2])  # Макс 2 тренировки
                
                # Ищем строки для записи в столбце B (столбец A теперь с номерами)
                col_b = worksheet.col_values(2)
                
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
                
                # Ищем строки по тексту в колонке B (только в пределах блока)
                # row_num - это уже 1-based индекс из enumerate
                # col_b - это список со строками, индексы с 0
                
                # Находим конец блока (следующий заголовок блока или конец таблицы)
                block_end = len(col_b)
                for next_idx in range(row_num, len(col_b)):
                    next_text = str(col_b[next_idx]).strip().upper() if next_idx < len(col_b) else ''
                    if next_text and any(kw in next_text for kw in ['RUN', 'BIKE', 'БЕГ', 'ВЕЛ', 'ПЛАВ']):
                        # Это следующий блок
                        block_end = next_idx
                        break
                
                print(f"  🔍 Ищем с row {row_num} до {block_end}")
                
                for search_idx in range(row_num - 1, min(block_end, len(col_b))):
                    cell_text = str(col_b[search_idx]).strip().lower() if search_idx < len(col_b) else ''
                    actual_row = search_idx + 1  # Реальный номер строки в Google Sheets
                    
                    # Для вторника/четверга: Время, Расстояние, Средний темп (=скорость для вела), Средняя ЧП
                    # Агрегируем данные через слеш если несколько тренировок (как для мощности)
                    if 'врем' in cell_text and 'длительност' not in cell_text:
                        # Время тренировки (агрегация)
                        if cycling_activities:
                            times = [format_time(act.get('duration', 0)) for act in cycling_activities[:2]]
                            times = [t for t in times if t]
                            if times:
                                time_str = '/'.join(times) if len(times) > 1 else times[0]
                                batch.add_update(actual_row, col_index + 1, time_str)
                                print(f"  ✓ Время: {time_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'расстоян' in cell_text:
                        # Расстояние (агрегация)
                        if cycling_activities:
                            distances = []
                            for act in cycling_activities[:2]:
                                dist = act.get('distance', 0)
                                if dist:
                                    distances.append(str(round(dist / 1000, 2)))
                            if distances:
                                dist_str = '/'.join(distances) if len(distances) > 1 else distances[0]
                                batch.add_update(actual_row, col_index + 1, dist_str)
                                print(f"  ✓ Расстояние: {dist_str} км → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'средн' in cell_text and 'темп' in cell_text:
                        # Средний темп для вело = скорость (уже агрегировано)
                        if speed_str:
                            batch.add_update(actual_row, col_index + 1, speed_str)
                            print(f"  ✓ Средний темп (скорость): {speed_str} км/ч → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'средн' in cell_text and 'ват' in cell_text:
                        if avg_power_str:
                            batch.add_update(actual_row, col_index + 1, avg_power_str)
                            print(f"  ✓ Средние ваты: {avg_power_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'normalized' in cell_text or ('power' in cell_text and 'norm' in cell_text):
                        if np_str:
                            batch.add_update(actual_row, col_index + 1, np_str)
                            print(f"  ✓ Normalized Power: {np_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'сред' in cell_text and 'скор' in cell_text:
                        if speed_str:
                            batch.add_update(actual_row, col_index + 1, speed_str)
                            print(f"  ✓ Средняя скорость: {speed_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'частот' in cell_text and 'вращ' in cell_text:
                        if cadence_str:
                            batch.add_update(actual_row, col_index + 1, cadence_str)
                            print(f"  ✓ Частота вращения: {cadence_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    # Различаем ЧП (каденс) и ЧСС (пульс)
                    elif ('средн' in cell_text or 'срадн' in cell_text) and 'чп' in cell_text and 'чсс' not in cell_text:
                        # Это каденс (ЧП = частота педалирования)
                        if cadence_str:
                            batch.add_update(actual_row, col_index + 1, cadence_str)
                            print(f"  ✓ Средняя ЧП (каденс): {cadence_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif ('средн' in cell_text or 'срадн' in cell_text) and 'чсс' in cell_text:
                        # Это пульс (ЧСС)
                        if hr_str:
                            batch.add_update(actual_row, col_index + 1, hr_str)
                            print(f"  ✓ Средняя ЧСС: {hr_str} → {chr(64+col_index+1)}{actual_row}")
        
        elif ('СТАНОВ' in name.upper() or 'ПЛАВ' in name.upper()) and 'ПН' in name.upper():
            # Это понедельник - становая + плавание
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
                
                # Ищем строки для записи длительности (динамически в пределах блока)
                # Столбец B теперь содержит названия (столбец A - порядковые номера)
                col_b = worksheet.col_values(2)
                
                # Находим конец блока
                block_end = len(col_b)
                for next_idx in range(row_num, len(col_b)):
                    next_text = str(col_b[next_idx]).strip().upper() if next_idx < len(col_b) else ''
                    if next_text and next_idx > row_num and any(kw in next_text for kw in ['RUN', 'BIKE', 'БЕГ', 'ВЕЛ', 'ПЛАВ', 'ЛОНГ', 'ИНТЕРВАЛ', 'КОРОТКИЕ', 'ДЛИН']):
                        block_end = next_idx
                        break
                
                # Ищем строки "Длительность первой/второй тренировки"
                for search_idx in range(row_num - 1, min(block_end, len(col_b))):
                    cell_text = str(col_b[search_idx]).strip().lower() if search_idx < len(col_b) else ''
                    actual_row = search_idx + 1
                    
                    if 'длительност' in cell_text and 'перв' in cell_text:
                        if len(durations) >= 1:
                            batch.add_update(actual_row, col_index + 1, durations[0])
                            print(f"  ✓ Длительность первой тренировки: {durations[0]} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'длительност' in cell_text and 'втор' in cell_text:
                        if len(durations) >= 2:
                            batch.add_update(actual_row, col_index + 1, durations[1])
                            print(f"  ✓ Длительность второй тренировки: {durations[1]} → {chr(64+col_index+1)}{actual_row}")
    
    # Отправляем все накопленные обновления одним batch запросом
    batch.flush()

def get_week_start(date_obj):
    """Получить субботу начала недели для данной даты"""
    # Неделя начинается с субботы (weekday 5)
    # Если сегодня суббота или позже - это текущая неделя
    # Если до субботы - берем предыдущую субботу
    days_since_saturday = (date_obj.weekday() + 2) % 7  # Суббота = 0
    week_start = date_obj - timedelta(days=days_since_saturday)
    return week_start

def parse_week_dates_from_block_rows(worksheet):
    """Парсит даты из строк блоков и возвращает словарь {столбец: дата_начала_недели}"""
    # Строки с датами блоков (по новой структуре с нумерацией в столбце A)
    date_rows = [20, 33, 38, 73]  # Воскресенье, Понедельник, Вторник, Пятница
    
    # Читаем все строки одним batch запросом для оптимизации
    ranges = [f"{row}:{row}" for row in date_rows]
    try:
        batch_data = worksheet.batch_get(ranges)
    except Exception as e:
        print(f"✗ Ошибка при чтении дат: {e}")
        return {}
    
    # Словарь для хранения дат по столбцам: {столбец: [список_дат]}
    column_dates = {}
    
    # Парсим даты из каждой строки
    for row_data in batch_data:
        if not row_data:
            continue
        
        row_values = row_data[0] if row_data else []
        
        for idx, cell in enumerate(row_values):
            if not cell or not isinstance(cell, str):
                continue
            
            col_letter = chr(65 + idx)  # A, B, C, D, E, F...
            
            # Ищем дату в формате dd.mm.yy
            date_match = re.search(r'\b(\d{2})\.(\d{2})\.(\d{2})\b', cell)
            if date_match:
                try:
                    day, month, year = date_match.groups()
                    year = '20' + year
                    date_obj = datetime.strptime(f"{day}.{month}.{year}", "%d.%m.%Y").date()
                    
                    if col_letter not in column_dates:
                        column_dates[col_letter] = []
                    column_dates[col_letter].append(date_obj)
                except:
                    continue
    
    # Определяем начало недели (суббота) для каждого столбца
    week_columns = {}
    for col_letter, dates in column_dates.items():
        if dates:
            # Берем первую дату из столбца и определяем субботу начала недели
            first_date = min(dates)
            week_start = get_week_start(first_date)
            week_columns[col_letter] = week_start
    
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

def export_all_data_to_source(garmin, sheet):
    """Выгружает все доступные данные тренировок на лист 'исходник' для диагностики"""
    try:
        # Открываем лист исходник (или создаем если нет)
        try:
            worksheet = sheet.worksheet("исходник")
        except:
            worksheet = sheet.add_worksheet("исходник", rows=100, cols=20)
        
        print("\n" + "="*60)
        print("Выгрузка всех данных на лист 'исходник'")
        print("="*60)
        
        # Получаем тренировки за последнюю неделю
        days = int(os.getenv('DAYS_TO_SYNC', '7'))
        activities = garmin.get_activities(0, days * 2)
        
        # Очищаем лист
        worksheet.clear()
        
        # Собираем все данные для batch update
        batch = BatchUpdater(worksheet)
        row = 1
        
        for activity in activities:
            activity_type = activity.get('activityType', {}).get('typeKey', 'unknown')
            activity_name = activity.get('activityName', 'Без названия')
            start_time = activity.get('startTimeLocal', '')
            
            # Записываем заголовок тренировки
            batch.add_update(row, 1, f"=== {activity_name} ===")
            batch.add_update(row, 2, start_time[:10] if start_time else '')
            batch.add_update(row, 3, activity_type)
            print(f"\n{activity_name} ({start_time[:10]}) - {activity_type}")
            row += 1
            
            # Получаем детали тренировки
            try:
                activity_id = activity.get('activityId')
                details = garmin.get_activity(activity_id)
                summary = details.get('summaryDTO', {})
                
                # Основные метрики
                data_to_export = {
                    'Длительность': format_time(activity.get('duration', 0)),
                    'Расстояние (км)': round(activity.get('distance', 0) / 1000, 2) if activity.get('distance') else None,
                    'Средняя скорость (км/ч)': round(activity.get('averageSpeed', 0) * 3.6, 1) if activity.get('averageSpeed') else None,
                    'Средняя ЧСС': summary.get('averageHR'),
                    'Калории': activity.get('calories'),
                }
                
                # Данные велосипеда
                if 'cycling' in activity_type.lower():
                    data_to_export.update({
                        'Средняя мощность (Вт)': summary.get('avgPower'),
                        'Normalized Power': summary.get('normPower') or summary.get('normalizedPower'),
                        'Средняя каденс': summary.get('avgBikeCadence') or summary.get('averageBikingCadenceInRevPerMinute'),
                    })
                
                # Данные бега
                if 'running' in activity_type.lower():
                    data_to_export.update({
                        'Средний темп (мин/км)': format_pace(activity.get('averageSpeed')),
                        'Средняя каденс (шаги/мин)': summary.get('averageRunningCadenceInStepsPerMinute'),
                    })
                
                # Записываем все доступные данные
                for key, value in data_to_export.items():
                    if value is not None and value != '':
                        batch.add_update(row, 1, key)
                        batch.add_update(row, 2, str(value))
                        print(f"  {key}: {value}")
                        row += 1
                
            except Exception as e:
                batch.add_update(row, 1, f"Ошибка: {str(e)}")
                row += 1
            
            # Пустая строка между тренировками
            row += 1
        
        # Отправляем все обновления одним запросом
        batch.flush()
        
        print(f"\n✓ Выгружено {len(activities)} тренировок на лист 'исходник'")
        print("="*60)
        
    except Exception as e:
        print(f"✗ Ошибка при выгрузке: {e}")
        import traceback
        traceback.print_exc()

def main():
    try:
        print("=== Garmin to Google Sheets Sync ===\n")
        
        # Подключение к Garmin
        garmin = connect_to_garmin()
        
        # Подключение к Google Sheets
        sheet = connect_to_google_sheets()
        
        # ДИАГНОСТИКА: выгружаем все данные на лист "исходник"
        export_all_data_to_source(garmin, sheet)
        
        worksheet = sheet.worksheet("ВЕЛ БЕГ")
        print(f"\n✓ Opened worksheet: {worksheet.title}")
        
        # Парсим даты недель из строк блоков (20, 33, 38, 73)
        week_columns = parse_week_dates_from_block_rows(worksheet)
        print(f"✓ Найдено {len(week_columns)} недель в таблице")
        
        # Получаем тренировки за последние N дней
        days_to_sync = int(os.getenv('DAYS_TO_SYNC', '7'))  # По умолчанию 1 неделя
        activities = garmin.get_activities(0, days_to_sync * 2)  # С запасом
        
        # Группируем тренировки по неделям
        activities_by_week = {}
        for activity in activities:
            start_time = activity.get('startTimeLocal', '')
            if start_time:
                activity_date = datetime.strptime(start_time[:10], '%Y-%m-%d').date()
                column = find_column_for_date(activity_date, week_columns)
                
                if column:
                    if column not in activities_by_week:
                        activities_by_week[column] = []
                    activities_by_week[column].append(activity)
        
        # Синхронизируем ВСЕ недели с тренировками
        if activities_by_week:
            # Сортируем недели по дате
            sorted_columns = sorted(activities_by_week.keys(), key=lambda col: week_columns.get(col, datetime.min.date()))
            
            for column in sorted_columns:
                week_activities = activities_by_week[column]
                week_date = week_columns.get(column)
                
                print(f"\n{'='*60}")
                if week_date:
                    print(f"Синхронизация недели {column} (начало: {week_date.strftime('%d.%m.%Y')})")
                else:
                    print(f"Синхронизация недели {column}")
                print(f"Найдено тренировок: {len(week_activities)}")
                print(f"{'='*60}")
                
                sync_to_sheet(garmin, worksheet, column)
        else:
            print("\nℹ️  Нет тренировок для синхронизации")
        
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
