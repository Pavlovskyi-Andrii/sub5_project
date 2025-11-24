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

def get_activities_for_date(garmin_client, target_date, all_activities=None):
    """Получить все тренировки за конкретную дату

    Args:
        garmin_client: Клиент Garmin
        target_date: Целевая дата
        all_activities: Опциональный список всех тренировок (для оптимизации)
    """
    # Если передан список тренировок, используем его
    if all_activities is not None:
        activities = all_activities
    else:
        # Иначе получаем тренировки из Garmin (увеличили лимит до 200)
        activities = garmin_client.get_activities(0, 200)

    # Если target_date - datetime, преобразуем в date
    if hasattr(target_date, 'date'):
        target_date = target_date.date()

    result = []
    for activity in activities:
        start_time_str = activity.get('startTimeLocal', '')
        if start_time_str:
            activity_date = datetime.strptime(start_time_str.split()[0], '%Y-%m-%d').date()
            if activity_date == target_date:
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
        if value and any(keyword in value.upper() for keyword in ['RUN', 'BIKE', 'БЕГ', 'ВЕЛ', 'ПЛАВ', 'FTP', 'ДЛИН']):
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
        'avg_hr': [],
        'tss': []
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
        tss = summary.get('trainingStressScore', '')
        
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
        if tss:
            data['tss'].append(str(int(tss)))
    
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

def get_activity_hr_zones(garmin_client, activity_id):
    """Получить данные о зонах пульса для тренировки

    Args:
        garmin_client: Клиент Garmin
        activity_id: ID тренировки

    Returns:
        list: Список зон с временем в каждой зоне
        Example: [
            {"zoneNumber": 1, "secsInZone": 84.933, "zoneLowBoundary": 97},
            {"zoneNumber": 2, "secsInZone": 199.296, "zoneLowBoundary": 116},
            ...
        ]
    """
    try:
        hr_zones = garmin_client.get_activity_hr_in_timezones(activity_id)
        return hr_zones if hr_zones else []
    except Exception as e:
        print(f"Warning: Could not fetch HR zones for activity {activity_id}: {e}")
        return []

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

def calculate_weekly_totals(garmin_client, week_activities, sunday_date, batch, col_index):
    """Подсчет недельных итогов для велосипеда и бега

    Args:
        garmin_client: Клиент Garmin
        week_activities: Список всех активностей недели
        sunday_date: Дата воскресенья из строки 20 (конец недели)
        batch: BatchUpdater для записи данных
        col_index: Индекс столбца
    """
    if not sunday_date or not week_activities:
        return

    # Неделя: ПОНЕДЕЛЬНИК - ВОСКРЕСЕНЬЕ (пн-вс)
    # Если воскресенье = 23.11, то понедельник = 17.11
    monday_date = sunday_date - timedelta(days=6)    # Понедельник = воскресенье - 6

    print(f"\n📊 Подсчет недельных итогов (пн-вс: {monday_date.strftime('%d.%m')} - {sunday_date.strftime('%d.%m')})...")

    # Фильтруем активности недели (пн-вс)
    week_filtered_activities = []
    for activity in week_activities:
        start_time_str = activity.get('startTimeLocal', '')
        if start_time_str:
            activity_date = datetime.strptime(start_time_str.split()[0], '%Y-%m-%d').date()
            if monday_date <= activity_date <= sunday_date:
                week_filtered_activities.append(activity)
    
    # Разделяем по типам
    cycling_activities = [a for a in week_filtered_activities if 'cycling' in a.get('activityType', {}).get('typeKey', '').lower()]
    running_activities = [a for a in week_filtered_activities if 'running' in a.get('activityType', {}).get('typeKey', '').lower()]
    
    # ВЕЛОСИПЕД
    total_cycling_distance = 0  # в км
    total_cycling_time = 0  # в секундах
    
    for activity in cycling_activities:
        distance = activity.get('distance')
        if distance:
            total_cycling_distance += distance / 1000  # метры -> км
        
        duration = activity.get('duration')
        if duration:
            total_cycling_time += duration
    
    # БЕГА
    total_running_distance = 0  # в км
    total_running_time = 0  # в секундах
    sunday_long_run_hrv = None
    
    # Ищем воскресные беговые тренировки для HRV
    sunday_runs = []
    for activity in running_activities:
        start_time_str = activity.get('startTimeLocal', '')
        if start_time_str:
            activity_date = datetime.strptime(start_time_str.split()[0], '%Y-%m-%d').date()
            if activity_date == sunday_date:
                sunday_runs.append(activity)
    
    for activity in running_activities:
        distance = activity.get('distance')
        if distance:
            total_running_distance += distance / 1000  # метры -> км
        
        duration = activity.get('duration')
        if duration:
            total_running_time += duration
    
    # Извлекаем HRV из воскресной Long Run
    if sunday_runs:
        try:
            # Получаем HRV данные для воскресенья
            hrv_data = garmin_client.get_hrv_data(sunday_date.strftime('%Y-%m-%d'))
            if hrv_data and 'hrvSummary' in hrv_data:
                # lastNightAvg - средняя HRV за последнюю ночь
                hrv_value = hrv_data['hrvSummary'].get('lastNightAvg')
                if hrv_value:
                    sunday_long_run_hrv = str(hrv_value)
        except Exception as e:
            print(f"  ⚠️ Не удалось получить HRV: {e}")
            pass
    
    # Форматируем данные
    # Строка 18: TVD dist (Bike) - недельный проезд в км
    if total_cycling_distance > 0:
        cycling_dist_str = f"{total_cycling_distance:.2f}"
        batch.add_update(18, col_index + 1, cycling_dist_str)
        print(f"  ✓ TVD dist (Bike): {cycling_dist_str} км → строка 18")
    
    # Строка 19: TVT time (Bike) - формат ЧЧ:ММ или ММ:СС
    if total_cycling_time > 0:
        hours = int(total_cycling_time // 3600)
        minutes = int((total_cycling_time % 3600) // 60)
        if hours > 0:
            cycling_time_str = f"{hours}:{minutes:02d}"
        else:
            seconds = int(total_cycling_time % 60)
            cycling_time_str = f"{minutes}:{seconds:02d}"
        batch.add_update(19, col_index + 1, cycling_time_str)
        print(f"  ✓ TVT time (Bike): {cycling_time_str} → строка 19")
    
    # Строка 29: TVD dist RUN - недельный пробег в км
    if total_running_distance > 0:
        running_dist_str = f"{total_running_distance:.2f}"
        batch.add_update(29, col_index + 1, running_dist_str)
        print(f"  ✓ TVD dist RUN: {running_dist_str} км → строка 29")
    
    # Строка 30: TVT time RUN
    if total_running_time > 0:
        hours = int(total_running_time // 3600)
        minutes = int((total_running_time % 3600) // 60)
        if hours > 0:
            running_time_str = f"{hours}:{minutes:02d}"
        else:
            seconds = int(total_running_time % 60)
            running_time_str = f"{minutes}:{seconds:02d}"
        batch.add_update(30, col_index + 1, running_time_str)
        print(f"  ✓ TVT time RUN: {running_time_str} → строка 30")
    
    # Строка 31: вариабельность СР из Long Run (вс)
    if sunday_long_run_hrv:
        batch.add_update(31, col_index + 1, sunday_long_run_hrv)
        print(f"  ✓ Вариабельность СР (HRV): {sunday_long_run_hrv} → строка 31")
    
    print(f"  📈 Итого вел: {total_cycling_distance:.2f} км, бег: {total_running_distance:.2f} км")

def sync_to_sheet(garmin_client, worksheet, column, week_start_date=None, training_blocks=None, week_activities=None):
    """Синхронизация данных в конкретный столбец
    
    Args:
        garmin_client: Клиент Garmin
        worksheet: Лист Google Sheets
        column: Столбец для синхронизации (A, B, C и т.д.)
        week_start_date: Дата начала недели (суббота), для блоков без даты
        training_blocks: Список блоков тренировок (для оптимизации API)
        week_activities: Список всех активностей недели (для оптимизации API)
    """
    print(f"\n{'='*60}")
    print(f"Синхронизация для столбца {column}")
    print(f"{'='*60}")
    
    # Создаем batch updater
    batch = BatchUpdater(worksheet)
    
    # Определяем индекс колонки (E = 5, значит индекс 4)
    col_index = ord(column.upper()) - ord('A')
    
    # ФИКСИРОВАННАЯ ОБРАБОТКА СУББОТЫ (строки 7-15)
    # Суббота ВСЕГДА обрабатывается если есть week_start_date
    if week_start_date:
        # ВАЖНО: week_start_date это ВОСКРЕСЕНЬЕ (из строки 20)
        # Для субботы нужно взять день раньше
        saturday_date = week_start_date - timedelta(days=1)
        print(f"\n📅 Суббота (Вел длинная + бег брик) - {saturday_date.strftime('%d.%m.%y')}")
        
        # Получаем тренировки за субботу (используем week_activities для оптимизации)
        saturday_activities = get_activities_for_date(garmin_client, saturday_date, week_activities)
        
        if saturday_activities:
            # Разделяем по типам
            cycling_activities = [a for a in saturday_activities if 'cycling' in a.get('activityType', {}).get('typeKey', '').lower()]
            running_activities = [a for a in saturday_activities if 'running' in a.get('activityType', {}).get('typeKey', '').lower()]
            
            print(f"  🚴 Велосипед: {len(cycling_activities)} тренировок")
            print(f"  🏃 Бег: {len(running_activities)} тренировок")
            
            # Обрабатываем велосипед (строки 7-11)
            if cycling_activities:
                cycle_data = process_cycling_data(garmin_client, cycling_activities[:2])
                
                def format_values(values_list):
                    if len(values_list) >= 2:
                        return f"{values_list[0]}/{values_list[1]}"
                    elif len(values_list) == 1:
                        return values_list[0]
                    return ''
                
                # Безопасное извлечение данных с проверкой длины списков
                def safe_get(data_list, index):
                    return data_list[index] if index < len(data_list) else None
                
                avg_power_list = cycle_data.get('avg_power', [])
                np_power_list = cycle_data.get('normalized_power', [])
                speed_list = cycle_data.get('avg_speed', [])
                cadence_list = cycle_data.get('avg_cadence', [])
                hr_list = cycle_data.get('avg_hr', [])
                
                avg_power = format_values([safe_get(avg_power_list, i) for i in range(min(2, len(cycling_activities)))])
                np_power = format_values([safe_get(np_power_list, i) for i in range(min(2, len(cycling_activities)))])
                speed = format_values([safe_get(speed_list, i) for i in range(min(2, len(cycling_activities)))])
                cadence = format_values([safe_get(cadence_list, i) for i in range(min(2, len(cycling_activities)))])
                hr = format_values([safe_get(hr_list, i) for i in range(min(2, len(cycling_activities)))])
                
                if avg_power:
                    batch.add_update(7, col_index + 1, avg_power)
                    print(f"  ✓ Средние ваты: {avg_power} → {column}7")
                if np_power:
                    batch.add_update(8, col_index + 1, np_power)
                    print(f"  ✓ Normalized Power: {np_power} → {column}8")
                if speed:
                    batch.add_update(9, col_index + 1, speed)
                    print(f"  ✓ Средняя скорость: {speed} → {column}9")
                if cadence:
                    batch.add_update(10, col_index + 1, cadence)
                    print(f"  ✓ Частота вращения: {cadence} → {column}10")
                if hr:
                    batch.add_update(11, col_index + 1, hr)
                    print(f"  ✓ Средняя ЧСС: {hr} → {column}11")
            
            # Обрабатываем бег брик (строки 13-15)
            if running_activities:
                run = running_activities[0]
                distance = run.get('distance', 0)
                if distance:
                    distance_km = round(distance / 1000, 2)
                    batch.add_update(13, col_index + 1, str(distance_km))
                    print(f"  ✓ Бег брик км: {distance_km} → {column}13")
                else:
                    batch.add_update(13, col_index + 1, "—")

                avg_speed = run.get('averageSpeed')
                if avg_speed:
                    pace_str = format_pace(avg_speed)
                    if pace_str:
                        batch.add_update(14, col_index + 1, pace_str)
                        print(f"  ✓ Бег брик темп: {pace_str} → {column}14")
                    else:
                        batch.add_update(14, col_index + 1, "—")
                else:
                    batch.add_update(14, col_index + 1, "—")

                avg_hr = run.get('averageHR')
                if avg_hr:
                    batch.add_update(15, col_index + 1, str(int(avg_hr)))
                    print(f"  ✓ Бег брик ЧСС: {int(avg_hr)} → {column}15")
                else:
                    batch.add_update(15, col_index + 1, "—")
            else:
                # Нет беговой тренировки - ставим прочерки
                batch.add_update(13, col_index + 1, "—")
                batch.add_update(14, col_index + 1, "—")
                batch.add_update(15, col_index + 1, "—")
                print(f"  ℹ️  Нет брик-бега в субботу → ставим прочерки в строках 13-15")
        else:
            print(f"  ℹ️  Нет тренировок за субботу {saturday_date.strftime('%d.%m.%y')}")
            # Ставим прочерки если нет тренировок вообще
            batch.add_update(13, col_index + 1, "—")
            batch.add_update(14, col_index + 1, "—")
            batch.add_update(15, col_index + 1, "—")
    
    # Используем переданные блоки тренировок или получаем их (для совместимости)
    blocks = training_blocks if training_blocks is not None else get_training_blocks(worksheet)
    
    # Для каждого блока ищем дату в нужном столбце
    for block in blocks:
        row_num = block['row']
        name = block['name']
        row_data = block['data']
        
        # Пропускаем блок субботы (строки 6-15) - он уже обработан фиксированной логикой
        if 6 <= row_num <= 15:
            continue
        
        # Определяем индекс колонки
        if col_index >= len(row_data):
            continue
        
        date_str = row_data[col_index]
        date_obj = parse_date(date_str)
        
        # СПЕЦИАЛЬНАЯ ЛОГИКА ДЛЯ СУББОТЫ: используем дату начала недели
        if not date_obj and 'сб' in name.lower() and week_start_date:
            date_obj = week_start_date
            print(f"📅 {name} - {date_obj.strftime('%d.%m.%y')} (начало недели)")
        
        # Если в строке блока нет даты, проверяем строку 1 (заголовки недель) - устаревшая логика
        elif not date_obj:
            row1 = worksheet.row_values(1)
            if col_index < len(row1):
                date_str = row1[col_index]
                date_obj = parse_date(date_str)
        
        if not date_obj:
            continue
        
        print(f"\n📅 {name} - {date_str}")
        
        # Получаем тренировки за эту дату (используем week_activities для оптимизации)
        activities = get_activities_for_date(garmin_client, date_obj, week_activities)
        
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
                tss_str = format_values(cycle_data['tss'])
                
                print(f"  📊 Данные вел: power={avg_power_str}, NP={np_str}, speed={speed_str}, cadence={cadence_str}, HR={hr_str}, TSS={tss_str}")
                
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
                
                # Строка 43: TSS (между субботними блоками и следующим блоком)
                # ВАЖНО: это строка находится ниже субботнего блока
                if tss_str:
                    batch.add_update(43, col_index + 1, tss_str)
                    print(f"  ✓ TSS: {tss_str} → {chr(64+col_index+1)}43")
            
            # Потом записываем бег брик (строки 13-15)
            # ВАЖНО: Для брик бега берем ПОСЛЕДНЮЮ беговую тренировку дня (по времени)
            # потому что брик бег всегда идет после велосипеда
            if running_activities:
                # Сортируем по времени начала и берем последнюю
                running_sorted = sorted(running_activities, key=lambda x: x.get('startTimeLocal', ''))
                brick_run = running_sorted[-1]  # Последняя по времени = брик бег
                run_data = process_running_data(garmin_client, brick_run)
                
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
        
        elif 'ВЕЛ' in name.upper() or 'BIKE' in name.upper() or 'FTP' in name.upper() or ('ЧТ' in name.upper() and 'ДЛИН' in name.upper()):
            # Это блок велосипеда
            is_ftp_block = 'FTP' in name.upper()  # Флаг для FTP
            is_thursday_block = 'ЧТ' in name.upper() and 'ДЛИН' in name.upper()  # Флаг для четверга
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
                tss_str = format_values(cycle_data['tss'])
                
                print(f"  📊 Данные вел: power={avg_power_str}, NP={np_str}, speed={speed_str}, cadence={cadence_str}, HR={hr_str}, TSS={tss_str}")
                
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
                    
                    elif 'tss' in cell_text or 'training stress' in cell_text:
                        # Строка 43: TSS
                        if tss_str:
                            batch.add_update(actual_row, col_index + 1, tss_str)
                            print(f"  ✓ TSS: {tss_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'сред' in cell_text and 'скор' in cell_text:
                        if speed_str:
                            batch.add_update(actual_row, col_index + 1, speed_str)
                            print(f"  ✓ Средняя скорость: {speed_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif 'частот' in cell_text and 'вращ' in cell_text:
                        if cadence_str:
                            batch.add_update(actual_row, col_index + 1, cadence_str)
                            print(f"  ✓ Частота вращения: {cadence_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    # Строка 42: "Средняя ЧП" - пользователь хочет сюда HR (несмотря на название)
                    # Строка 48: "средний каденс" - сюда каденс
                    elif 'каденс' in cell_text:
                        # Строка 48: средний каденс
                        if cadence_str:
                            batch.add_update(actual_row, col_index + 1, cadence_str)
                            print(f"  ✓ Средний каденс: {cadence_str} → {chr(64+col_index+1)}{actual_row}")
                    
                    elif ('средн' in cell_text or 'срадн' in cell_text) and 'чп' in cell_text:
                        # Строка 42: Средняя ЧП - пишем HR (несмотря на название)
                        if hr_str:
                            batch.add_update(actual_row, col_index + 1, hr_str)
                            print(f"  ✓ Средняя ЧП (HR): {hr_str} → {chr(64+col_index+1)}{actual_row}")
                
                # Если это FTP блок (строка 3) или четверг (строка 60), TSS записываем в строку 43
                if (is_ftp_block or is_thursday_block) and tss_str:
                    batch.add_update(43, col_index + 1, tss_str)
                    print(f"  ✓ TSS: {tss_str} → {chr(64+col_index+1)}43")
        
        elif ('СТАНОВ' in name.upper() or 'ПЛАВ' in name.upper()) and 'ПН' in name.upper():
            # Это понедельник - становая + плавание
            # Записываем длительность тренировок
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

                # ВАЖНО: строка 33 это заголовок блока, не заполняем его
                if actual_row == 33:
                    continue

                if 'длительност' in cell_text and 'перв' in cell_text:
                    # Строка 34: длительность первой тренировки
                    if len(durations) >= 1:
                        batch.add_update(actual_row, col_index + 1, durations[0])
                        print(f"  ✓ Длительность первой тренировки: {durations[0]} → {chr(64+col_index+1)}{actual_row}")
                    else:
                        batch.add_update(actual_row, col_index + 1, "—")
                        print(f"  ✓ Длительность первой тренировки: — → {chr(64+col_index+1)}{actual_row}")

                elif 'длительност' in cell_text and 'втор' in cell_text:
                    # Строка 35: длительность второй тренировки
                    if len(durations) >= 2:
                        batch.add_update(actual_row, col_index + 1, durations[1])
                        print(f"  ✓ Длительность второй тренировки: {durations[1]} → {chr(64+col_index+1)}{actual_row}")
                    else:
                        batch.add_update(actual_row, col_index + 1, "—")
                        print(f"  ✓ Длительность второй тренировки: — → {chr(64+col_index+1)}{actual_row}")
    
    # Подсчитываем недельные итоги (строки 18, 19, 29, 30, 31)
    # Параметр week_start_date на самом деле содержит sunday_date (из строки 20)
    if week_activities and week_start_date:
        calculate_weekly_totals(garmin_client, week_activities, week_start_date, batch, col_index)
    
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
    """Парсит даты из строки 20 (воскресенья) и возвращает словарь {столбец: дата_воскресенья}"""
    # Строка 20: "Лонг RUN (вс)" - содержит даты воскресений (конец недели)
    try:
        row_20 = worksheet.row_values(20)
    except Exception as e:
        print(f"✗ Ошибка при чтении строки 20: {e}")
        return {}
    
    # Словарь для хранения воскресений по столбцам: {столбец: дата_воскресенья}
    sunday_columns = {}
    
    # Парсим даты воскресений
    for idx, cell in enumerate(row_20):
        if idx < 2:  # Пропускаем столбцы A и B
            continue
        
        if not cell or not isinstance(cell, str):
            continue
        
        col_letter = chr(65 + idx)  # A, B, C, D, E, F...
        
        # Ищем дату в формате dd.mm.yy
        date_match = re.search(r'\b(\d{2})\.(\d{2})\.(\d{2})\b', cell)
        if date_match:
            try:
                day, month, year = date_match.groups()
                year = '20' + year
                sunday_date = datetime.strptime(f"{day}.{month}.{year}", "%d.%m.%Y").date()
                sunday_columns[col_letter] = sunday_date
            except:
                continue
    
    return sunday_columns

def find_column_for_date(activity_date, sunday_columns):
    """Находит столбец для записи данных на основе даты тренировки

    Неделя: суббота - пятница (сб-пт)
    Если тренировка 25.10, а воскресенье = 26.10, то суббота = 25.10, пятница = 31.10
    Значит тренировка 25.10 попадает в неделю 25.10-31.10
    """
    # Ищем подходящий столбец по воскресенью
    for col_letter, sunday_date in sunday_columns.items():
        # Суббота = воскресенье - 1 день (начало недели)
        saturday_date = sunday_date - timedelta(days=1)
        # Пятница = суббота + 6 дней (конец недели)
        friday_date = saturday_date + timedelta(days=6)

        # Проверяем попадает ли тренировка в неделю сб-пт
        if saturday_date <= activity_date <= friday_date:
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
        
        # ДИАГНОСТИКА: выгружаем все данные на лист "исходник" (раскомментируйте при необходимости)
        # export_all_data_to_source(garmin, sheet)
        
        worksheet = sheet.worksheet("ВЕЛ БЕГ")
        print(f"\n✓ Opened worksheet: {worksheet.title}")
        
        # Парсим даты недель из строк блоков (20, 33, 38, 73)
        week_columns = parse_week_dates_from_block_rows(worksheet)
        print(f"✓ Найдено {len(week_columns)} недель в таблице")
        
        # Диагностика: показываем все найденные недели
        print("\n📅 Найденные недели:")
        for col, date in sorted(week_columns.items()):
            print(f"  Столбец {col}: {date.strftime('%d.%m.%Y')}")
        
        # Получаем тренировки за последние N дней
        days_to_sync = int(os.getenv('DAYS_TO_SYNC', '14'))  # По умолчанию 30 дней
        activities = garmin.get_activities(0, days_to_sync * 2)  # С запасом
        
        # Группируем тренировки по неделям
        print("\n📊 Тренировки из Garmin:")
        activities_by_week = {}
        for activity in activities:
            # Проверяем что activity это dict
            if not isinstance(activity, dict):
                continue
                
            start_time = activity.get('startTimeLocal', '')
            if start_time:
                activity_date = datetime.strptime(start_time[:10], '%Y-%m-%d').date()
                activity_name = activity.get('activityName', 'Без названия')
                print(f"  {activity_date.strftime('%d.%m.%Y')} - {activity_name}")
                
                column = find_column_for_date(activity_date, week_columns)
                
                if column:
                    if column not in activities_by_week:
                        activities_by_week[column] = []
                    activities_by_week[column].append(activity)
                else:
                    print(f"    ⚠️ Не найден столбец для даты {activity_date.strftime('%d.%m.%Y')}")
        
        # Синхронизируем ВСЕ недели с тренировками
        if activities_by_week:
            # Получаем блоки тренировок ОДИН РАЗ для оптимизации API
            training_blocks = get_training_blocks(worksheet)
            
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
                
                # Передаем дату начала недели, блоки и активности для оптимизации API
                sync_to_sheet(garmin, worksheet, column, week_start_date=week_date, training_blocks=training_blocks, week_activities=week_activities)
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
