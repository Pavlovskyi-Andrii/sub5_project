#!/usr/bin/env python3
import os
import json
from pathlib import Path
from datetime import datetime, timedelta
from garminconnect import Garmin
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

def connect_to_garmin():
    """Подключение к Garmin Connect через Garth"""
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

def get_cycling_activities(garmin_client, days=7):
    """Получение велосипедных тренировок за последние дни"""
    activities = garmin_client.get_activities(0, days * 5)
    cycling_activities = []
    
    cutoff_date = (datetime.today() - timedelta(days=days)).date()
    
    for activity in activities:
        activity_type = activity.get('activityType', {}).get('typeKey', '')
        start_time_str = activity.get('startTimeLocal', '')
        
        if start_time_str:
            activity_date = datetime.strptime(start_time_str.split()[0], '%Y-%m-%d').date()
            if activity_date >= cutoff_date and 'cycling' in activity_type.lower():
                cycling_activities.append(activity)
    
    return cycling_activities

def get_running_activities(garmin_client, days=7):
    """Получение беговых тренировок за последние дни"""
    activities = garmin_client.get_activities(0, days * 5)
    running_activities = []
    
    cutoff_date = (datetime.today() - timedelta(days=days)).date()
    
    for activity in activities:
        activity_type = activity.get('activityType', {}).get('typeKey', '')
        start_time_str = activity.get('startTimeLocal', '')
        
        if start_time_str:
            activity_date = datetime.strptime(start_time_str.split()[0], '%Y-%m-%d').date()
            if activity_date >= cutoff_date and 'running' in activity_type.lower():
                running_activities.append(activity)
    
    return running_activities

def format_time(seconds):
    """Форматирование времени из секунд"""
    if not seconds:
        return ""
    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    secs = int(seconds % 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"

def format_pace(speed_mps):
    """Форматирование темпа из м/с в мин/км"""
    if not speed_mps or speed_mps == 0:
        return ""
    pace_min_per_km = 1000 / (speed_mps * 60)
    minutes = int(pace_min_per_km)
    seconds = int((pace_min_per_km - minutes) * 60)
    return f"{minutes}:{seconds:02d}"

def add_cycling_data(sheet, activity, garmin_client):
    """Добавление данных велосипеда в таблицу"""
    try:
        cycling_sheet = sheet.worksheet("Вел")
    except:
        print("Sheet 'Вел' not found, creating it...")
        cycling_sheet = sheet.add_worksheet(title="Вел", rows=100, cols=15)
        headers = ["Дата", "Средние ваты", "Normalized Power", "Сред.скор", 
                   "Частота вращения", "Средняя ЧСС", "Субъективное ощущение",
                   "Бег брик", "ЧСС бег", "TTV dist", "TTV time", "ActivityID"]
        cycling_sheet.append_row(headers)
    
    activity_id = str(activity['activityId'])
    
    try:
        existing_ids = cycling_sheet.col_values(12)[1:]
        if activity_id in existing_ids:
            date_str = activity.get('startTimeLocal', '').split()[0]
            print(f"Skipping cycling activity from {date_str} (already exists)")
            return
    except:
        pass
    
    details = garmin_client.get_activity(int(activity_id))
    date_str = activity.get('startTimeLocal', '').split()[0]
    
    avg_power = details.get('avgPower', '')
    normalized_power = details.get('normalizedPower', '')
    avg_speed = round(details.get('averageSpeed', 0) * 3.6, 2) if details.get('averageSpeed') else ''
    avg_cadence = details.get('averageBikingCadenceInRevPerMinute', '')
    avg_hr = details.get('averageHR', '')
    
    row = [
        date_str,
        avg_power,
        normalized_power,
        avg_speed,
        avg_cadence,
        avg_hr,
        '',
        '',
        '',
        '',
        '',
        activity_id
    ]
    
    cycling_sheet.append_row(row)
    print(f"✓ Added cycling activity from {date_str}")

def add_running_data(sheet, activity, garmin_client):
    """Добавление данных бега в таблицу"""
    try:
        running_sheet = sheet.worksheet("Бег")
    except:
        print("Sheet 'Бег' not found, creating it...")
        running_sheet = sheet.add_worksheet(title="Бег", rows=100, cols=15)
        headers = ["Дата", "Время", "Расстояние", "Средний темп", "Средняя ЧСС",
                   "Усталость в течении дня", "Во время тренировки", 
                   "После тренировки", "TTR dist", "TTR time", "Вариабельность СР", "ActivityID"]
        running_sheet.append_row(headers)
    
    activity_id = str(activity['activityId'])
    
    try:
        existing_ids = running_sheet.col_values(12)[1:]
        if activity_id in existing_ids:
            date_str = activity.get('startTimeLocal', '').split()[0]
            print(f"Skipping running activity from {date_str} (already exists)")
            return
    except:
        pass
    
    details = garmin_client.get_activity(int(activity_id))
    date_str = activity.get('startTimeLocal', '').split()[0]
    
    duration = format_time(details.get('duration', 0))
    distance = round(details.get('distance', 0) / 1000, 2) if details.get('distance') else ''
    avg_speed = details.get('averageSpeed', 0)
    avg_pace = format_pace(avg_speed) if avg_speed else ''
    avg_hr = details.get('averageHR', '')
    
    row = [
        date_str,
        duration,
        distance,
        avg_pace,
        avg_hr,
        '',
        '',
        '',
        '',
        '',
        '',
        activity_id
    ]
    
    running_sheet.append_row(row)
    print(f"✓ Added running activity from {date_str}")

def main():
    try:
        print("=== Garmin to Google Sheets Sync ===\n")
        
        garmin_client = connect_to_garmin()
        sheet = connect_to_google_sheets()
        
        days_to_sync = int(os.getenv('DAYS_TO_SYNC', 7))
        
        print(f"\nFetching cycling activities from last {days_to_sync} days...")
        cycling_activities = get_cycling_activities(garmin_client, days_to_sync)
        print(f"Found {len(cycling_activities)} cycling activities")
        
        for activity in cycling_activities:
            add_cycling_data(sheet, activity, garmin_client)
        
        print(f"\nFetching running activities from last {days_to_sync} days...")
        running_activities = get_running_activities(garmin_client, days_to_sync)
        print(f"Found {len(running_activities)} running activities")
        
        for activity in running_activities:
            add_running_data(sheet, activity, garmin_client)
        
        print("\n✓ Sync completed successfully!")
        
    except Exception as e:
        print(f"\n✗ Error: {str(e)}")
        raise

if __name__ == "__main__":
    main()
