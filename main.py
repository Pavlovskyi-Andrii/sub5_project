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
    
    if not email or not password:
        raise ValueError("Garmin credentials not found. Please set GARMIN_EMAIL and GARMIN_PASSWORD")
    
    print("Connecting to Garmin Connect...")
    client = Garmin(email, password)
    client.login()
    print("Successfully connected to Garmin!")
    return client

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
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        print("Connecting to Google Sheets with service account...")
    else:
        print("Error: SERVICE_ACCOUNT_JSON not found in environment variables")
        raise ValueError("SERVICE_ACCOUNT_JSON is required")
    
    client = gspread.authorize(creds)
    sheet = client.open_by_url(spreadsheet_url)
    print(f"Successfully connected to Google Sheet!")
    return sheet

def get_cycling_activities(garmin_client, days=7):
    """Получение велосипедных тренировок за последние дни"""
    activities = garmin_client.get_activities(0, days * 5)
    cycling_activities = []
    
    for activity in activities:
        activity_type = activity.get('activityType', {}).get('typeKey', '')
        if 'cycling' in activity_type.lower():
            cycling_activities.append(activity)
    
    return cycling_activities

def get_running_activities(garmin_client, days=7):
    """Получение беговых тренировок за последние дни"""
    activities = garmin_client.get_activities(0, days * 5)
    running_activities = []
    
    for activity in activities:
        activity_type = activity.get('activityType', {}).get('typeKey', '')
        if 'running' in activity_type.lower():
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
                   "Бег брик", "ЧСС бег", "TTV dist", "TTV time"]
        cycling_sheet.append_row(headers)
    
    activity_id = activity['activityId']
    details = garmin_client.get_activity(activity_id)
    
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
        ''
    ]
    
    cycling_sheet.append_row(row)
    print(f"Added cycling activity from {date_str}")

def add_running_data(sheet, activity, garmin_client):
    """Добавление данных бега в таблицу"""
    try:
        running_sheet = sheet.worksheet("Бег")
    except:
        print("Sheet 'Бег' not found, creating it...")
        running_sheet = sheet.add_worksheet(title="Бег", rows=100, cols=15)
        headers = ["Дата", "Время", "Расстояние", "Средний темп", "Средняя ЧСС",
                   "Усталость в течении дня", "Во время тренировки", 
                   "После тренировки", "TTR dist", "TTR time", "Вариабельность СР"]
        running_sheet.append_row(headers)
    
    activity_id = activity['activityId']
    details = garmin_client.get_activity(activity_id)
    
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
        ''
    ]
    
    running_sheet.append_row(row)
    print(f"Added running activity from {date_str}")

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
