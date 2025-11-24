#!/usr/bin/env python3
"""
Road to SUB5 Web UI - Flask application for visualizing training data
"""
import os
import json
from datetime import datetime, timedelta
from flask import Flask, render_template, jsonify, request, session
from flask_cors import CORS
import sqlite3
from dotenv import load_dotenv
from garminconnect import Garmin
import gspread
from google.oauth2.service_account import Credentials
import threading
import logging
from collections import defaultdict

# Import functions from main.py
from main import (
    connect_to_garmin,
    connect_to_google_sheets,
    get_activities_for_date,
    parse_date,
    format_time,
    format_pace,
    process_cycling_data,
    process_running_data,
    get_training_blocks,
    get_activity_hr_zones
)

load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'your-secret-key-here')
CORS(app)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Database setup
DB_PATH = 'training_data.db'

def init_db():
    """Initialize SQLite database for storing training data"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Create activities table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS activities (
            id TEXT PRIMARY KEY,
            date DATE,
            type TEXT,
            name TEXT,
            duration INTEGER,
            distance REAL,
            avg_speed REAL,
            avg_hr INTEGER,
            avg_power INTEGER,
            normalized_power INTEGER,
            avg_cadence INTEGER,
            tss INTEGER,
            calories INTEGER,
            hr_zones JSON,
            data JSON,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Add hr_zones column if it doesn't exist (migration)
    try:
        cursor.execute('ALTER TABLE activities ADD COLUMN hr_zones JSON')
        conn.commit()
    except sqlite3.OperationalError:
        # Column already exists
        pass
    
    # Create sync_logs table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sync_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sync_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            status TEXT,
            activities_synced INTEGER,
            error_message TEXT,
            details JSON
        )
    ''')
    
    # Create weekly_stats table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS weekly_stats (
            week_start DATE PRIMARY KEY,
            week_end DATE,
            total_cycling_km REAL,
            total_cycling_time INTEGER,
            total_running_km REAL,
            total_running_time INTEGER,
            avg_hrv REAL,
            total_activities INTEGER,
            data JSON,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    conn.commit()
    conn.close()

def save_activity_to_db(activity_data):
    """Save activity data to database"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''
        INSERT OR REPLACE INTO activities 
        (id, date, type, name, duration, distance, avg_speed, avg_hr, 
         avg_power, normalized_power, avg_cadence, tss, calories, data, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
    ''', (
        activity_data.get('activityId'),
        activity_data.get('startTimeLocal', '')[:10],
        activity_data.get('activityType', {}).get('typeKey', ''),
        activity_data.get('activityName', ''),
        activity_data.get('duration'),
        activity_data.get('distance'),
        activity_data.get('averageSpeed'),
        activity_data.get('averageHR'),
        activity_data.get('avgPower'),
        activity_data.get('normalizedPower'),
        activity_data.get('averageBikingCadenceInRevPerMinute'),
        activity_data.get('trainingStressScore'),
        activity_data.get('calories'),
        json.dumps(activity_data)
    ))
    
    conn.commit()
    conn.close()

def log_sync(status, activities_synced=0, error_message=None, details=None):
    """Log synchronization attempt"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''
        INSERT INTO sync_logs (status, activities_synced, error_message, details)
        VALUES (?, ?, ?, ?)
    ''', (status, activities_synced, error_message, json.dumps(details) if details else None))
    
    conn.commit()
    conn.close()

@app.route('/')
def index():
    """Main dashboard"""
    return render_template('dashboard.html')

@app.route('/weekly-calendar')
def weekly_calendar():
    """Weekly calendar view"""
    return render_template('weekly_calendar.html')

@app.route('/api/activities')
def get_activities():
    """Get activities from database"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Get query parameters
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    activity_type = request.args.get('type')
    limit = request.args.get('limit', 100)
    
    query = 'SELECT * FROM activities WHERE 1=1'
    params = []
    
    if start_date:
        query += ' AND date >= ?'
        params.append(start_date)
    
    if end_date:
        query += ' AND date <= ?'
        params.append(end_date)
    
    if activity_type:
        query += ' AND type LIKE ?'
        params.append(f'%{activity_type}%')
    
    query += ' ORDER BY date DESC LIMIT ?'
    params.append(limit)
    
    cursor.execute(query, params)
    columns = [description[0] for description in cursor.description]
    activities = []
    for row in cursor.fetchall():
        activity = dict(zip(columns, row))
        if activity['data']:
            activity['data'] = json.loads(activity['data'])
        activities.append(activity)
    
    conn.close()
    return jsonify(activities)

@app.route('/api/weekly-calendar')
def get_weekly_calendar():
    """Get weekly calendar view with all activities and HR zones

    Query params:
        week_start: Start date of the week (YYYY-MM-DD), defaults to current week's Monday
    """
    try:
        # Get week start date from query params or default to current week's Monday
        week_start_str = request.args.get('week_start')
        if week_start_str:
            week_start = datetime.strptime(week_start_str, '%Y-%m-%d').date()
        else:
            today = datetime.now().date()
            # Get Monday of current week (0 = Monday, 6 = Sunday)
            week_start = today - timedelta(days=today.weekday())

        # Calculate week end (Sunday)
        week_end = week_start + timedelta(days=6)

        # Connect to Garmin
        garmin_client = connect_to_garmin()
        if not garmin_client:
            return jsonify({'error': 'Failed to connect to Garmin'}), 500

        # Get all activities for the week
        all_activities = garmin_client.get_activities(0, 200)

        # Structure to hold weekly data
        weekly_data = {
            'week_start': week_start.isoformat(),
            'week_end': week_end.isoformat(),
            'days': []
        }

        # Iterate through each day of the week
        for day_offset in range(7):
            current_date = week_start + timedelta(days=day_offset)
            day_activities = get_activities_for_date(garmin_client, current_date, all_activities)

            # Process each activity for the day
            activities_data = []
            for activity in day_activities:
                activity_id = activity['activityId']
                activity_type = activity.get('activityType', {}).get('typeKey', 'unknown')

                # Get activity details
                try:
                    details = garmin_client.get_activity(activity_id)
                    summary = details.get('summaryDTO', {})

                    # Get HR zones
                    hr_zones = get_activity_hr_zones(garmin_client, activity_id)

                    # Build activity data
                    activity_data = {
                        'id': activity_id,
                        'name': activity.get('activityName', ''),
                        'type': activity_type,
                        'duration': summary.get('duration', 0),
                        'distance': summary.get('distance', 0) / 1000 if summary.get('distance') else 0,  # Convert to km
                        'avg_hr': summary.get('averageHR', None),
                        'max_hr': summary.get('maxHR', None),
                        'avg_power': summary.get('averagePower', None),
                        'normalized_power': summary.get('normalizedPower', None),
                        'avg_cadence': summary.get('averageBikeCadence') or summary.get('averageRunCadence', None),
                        'tss': summary.get('trainingStressScore', None),
                        'calories': summary.get('calories', None),
                        'avg_speed': summary.get('averageSpeed', None),
                        'hr_zones': hr_zones,
                        'start_time': activity.get('startTimeLocal', '')
                    }

                    activities_data.append(activity_data)

                except Exception as e:
                    logger.error(f"Error processing activity {activity_id}: {e}")
                    continue

            # Calculate daily summary
            total_duration = sum(a['duration'] for a in activities_data)
            total_distance = sum(a['distance'] for a in activities_data)
            avg_hr = sum(a['avg_hr'] for a in activities_data if a['avg_hr']) / len([a for a in activities_data if a['avg_hr']]) if activities_data else 0

            day_data = {
                'date': current_date.isoformat(),
                'day_name': current_date.strftime('%A'),
                'activities': activities_data,
                'summary': {
                    'total_duration': total_duration,
                    'total_distance': round(total_distance, 2),
                    'avg_hr': round(avg_hr) if avg_hr else None,
                    'activity_count': len(activities_data)
                }
            }

            weekly_data['days'].append(day_data)

        return jsonify(weekly_data)

    except Exception as e:
        logger.error(f"Error in weekly calendar endpoint: {e}", exc_info=True)
        return jsonify({'error': str(e)}), 500

@app.route('/api/weekly-stats')
def get_weekly_stats():
    """Get weekly statistics"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''
        SELECT * FROM weekly_stats 
        ORDER BY week_start DESC 
        LIMIT 12
    ''')
    
    columns = [description[0] for description in cursor.description]
    stats = []
    for row in cursor.fetchall():
        stat = dict(zip(columns, row))
        if stat['data']:
            stat['data'] = json.loads(stat['data'])
        stats.append(stat)
    
    conn.close()
    return jsonify(stats)

@app.route('/api/sync-logs')
def get_sync_logs():
    """Get synchronization logs"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''
        SELECT * FROM sync_logs 
        ORDER BY sync_date DESC 
        LIMIT 50
    ''')
    
    columns = [description[0] for description in cursor.description]
    logs = []
    for row in cursor.fetchall():
        log = dict(zip(columns, row))
        if log['details']:
            log['details'] = json.loads(log['details'])
        logs.append(log)
    
    conn.close()
    return jsonify(logs)

@app.route('/api/sync', methods=['POST'])
def sync_data():
    """Trigger data synchronization from Garmin"""
    try:
        # Run sync in background thread
        thread = threading.Thread(target=perform_sync)
        thread.start()

        return jsonify({'status': 'started', 'message': 'Синхронизация запущена в фоновом режиме'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/sync-google-sheets', methods=['POST'])
def sync_google_sheets_only():
    """Trigger Google Sheets synchronization only (without Garmin sync)"""
    try:
        # Run sync in background thread
        thread = threading.Thread(target=perform_google_sheets_sync)
        thread.start()

        return jsonify({'status': 'started', 'message': 'Синхронизация с Google Таблицей запущена'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/sync-google-sheets-full', methods=['POST'])
def sync_google_sheets_full():
    """Trigger FULL Google Sheets synchronization from 13.09.2025"""
    try:
        # Run full sync in background thread
        thread = threading.Thread(target=perform_full_google_sheets_sync)
        thread.start()

        return jsonify({'status': 'started', 'message': 'Полная синхронизация таблицы с 13.09.2025 запущена'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

def perform_sync():
    """Perform the actual synchronization"""
    try:
        logger.info("Starting synchronization...")

        # Connect to Garmin
        garmin = connect_to_garmin()

        # Get activities (увеличили лимит для надежности)
        days = int(os.getenv('DAYS_TO_SYNC', '14'))
        # Берем больше тренировок, чтобы охватить весь период
        # Обычно 5-10 тренировок в неделю, за 14 дней это ~20-40 тренировок
        # Увеличиваем до 200 для надежности
        activities = garmin.get_activities(0, 200)

        activities_saved = 0
        for activity in activities:
            try:
                # Get detailed activity data
                activity_id = activity['activityId']
                details = garmin.get_activity(activity_id)
                summary = details.get('summaryDTO', {})

                # Merge data
                activity_full = {**activity, **summary}

                # Save to database
                save_activity_to_db(activity_full)
                activities_saved += 1
            except Exception as e:
                logger.error(f"Error saving activity {activity.get('activityName')}: {e}")

        # Calculate weekly stats
        calculate_and_save_weekly_stats()

        # Sync to Google Sheets
        try:
            logger.info("Starting Google Sheets synchronization...")
            sync_to_google_sheets(garmin, activities)
            logger.info("Google Sheets synchronization complete.")
        except Exception as e:
            logger.error(f"Google Sheets sync failed: {e}")
            # Continue even if Google Sheets sync fails

        # Log successful sync
        log_sync('success', activities_saved, details={'days_synced': days})
        logger.info(f"Synchronization complete. Saved {activities_saved} activities.")

    except Exception as e:
        logger.error(f"Sync failed: {e}")
        log_sync('error', 0, str(e))

def perform_google_sheets_sync():
    """Perform Google Sheets synchronization using data from database"""
    try:
        logger.info("Starting Google Sheets synchronization from database...")

        # Get activities from database
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        # Get recent activities (last 30 days to cover multiple weeks)
        cursor.execute('''
            SELECT data FROM activities
            WHERE date >= date('now', '-30 days')
            ORDER BY date DESC
        ''')

        activities = []
        for row in cursor.fetchall():
            if row[0]:
                activity_data = json.loads(row[0])
                activities.append(activity_data)

        conn.close()

        if not activities:
            logger.warning("No activities found in database to sync")
            log_sync('error', 0, 'No activities in database')
            return

        logger.info(f"Found {len(activities)} activities in database")

        # Connect to Garmin (needed for sync_to_sheet function)
        garmin = connect_to_garmin()

        # Sync to Google Sheets
        sync_to_google_sheets(garmin, activities)

        # Log successful sync
        log_sync('success', len(activities), details={'source': 'database', 'type': 'google_sheets_only'})
        logger.info(f"Google Sheets synchronization complete. Synced {len(activities)} activities.")

    except Exception as e:
        logger.error(f"Google Sheets sync failed: {e}")
        log_sync('error', 0, str(e))

def perform_full_google_sheets_sync():
    """Perform FULL Google Sheets synchronization from 13.09.2024"""
    try:
        logger.info("Starting FULL Google Sheets synchronization from 13.09.2024...")

        # Connect to Garmin
        garmin = connect_to_garmin()

        # Получаем ВСЕ тренировки с 13.09.2024
        # Рассчитываем количество дней с 13.09.2024 до сегодня
        start_date = datetime(2024, 9, 13).date()
        today = datetime.now().date()
        days_diff = (today - start_date).days

        logger.info(f"Fetching activities from {start_date.strftime('%d.%m.%Y')} ({days_diff} days)")

        # Получаем тренировки - увеличиваем лимит до 800
        # За 11 недель (~77 дней) при 7-10 тренировок в неделю = ~77-110 тренировок
        # Берем с запасом 800 чтобы точно все получить
        max_activities = max(800, days_diff * 3)
        logger.info(f"Requesting up to {max_activities} activities from Garmin...")
        activities = garmin.get_activities(0, max_activities)

        # Фильтруем только те, что с 13.09.2024
        filtered_activities = []
        for activity in activities:
            if not isinstance(activity, dict):
                continue

            start_time = activity.get('startTimeLocal', '')
            if start_time:
                activity_date = datetime.strptime(start_time[:10], '%Y-%m-%d').date()
                if activity_date >= start_date:
                    filtered_activities.append(activity)
                    logger.info(f"  + Activity: {activity_date.strftime('%d.%m.%Y')} - {activity.get('activityName', 'No name')[:40]}")

        logger.info(f"\n{'='*60}")
        logger.info(f"Total activities fetched from Garmin: {len(activities)}")
        logger.info(f"Filtered activities since {start_date.strftime('%d.%m.%Y')}: {len(filtered_activities)}")
        logger.info(f"{'='*60}\n")

        if not filtered_activities:
            logger.warning("No activities found to sync")
            log_sync('error', 0, 'No activities found')
            return

        # Sync to Google Sheets
        sync_to_google_sheets(garmin, filtered_activities)

        # Log successful sync
        log_sync('success', len(filtered_activities), details={'source': 'garmin', 'type': 'full_sync', 'start_date': '13.09.2024'})
        logger.info(f"FULL Google Sheets synchronization complete. Synced {len(filtered_activities)} activities.")

    except Exception as e:
        logger.error(f"FULL Google Sheets sync failed: {e}")
        import traceback
        traceback.print_exc()
        log_sync('error', 0, str(e))

def sync_to_google_sheets(garmin, activities):
    """Sync data to Google Sheets using logic from main.py"""
    from datetime import datetime, timedelta
    from main import (
        parse_week_dates_from_block_rows,
        find_column_for_date,
        get_training_blocks,
        sync_to_sheet
    )

    # Connect to Google Sheets
    logger.info("Connecting to Google Sheets...")
    sheet = connect_to_google_sheets()
    worksheet = sheet.worksheet("ВЕЛ БЕГ")
    logger.info("Connected to worksheet: ВЕЛ БЕГ")

    # Parse week dates from block rows
    logger.info("Parsing week dates from row 20...")
    week_columns = parse_week_dates_from_block_rows(worksheet)
    logger.info(f"Found {len(week_columns)} weeks in spreadsheet:")
    for col, date in sorted(week_columns.items()):
        saturday = date - timedelta(days=1)
        friday = saturday + timedelta(days=6)
        logger.info(f"  Column {col}: {saturday.strftime('%d.%m.%Y')} (Sat) - {friday.strftime('%d.%m.%Y')} (Fri)")

    # Group activities by week
    logger.info(f"Grouping {len(activities)} activities by week...")
    activities_by_week = {}
    activities_without_column = []

    for activity in activities:
        if not isinstance(activity, dict):
            continue

        start_time = activity.get('startTimeLocal', '')
        if start_time:
            activity_date = datetime.strptime(start_time[:10], '%Y-%m-%d').date()
            activity_name = activity.get('activityName', 'No name')
            activity_type = activity.get('activityType', {}).get('typeKey', 'unknown')

            column = find_column_for_date(activity_date, week_columns)

            if column:
                if column not in activities_by_week:
                    activities_by_week[column] = []
                activities_by_week[column].append(activity)
                logger.info(f"  ✓ {activity_date.strftime('%d.%m.%Y')}: {activity_name[:40]} → Column {column}")
            else:
                activities_without_column.append({
                    'date': activity_date,
                    'name': activity_name,
                    'type': activity_type
                })
                logger.warning(f"  ✗ {activity_date.strftime('%d.%m.%Y')}: {activity_name[:40]} → NO COLUMN FOUND")

    if activities_without_column:
        logger.warning(f"WARNING: {len(activities_without_column)} activities without column:")
        for act in activities_without_column[:5]:
            logger.warning(f"  - {act['date'].strftime('%d.%m.%Y')}: {act['name'][:40]}")

    # Sync all weeks with activities
    if activities_by_week:
        logger.info(f"Syncing {len(activities_by_week)} weeks to Google Sheets...")

        # Get training blocks once for optimization
        training_blocks = get_training_blocks(worksheet)

        # Sort columns by date
        sorted_columns = sorted(activities_by_week.keys(),
                              key=lambda col: week_columns.get(col, datetime.min.date()))

        for column in sorted_columns:
            week_activities = activities_by_week[column]
            week_date = week_columns.get(column)

            logger.info(f"Syncing column {column}: {len(week_activities)} activities ({week_date.strftime('%d.%m.%Y') if week_date else 'unknown'})")

            # Sync to sheet
            sync_to_sheet(garmin, worksheet, column,
                        week_start_date=week_date,
                        training_blocks=training_blocks,
                        week_activities=week_activities)

            logger.info(f"✓ Column {column} synced successfully")
    else:
        logger.warning("No activities to sync to Google Sheets!")

def calculate_and_save_weekly_stats():
    """Calculate and save weekly statistics"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Get all activities grouped by week
    cursor.execute('''
        SELECT 
            strftime('%Y-%W', date) as week,
            type,
            SUM(distance) as total_distance,
            SUM(duration) as total_duration,
            AVG(avg_hr) as avg_hr,
            COUNT(*) as count
        FROM activities
        WHERE date >= date('now', '-84 days')
        GROUP BY week, type
    ''')
    
    # Organize data by week
    weeks_data = defaultdict(lambda: {
        'cycling_km': 0,
        'cycling_time': 0,
        'running_km': 0,
        'running_time': 0,
        'total_activities': 0
    })
    
    for row in cursor.fetchall():
        week, activity_type, distance, duration, avg_hr, count = row
        week_data = weeks_data[week]
        
        if 'cycling' in (activity_type or '').lower():
            week_data['cycling_km'] += (distance or 0) / 1000
            week_data['cycling_time'] += (duration or 0)
        elif 'running' in (activity_type or '').lower():
            week_data['running_km'] += (distance or 0) / 1000
            week_data['running_time'] += (duration or 0)
        
        week_data['total_activities'] += count
    
    # Save weekly stats
    for week, data in weeks_data.items():
        # Convert week format to dates
        year, week_num = week.split('-')
        week_start = datetime.strptime(f'{year}-W{week_num}-1', '%Y-W%W-%w').date()
        week_end = week_start + timedelta(days=6)
        
        cursor.execute('''
            INSERT OR REPLACE INTO weekly_stats
            (week_start, week_end, total_cycling_km, total_cycling_time,
             total_running_km, total_running_time, total_activities, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        ''', (
            week_start,
            week_end,
            data['cycling_km'],
            data['cycling_time'],
            data['running_km'],
            data['running_time'],
            data['total_activities']
        ))
    
    conn.commit()
    conn.close()

@app.route('/api/summary')
def get_summary():
    """Get overall summary statistics"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Get totals for current week (last 7 days)
    cursor.execute('''
        SELECT
            COUNT(*) as total_activities,
            SUM(CASE WHEN type LIKE '%cycling%' THEN distance ELSE 0 END) / 1000 as total_cycling_km,
            SUM(CASE WHEN type LIKE '%running%' THEN distance ELSE 0 END) / 1000 as total_running_km,
            SUM(duration) as total_duration,
            AVG(avg_hr) as avg_hr,
            SUM(calories) as total_calories
        FROM activities
        WHERE date >= date('now', '-7 days')
    ''')

    week_stats = dict(zip(
        ['total_activities', 'total_cycling_km', 'total_running_km', 'total_duration', 'avg_hr', 'total_calories'],
        cursor.fetchone()
    ))
    
    # Get last sync info
    cursor.execute('''
        SELECT sync_date, status, activities_synced
        FROM sync_logs
        ORDER BY sync_date DESC
        LIMIT 1
    ''')
    
    last_sync = cursor.fetchone()
    if last_sync:
        last_sync_info = {
            'date': last_sync[0],
            'status': last_sync[1],
            'activities_synced': last_sync[2]
        }
    else:
        last_sync_info = None
    
    conn.close()

    return jsonify({
        'week_stats': week_stats,
        'last_sync': last_sync_info
    })

@app.route('/api/export/<format>')
def export_data(format):
    """Export data in various formats"""
    if format not in ['json', 'csv']:
        return jsonify({'error': 'Invalid format'}), 400
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('SELECT * FROM activities ORDER BY date DESC')
    columns = [description[0] for description in cursor.description]
    activities = cursor.fetchall()
    
    if format == 'json':
        data = []
        for row in activities:
            activity = dict(zip(columns, row))
            if activity['data']:
                activity['data'] = json.loads(activity['data'])
            data.append(activity)
        return jsonify(data)
    
    elif format == 'csv':
        import csv
        from io import StringIO
        from flask import make_response
        
        si = StringIO()
        cw = csv.writer(si)
        cw.writerow(columns)
        cw.writerows(activities)
        
        output = make_response(si.getvalue())
        output.headers["Content-Disposition"] = "attachment; filename=training_data.csv"
        output.headers["Content-type"] = "text/csv"
        return output
    
    conn.close()

# Initialize database on import (for gunicorn)
if not os.path.exists(DB_PATH):
    logger.info("Database not found, initializing...")
    init_db()

if __name__ == '__main__':
    # Initialize database if it doesn't exist
    if not os.path.exists(DB_PATH):
        init_db()

    port = int(os.getenv('PORT', 5000))
    debug = os.getenv('FLASK_ENV', 'production') == 'development'
    app.run(debug=debug, host='0.0.0.0', port=port)