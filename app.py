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
    get_training_blocks
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
            data JSON,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
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

def perform_sync():
    """Perform the actual synchronization"""
    try:
        logger.info("Starting synchronization...")
        
        # Connect to Garmin
        garmin = connect_to_garmin()
        
        # Get activities
        days = int(os.getenv('DAYS_TO_SYNC', '14'))
        activities = garmin.get_activities(0, days * 2)
        
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
        
        # Log successful sync
        log_sync('success', activities_saved, details={'days_synced': days})
        logger.info(f"Synchronization complete. Saved {activities_saved} activities.")
        
    except Exception as e:
        logger.error(f"Sync failed: {e}")
        log_sync('error', 0, str(e))

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