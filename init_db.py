#!/usr/bin/env python3
"""
Database initialization script for deployment
Run this before starting the application to create database tables
"""
import sqlite3
import os

DB_PATH = 'training_data.db'

def init_db():
    """Initialize SQLite database for storing training data"""
    print("Initializing database...")

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

    print(f"✅ Database initialized successfully at {DB_PATH}")
    print(f"   - activities table created")
    print(f"   - sync_logs table created")
    print(f"   - weekly_stats table created")

if __name__ == '__main__':
    init_db()
