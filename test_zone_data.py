#!/usr/bin/env python3
"""
Test script to check what zone data is available from Garmin API
"""
import os
import json
from datetime import datetime, timedelta
from dotenv import load_dotenv
from garminconnect import Garmin

load_dotenv()

def connect_to_garmin():
    """Connect to Garmin"""
    email = os.getenv('GARMIN_EMAIL')
    password = os.getenv('GARMIN_PASSWORD')

    try:
        client = Garmin(email, password)
        client.login()
        print("✓ Successfully connected to Garmin")
        return client
    except Exception as e:
        print(f"✗ Failed to connect to Garmin: {e}")
        return None

def test_zone_data():
    """Test what zone data is available from recent activities"""
    client = connect_to_garmin()
    if not client:
        return

    print("\n" + "="*80)
    print("Fetching recent activities...")
    print("="*80)

    # Get recent activities
    activities = client.get_activities(0, 10)

    for i, activity in enumerate(activities[:3], 1):
        activity_id = activity['activityId']
        activity_name = activity.get('activityName', 'Unknown')
        activity_type = activity.get('activityType', {}).get('typeKey', 'Unknown')
        start_time = activity.get('startTimeLocal', 'Unknown')

        print(f"\n{'='*80}")
        print(f"Activity #{i}: {activity_name}")
        print(f"Type: {activity_type}")
        print(f"Date: {start_time}")
        print(f"Activity ID: {activity_id}")
        print(f"{'='*80}")

        # Get detailed activity data
        try:
            details = client.get_activity(activity_id)

            # Check summary data
            summary = details.get('summaryDTO', {})

            print("\n📊 Summary Data:")
            print(f"  Duration: {summary.get('duration', 0)} seconds")
            print(f"  Distance: {summary.get('distance', 0)/1000:.2f} km")
            print(f"  Avg HR: {summary.get('averageHR', 'N/A')}")
            print(f"  Max HR: {summary.get('maxHR', 'N/A')}")

            # Heart Rate Zones
            print("\n❤️  Heart Rate Zones:")
            if 'heartRateZones' in summary:
                hr_zones = summary['heartRateZones']
                print(f"  Raw data: {json.dumps(hr_zones, indent=2)}")
            else:
                print("  No heartRateZones found in summary")

            # Time in HR zones
            if 'timeInHRZone' in summary:
                time_in_zones = summary['timeInHRZone']
                print(f"\n⏱️  Time in HR Zones:")
                for zone_num, time_seconds in enumerate(time_in_zones):
                    if time_seconds > 0:
                        minutes = int(time_seconds / 60)
                        seconds = int(time_seconds % 60)
                        print(f"    Zone {zone_num}: {minutes}m {seconds}s ({time_seconds}s)")
            else:
                print("\n  No timeInHRZone data found")

            # Power Zones (for cycling)
            print("\n⚡ Power Data:")
            if 'averagePower' in summary:
                print(f"  Avg Power: {summary.get('averagePower', 'N/A')} W")
                print(f"  Normalized Power: {summary.get('normalizedPower', 'N/A')} W")
                print(f"  Max Power: {summary.get('maxPower', 'N/A')} W")

            # Time in Power zones
            if 'timeInPowerZone' in summary:
                time_in_power_zones = summary['timeInPowerZone']
                print(f"\n⏱️  Time in Power Zones:")
                for zone_num, time_seconds in enumerate(time_in_power_zones):
                    if time_seconds > 0:
                        minutes = int(time_seconds / 60)
                        seconds = int(time_seconds % 60)
                        print(f"    Zone {zone_num}: {minutes}m {seconds}s ({time_seconds}s)")
            else:
                print("  No timeInPowerZone data found")

            # Training Effect
            print("\n🎯 Training Metrics:")
            print(f"  TSS: {summary.get('trainingStressScore', 'N/A')}")
            print(f"  Aerobic TE: {summary.get('aerobicTrainingEffect', 'N/A')}")
            print(f"  Anaerobic TE: {summary.get('anaerobicTrainingEffect', 'N/A')}")

            # Check for zone distribution in activity zones
            if 'activityZones' in details:
                print(f"\n📈 Activity Zones data found:")
                print(json.dumps(details['activityZones'], indent=2))

            # Save full details for inspection
            output_file = f'activity_{activity_id}_details.json'
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(details, f, indent=2, ensure_ascii=False)
            print(f"\n💾 Full activity details saved to: {output_file}")

        except Exception as e:
            print(f"✗ Error getting details for activity {activity_id}: {e}")

        print("\n")

if __name__ == "__main__":
    test_zone_data()
