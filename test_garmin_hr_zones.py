#!/usr/bin/env python3
"""
Test script to fetch HR zones and power zones from Garmin
"""
import os
import json
from datetime import datetime
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

def test_zone_methods():
    """Test different methods to get zone data"""
    client = connect_to_garmin()
    if not client:
        return

    # Get a recent activity ID
    activities = client.get_activities(0, 5)
    activity_id = activities[0]['activityId']
    activity_name = activities[0].get('activityName', 'Unknown')

    print(f"\nTesting with activity: {activity_name} (ID: {activity_id})")
    print("="*80)

    # List all available methods
    print("\n📋 Available methods in Garmin client:")
    methods = [method for method in dir(client) if not method.startswith('_') and callable(getattr(client, method))]
    zone_methods = [m for m in methods if 'zone' in m.lower() or 'hr' in m.lower() or 'power' in m.lower() or 'split' in m.lower()]

    print("\n🎯 Zone/HR/Power related methods:")
    for method in zone_methods:
        print(f"  - {method}")

    # Try get_activity_splits
    try:
        print("\n\n1️⃣ Testing get_activity_splits():")
        splits = client.get_activity_splits(activity_id)
        print(json.dumps(splits, indent=2))

        with open(f'activity_{activity_id}_splits.json', 'w', encoding='utf-8') as f:
            json.dump(splits, f, indent=2, ensure_ascii=False)
        print(f"✓ Saved to activity_{activity_id}_splits.json")
    except AttributeError:
        print("  ✗ Method get_activity_splits() not available")
    except Exception as e:
        print(f"  ✗ Error: {e}")

    # Try get_activity_hr_in_timezones
    try:
        print("\n2️⃣ Testing get_activity_hr_in_timezones():")
        hr_zones = client.get_activity_hr_in_timezones(activity_id)
        print(json.dumps(hr_zones, indent=2))

        with open(f'activity_{activity_id}_hr_zones.json', 'w', encoding='utf-8') as f:
            json.dump(hr_zones, f, indent=2, ensure_ascii=False)
        print(f"✓ Saved to activity_{activity_id}_hr_zones.json")
    except AttributeError:
        print("  ✗ Method get_activity_hr_in_timezones() not available")
    except Exception as e:
        print(f"  ✗ Error: {e}")

    # Try downloading FIT file (contains all data including zones)
    try:
        print("\n3️⃣ Testing download_activity() for FIT file:")
        fit_data = client.download_activity(activity_id, dl_fmt=client.ActivityDownloadFormat.ORIGINAL)
        output_file = f'activity_{activity_id}.fit'
        with open(output_file, 'wb') as f:
            f.write(fit_data)
        print(f"✓ FIT file saved to {output_file}")
        print(f"  Size: {len(fit_data)} bytes")
    except Exception as e:
        print(f"  ✗ Error: {e}")

    # Try getting weather data (it's in the API response sometimes)
    try:
        print("\n4️⃣ Testing get_activity_weather():")
        weather = client.get_activity_weather(activity_id)
        print(json.dumps(weather, indent=2))
    except AttributeError:
        print("  ✗ Method get_activity_weather() not available")
    except Exception as e:
        print(f"  ✗ Error: {e}")

    # Try activity evaluation (may contain zone info)
    try:
        print("\n5️⃣ Testing get_activity_evaluation():")
        evaluation = client.get_activity_evaluation(activity_id)
        print(json.dumps(evaluation, indent=2))

        with open(f'activity_{activity_id}_evaluation.json', 'w', encoding='utf-8') as f:
            json.dump(evaluation, f, indent=2, ensure_ascii=False)
        print(f"✓ Saved to activity_{activity_id}_evaluation.json")
    except AttributeError:
        print("  ✗ Method get_activity_evaluation() not available")
    except Exception as e:
        print(f"  ✗ Error: {e}")

    # Direct API call to zones endpoint
    try:
        print("\n6️⃣ Testing direct Garmin Connect API for HR zones:")
        # Garmin Connect API endpoints
        url = f"https://connect.garmin.com/modern/proxy/activity-service/activity/{activity_id}/hrTimeInZones"
        print(f"  Trying: {url}")

        # The garminconnect library uses requests under the hood
        # We can try to access the session directly
        if hasattr(client, 'session') or hasattr(client, 'req'):
            session = getattr(client, 'session', getattr(client, 'req', None))
            if session:
                response = session.get(url)
                if response.status_code == 200:
                    hr_zone_data = response.json()
                    print(json.dumps(hr_zone_data, indent=2))

                    with open(f'activity_{activity_id}_hr_zones_direct.json', 'w', encoding='utf-8') as f:
                        json.dump(hr_zone_data, f, indent=2, ensure_ascii=False)
                    print(f"✓ Saved to activity_{activity_id}_hr_zones_direct.json")
                else:
                    print(f"  ✗ Status code: {response.status_code}")
            else:
                print("  ✗ No session found in client")
        else:
            print("  ✗ Cannot access session")
    except Exception as e:
        print(f"  ✗ Error: {e}")

    # Direct API call to power zones endpoint
    try:
        print("\n7️⃣ Testing direct Garmin Connect API for Power zones:")
        url = f"https://connect.garmin.com/modern/proxy/activity-service/activity/{activity_id}/powerTimeInZones"
        print(f"  Trying: {url}")

        if hasattr(client, 'session') or hasattr(client, 'req'):
            session = getattr(client, 'session', getattr(client, 'req', None))
            if session:
                response = session.get(url)
                if response.status_code == 200:
                    power_zone_data = response.json()
                    print(json.dumps(power_zone_data, indent=2))

                    with open(f'activity_{activity_id}_power_zones_direct.json', 'w', encoding='utf-8') as f:
                        json.dump(power_zone_data, f, indent=2, ensure_ascii=False)
                    print(f"✓ Saved to activity_{activity_id}_power_zones_direct.json")
                else:
                    print(f"  ✗ Status code: {response.status_code}")
    except Exception as e:
        print(f"  ✗ Error: {e}")

if __name__ == "__main__":
    test_zone_methods()
