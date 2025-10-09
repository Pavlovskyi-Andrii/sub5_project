import os
import json
from datetime import datetime
from garminconnect import Garmin
from dotenv import load_dotenv

load_dotenv()

email = os.getenv('GARMIN_EMAIL')
password = os.getenv('GARMIN_PASSWORD')

client = Garmin(email, password)
client.login()

# Берем одну из велотренировок субботы
activity_id = 20589014181  # "2 по 185 (15 м ) Sweet spot 1час"

details = client.get_activity(activity_id)

print(f"Все поля тренировки ID {activity_id}:\n")
print(json.dumps(details, indent=2, ensure_ascii=False))
