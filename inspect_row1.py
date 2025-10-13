import os
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Google Sheets подключение
service_account_info = json.loads(os.environ['SERVICE_ACCOUNT_JSON'])
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)

# Открываем таблицу
sheet_url = os.environ['GOOGLE_SHEET_URL']
spreadsheet = client.open_by_url(sheet_url)
worksheet = spreadsheet.worksheet('ВЕЛ БЕГ')

# Читаем строку 1
row1 = worksheet.row_values(1)

print("=" * 60)
print("СТРОКА 1 - ДАТЫ НАЧАЛА НЕДЕЛЬ:")
print("=" * 60)
for idx, cell in enumerate(row1):
    col_letter = chr(65 + idx)  # A, B, C, D, E, F, G...
    if cell.strip():
        print(f"{col_letter}: {cell}")

print("\n" + "=" * 60)
print("СТРОКИ 32-34 (ПОНЕДЕЛЬНИК):")
print("=" * 60)
for row_num in [32, 33, 34]:
    row_data = worksheet.row_values(row_num)
    if row_data:
        print(f"Строка {row_num}: {row_data[0]}")
