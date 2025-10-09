import os
import json
import gspread
from google.oauth2.service_account import Credentials

spreadsheet_url = os.getenv('GOOGLE_SHEET_URL')
service_account_json = os.getenv('SERVICE_ACCOUNT_JSON')

creds_dict = json.loads(service_account_json)
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
client = gspread.authorize(creds)

sheet = client.open_by_url(spreadsheet_url)
worksheet = sheet.worksheet("ВЕЛ БЕГ")

# Смотрим на область для "Короткие интервалы BIKE (вт)" - строка 31
# Дата 07.10.25 находится в E31
print("=" * 80)
print("📊 Структура для BIKE блока (строка 31 и ниже, колонка E):")
print("=" * 80)

for row_num in range(31, 42):
    row_data = worksheet.row_values(row_num)
    # Колонка A (название) и E (данные)
    col_a = row_data[0] if len(row_data) > 0 else ''
    col_e = row_data[4] if len(row_data) > 4 else ''
    
    if col_a or col_e:
        print(f"Row {row_num}: A='{col_a}' | E='{col_e}'")

# Также проверим субботу (строка 4)
print("\n" + "=" * 80)
print("📊 Структура для СУББОТЫ (строка 4 и ниже, колонка E):")
print("=" * 80)

for row_num in range(4, 15):
    row_data = worksheet.row_values(row_num)
    col_a = row_data[0] if len(row_data) > 0 else ''
    col_e = row_data[4] if len(row_data) > 4 else ''
    
    if col_a or col_e:
        print(f"Row {row_num}: A='{col_a}' | E='{col_e}'")
