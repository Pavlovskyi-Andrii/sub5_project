import os
import json
import gspread
from google.oauth2.service_account import Credentials

def inspect_detailed():
    """Детальное изучение структуры"""
    
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
    
    # Смотрим на область вокруг 05.10.25 (E19)
    print("=" * 80)
    print("📊 Структура вокруг даты 05.10.25 (строка 19):")
    print("=" * 80)
    
    for row_num in range(19, 29):
        row_data = worksheet.row_values(row_num)
        print(f"\nRow {row_num}:")
        for col_idx, value in enumerate(row_data[:10], 1):  # Первые 10 колонок
            col_letter = chr(64 + col_idx)  # A=65, B=66, etc
            if value:
                print(f"  {col_letter}{row_num}: {value}")
    
    # Смотрим на область вокруг 07.10.25 (E31)
    print("\n" + "=" * 80)
    print("📊 Структура вокруг даты 07.10.25 (строка 31):")
    print("=" * 80)
    
    for row_num in range(31, 42):
        row_data = worksheet.row_values(row_num)
        print(f"\nRow {row_num}:")
        for col_idx, value in enumerate(row_data[:10], 1):
            col_letter = chr(64 + col_idx)
            if value:
                print(f"  {col_letter}{row_num}: {value}")
    
    # Получаем заголовки, если есть
    print("\n" + "=" * 80)
    print("📋 Заголовки колонок (первые 3 строки):")
    print("=" * 80)
    for row_num in range(1, 4):
        row_data = worksheet.row_values(row_num)
        print(f"\nRow {row_num}: {row_data[:10]}")

if __name__ == "__main__":
    inspect_detailed()
