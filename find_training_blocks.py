import os
import json
import gspread
from google.oauth2.service_account import Credentials

def find_blocks():
    """Найти все блоки с тренировками"""
    
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
    
    # Читаем столбец A (заголовки тренировок)
    col_a = worksheet.col_values(1)
    
    print("=" * 80)
    print("🏃 Блоки тренировок в таблице (столбец A):")
    print("=" * 80)
    
    training_blocks = []
    
    for row_num, value in enumerate(col_a, 1):
        value = value.strip()
        # Ищем блоки с названиями тренировок (содержат RUN, BIKE, и т.д.)
        if value and any(keyword in value.upper() for keyword in ['RUN', 'BIKE', 'БЕГ', 'ВЕЛ']):
            training_blocks.append((row_num, value))
            print(f"\nRow {row_num}: {value}")
            
            # Показываем даты в этой строке
            row_data = worksheet.row_values(row_num)
            dates = []
            for col_idx, cell_value in enumerate(row_data[1:11], 2):  # Колонки B-K
                if cell_value and '.' in cell_value and len(cell_value.split('.')) >= 2:
                    col_letter = chr(64 + col_idx)
                    dates.append(f"{col_letter}{row_num}={cell_value}")
            if dates:
                print(f"  Даты: {', '.join(dates)}")
    
    print(f"\n\n📋 Всего найдено блоков: {len(training_blocks)}")
    
    return training_blocks

if __name__ == "__main__":
    find_blocks()
