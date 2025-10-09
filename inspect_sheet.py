import os
import json
import gspread
from google.oauth2.service_account import Credentials

def inspect_sheet():
    """Изучение структуры листа ВЕЛ БЕГ"""
    
    spreadsheet_url = os.getenv('GOOGLE_SHEET_URL')
    service_account_json = os.getenv('SERVICE_ACCOUNT_JSON')
    
    if not service_account_json:
        print("Error: SERVICE_ACCOUNT_JSON not found")
        return
    
    creds_dict = json.loads(service_account_json)
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    
    try:
        sheet = client.open_by_url(spreadsheet_url)
        print(f"✓ Connected to sheet: {sheet.title}")
        print(f"\nAvailable worksheets:")
        for ws in sheet.worksheets():
            print(f"  - {ws.title}")
        
        # Ищем лист "ВЕЛ БЕГ"
        worksheet = None
        for ws in sheet.worksheets():
            if "ВЕЛ" in ws.title.upper() and "БЕГ" in ws.title.upper():
                worksheet = ws
                break
        
        if not worksheet:
            print("\n❌ Worksheet 'ВЕЛ БЕГ' not found!")
            return
        
        print(f"\n✓ Found worksheet: {worksheet.title}")
        
        # Читаем столбец E (колонка 5)
        print(f"\n📋 Column E content (first 50 rows):")
        col_e = worksheet.col_values(5)
        
        for i, value in enumerate(col_e[:50], 1):
            if value.strip():
                print(f"  E{i}: {value}")
        
        # Ищем даты в столбце E
        print(f"\n📅 Dates found in column E:")
        for i, value in enumerate(col_e, 1):
            if value.strip():
                # Проверяем формат даты (DD.MM.YY или DD.MM.YYYY)
                if '.' in value and len(value.split('.')) >= 2:
                    print(f"  E{i}: {value}")
        
        # Показываем область вокруг примера (E31)
        print(f"\n📊 Area around E31 (example from user):")
        for row in range(28, 38):
            try:
                values = worksheet.row_values(row)
                if len(values) >= 5:
                    print(f"  Row {row}: E={values[4] if len(values) > 4 else ''}")
            except:
                pass
                
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    inspect_sheet()
