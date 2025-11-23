#!/usr/bin/env python3
"""
Диагностика парсинга дат из таблицы
"""
from datetime import datetime

def test_date_parsing():
    """Тестирует парсинг дат из строки 20"""
    import re

    # Примеры ячеек из строки 20
    test_cells = [
        "08.11.25",  # Столбец K
        "15.11.25",  # Столбец L
        "01.11.24",  # Ноябрь 2024
        "25.10.24",  # Октябрь 2024
    ]

    print("=== Тест парсинга дат ===\n")

    for cell in test_cells:
        # Текущая логика
        date_match = re.search(r'\b(\d{2})\.(\d{2})\.(\d{2})\b', cell)
        if date_match:
            day, month, year = date_match.groups()
            year_old = '20' + year  # Старая логика - всегда 20XX

            try:
                date_old = datetime.strptime(f"{day}.{month}.{year_old}", "%d.%m.%Y").date()
                print(f"Ячейка: {cell}")
                print(f"  Старая логика: {date_old.strftime('%d.%m.%Y')} ({year_old})")

                # Новая логика - определяем год умнее
                year_int = int(year)
                current_year = datetime.now().year
                current_year_short = current_year % 100  # Последние 2 цифры текущего года

                # Если год больше текущего короткого года + 5, считаем прошлым веком
                # Иначе - текущим веком
                if year_int > current_year_short + 5:
                    year_full = 1900 + year_int
                else:
                    year_full = 2000 + year_int

                date_new = datetime.strptime(f"{day}.{month}.{year_full}", "%d.%m.%Y").date()
                print(f"  Новая логика: {date_new.strftime('%d.%m.%Y')} ({year_full})")
                print()

            except Exception as e:
                print(f"  Ошибка: {e}\n")

    print("\n=== Анализ проблемы ===")
    print("Текущий год: 2024")
    print("Короткий формат: 24")
    print()
    print("Дата 08.11.25:")
    print("  Старая логика: 08.11.2025 (будущее!)")
    print("  Проблема: тренировки из 2024 не попадут в 2025!")
    print()
    print("Решение: Нужно определять год относительно текущей даты:")
    print("  - Если 25 близко к текущему году (24), то это 2025")
    print("  - Если разница большая, то это прошлый век")

if __name__ == "__main__":
    test_date_parsing()
