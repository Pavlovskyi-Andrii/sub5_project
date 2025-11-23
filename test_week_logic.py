#!/usr/bin/env python3
"""
Тест для проверки правильности логики определения недели
"""
from datetime import datetime, timedelta

def test_week_logic():
    """Тестирует логику определения недели суббота-пятница"""

    print("=== Тест логики определения недели (суббота-пятница) ===\n")

    # Пример: воскресенье = 19.10.2024
    sunday_date = datetime(2024, 10, 19).date()

    # Старая логика (неправильная): понедельник-воскресенье
    print("СТАРАЯ ЛОГИКА (неправильная):")
    old_monday = sunday_date - timedelta(days=6)
    print(f"  Воскресенье: {sunday_date.strftime('%d.%m.%Y (%A)')}")
    print(f"  Неделя: {old_monday.strftime('%d.%m.%Y')} (пн) - {sunday_date.strftime('%d.%m.%Y')} (вс)")
    print(f"  Диапазон: {old_monday} <= дата <= {sunday_date}\n")

    # Новая логика (правильная): суббота-пятница
    print("НОВАЯ ЛОГИКА (правильная):")
    saturday_date = sunday_date - timedelta(days=1)  # Суббота = воскресенье - 1
    friday_date = saturday_date + timedelta(days=6)  # Пятница = суббота + 6
    print(f"  Воскресенье: {sunday_date.strftime('%d.%m.%Y (%A)')}")
    print(f"  Неделя: {saturday_date.strftime('%d.%m.%Y (%A)')} (сб) - {friday_date.strftime('%d.%m.%Y (%A)')} (пт)")
    print(f"  Диапазон: {saturday_date} <= дата <= {friday_date}\n")

    # Тестовые даты
    test_dates = [
        datetime(2024, 10, 12).date(),  # Суббота предыдущей недели
        datetime(2024, 10, 13).date(),  # Воскресенье предыдущей недели (старая неделя начинается здесь)
        datetime(2024, 10, 18).date(),  # Пятница текущей недели
        datetime(2024, 10, 19).date(),  # Воскресенье (из строки 20)
        datetime(2024, 10, 24).date(),  # Пятница (конец недели)
        datetime(2024, 10, 25).date(),  # Суббота СЛЕДУЮЩЕЙ недели
    ]

    print("ТЕСТИРОВАНИЕ ДАТ:")
    print(f"{'Дата':<25} {'Старая логика':<20} {'Новая логика':<20}")
    print("-" * 65)

    for test_date in test_dates:
        # Старая логика
        old_match = old_monday <= test_date <= sunday_date

        # Новая логика
        new_match = saturday_date <= test_date <= friday_date

        old_result = "✓ Попадает" if old_match else "✗ Не попадает"
        new_result = "✓ Попадает" if new_match else "✗ Не попадает"

        print(f"{test_date.strftime('%d.%m.%Y (%A)'):<25} {old_result:<20} {new_result:<20}")

    print("\n=== КРИТИЧЕСКИЙ ТЕСТ: Суббота 25.10 ===")
    saturday_next_week = datetime(2024, 10, 25).date()

    # Старая логика - НЕ должна попадать в текущую неделю
    old_match = old_monday <= saturday_next_week <= sunday_date
    print(f"Старая логика (пн-вс): суббота 25.10 попадает в неделю 13.10-19.10? {old_match}")
    if old_match:
        print("  ❌ ОШИБКА: суббота следующей недели попала в текущую!")
    else:
        print("  ✗ Суббота НЕ попадает (это и есть баг!)")

    # Новая логика - НЕ должна попадать в текущую неделю (18.10-24.10)
    new_match = saturday_date <= saturday_next_week <= friday_date
    print(f"\nНовая логика (сб-пт): суббота 25.10 попадает в неделю 18.10-24.10? {new_match}")
    if new_match:
        print("  ❌ ОШИБКА: суббота следующей недели попала в текущую!")
    else:
        print("  ✓ Правильно! Суббота 25.10 - это начало СЛЕДУЮЩЕЙ недели (25.10-31.10)")

    # Проверяем следующую неделю
    print("\n=== Следующая неделя ===")
    next_sunday = sunday_date + timedelta(days=7)  # 26.10
    next_saturday = next_sunday - timedelta(days=1)  # 25.10
    next_friday = next_saturday + timedelta(days=6)  # 31.10

    print(f"Воскресенье: {next_sunday.strftime('%d.%m.%Y')}")
    print(f"Неделя: {next_saturday.strftime('%d.%m.%Y')} (сб) - {next_friday.strftime('%d.%m.%Y')} (пт)")

    # Проверяем что суббота 25.10 попадает в следующую неделю
    next_week_match = next_saturday <= saturday_next_week <= next_friday
    print(f"\nСуббота 25.10 попадает в следующую неделю 25.10-31.10? {next_week_match}")
    if next_week_match:
        print("  ✓ Правильно!")
    else:
        print("  ❌ ОШИБКА!")

if __name__ == "__main__":
    test_week_logic()
