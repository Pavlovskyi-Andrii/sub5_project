# Road to SUB5 Web UI - Инструкция по запуску

## Описание
Flask веб-приложение для визуализации тренировочных данных из Garmin Connect.
Цель проекта - подготовка к марафону с результатом Sub-5 (менее 5 часов).

## Функции
- Синхронизация данных с Garmin Connect
- Визуализация активностей (велосипед, бег)
- Еженедельная статистика
- Экспорт данных (JSON, CSV)
- Логи синхронизации
- Локальная база данных SQLite

## Требования
- Python 3.8+
- Аккаунт Garmin Connect
- (Опционально) Google Sheets API credentials

## Установка

### 1. Создание виртуального окружения
```bash
python3 -m venv .venv
source .venv/bin/activate  # Linux/Mac
# или
.venv\Scripts\activate  # Windows
```

### 2. Установка зависимостей
```bash
pip install --upgrade pip
pip install -r requirements.txt
```

Если установка зависает на pandas/plotly, попробуйте установить их отдельно:
```bash
pip install Flask Flask-CORS garminconnect google-auth gspread python-dotenv requests
pip install pandas plotly APScheduler
```

### 3. Настройка переменных окружения

Создайте файл `.env` в корне проекта (используйте `.env.example` как шаблон):

```env
# Обязательные переменные
GARMIN_EMAIL=your_email@example.com
GARMIN_PASSWORD=your_password

# Опциональные
FLASK_SECRET_KEY=your-secret-key-here
DAYS_TO_SYNC=14

# Google Sheets (если используете)
SPREADSHEET_ID=your_spreadsheet_id
SERVICE_ACCOUNT_FILE=path/to/credentials.json
```

### 4. Проверка импортов
```bash
python3 -c "from flask import Flask; from garminconnect import Garmin; print('OK')"
```

## Запуск приложения

### Обычный режим
```bash
python3 app.py
```

Приложение будет доступно по адресу: `http://localhost:5000`

### Режим разработки
```bash
export FLASK_ENV=development  # Linux/Mac
# или
set FLASK_ENV=development  # Windows

python3 app.py
```

### Production режим (с Gunicorn)
```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

## Использование

### Главная страница
Откройте `http://localhost:5000` - вы увидите дашборд с:
- Статистикой за текущий месяц
- Последними активностями
- Еженедельной статистикой
- Графиками

### API Endpoints

#### Получить активности
```bash
GET /api/activities?start_date=2024-01-01&end_date=2024-01-31&type=cycling&limit=100
```

#### Получить еженедельную статистику
```bash
GET /api/weekly-stats
```

#### Получить общую статистику
```bash
GET /api/summary
```

#### Синхронизировать данные
```bash
POST /api/sync
```

#### Получить логи синхронизации
```bash
GET /api/sync-logs
```

#### Экспорт данных
```bash
GET /api/export/json  # Экспорт в JSON
GET /api/export/csv   # Экспорт в CSV
```

## Структура проекта

```
Road to SUB5/
├── app.py              # Flask приложение
├── main.py             # Основные функции для работы с Garmin
├── requirements.txt    # Python зависимости
├── .env                # Переменные окружения (не в git)
├── templates/
│   └── dashboard.html  # Главная страница
├── static/
│   └── css/
│       └── style.css   # Стили
└── training_data.db    # SQLite база данных (создается автоматически)
```

## База данных

База данных SQLite создается автоматически при первом запуске.

### Таблицы:
- `activities` - тренировочные активности
- `weekly_stats` - еженедельная статистика
- `sync_logs` - логи синхронизации

## Решение проблем

### Ошибка импорта модулей
```bash
# Убедитесь что виртуальное окружение активировано
source .venv/bin/activate

# Переустановите зависимости
pip install -r requirements.txt
```

### Ошибка подключения к Garmin
1. Проверьте правильность email и пароля в `.env`
2. Если появляется капча - залогиньтесь в Garmin вручную в браузере
3. Используйте SESSION_SECRET для сохранения сессии (см. вывод после первого входа)

### База данных заблокирована
```bash
# Остановите все запущенные экземпляры app.py
# Удалите файл блокировки (если есть)
rm training_data.db-journal
```

### Порт 5000 занят
```python
# В app.py измените порт:
app.run(debug=True, host='0.0.0.0', port=8080)
```

## Полезные команды

### Проверить версии зависимостей
```bash
pip list
```

### Обновить зависимости
```bash
pip install --upgrade -r requirements.txt
```

### Просмотр базы данных
```bash
sqlite3 training_data.db
.tables
SELECT * FROM activities LIMIT 5;
```

### Очистить базу данных
```bash
rm training_data.db
# При следующем запуске создастся новая
```

## Поддержка

Если возникли проблемы:
1. Проверьте логи в консоли
2. Убедитесь что все зависимости установлены
3. Проверьте файл `.env`
4. Посмотрите логи синхронизации в UI

## Лицензия

Проект для личного использования.
