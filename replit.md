# Garmin Connect to Google Sheets Sync

## Overview
This Python script automates the synchronization of training data from Garmin Connect to Google Sheets. It is designed for athletes who want to collect and analyze their training data in a convenient and structured format, specifically catering to triathlon metrics. The project aims to provide detailed insights into training performance by integrating key metrics like power, cadence, pace, and heart rate directly into a customizable spreadsheet.

## User Preferences
- Язык: Русский
- Формат данных: специфичные поля для триатлона (ваты, каденс, темп и т.д.)
- Требуются поля для ручного заполнения субъективных данных

## System Architecture
The system is built around a `main.py` script that orchestrates data flow between Garmin Connect and Google Sheets.
-   **UI/UX Decisions:** The project integrates with an existing Google Sheet structure that includes specific rows for training blocks (e.g., "Лонг RUN (вс)", "Становая + плаванье (пн)") and dedicated cells for various metrics. Data is formatted without units (e.g., "5.91" instead of "5.91 км") to maintain consistency.
-   **Technical Implementations:**
    -   **Garmin Connect Integration:** Utilizes the `garminconnect` library for authentication (with session caching) and data retrieval.
    -   **Google Sheets Integration:** Employs `gspread` and `oauth2client` for secure interaction with Google Sheets via a Service Account.
    -   **Data Processing:** Includes functions for parsing dates from specific block rows, calculating weekly totals for cycling and running, and safely extracting metrics from activity data.
    -   **Optimization:** A `BatchUpdater` class is used to consolidate multiple cell updates into single batch requests, minimizing API quota consumption. Date reading is also optimized using `batch_get`.
    -   **Data Mapping:** Specific logic handles the mapping of Garmin activity data to predefined cells in the Google Sheet for various training types (bike, run, strength, swim), including special handling for combined Saturday activities (bike + brick run).
    -   **Data Conversion:** Handles conversions for pace (m/s to min/km), time (seconds to HH:MM:SS), and speed (m/s to km/h).
-   **Feature Specifications:**
    -   Automated synchronization of training data (bike, run, strength, swim).
    -   Extraction of key metrics: average watts, Normalized Power, speed, cadence, heart rate, distance, pace, and HRV (planned).
    -   Dynamic identification of training block rows and corresponding dates.
    -   Calculation of weekly distance and time totals for cycling and running.
    -   Error handling and informative logging.
    -   Secure credential management via environment variables.

## External Dependencies
-   **Garmin Connect API:** Used for fetching user activity data.
-   **Google Sheets API:** Used for reading from and writing to Google Spreadsheets.
-   **Google Cloud Platform:** Required for managing API access and Service Account credentials.
-   `garminconnect` library: Python wrapper for Garmin Connect.
-   `gspread` library: Python client for Google Sheets API.
-   `oauth2client` library: OAuth2 client for Google APIs.
-   `python-dotenv` library: For loading environment variables.