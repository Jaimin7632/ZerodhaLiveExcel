# live_tick_excel_app/config.py

# --- Kite Connect API Credentials ---
API_KEY = "<KEY>"
API_SECRET = "<SECRET>"

# --- Excel Configuration for Instrument List Input ---
EXCEL_INPUT_FILE_PATH = 'instruments_to_track.xlsx' # File where you list symbols
EXCEL_INPUT_SHEET_NAME = 'Symbols'                  # Sheet name in input file
EXCEL_INPUT_HEADER = ['Instrument_Name']# Expected headers in input file
EXCEL_INPUT_DATA_START_ROW = 1                      # Data starts from row 2

# --- Server Configuration ---
FLASK_HOST = '127.0.0.1'
FLASK_PORT = 5000

# --- Data Buffer Configuration ---
# Maximum ticks to store in memory for each symbol.
# For 'last price only' display, we technically only need the latest, so 1 is fine.
# If you wanted to serve a history, this would be higher.
MAX_TICKS_IN_MEMORY_PER_SYMBOL = 1