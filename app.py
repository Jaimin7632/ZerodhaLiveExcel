# live_tick_excel_app/app.py
import asyncio
import io
import json

import pandas as pd
from flask import Flask, Response

from excel_handler import read_instrument_list_from_excel
import kite_data_manager  # Import the data manager module
from config import FLASK_HOST, FLASK_PORT

app = Flask(__name__)


# --- Flask Route to serve CSV data for Excel ---
@app.route('/live_prices.csv')
def get_live_prices_csv_for_excel():
    """
    Serves the latest tick data for all subscribed instruments as a CSV file.
    Excel will read from this URL.
    """
    # Get the processed data from the kite_data_manager
    rows = kite_data_manager.get_live_prices_for_csv()

    # Create an in-memory CSV string
    if len(rows) > 1:  # If there are actual data rows beyond just headers
        df = pd.DataFrame(rows[1:], columns=rows[0])
    else:  # Only headers if no data yet
        df = pd.DataFrame(columns=rows[0])

    buffer = io.StringIO()
    df.to_csv(buffer, index=False)
    csv_data = buffer.getvalue()

    response = Response(csv_data, mimetype="text/csv")
    response.headers["Content-Disposition"] = "attachment; filename=live_prices.csv"
    # print(f"[Flask] Served live_prices.csv. Contains {len(rows) - 1} data rows.")
    return response


# --- Application Startup Logic ---
async def startup_sequence():
    """Performs initial setup (Excel read, Kite init, Ticker start) asynchronously."""
    print("[App] Initiating startup sequence...")

    # 1. Read instruments from Excel
    instruments_from_excel = read_instrument_list_from_excel()
    if not instruments_from_excel:
        print("[App] No instruments loaded from Excel. Please check the file and restart.")
        return False  # Indicate failure, main script will exit

    # 2. Initialize KiteConnect and map instruments to tokens
    success = await kite_data_manager.initialize_kite_and_instruments(instruments_from_excel)
    if not success:
        print("[App] Failed to initialize KiteConnect or map instruments. Exiting.")
        return False

    # 3. Start the KiteTicker WebSocket thread
    if not kite_data_manager.start_kite_ticker_thread():
        print("[App] Failed to start KiteTicker WebSocket. Exiting.")
        return False

    print("[App] Startup sequence completed successfully.")
    return True


if __name__ == '__main__':
    print("--- Starting Live Tick Data for Excel Application ---")

    # Run the asynchronous startup sequence
    if asyncio.run(startup_sequence()):
        print(f"[App] Flask server will start on http://{FLASK_HOST}:{FLASK_PORT}/")
        print(f"[App] Excel should connect to http://{FLASK_HOST}:{FLASK_PORT}/live_prices.csv")
        # Start Flask web server (synchronous)
        app.run(host=FLASK_HOST, port=FLASK_PORT, debug=False, use_reloader=False)
    else:
        print("[App] Application startup failed. Exiting.")