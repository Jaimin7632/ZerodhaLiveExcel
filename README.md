# Zerodha Live Excel Dashboard

This project provides a Python-based backend to fetch live tick data from Zerodha Kite Connect API and serve it to an Excel dashboard, allowing for near real-time price updates within Excel.

## Features

  * Fetches live tick data from Zerodha Kite Connect.
  * Allows defining instruments to track from an Excel input file.
  * Dynamically maps instrument names to their respective tokens.
  * Serves the latest tick data for all tracked instruments via a local web server (CSV format).
  * Enables an Excel workbook to automatically refresh and display live prices.



## Setup Instructions

Follow these steps to get your live Excel dashboard running.

### 1\. Python Installation

If you don't have Python installed, download it from the official website:

  * Go to [Python.org](https://www.python.org/downloads/).
  * Download the latest stable version of Python 3.x.
  * During installation, **make sure to check "Add Python X.Y to PATH"** (this is very important for running commands from CMD).

### 2\. Install Required Packages

Open your Command Prompt (CMD) or PowerShell and navigate to the `ZerodhaLiveExcel` project directory. Then run the following command:

```bash
pip install Flask openpyxl kiteconnect pandas websockets asgiref
```

### 3\. Change API Key and Secret in `config.py`

Open the `config.py` file in your project directory using a text editor (like Notepad, VS Code, Sublime Text, etc.).

Update the following lines with your Zerodha Kite Connect API credentials:

```python
# config.py

# --- Kite Connect API Credentials ---
API_KEY = "YOUR_KITE_API_KEY"         # Get this from your Kite Connect developer console
API_SECRET = "YOUR_KITE_API_SECRET"   # Get this from your Kite Connect developer console
                                       
```

### 4\. Update Needed Symbols in `instruments_to_track.xlsx`

This Excel file tells your Python script which instruments to fetch data for.

1.  Locate `instruments_to_track.xlsx` in your project directory.

2.  Open `instruments_to_track.xlsx`.

3.  Starting from **Row 2**, list the instruments you want to track:

      * **For simple Equity/Futures/Options symbols:**
    
        | Instrument\_Name      |
        |:----------------------|
        | RELIANCE              |
        | NIFTY25FEB            |
        | BSE:TATAMOTORS        |
        | CRUDEOIL25MAR         |
        | INFY                  |
        | BANKNIFTY25FEB47000CE |

      * **For specific exchange/instrument combinations (recommended for clarity and avoiding ambiguity):**
        You can explicitly state the exchange in the `Instrument_Name` column using the `EXCHANGE:TRADINGSYMBOL` format. This can be helpful if `TRADINGSYMBOL` might exist on multiple exchanges or for complex F\&O names. The script will prioritize finding this exact match.

      * **Ensure `Instrument_Name` is the exact `tradingsymbol` used by Kite.** Check Kite's instrument master if unsure.

6.  **Save and Close `instruments_to_track.xlsx`** before running `app.py`.

### 5\. Run `app.py` (Guide for Windows using CMD)

This will start your Python backend, connecting to Kite and running the local web server.

1.  Open your Command Prompt (CMD) or PowerShell.
2.  Navigate to your project directory (`ZerodhaLiveExcel`). For example, if your folder is on your Desktop:
    ```cmd
    cd %USERPROFILE%\Desktop\ZerodhaLiveExcel
    ```
    (Adjust path as per your setup)
3.  Run the application:
    ```bash
    python app.py
    ```
4.  Keep this CMD window open. You will see logs from the Python script indicating connection status, instruments mapped, and when data is served. **Do not close this window.**



## Running the Live Dashboard

1.  **Ensure your Python `app.py` is running** in its CMD window.
2.  **Open `Live.xlsm`**.
3.  When prompted by Excel, **click "Enable Content" or "Enable Macros"**.

Your Excel sheet should now begin refreshing automatically, approximately every second, displaying live prices from your Python server.

## Troubleshooting Common Issues

  * **"Sub or Function not defined" (VBA Error):**
      * **Cause:** Code is in the wrong place or a typo.
      * **Fix:** Ensure `StartLivePriceRefresh`, `RefreshAndScheduleNext`, `StopLivePriceRefresh` are in a standard module (`Module1`). `Workbook_Open` and `Workbook_BeforeClose` must be in the `ThisWorkbook` object. Double-check spelling. Run `Debug > Compile VBAProject` in the VBA editor.
  * **"Query '[Name]' not found" (VBA Debug.Print message):**
      * **Cause:** The `Const LIVE_DATA_QUERY_NAME` in your VBA module does not exactly match the query name in Excel.
      * **Fix:** In Excel, go `Data` tab -\> `Queries & Connections` pane. Note the *exact* query name (e.g., "Query1", "Query - live\_prices"). Update `Const LIVE_DATA_QUERY_NAME` in VBA to this exact string.
  * **"Object does not support this property or method" (VBA Error):**
      * **Cause:** The connection object retrieved by VBA isn't of a type that supports `.Refresh` (usually when it's not a proper Power Query `WorkbookConnection`).
      * **Fix:** The provided VBA code includes checks for `WorkbookConnection` and `QueryTable` types. If the error persists, ensure your data was imported via `Data > From Web` in the current Excel version, not an older method. If the error points to `conn.Refresh` after these checks, inspect `TypeName(conn)` in the Immediate Window during debug to see what kind of object it is.
  * **Data not refreshing / Not "second-wise":**
      * **Check Python CMD:** Is `app.py` running without errors? Is it logging `[Flask] Served live_prices.csv...` every time Excel refreshes (approx. every second)?
      * **Check Python `kite_data_manager` logs:** Are `[Kite Data Manager] Received tick...` messages appearing frequently? If not, Kite may not be sending new ticks for your selected instruments, or your API access token may be expired.
      * **Check Excel VBA Immediate Window:** Is it logging `Refreshed query...` every second? If not, the VBA `Application.OnTime` schedule might be broken.
      * **Data Liquidity:** The instrument might not receive new ticks every second, especially during off-market hours or for illiquid stocks.
      * **Excel Performance:** For very large numbers of instruments or complex dashboards, Excel might struggle to re-process data every second. Consider increasing `REFRESH_INTERVAL_SECONDS` to 2 or 3.