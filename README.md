

# Zerodha Live Excel Dashboard (Python + VBA HTTP)

This project provides a **Python + Excel VBA** solution to stream **live market data from Zerodha Kite Connect** into Microsoft Excel.

Live ticks are received via **Zerodha WebSocket**, stored in memory by Python, and exposed through a **local HTTP snapshot API**.
Excel uses a **VBA timer (1-second refresh)** to fetch this snapshot and update only the required cells.

---

## Features

* Live tick data via **Zerodha Kite WebSocket**
* Dynamic symbol → instrument token mapping
* Automatic subscription when a new symbol is added in Excel
* 1-second refresh using VBA timer
* Updates **only required cells**, not the full workbook

---

## Architecture

```
Zerodha Kite WebSocket
        ↓
    Python Backend
 (in-memory LIVE_DATA)
        ↓
   /snapshot (HTTP JSON)
        ↓
 Excel VBA (1 sec timer)
        ↓
     Excel Cells
```

---

## Excel Layout (Required)

Create your Excel sheet like this:

| Column | Header     |
| ------ | ---------- |
| A      | Symbol     |
| B      | Open       |
| C      | High       |
| D      | Low        |
| E      | Close      |
| F      | Last Price |
| G      | Volume     |
| H      | OI         |

Example:

```
A2: RELIANCE
A3: INFY
A4: TCS
```

⚠️ **You only type symbols in Column A**
All other columns are filled automatically by VBA.

---

## Setup Instructions

### 1️⃣ Install Python

Download **64-bit Python** from:

[https://www.python.org/downloads/](https://www.python.org/downloads/)

During installation:

✔ Add Python to PATH
✔ Install for all users

---

### 2️⃣ Install Python Dependencies

Open Command Prompt and run:

```bash
pip install kiteconnect flask pandas
```

---

### 3️⃣ Configure API Credentials

Edit `config.py`:

```python
API_KEY = "YOUR_KITE_API_KEY"
API_SECRET = "YOUR_KITE_API_SECRET"
```

Login logic should be handled inside `utils.get_client()`.

---

### 4️⃣ Run Python Server

From project directory:

```bash
python app.py
```

Expected output:

```
[INIT] Initializing Kite client
[INIT] Loaded XXXX instruments
[Kite] WebSocket connected
```

⚠️ **Keep this terminal open** while Excel is running.

---

### 5️⃣ Instrument Reference File

After Python starts, an **`instruments.csv`** file is created.

Use this file to:

* Verify correct `tradingsymbol`
* Check exchange & instrument availability
* Avoid spelling mistakes

Only symbols present in `instruments.csv` will work.

---

### 6️⃣ Excel Setup (Already provided as Live.xlsm)

1. Open Excel
2. Press **ALT + F11**
3. Insert a **Standard Module**
4. Paste the VBA code you provided
5. Paste `Workbook_Open` and `Workbook_BeforeClose` into **ThisWorkbook**
6. Save file as:

```
Live.xlsm
```

---

## Running the Live Feed

### Start Automatically on Excel Open

When `Live.xlsm` opens:

* VBA automatically starts live feed
* Snapshot is fetched every 1 second
* Symbols are subscribed dynamically

### Manual Control (Optional)

Press **ALT + F8** and run:

* `StartLiveFeed`
* `StopLiveFeed`

---

## Dynamic Symbol Subscription

✔ Add a new symbol in Column A
✔ VBA automatically calls:

```
/subscribe/<symbol>
```

✔ Python subscribes via WebSocket
✔ Data appears when snapshot updates

No restart required in most cases.

---

## Troubleshooting

### ❌ New symbol subscribed but data not showing

**Cause:**
Excel cached blank values before data arrived.

**Fix:**

1. Save `Live.xlsm`
2. Close Excel
3. Re-open Excel
4. Data will populate

---

### ❌ Symbol not updating at all

Check:

* Symbol exists in `instruments.csv`
* Correct spelling and case
* Market is open
* Instrument is liquid

---

### ❌ No live updates

Check:

* Python server is running
* `/snapshot` works in browser:

  ```
  http://127.0.0.1:5001/snapshot
  ```
* Excel macros are enabled
* Excel calculation mode is **Automatic**


---

## Performance & Limits

| Item        | Value                   |
| ----------- | ----------------------- |
| Refresh     | 1 second (configurable) |
| Latency     | ~300–700 ms             |
| Max symbols | ~500 practical          |
| Stability   | High                    |
| Best use    | Monitoring & dashboards |

---
