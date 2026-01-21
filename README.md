Here is a **final, clean, copy-paste-ready `README.md`** for your **xlwings-based Zerodha Live Excel Dashboard**.

This version is **fully aligned with your final Python code**, removes all Flask/VBA/CSV logic, and clearly documents **requirements, setup, usage, and limitations**.

---

# Zerodha Live Excel Dashboard (xlwings RTD)

This project provides a **Python + xlwings Real-Time Data (RTD) bridge** between the Zerodha Kite Connect API and Microsoft Excel.

Live market data is streamed from Zerodha via WebSocket and consumed **directly inside Excel formulas** using xlwings User Defined Functions (UDFs), allowing near real-time updates without CSV files, Power Query, or VBA refresh loops.

---

## ⚠️ IMPORTANT PLATFORM REQUIREMENT (READ FIRST)

This project **ONLY works on**:

✅ **Windows Desktop Excel** (Microsoft 365 / Office 2019 / 2021 – non-Store)
❌ **Microsoft Store Excel** – NOT supported
❌ **Excel on macOS** – NOT supported

> xlwings UDFs require **Windows COM automation**, which is unavailable on Store Excel and macOS.

---

## Features

* Live tick data via **Zerodha Kite WebSocket**
* Dynamic symbol → instrument token mapping
* On-demand subscription (symbol subscribes when first used in Excel)
* Excel formulas update automatically (RTD-style)
* No CSV files
* No Power Query
* No VBA refresh loops
* No Excel polling
* Single Python process feeds Excel in real time

---

## Architecture

```
Zerodha Kite WebSocket
        ↓
    Python Backend
        ↓
 xlwings COM RTD Server
        ↓
     Excel Formulas
```

---

## Excel Usage Examples

```excel
=hello()

=kite_rtd("RELIANCE","last_price")

=kite_rtd(A2,"bid_price")
```

### Supported Fields (examples)

* `last_price`
* `volume`
* `oi`
* `open`
* `high`
* `low`
* `close`
* `bid_price`
* `bid_qty`
* `ask_price`
* `ask_qty`
* `exchange_timestamp`
* `last_trade_time`

---

## Setup Instructions

### 1. Install Python (Windows)

Download **64-bit Python** from:

[https://www.python.org/downloads/](https://www.python.org/downloads/)

During installation:

* ✅ Check **“Add Python to PATH”**
* ✅ Install for all users (recommended)

---

### 2. Install Required Packages

Open **Command Prompt (CMD)** and run:

```bash
pip install xlwings kiteconnect pandas
```

---

### 3. Install xlwings Excel Add-in

Run **once**:

```bash
xlwings addin install
```

Restart Excel and confirm the **xlwings** tab appears.

---

### 4. Verify xlwings COM Add-in (CRITICAL)

In Excel:

```
File → Options → Add-ins
Manage: COM Add-ins → Go
```

You **must see and enable**:

```
☑ xlwings
```

If xlwings does not appear here, UDFs will **not work**.

---

### 5. Enable VBA Trust Access

In Excel:

```
File → Options → Trust Center → Trust Center Settings
→ Macro Settings
```

Enable:

* ✅ Enable all macros
* ✅ Trust access to the VBA project model

Restart Excel.

---

### 6. Configure API Credentials (`config.py`)

Edit `config.py`:

```python
API_KEY = "YOUR_KITE_API_KEY"
API_SECRET = "YOUR_KITE_API_SECRET"
```

Access token handling should be implemented in `utils.get_client()` as per your login flow.

---

### 7. Run the Python Backend

Open CMD, navigate to the project directory:

```cmd
cd path\to\ZerodhaLiveExcel
```

Run:

```bash
python app.py
```

Expected logs:

```
[BOOTSTRAP] Starting Kite RTD
[INIT] Initializing Kite client
[INIT] Loaded XXXX instruments
[Kite] WebSocket connected
```

⚠️ **Keep this window open** while Excel is running.

---

### 8. Import Functions into Excel

1. Open Excel
2. Go to **xlwings → Import Functions**
3. Wait for import to complete

---

## Running the Live Dashboard

1. Ensure `app.py` is running
2. In Excel, test:

   ```excel
   =hello()
   ```

   Expected output:

   ```
   hello
   ```
3. Then use:

   ```excel
   =kite_rtd("RELIANCE","last_price")
   ```

Prices will update automatically as ticks arrive.

---

## Troubleshooting

### ❌ `Object required` in Excel

Cause:

* Excel does not support COM automation
* xlwings COM add-in not loaded

Fix:

* Use **Desktop Excel only**
* Ensure xlwings appears under **COM Add-ins**
* Restart Excel and re-import functions

---

### ❌ Function name appears but returns error

Cause:

* Function registered but COM server unavailable

Fix:

* Excel environment issue (see above)

---

### ❌ No live updates

Check:

* Python console shows `[Kite] WebSocket connected`
* Excel calculation mode:

  ```
  Formulas → Calculation Options → Automatic
  ```
* Market is open and instrument is liquid

---

### ❌ Excel freezes or lags

Excel is not designed for high-frequency RTD:

Recommendations:

* Limit to ~50–100 symbols
* Avoid heavy formulas
* Restart Excel daily

---

## Performance & Limitations

| Aspect           | Notes                   |
| ---------------- | ----------------------- |
| Latency          | ~200–500 ms             |
| Update type      | Tick-driven             |
| Max symbols      | ~100 practical          |
| Use case         | Monitoring & dashboards |
| Not suitable for | HFT / scalping          |

---
