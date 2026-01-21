import time
import threading
import os
import xlwings as xw
from kiteconnect import KiteTicker

import config
import utils

# =========================================================
# FIX WORKING DIRECTORY (IMPORTANT FOR XLWINGS)
# =========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

# =========================================================
# GLOBAL STORES
# =========================================================
LIVE_DATA = {}              # { SYMBOL: { field: value } }
TOKEN_TO_SYMBOL = {}        # { token: symbol }
SYMBOL_TO_TOKEN = {}        # { symbol: token }
SUBSCRIBED_TOKENS = set()

kws = None
kite = None
lock = threading.Lock()

STARTED = False   # <-- IMPORTANT GUARD


# =========================================================
# KITE CALLBACKS
# =========================================================
def on_ticks(ws, ticks):
    for tick in ticks:
        token = tick.get("instrument_token")
        symbol = TOKEN_TO_SYMBOL.get(token)

        if not symbol:
            continue

        LIVE_DATA[symbol] = {
            "last_price": tick.get("last_price"),
            "volume": tick.get("volume_traded"),
            "oi": tick.get("oi"),

            "open": tick.get("ohlc", {}).get("open"),
            "high": tick.get("ohlc", {}).get("high"),
            "low": tick.get("ohlc", {}).get("low"),
            "close": tick.get("ohlc", {}).get("close"),

            "bid_price": tick.get("depth", {}).get("buy", [{}])[0].get("price"),
            "bid_qty": tick.get("depth", {}).get("buy", [{}])[0].get("quantity"),
            "ask_price": tick.get("depth", {}).get("sell", [{}])[0].get("price"),
            "ask_qty": tick.get("depth", {}).get("sell", [{}])[0].get("quantity"),

            "exchange_timestamp": tick.get("exchange_timestamp"),
            "last_trade_time": tick.get("last_trade_time"),
        }


def on_connect(ws, response):
    print("[Kite] WebSocket connected")


# =========================================================
# DYNAMIC SUBSCRIBE LOGIC
# =========================================================
def subscribe_symbol(symbol):
    global kws

    if kws is None:
        return

    symbol = symbol.upper()

    with lock:
        token = SYMBOL_TO_TOKEN.get(symbol)

        if not token or token in SUBSCRIBED_TOKENS:
            return

        SUBSCRIBED_TOKENS.add(token)
        TOKEN_TO_SYMBOL[token] = symbol
        LIVE_DATA[symbol] = {}

        kws.subscribe([token])
        kws.set_mode(kws.MODE_FULL, [token])

        print(f"[SUBSCRIBE] {symbol} ({token})")


# =========================================================
# EXCEL FUNCTIONS
# =========================================================
@xw.func
def hello():
    return "hello"


@xw.func(rtd=True, volatile=False)
def kite_rtd(symbol, field):
    """
    Excel:
    =kite_rtd("RELIANCE","last_price")
    =kite_rtd(A2,"bid_price")
    """
    if not symbol or not field:
        return ""

    symbol = str(symbol).strip().upper()
    field = str(field).strip().lower()

    if symbol not in SYMBOL_TO_TOKEN:
        return "INVALID SYMBOL"

    if symbol not in LIVE_DATA:
        subscribe_symbol(symbol)

    return LIVE_DATA.get(symbol, {}).get(field, "")


# =========================================================
# START KITE WEBSOCKET
# =========================================================
def start_kite_ws():
    global kws, kite

    if kite is None:
        print("[ERROR] Kite not initialized")
        return

    kws = KiteTicker(config.API_KEY, kite.access_token)
    kws.on_ticks = on_ticks
    kws.on_connect = on_connect

    kws.connect(threaded=True)


# =========================================================
# INIT KITE + LOAD INSTRUMENT MASTER
# =========================================================
def init():
    global kite, SYMBOL_TO_TOKEN

    print("[INIT] Initializing Kite client")
    kite = utils.get_client(api_key=config.API_KEY)

    print("[INIT] Loading instrument master...")
    instruments = kite.instruments()

    SYMBOL_TO_TOKEN = {
        i["tradingsymbol"].upper(): i["instrument_token"]
        for i in instruments
    }

    print(f"[INIT] Loaded {len(SYMBOL_TO_TOKEN)} instruments")


# =========================================================
# BOOTSTRAP (RUNS ON XLWINGS IMPORT)
# =========================================================
def bootstrap():
    global STARTED

    if STARTED:
        return

    STARTED = True
    print("[BOOTSTRAP] Starting Kite RTD")

    init()
    threading.Thread(target=start_kite_ws, daemon=True).start()


# IMPORTANT: this runs when Excel imports the file
bootstrap()


# =========================================================
# TERMINAL MODE (OPTIONAL)
# =========================================================
if __name__ == "__main__":
    print("=== ZERODHA EXCEL RTD (TERMINAL MODE) ===")
    while True:
        time.sleep(1)
