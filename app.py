import time
import threading
import os

import pandas as pd
from flask import Flask, jsonify
from kiteconnect import KiteTicker

import config
import utils

# =========================================================
# WORKING DIRECTORY
# =========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

# =========================================================
# GLOBAL STORES
# =========================================================
LIVE_DATA = {}              # { SYMBOL: { field: value } }
TOKEN_TO_SYMBOL = {}
SYMBOL_TO_TOKEN = {}
SUBSCRIBED_TOKENS = set()

kws = None
kite = None
lock = threading.Lock()

# =========================================================
# KITE CALLBACKS
# =========================================================
def on_ticks(ws, ticks):
    with lock:
        for tick in ticks:
            token = tick.get("instrument_token")
            symbol = TOKEN_TO_SYMBOL.get(token)
            if not symbol:
                continue

            LIVE_DATA[symbol] = {
                "last_price": tick.get("last_price"),
                "volume": tick.get("volume_traded"),
                "vwap": tick.get("average_traded_price"),
                "oi": tick.get("oi"),
                "open": tick.get("ohlc", {}).get("open"),
                "high": tick.get("ohlc", {}).get("high"),
                "low": tick.get("ohlc", {}).get("low"),
                "close": tick.get("ohlc", {}).get("close"),
                "bid_price": tick.get("depth", {}).get("buy", [{}])[0].get("price"),
                "ask_price": tick.get("depth", {}).get("sell", [{}])[0].get("price"),
            }


def on_connect(ws, response):
    print("[KITE] WebSocket connected")


# =========================================================
# SUBSCRIBE SYMBOL
# =========================================================
def subscribe_symbol(symbol):
    symbol = symbol.upper()

    with lock:
        token = SYMBOL_TO_TOKEN.get(symbol)
        if not token or token in SUBSCRIBED_TOKENS:
            return

        SUBSCRIBED_TOKENS.add(token)
        TOKEN_TO_SYMBOL[token] = symbol
        kws.subscribe([token])
        kws.set_mode(kws.MODE_FULL, [token])

        print(f"[SUBSCRIBE] {symbol}")


# =========================================================
# START KITE WS
# =========================================================
def start_kite_ws():
    global kws
    kws = KiteTicker(config.API_KEY, kite.access_token)
    kws.on_ticks = on_ticks
    kws.on_connect = on_connect
    kws.connect(threaded=True)


# =========================================================
# INIT
# =========================================================
def init():
    global kite, SYMBOL_TO_TOKEN
    kite = utils.get_client(api_key=config.API_KEY)

    print("[INIT] Loading instruments...")
    instruments = kite.instruments()
    instrument_df = pd.DataFrame(instruments)
    instrument_df.to_csv(os.path.join('instuments.csv'))

    SYMBOL_TO_TOKEN = {
        i["tradingsymbol"].upper(): i["instrument_token"]
        for i in instruments
    }
    print(f"[INIT] {len(SYMBOL_TO_TOKEN)} instruments loaded")


# =========================================================
# HTTP SERVER (SNAPSHOT)
# =========================================================
app = Flask(__name__)

@app.route("/snapshot")
def snapshot():
    with lock:
        return jsonify(LIVE_DATA)

@app.route("/subscribe/<symbol>")
def subscribe(symbol):
    subscribe_symbol(symbol)
    return jsonify({"status": "ok", "symbol": symbol.upper()})


def start_http():
    app.run(host="127.0.0.1", port=5001, threaded=True)


# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":
    print("=== ZERODHA LIVE SNAPSHOT SERVER ===")
    init()

    threading.Thread(target=start_kite_ws, daemon=True).start()
    threading.Thread(target=start_http, daemon=True).start()

    while True:
        time.sleep(1)


# =RTD("clsZerodhaRTD","", "RELIANCE", "last_price")
# =RTD("clsZerodhaRTD","", A2, "volume")
# =RTD("clsZerodhaRTD","", "INFY", "bid_price")