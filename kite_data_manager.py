# live_tick_excel_app/kite_data_manager.py
import asyncio
import threading
import time
from datetime import datetime
from kiteconnect import KiteConnect, KiteTicker

import config
import utils
from config import API_KEY

# Global storage for subscribed symbols and their latest tick data
# This is accessed by both the WebSocket thread and the Flask thread
_subscribed_symbols_data = {}  # Key: instrument_token, Value: Symbol object
_symbols_lock = threading.Lock()  # Lock for thread-safe access to _subscribed_symbols_data

# Kite Connect and Ticker instances
_kite = None
_kws = None


# --- Symbol Class Definition ---
class Symbol:
    def __init__(self, exchange, instrument_name, instrument_token):
        self.exchange = exchange
        self.instrument_name = instrument_name
        self.instrument_token = instrument_token
        # Initialize last_tick_data with None for all fields
        self.last_tick_data = {
            'timestamp': None,
            'last_price': None,
            'ohlc_open': None,
            'ohlc_high': None,
            'ohlc_low': None,
            'ohlc_close': None,
            'volume': None,
            'average_price': None,
            'oi': None,
            'buy_quantity': None,
            'sell_quantity': None,
        }

    # Method to update the symbol's last tick data
    def update_tick(self, tick):
        # Update only the relevant fields from the tick, handling missing keys gracefully
        self.last_tick_data['timestamp'] = datetime.fromtimestamp(tick.get('timestamp', time.time())).strftime(
            '%Y-%m-%d %H:%M:%S.%f')
        self.last_tick_data['last_price'] = tick.get('last_price')

        ohlc = tick.get('ohlc', {})
        self.last_tick_data['ohlc_open'] = ohlc.get('open')
        self.last_tick_data['ohlc_high'] = ohlc.get('high')
        self.last_tick_data['ohlc_low'] = ohlc.get('low')
        self.last_tick_data['ohlc_close'] = ohlc.get('close')

        self.last_tick_data['volume'] = tick.get('volume')
        self.last_tick_data['average_price'] = tick.get('average_price')
        self.last_tick_data['oi'] = tick.get('oi')
        self.last_tick_data['buy_quantity'] = tick.get('buy_quantity')
        self.last_tick_data['sell_quantity'] = tick.get('sell_quantity')


# --- WebSocket Event Handlers ---
def _on_ticks(ws, ticks):
    """Callback to receive ticks and update Symbol objects in thread-safe manner."""
    with _symbols_lock:
        for tick in ticks:
            token = tick.get('instrument_token')
            if token in _subscribed_symbols_data:
                _subscribed_symbols_data[token].update_tick(tick)
            # else:
            #     print(f"[Kite Data Manager] Received tick for unsubscribed token: {token}")


def _on_connect(ws, response):
    """Callback on successful WebSocket connection."""
    print("[Kite Data Manager] WebSocket connected.")
    with _symbols_lock:
        if _subscribed_symbols_data:
            tokens = list(_subscribed_symbols_data.keys())
            ws.subscribe(tokens)
            ws.set_mode(ws.MODE_FULL, tokens)  # MODE_FULL for detailed ticks
            print(f"[Kite Data Manager] Subscribed to {len(tokens)} instruments.")
        else:
            print("[Kite Data Manager] No instruments to subscribe to. Check Excel config.")


def _on_close(ws, code, reason):
    """Callback on WebSocket connection close."""
    print(f"[Kite Data Manager] WebSocket closed with code {code}: {reason}")


def _on_error(ws, code, reason):
    """Callback on WebSocket connection error."""
    print(f"[Kite Data Manager] WebSocket error with code {code}: {reason}")


def _on_reconnect(ws, attempts_count):
    """Callback on WebSocket reconnection attempt."""
    print(f"[Kite Data Manager] WebSocket attempting to reconnect... (Attempt {attempts_count})")


def _on_noreconnect(ws):
    """Callback when WebSocket reconnection attempts fail."""
    print("[Kite Data Manager] WebSocket reconnection failed. Please check connectivity.")


# --- Initialization and Data Retrieval Functions ---
async def initialize_kite_and_instruments(instruments_to_find):
    """
    Initializes KiteConnect, fetches instrument tokens based on Excel input,
    and populates _subscribed_symbols_data with Symbol objects.
    Returns True on success, False on failure.
    """
    global _kite, _kws

    # Initialize KiteConnect
    try:
        _kite = utils.get_client(api_key=config.API_KEY)
        profile = _kite.profile()
        print(f"[Kite Data Manager] KiteConnect initialized. Logged in as: {profile.get('user_name')}")
    except Exception as e:
        print(f"[Kite Data Manager] Error initializing KiteConnect. Ensure ACCESS_TOKEN is valid. Error: {e}")
        print(f"[Kite Data Manager] Login URL for new request_token: {_kite.login_url()}")
        return False

    # Fetch all instruments. This can be memory-intensive.
    print("[Kite Data Manager] Fetching full instrument master from KiteConnect (might take a moment)...")
    try:
        all_instruments = _kite.instruments()
        print(f"[Kite Data Manager] Fetched {len(all_instruments)} instruments from KiteConnect.")
    except Exception as e:
        print(f"[Kite Data Manager] Error fetching instrument master: {e}")
        return False

    # Map instruments to tokens and create Symbol objects
    tokens_for_subscription_list = []  # List of tokens to pass to kws.subscribe()
    with _symbols_lock:  # Ensure thread safety when modifying _subscribed_symbols_data
        for instrument_name in instruments_to_find:
            instrument_token = None

            # 1. Try to find by exact Exchange and Instrument_Name (trading symbol)
            # This is robust for F&O, equity etc.
            exchange = None
            if ':' in instrument_name:
                exchange, instrument_name = instrument_name.split(':')
            if exchange:
                found_insts = [
                    i for i in all_instruments
                    if i['exchange'] == exchange and i['tradingsymbol'] == instrument_name
                ]
            else:
                found_insts = [
                    i for i in all_instruments
                    if i['tradingsymbol'] == instrument_name
                ]

            if found_insts:
                if len(found_insts) > 1:
                    print(
                        f"[Kite Data Manager] Warning: Multiple matches for '{exchange}:{instrument_name}'. Picking first by (instrument_type, expiry, strike).")
                    found_insts.sort(
                        key=lambda x: (x.get('instrument_type', ''), x.get('expiry', ''), x.get('strike', 0)))

                instrument_token = found_insts[0]['instrument_token']
                print(f"[Kite Data Manager] Mapped '{exchange}:{instrument_name}' to token: {instrument_token}")
            else:
                print(f"[Kite Data Manager] Could not find '{exchange}:{instrument_name}'. Skipping.")
                continue  # Move to next instrument if not found

            if instrument_token:
                _subscribed_symbols_data[instrument_token] = Symbol(exchange, instrument_name, instrument_token)
                tokens_for_subscription_list.append(instrument_token)

    if not tokens_for_subscription_list:
        print("[Kite Data Manager] No valid instrument tokens found for subscription.")
        return False

    print(f"[Kite Data Manager] Ready to subscribe to {len(tokens_for_subscription_list)} instruments.")
    return True  # Indicate success


def start_kite_ticker_thread():
    """Starts the KiteTicker WebSocket connection in a separate daemon thread."""
    global _kws
    if not _kite:
        print("[Kite Data Manager] KiteConnect not initialized. Cannot start Ticker.")
        return False
    kite = utils.get_client(api_key=API_KEY)
    _kws = KiteTicker(api_key=API_KEY, access_token=kite.access_token)

    # Assign all event handlers
    _kws.on_ticks = _on_ticks
    _kws.on_connect = _on_connect
    _kws.on_close = _on_close
    _kws.on_error = _on_error
    _kws.on_reconnect = _on_reconnect
    _kws.on_noreconnect = _on_noreconnect

    # Start the WebSocket connection. The _on_connect will handle actual subscriptions.
    threading.Thread(target=_kws.connect, daemon=True).start()
    print("[Kite Data Manager] KiteTicker WebSocket thread initiated.")
    return True


def get_live_prices_for_csv():
    """
    Returns a list of lists representing rows for the CSV output,
    with only the latest price and necessary details for Excel.
    """
    output_rows = []
    # CSV Headers for Excel
    csv_headers = [
        'Instrument_Name', 'Instrument_Token',
        'Last_Price', 'Timestamp', 'OHLC_Open', 'OHLC_High', 'OHLC_Low', 'OHLC_Close',
        'Volume', 'Average_Price', 'OI', 'Buy_Quantity', 'Sell_Quantity'
    ]
    output_rows.append(csv_headers)

    with _symbols_lock:  # Ensure thread safety when reading _subscribed_symbols_data
        for token, symbol_obj in _subscribed_symbols_data.items():
            # Get latest tick data from the Symbol object
            tick_data = symbol_obj.last_tick_data

            # Format the row for CSV
            row = [
                # symbol_obj.exchange,
                symbol_obj.instrument_name,
                symbol_obj.instrument_token,
                f"{tick_data['last_price']:.2f}" if tick_data['last_price'] is not None else '',
                # Format to 2 decimal places
                tick_data['timestamp'] if tick_data['timestamp'] is not None else '',
                f"{tick_data['ohlc_open']:.2f}" if tick_data['ohlc_open'] is not None else '',
                f"{tick_data['ohlc_high']:.2f}" if tick_data['ohlc_high'] is not None else '',
                f"{tick_data['ohlc_low']:.2f}" if tick_data['ohlc_low'] is not None else '',
                f"{tick_data['ohlc_close']:.2f}" if tick_data['ohlc_close'] is not None else '',
                tick_data['volume'] if tick_data['volume'] is not None else '',
                f"{tick_data['average_price']:.2f}" if tick_data['average_price'] is not None else '',
                tick_data['oi'] if tick_data['oi'] is not None else '',
                tick_data['buy_quantity'] if tick_data['buy_quantity'] is not None else '',
                tick_data['sell_quantity'] if tick_data['sell_quantity'] is not None else '',
            ]
            output_rows.append(row)

    return output_rows