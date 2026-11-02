import { Component } from '@angular/core';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
  
export class AppComponent {
  title = 'angular_project';
}
// code
// upstoxlogin.py
from playwright.sync_api import Playwright, sync_playwright, expect
from urllib.parse import parse_qs,urlparse,quote
import pyotp
import requests
import json
import time
import datetime
import pandas as pd
import numpy as np
import os
from urllib.parse import quote
from dateutil import parser
# from getLtp import get_ltp
# python -m playwright install
# python -m  playwright codegen demo.playwright.dev/todomvc
# from tradestock import trade_vectorized, print
# from test import smsndbdmsfbd, calculate_dema, calculate_macd

API_KEY = '4de0c889-4150-4f47-94f5-42b3e26199de'
SECRET_KEY = 'eoo5a9r0zt'
RURL = 'https://127.0.0.1:5000/'

TOTP_KEY = ''
MOBILE_NO = ''
PIN   =    ''

file_path = "hdfcbank_intraday_signals.xlsx"
rurlEncode = quote(RURL,safe="")
symbol = 'NSE_EQ|HDFCBANK'
instrument_token = "INE040A01034" 
access_token = "eyJ0eXAiOiJKV1QiLCJrZXlfaWQiOiJza192MS4wIiwiYWxnIjoiSFMyNTYifQ.eyJzdWIiOiI0TENQS0siLCJqdGkiOiI2YWRjZjUyMWMyN2RkNDBlYzA2MWNiZjEiLCJpc011bHRpQ2xpZW50IjpmYWxzZSwiaXNQbHVzUGxhbiI6ZmFsc2UsImlhdCI6MTc5Mjg2NTU2OSwiaXNzIjoidWRhcGktZ2F0ZXdheS1zZXJ2aWNlIiwiZXhwIjoxNzkyODc5MjAwfQ._v8g3vt-HFQvAglE85hpAwoGJsAQP42iDLAXxtjNmmY"
AUTH_URL = f'https://api-v2.upstox.com/login/authorization/dialog?response_type=code&client_id={API_KEY}&redirect_uri={rurlEncode}'

def getAccessToken(code):
    url = 'https://api-v2.upstox.com/login/authorization/token'

    headers = {
        'accept': 'application/json',
        # 'Api-Version': '2.0',
        'Content-Type': 'application/x-www-form-urlencoded'
    }

    data1 = {
        'code': code,
        'client_id': API_KEY,
        'client_secret': SECRET_KEY,
        'redirect_uri': RURL,
        'grant_type': 'authorization_code'
    }

    response = requests.post(url, headers=headers, data=data1)
    json_response = response.json()
    # print("jj",json_response)
    value = json_response['access_token']
    # print(f"access_token:  {json_response['access_token']}")
    return json_response, value

def save_access_token(token):
    with open("access_token.txt", "w") as f:
        f.write(token)

def load_access_token():
    with open("access_token.txt", "r") as f:
        return f.read().strip()

def fetch_intraday_candles(instrument_key):
    access_token = load_access_token()
    url = f'https://api.upstox.com/v3/historical-candle/intraday/{instrument_key}/minutes/5'

    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Authorization': f'Bearer {access_token}'
    }

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    df = response.json()
    # print(df)
    return df
    # candles = df['data']['candles']
    # df = pd.DataFrame(candles, columns=['timestamp', 'open', 'high', 'low', 'close', 'volume'])

def fetch_intraday_candles1(instrument_key):
    access_token = load_access_token()
    url = f'https://api.upstox.com/v3/historical-candle/intraday/{instrument_key}/minutes/1'

    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Authorization': f'Bearer {access_token}'
    }

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    df = response.json()
    return df

def save_intraday_to_excel1(upstox_response, file_name):
    candles = upstox_response["data"]["candles"]

    df = pd.DataFrame(candles, columns=["Timestamp", "Open", "High", "Low", "Close", "Volume", "Oi"])

    df["time"] = df["Timestamp"].apply(
        lambda x: parser.parse(x).strftime("%I:%M").lstrip("0")
    )

    df = df[["time", "Open", "High", "Low", "Close"]]

    df = df[::-1].reset_index(drop=True)

    df.to_excel(file_name, index=False)
    print(f"Saved to Excel → {file_name}")

def upstox_candles_to_df(resp):
    candles = resp['data']['candles']

    df = pd.DataFrame(
        candles,
        columns=['time', 'open', 'high', 'low', 'close', 'volume', 'oi']
    )
    # print("df", df)
    df['time'] = pd.to_datetime(df['time']).dt.tz_localize(None)
    return df

def save_intraday_to_excel(upstox_response, file_name):
    new_df = upstox_candles_to_df(upstox_response)
    # print("new_df", new_df)
    try:
        old_df = pd.read_excel(file_name)
        old_df['time'] = pd.to_datetime(old_df['time']).dt.tz_localize(None)
    except FileNotFoundError:
        new_df.to_excel(file_name, index=False)
        print("[Excel] First candle file created")
        return

    # ------------------------------------------------
    # Append only new candles (by timestamp)
    # ------------------------------------------------
    merged = pd.concat([old_df, new_df])
    merged = merged.drop_duplicates(subset=['time'])
    merged = merged.sort_values('time')
    if len(merged) > len(old_df):
        merged.to_excel(file_name, index=False)
        # print(f"[Excel] Appended {len(merged) - len(old_df)} new candle(s)")
    else:
        print("[Excel] No new candles to append")

def get_historical_candles(access_token, symbol_token, interval, start_date, end_date):
    """
    Fetch historical candle data from Upstox API.

    Parameters:
        access_token (str)  : Your Upstox access token
        symbol_token (str)  : Example -> 'NSE_EQ|INE848E01016'
        interval (str)      : '1', '5', '15', '30', '60'
        start_date (str)    : 'YYYY-MM-DD'
        end_date (str)      : 'YYYY-MM-DD'

    Returns:
        dict JSON response from Upstox OR error message
    """

    encoded_symbol = symbol_token.replace("|", "%7C")

    url = f"https://api.upstox.com/v3/historical-candle/{encoded_symbol}/minutes/{interval}/{start_date}/{end_date}"

    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error: {response.status_code} - {response.text}")
        return None

def run(playwright: Playwright) -> str:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()   
    with page.expect_request(f"*{RURL}?code*") as request:
        page.goto(AUTH_URL)
        page.locator("#mobileNum").click()
        page.locator("#mobileNum").fill(MOBILE_NO)
        page.get_by_role("button", name="Get OTP").click()
        page.locator("#otpNum").click()
        otp = pyotp.TOTP(TOTP_KEY).now()
        page.locator("#otpNum").fill(otp)
        page.get_by_role("button", name="Continue").click()
        page.get_by_label("Enter 6-digit PIN").click()
        page.get_by_label("Enter 6-digit PIN").fill(PIN)
        res = page.get_by_role("button", name="Continue").click()
        page.wait_for_load_state()

    url =    request.value.url 
    parsed = urlparse(url)
    code = parse_qs(parsed.query)['code'][0]
    context.close()
    browser.close()
    return code

with sync_playwright() as playwright:
    code = run(playwright)

access_token1 = getAccessToken(code)[1]
# print("a", access_token1)
save_access_token(access_token1)

def fetch_upstox_ltp(access_token, instrument_keys):
    """
    Fetch LTP for one or more Upstox instrument keys.

    Args:
        access_token (str): Upstox OAuth access token
        instrument_keys (str | list[str]): Single key or list of keys

    Returns:
        dict: {instrument_key: ltp}
    """

    if isinstance(instrument_keys, str):
        instrument_keys = [instrument_keys]

    url = "https://api.upstox.com/v3/market-quote/ltp"

    params = {
        "instrument_key": ",".join(instrument_keys)
    }

    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers, params=params, timeout=5)

    if response.status_code == 401:
        raise Exception("401 Unauthorized → Invalid or expired Upstox access token")

    response.raise_for_status()

    data = response.json().get("data", {})

    # Return simplified mapping: {instrument_key: ltp}
    return {
        key: info["last_price"]
        for key, info in data.items()
    }

# ltp_data = fetch_upstox_ltp(
#     access_token=access_token,
#     instrument_keys="NSE_EQ|INE117A01022"
# )

# print(ltp_data)

def run_5min_scheduler(instrument_key):
    print("\n=== CLOCK-DRIVEN 5-MIN SCHEDULER STARTED ===")

    market_open = datetime.time(9, 20, 1)
    market_close = datetime.time(23, 59, 1)

    last_fired_slot = None  # prevents duplicate execution

    while True:
        now = datetime.datetime.now()
        now_time = now.time()

        # Stop after market close
        if now_time >= market_close:
            print("Market closed. Scheduler stopped.")
            break

        # -----------------------------------------
        # EXACT TRIGGER CONDITION
        # -----------------------------------------
        if (
            now_time >= market_open
            and now.second == 1
            and now.minute % 1 == 0
        ):
            slot_key = now.replace(second=1, microsecond=0)

            if last_fired_slot != slot_key:
                last_fired_slot = slot_key

                print(f"\n[Scheduler] Triggered at {slot_key.time()}")

                intraday_candles = fetch_intraday_candles(instrument_key="NSE_EQ|INE356A01018")
                # candles = get_historical_candles(access_token=access_token, symbol_token="NSE_EQ|INE356A01018", interval="5", start_date="2025-12-26", end_date="2025-12-26")
                save_intraday_to_excel(upstox_response=intraday_candles, file_name="intraday_data.xlsx")
                # smsndbdmsfbd()

        # -----------------------------------------
        # CPU FRIENDLY PAUSE (IMPORTANT)
        # -----------------------------------------
        time.sleep(1)   # <-- ADD HERE

def run_1min_scheduler(instrument_key):
    print("\n=== 1 MINUTE SCHEDULER STARTED ===")

    market_open = datetime.time(9, 15, 1)
    market_close = datetime.time(15, 5, 1)

    while True:
        now = datetime.datetime.now()

        # Stop after market close
        if now.time() >= market_close:
            print("Market closed. Scheduler stopped.")
            break

        # Wait until market opens
        if now.time() < market_open:
            time.sleep(1)
            continue

        # --------------------------------------------------
        # Compute NEXT 1-minute boundary (HH:MM:01)
        # --------------------------------------------------
        next_run = (now + datetime.timedelta(minutes=1)).replace(
            second=1,
            microsecond=0
        )

        sleep_seconds = (next_run - now).total_seconds()
        if sleep_seconds > 0:
            time.sleep(sleep_seconds)

        # --------------------------------------------------
        # FIRE EVENT — fetch completed 1-min candle
        # --------------------------------------------------
        fire_time = datetime.datetime.now()
        print(f"[Scheduler] Fetching 1-min candle at {fire_time.time()}")

        intraday_candles = fetch_intraday_candles1(
            instrument_key=instrument_key
        )
        print("inter", intraday_candles)
        save_intraday_to_excel(
            upstox_response=intraday_candles,
            file_name="intraday_data1.xlsx"
        )

# run_1min_scheduler(instrument_key="NSE_EQ|INE117A01022") 
# run_5min_scheduler(instrument_key="NSE_EQ|INE117A01022")


class LoginManager:
    def __init__(self, api_key, api_secret, redirect_uri, token_file="access_token.json"):
        self.api_key = api_key
        self.api_secret = api_secret
        self.redirect_uri = redirect_uri
        self.token_file = token_file
        self.access_token = None

    # ----------------------------------
    # STEP 1: Generate login URL
    # ----------------------------------
    def get_login_url(self):
        return (
            "https://api.upstox.com/v2/login/authorization/dialog"
            f"?response_type=code&client_id={self.api_key}"
            f"&redirect_uri={self.redirect_uri}"
        )

    # ----------------------------------
    # STEP 2: Generate access token
    # ----------------------------------
    def generate_access_token(self, auth_code):
        url = "https://api.upstox.com/v2/login/authorization/token"

        payload = {
            "code": auth_code,
            "client_id": self.api_key,
            "client_secret": self.api_secret,
            "redirect_uri": self.redirect_uri,
            "grant_type": "authorization_code"
        }

        headers = {"Content-Type": "application/x-www-form-urlencoded"}

        response = requests.post(url, data=payload, headers=headers)
        response.raise_for_status()

        data = response.json()
        self.access_token = data["access_token"]
        

        self._save_token(data)
        return self.access_token

    # ----------------------------------
    # STEP 3: Save token
    # ----------------------------------
    def _save_token(self, data):
        with open("access_token.txt", "w") as f:
            f.write(data["access_token"])


# login = LoginManager(API_KEY, SECRET_KEY, RURL)

# print("Login URL:")
# print(login.get_login_url())

# auth_code = input("Enter auth code: ")

# token = login.generate_access_token(auth_code)
# print("Access token generated")


// hello.py
import time
import requests
import csv
from datetime import datetime
from zoneinfo import ZoneInfo
import numpy as np
import pandas as pd
from collections import deque

ACCESS_TOKEN = ""

prev_snapshot = None

def process_oi_and_save(curr_snapshot, nifty_data, futures_data, timestamp, csv_file, strike_filter):
    global prev_snapshot
    nifty = next(iter(nifty_data['data'].values()))
    fut   = next(iter(futures_data['data'].values()))

    ltp = nifty['last_price']
    ltp_fut = fut['last_price']
    buy_qty = fut['total_buy_quantity']
    sell_qty = fut['total_sell_quantity']
    diff_quantity = buy_qty - sell_qty
    diff_price = ltp_fut - ltp

    # --- First snapshot (no delta possible) ---
    if not prev_snapshot:
        prev_snapshot = curr_snapshot.copy()
        # print("prev", prev_snapshot)
        return

    # --- Build one row (IMPORTANT FIX) ---
    row = [timestamp]

    curr_snapshot1 = {
        item['strike_price']: (item['call_oi'], item['put_oi'])
        for item in curr_snapshot['data']['call_put_oi_data_list']
    }

    prev_snapshot1 = {
        item['strike_price']: (item['call_oi'], item['put_oi'])
        for item in prev_snapshot['data']['call_put_oi_data_list']
    }

    for strike in strike_filter:
        call_oi, put_oi = curr_snapshot1.get(strike, (0, 0))
        prev_call, prev_put = prev_snapshot1.get(strike, (0, 0))

        call_delta = call_oi - prev_call
        put_delta  = put_oi - prev_put

        row.extend([call_oi, put_oi, call_delta, put_delta])
    
    row.extend([ltp_fut, buy_qty, sell_qty, ltp, diff_quantity, diff_price])
    # print("row", row)

    # with open(csv_file, "a") as f:
    #     f.write(",".join(map(str, row)) + "\n")
    with open(csv_file, "a", newline="") as f:
        f.write(",".join(map(str, row)) + "\n")

    prev_snapshot.clear()
    prev_snapshot = curr_snapshot.copy()
    # print("prevvvv", prev_snapshot)

def get_market_quote(instrument_key, access_token): 

    url = "https://api.upstox.com/v2/market-quote/quotes"

    params = {
        "instrument_key": instrument_key
    }

    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }

    try:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()

        data = response.json()
        # print("data",data)

        return data

    except requests.exceptions.HTTPError as err:
        print("HTTP Error:", err)
        print("Response:", response.text)
        return None

    except Exception as e:
        print("Error:", str(e))
        return None

def get_market_quote1(instrument_keys, timestamp, access_token): 

    url = "https://api.upstox.com/v2/market-quote/quotes"

    params = {
        "instrument_key": ",".join(instrument_keys)
    }

    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }

    try:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()

        data = response.json()
        # with open("oi_data-copy.csv", "a", newline="") as f:
        #     writer = csv.writer(f)
        #     writer.writerow(row)
        instrument_prices = {
            value["instrument_token"]: value["last_price"]
            for value in data["data"].values()
        }

        last_prices = [
            instrument_prices[instrument_key]
            for instrument_key in instrument_keys
            if instrument_key in instrument_prices
        ]
        row = [timestamp] + last_prices

        with open("oi_data-copy.csv", "a", newline="") as f:
            f.write(",".join(map(str, row)) + "\n")

        return data

    except requests.exceptions.HTTPError as err:
        print("HTTP Error:", err)
        print("Response:", response.text)
        return None

    except Exception as e:
        print("Error:", str(e))
        return None
    
def get_option_chain_oi(access_token, instrument_key, expiry, date):

    url = "https://api.upstox.com/v2/market/oi"

    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Bearer {ACCESS_TOKEN}"
    }

    params = {
        "instrument_key": instrument_key,
        "expiry": expiry,
        "date": date
    }

    try:
        response = requests.get(url, headers=headers, params=params, timeout=10)
        response.raise_for_status()

        return response.json()

    except requests.exceptions.RequestException as e:
        print(f"API Error: {e}")
        return None

def run():
    while True:
        start = time.time()

        strike_filter = [22200, 22250, 22300, 22350, 22400, 22450, 22500, 22550, 22600, 22650, 22700, 22750, 22800, 22850, 22900, 22950, 23000, 23050, 23100, 23150, 23200]

        instrument_key = [
            "NSE_FO|53636", "NSE_FO|53655", "NSE_FO|53656", "NSE_FO|53661", "NSE_FO|53662", "NSE_FO|53663", "NSE_FO|53664", "NSE_FO|53667", "NSE_FO|53668", "NSE_FO|53679", "NSE_FO|53680", "NSE_FO|53681", "NSE_FO|53682", "NSE_FO|53701", "NSE_FO|53702", "NSE_FO|53719", "NSE_FO|53720", "NSE_FO|53749", "NSE_FO|53750", "NSE_FO|53781", "NSE_FO|53782", "NSE_FO|53791", "NSE_FO|53792", "NSE_FO|53793", "NSE_FO|53794", "NSE_FO|53795", "NSE_FO|53796", "NSE_FO|53835", "NSE_FO|53836", "NSE_FO|53859", "NSE_FO|53860", "NSE_FO|53871", "NSE_FO|53872", "NSE_FO|53923", "NSE_FO|53924", "NSE_FO|53925", "NSE_FO|53926", "NSE_FO|53927", "NSE_FO|53928", "NSE_FO|53929", "NSE_FO|53930", "NSE_FO|53931"
        ]

        curr_snapshot = get_option_chain_oi(
            access_token=ACCESS_TOKEN,
            instrument_key="NSE_INDEX|Nifty 50",
            expiry="2026-11-03",
            date="2026-10-28"
        )
    
        data = get_market_quote(instrument_key="NSE_FO|61471", access_token=ACCESS_TOKEN)
        # ts = data['data']['NSE_FO:NIFTY26AUGFUT']['last_trade_time']
        # timestamp = time.strftime("%H:%M:%S", time.localtime(int(ts)/1000))
        timestamp = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%H:%M:%S")
        get_market_quote1(instrument_keys=instrument_key, timestamp=timestamp, access_token=ACCESS_TOKEN)
        data2 = get_market_quote(instrument_key="NSE_INDEX|Nifty 50", access_token=ACCESS_TOKEN)

        process_oi_and_save(curr_snapshot, data2, data, timestamp, "oi_data.csv", strike_filter)

        elapsed = time.time() - start
        sleep_time = max(0, 150 - elapsed)
        time.sleep(sleep_time)

run()


strike_filter = [22200, 22250, 22300, 22350, 22400, 22450, 22500, 22550, 22600, 22650, 22700, 22750, 22800, 22850, 22900, 22950, 23000, 23050, 23100, 23150, 23200]
def write_header(strikes, csv_file):
    header = ["Timestamp"]

    for strike in sorted(strikes):
        header.extend([
            f"{strike}_Call_Price",
            f"{strike}_Put_Price",
            # f"{strike}_Call",
            # f"{strike}_Put",
            # f"{strike}_Call_Delta",
            # f"{strike}_Put_Delta"
        ])

    with open(csv_file, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(header)

# csv_file  = 'oi_data.csv'
csv_file  = 'oi_data-copy.csv'
strikes = strike_filter

# write_header(strikes, csv_file)

def calculate_ma(df, period):
    df[f'{period}_MA'] = df['close'].rolling(period).mean()
    return df

# def check_negative_trend(diff_values):
    # Need exactly 7 values
    if len(diff_values) < 7:
        return False

    x = np.asarray(diff_values[-7:], dtype=np.float64)

    # 1. Every diff must be negative
    if not np.all(x < 0):
        return False

    # 2. Consecutive changes
    changes = np.diff(x)

    # 3. At most ONE positive change
    if np.count_nonzero(changes > 0) > 1:
        return False

    # 4. At least one change <= -25,000
    if np.min(changes) > -25000:
        return False

    return True

def dataframe():
    df = pd.read_csv("customers-100.csv")

    data = {
        'open': df["open"].tolist(),
        'high': df["high"].tolist(),
        'low': df["low"].tolist(),
        'close': df["close"].tolist()
    }

    df1 = pd.DataFrame(data).tail(20)['close'].to_numpy()
    return df1
# # # ten = calculate_ma(df1, period=10)
# # # ten = pd.DataFrame(ten)
# # # sma_10 = ten["10_MA"].to_numpy()
# # # sma_10 = np.round(sma_10, 2)
# # # twenty = calculate_ma(df1, period=20)
# # # sma_20 = twenty["20_MA"].to_numpy()
# # # sma_20 = np.round(sma_20, 2)
# # # print(sma_20)

# # # closes = df1['close'].to_numpy()

def sma(sum10, sum20):
    sma10 = sum10 * 0.1
    sma20 = sum20 * 0.05
    sma10 = sma10.round(2)
    sma20 = sma20.round(2)
    return sma10, sma20

# nifty_ltp = 23947
# CSV_FILE = "customers-100.csv"
# # # diff_values = np.empty(7, dtype=np.float64)

# # # # def update_nifty_sma_direction(sma10_list, sma20_list):
# # # #     prev_sma10 = sma10_list[-2]
# # # #     prev_sma20 = sma20_list[-2]
# # # #     new_sma10 = sma10_list[-1]
# # # #     new_sma20 = sma20_list[-1]

# # # #     if new_sma10 < new_sma20:

# # # #         if (
# # # #             (new_sma10 < prev_sma10) and (new_sma20 < prev_sma20)
# # # #             and nifty_ltp < new_sma10
# # # #         ):
# # # #             direction = -1
# # # #             print("d",direction)

# # # #         else:
# # # #             direction = 0
# # # #             print("dd", direction)

# # # #     elif new_sma10 > new_sma20:

# # # #         if (
# # # #             (new_sma10 > prev_sma10) and (new_sma20 > prev_sma20)
# # # #             and nifty_ltp > new_sma10
# # # #         ):
# # # #             direction = 1
# # # #             print("ddd", direction)

# # # #         else:
# # # #             direction = 0
# # # #             print("dddd", direction)
# # # #     else:
# # # #         direction = 0

# # # #     return direction

# NUM_COLS = 7
# MAX_ROWS = 300
# sma10_list = deque(maxlen=10)
# sma20_list = deque(maxlen=10)
# diff_triggered = False
# diff_list = deque(maxlen=10)

# dtype = [
#     ('Timestamp', 'U20'),
#     ('data', 'f8', (6))
# ]

# data = np.empty(MAX_ROWS, dtype=dtype)

# row_count = 0

# with open(CSV_FILE, "r", newline="") as f:

#     reader = csv.reader(f)

#     # Skip header
#     next(reader)

#     for row in reader:

#         data[row_count]['Timestamp'] = row[0]

#         try:
#             data[row_count]['data'] = np.asarray(
#                 row[1:],
#                 dtype=np.float64
#             )

#         except ValueError:
#             data[row_count]['data'] = np.asarray(
#                 [0.0 if x == '' else x for x in row[1:]],
#                 dtype=np.float64
#             )

#         row_count += 1

# print("Initial rows:", row_count)

# with open(CSV_FILE, "r", newline="") as f:

#     f.seek(0, 2)

#     file_position = f.tell()

# def trend(x):
#     d = np.diff(x)
#     return np.all(d > 0), np.all(d < 0)

# last_check = time.monotonic()
# while True:
#     time.sleep(2)

#     with open(CSV_FILE, "r", newline="") as f: 
#         f.seek(file_position)
#         reader = csv.reader(f)

#         for row in reader:

#             if not row:
#                 continue
#             data[row_count]['data'] = np.asarray(
#                 # row[1:],
#                 [float(v) if v else 0.0 for v in row[1:]],
#                 dtype=np.float64
#             )
#             nifty_ltp = 23947

#             row_count += 1
#             latest = data[row_count - 1]
#             closes = dataframe()
#             sum10 = closes[-10:].sum()
#             sum20 = closes[-20:].sum()
#             sma10, sma20 = sma(sum10, sum20)
#             print(sma10, sma20)
#             sma10_list.append(float(sma10))
#             sma20_list.append(float(sma20))
            
#             if len(sma10_list) >= 3:
#                 p10, p20 = sma10_list[-2], sma20_list[-2]
#                 n10, n20 = sma10_list[-1], sma20_list[-1]
#                 prev_diff = p10 - p20
#                 new_diff  = n10 - n20
#                 print("s", sma10_list)
#                 trends = trend(sma10_list)
#                 print("t", trends)

#                 if prev_diff < 0 and new_diff > 0:
#                     print("BUY CROSS")

#                 if prev_diff > 0 and new_diff < 0:
#                     print("SELL CROSS")
#                 # direction = (
#                 #     (n10 > n20 and n10 > p10 and n20 > p20 and nifty_ltp > n10)
#                 #     - (n10 < n20 and n10 < p10 and n20 < p20 and nifty_ltp < n10)
#                 # )
#                 # print("dir", direction)
#         file_position = f.tell()
#     now = time.monotonic()

#     if now - last_check >= 30:

#         last_check = now

#         if len(sma10_list) >= 1:

#             n10 = sma10_list[-1]

#             diff = nifty_ltp - n10

#             if diff >= 50:
#                 nifty = -1
#                 diff_triggered = True
#                 print("sell nifty", nifty, diff)
#             elif diff_triggered and 0 < diff <= 5:

#                 print("hi")
#                 diff_list.append(float(diff))

#                 # get_last5_call_put_delta_sums(
#                 #     data,
#                 #     row_count,
#                 #     nifty_ltp
#                 # )

#                 diff_triggered = False

#     changes = np.diff(diff_values)

#     if (
#         np.all(diff_values < 0)
#         and np.count_nonzero(changes > 0) <= 1
#         and changes.min() <= -25000
#     ):
#         print("negative trend")

# intraday_low = float('inf')
# intraday_high = float('-inf')

# def get_ltp(instrument_key):
#     url = "https://api.upstox.com/v3/market-quote/ltp"

#     params = {"instrument_key": instrument_key}

#     headers = {
#         "Content-Type": "application/json",
#         "Accept": "application/json",
#         "Authorization": f"Bearer {ACCESS_TOKEN}"
#     }

#     response = requests.get(url, headers=headers, params=params)
#     response.raise_for_status()

#     data = response.json()
#     ltp = data["data"]["NSE_INDEX:Nifty 50"]["last_price"]

#     return ltp

# while True:
#     print("hiiii")

#     ltp = get_ltp(instrument_key="NSE_INDEX|Nifty 50")
#     print("ltp", ltp)

#     old_low = intraday_low
#     old_high = intraday_high

#     if ltp < intraday_low:
#         intraday_low = ltp

#     if ltp > intraday_high:
#         intraday_high = ltp

#     if ltp == old_low or ltp == old_high:
#         print("hi", ltp)

# ltp_list = deque(maxlen=10)

# sample_ltp = [
#     23900, 23908, 23915, 23940, 23945,
#     23920, 23950, 23975, 23980, 24005,
#     24010, 23990, 23970
# ]

# def nifty_direction(nifty_ltp):

#     ltp_list.append(nifty_ltp)

#     if len(ltp_list) < 10:
#         return 0, 0

#     it = iter(ltp_list)
#     prev = next(it)

#     up = down = fast_up = fast_down = 0

#     for ltp in it:
#         diff = ltp - prev

#         up += diff > 0
#         down += diff < 0

#         fast_up += diff > 20
#         fast_down += diff < -20

#         prev = ltp

#     direction = (up > down) - (down > up)
#     fast_direction = (fast_up >= 2) - (fast_down >= 2)

#     return direction, fast_direction

# def test_ltp(sample_ltp):

#     for ltp in sample_ltp:

#         direction, fast_direction = nifty_direction(ltp)

#         print(
#             f"LTP: {ltp} | "
#             f"Stored: {list(ltp_list)} | "
#             f"Direction: {direction} | "
#             f"Fast: {fast_direction}"
#         )

#         time.sleep(10)


  // trade.py
  import numpy as np
from openpyxl.utils import column_index_from_string
from datetime import datetime, time
from zoneinfo import ZoneInfo
import pandas as pd
import csv
import os
from time import monotonic
import time
import requests
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from collections import deque

ACCESS_TOKEN = "eyJ0eXAiOiJKV1QiLCJrZXlfaWQiOiJza192MS4wIiwiYWxnIjoiSFMyNTYifQ.eyJzdWIiOiI0TENQS0siLCJqdGkiOiI2YWQxOWY2ZThjYTQ1MTIxYmE1YWEyZDciLCJpc011bHRpQ2xpZW50IjpmYWxzZSwiaXNQbHVzUGxhbiI6ZmFsc2UsImlhdCI6MTc5MjEyMjczNCwiaXNzIjoidWRhcGktZ2F0ZXdheS1zZXJ2aWNlIiwiZXhwIjoxNzkyMTg4MDAwfQ.IpHzoKllH7YbmcAQA4JxUcFZoUFwK0xLaXAjDu7SqWI"

FIRST_STRIKE = 22000

BASE_DIR = "/home/ec2-user"

# CSV_FILE = os.path.join(BASE_DIR, "oi_data.csv")

def check_signal(call_total_itm, put_total_otm,
                 call_total_otm, put_total_itm):
    

    c_itm = np.asarray(call_total_itm)
    p_otm = np.asarray(put_total_otm)
    c_otm = np.asarray(call_total_otm[:3])
    p_itm = np.asarray(put_total_itm[:3])
    if (np.all((c_itm < 25000) & (p_otm > c_itm)) and \
           np.all((c_otm < 25000) & (p_itm > c_otm))):
        print("up")
    elif (np.all((c_itm > 0) & (p_otm < 0)) and \
           np.all((c_otm > 25000) & (p_itm < c_otm))):
        print("down")

def get_all_indexes(data, row_count, no_of_rows, ltp):
# Last 5 rows, numeric data only
    x = data[row_count - no_of_rows:row_count]['data']
    # print("x", x)

    # ATM
    atm = int((ltp + 25) // 50 * 50)


    # First strike = 1 step below ATM
    # first = atm + 50
    first = atm
    # print("first", first)


    # Starting strike index
    k = (first - FIRST_STRIKE) // 50
    # print("k", k)


    # Call Delta indexes:
    # k*4+2, k*4+6, k*4+10, ...
    start = k * 4 + 2
    startp = k * 2
    startp1 = k * 2 + 1
    return k, start, x, startp, startp1

def get_call_put_delta_sum(k ,start, x):
    call_indexes = start - np.arange(k) * 4
    # call_indexess = np.append(call_indexes, call_indexes[-1] - 4)
    call_indexess = start - np.arange(k + 1) * 4

    # Five consecutive Call Delta columns
    call_totals_otm = x[:, start:start + 20:4].sum(axis=0)
    call_totals_itm = x[:, call_indexess[1:]].sum(axis=0)

    # Five consecutive Put Delta columns
    put_totals_itm = x[:, start + 1:start + 21:4].sum(axis=0)
    put_indexes = call_indexes - 3
    put_totals_otm = x[:, put_indexes].sum(axis=0)
    print(call_totals_itm, put_totals_otm)
    print(call_totals_otm, put_totals_itm)
    return call_totals_itm, call_totals_otm, put_totals_itm, put_totals_otm

def get_call_put_price_list(k ,startp, x):
    print("hi")
    call_indexes = startp - np.arange(k) * 2
    call_indexess = startp - np.arange(k + 1) * 2
    call_price_otm = x[:, startp:startp + 12:2].T
    call_price_itm = x[:, call_indexess[1:]].T[:6]
    put_price_itm = x[:, startp + 1:startp + 13:2].T
    put_indexes = call_indexes - 1
    put_price_otm = x[:, put_indexes].T
    # print(call_price_otm, call_price_itm)
    # print(put_price_otm, put_price_itm)
    return call_price_otm, call_price_itm, put_price_otm, put_price_itm

def get_call_put_delta_sums(data, row_count, no_of_rows, ltp):

    # Last 5 rows, numeric data only
    x = data[row_count - no_of_rows:row_count]['data']
    # print("x", x)

    # ATM
    atm = int((ltp + 25) // 50 * 50)


    # First strike = 1 step below ATM
    # first = atm + 50
    first = atm
    # print("first", first)


    # Starting strike index
    k = (first - FIRST_STRIKE) // 50


    # Call Delta indexes:
    # k*4+2, k*4+6, k*4+10, ...
    start = k * 4 + 2
    call_indexes = start - np.arange(k) * 4
    call_indexess = np.append(call_indexes, call_indexes[-1] - 4)
    # call_indexess = start - np.arange(k + 1) * 4

    # Five consecutive Call Delta columns
    call_totals_otm = x[:, start:start + 20:4].sum(axis=0)
    call_totals_itm = x[:, call_indexess[1:]].sum(axis=0)

    # Five consecutive Put Delta columns
    put_totals_itm = x[:, start + 1:start + 21:4].sum(axis=0)
    put_indexes = call_indexes - 3

    put_totals_otm = x[:, put_indexes].sum(axis=0)

    # print(call_totals_otm, put_totals_itm)
    # print(call_totals_itm, put_totals_otm)
    # diff = np.asarray(call_totals_itm, dtype=np.int64) - np.asarray(put_totals_otm, dtype=np.int64)
    # diff = np.asarray(call_totals_otm, dtype=np.int64) - np.asarray(put_totals_itm, dtype=np.int64)
    # check_signal(call_totals_itm, put_totals_otm, call_totals_otm, put_totals_itm)
        # print("signal")
    # print(diff)
    # timestamp = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%H:%M:%S")
    # if np.all(diff < 0):
    #     print(f"Buy NIFTY at {timestamp}")
    # first - 50 * np.arange(5)

    return call_totals_otm, put_totals_itm, call_totals_itm, put_totals_otm

# nifty_ltp = 23240

NUM_COLS = 43
NUM_COLS1 = 91
MAX_ROWS = 300
MAX_ROWS1 = 300

CSV_FILE = "oi_data-copy.csv"
# CSV_FILE = "oi_data-copy_0710.csv"
CSV_FILE1 = "oi_data.csv"
# CSV_FILE1 = "oi_data_0710.csv"

dtype = [
    ('Timestamp', 'U20'),
    ('data', 'f8', (42))
]

dtype1 = [
    ('Timestamp', 'U20'),
    ('data', 'f8', (90))
]

data = np.empty(MAX_ROWS, dtype=dtype)
data1 = np.empty(MAX_ROWS1, dtype=dtype1)

row_count = 0
row_count1 = 0

with open(CSV_FILE, "r", newline="") as f:

    reader = csv.reader(f)

    # Skip header
    next(reader)

    for row in reader:

        data[row_count]['Timestamp'] = row[0]

        try:
            data[row_count]['data'] = np.asarray(
                row[1:],
                dtype=np.float64
            )

        except ValueError:
            data[row_count]['data'] = np.asarray(
                [0.0 if x == '' else x for x in row[1:]],
                dtype=np.float64
            )

        row_count += 1

    f.seek(0, 2)
    
    file_position = f.tell()

print("Initial rows:", row_count)

with open(CSV_FILE1, "r", newline="") as f:

    reader = csv.reader(f)

    # Skip header
    next(reader)

    for row in reader:

        data1[row_count1]['Timestamp'] = row[0]

        try:
            data1[row_count1]['data'] = np.asarray(
                row[1:],
                dtype=np.float64
            )

        except ValueError:
            data1[row_count1]['data'] = np.asarray(
                [0.0 if x == '' else x for x in row[1:]],
                dtype=np.float64
            )

        row_count += 1                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         ow_count1 += 1

    f.seek(0, 2)
    
    file_position1 = f.tell()

print("Initial rows1:", row_count1)

# with open(CSV_FILE, "r", newline="") as f:

    # f.seek(0, 2)

    # file_position = f.tell()

def trend(x):
    d = np.diff(x)
    return np.all(d > 0), np.all(d < 0)


sma10_list = deque(maxlen=10)
sma20_list = deque(maxlen=10)

def get_ltp(instrument_key):
    url = "https://api.upstox.com/v3/market-quote/ltp"

    params = {"instrument_key": instrument_key}

    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Bearer {ACCESS_TOKEN}"
    }

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()

    data = response.json()
    ltp = data["data"]["NSE_INDEX:Nifty 50"]["last_price"]

    return ltp

def fetch_intraday_candles(instrument_key):
    access_token = ACCESS_TOKEN
    url = f'https://api.upstox.com/v3/historical-candle/intraday/{instrument_key}/minutes/5'

    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Authorization': f'Bearer {access_token}'
    }

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    df = response.json()
    return df

def upstox_candles_to_df(resp):
    candles = resp['data']['candles']

    df = pd.DataFrame(
        candles,
        columns=['time', 'open', 'high', 'low', 'close', 'volume', 'oi']
    )
    df['time'] = pd.to_datetime(df['time']).dt.tz_localize(None)
    return df

def dataframe():
    df = pd.read_csv("customers-100.csv")

    data = {
        'open': df["open"].tolist(),
        'high': df["high"].tolist(),
        'low': df["low"].tolist(),
        'close': df["close"].tolist()
    }

    df1 = pd.DataFrame(data).tail(20)['close'].to_numpy()
    return df1

def save_intraday_to_csv(upstox_response, file_name):
    new_df = upstox_candles_to_df(upstox_response)
    try:
        old_df = pd.read_csv(file_name)
        old_df['time'] = pd.to_datetime(old_df['time']).dt.tz_localize(None)
    except FileNotFoundError:
        new_df.to_csv(file_name, index=False)
        print("[Excel] First candle file created")
        return

    merged = pd.concat([old_df, new_df])
    merged = merged.drop_duplicates(subset=['time'])
    merged = merged.sort_values('time')
    if len(merged) > len(old_df):
        merged.to_csv(file_name, index=False)
    else:
        print("[Excel] No new candles to append")

class CSVHandler(FileSystemEventHandler):

    def __init__(self, row_count, file_position, row_count1, file_position1):
        self.row_count = row_count
        self.row_count1 = row_count1
        self.file_position = file_position
        self.file_position1 = file_position1
        self.long_entry_price = None
        self.long_entry_time = None
        self.short_entry_price = None
        self.short_entry_time = None

    def enter_long(self, ltp):
        self.long_entry_price = float(ltp)
        self.long_entry_time = monotonic()
        print("LONG ENTRY:", ltp)

    def enter_short(self, ltp):
        self.short_entry_price = float(ltp)
        self.short_entry_time = monotonic()
        print("SHORT ENTRY:", ltp)

    def check_long_exit(self, ltp, now):
        if self.long_entry_price is None:
            return False

        if (
            ltp >= self.long_entry_price + 10
            or monotonic() - self.long_entry_time >= 300
        ):
            with open("trade3.txt", "a") as file:
                file.write(f"SELL {now}\n")

            self.long_entry_price = None
            self.long_entry_time = None
            return True

        return False

    def check_short_exit(self, ltp, now):
        if self.short_entry_price is None:
            return False

        if (
            ltp <= self.short_entry_price - 10
            or monotonic() - self.short_entry_time >= 300
        ):
            with open("trade3.txt", "a") as file:
                file.write(f"Buy {now}\n")

            self.short_entry_price = None
            self.short_entry_time = None
            return True

        return False

    def check_trend(self, arr):
        diff = np.diff(arr, axis=1)

        if np.all(diff > 0):
            return True
        elif np.all(diff < 0):
            return True
        
        return False
        
    def on_modified(self, event):
        

        # if event.src_path.endswith(CSV_FILE):

        #     self.process_file_change()

        if os.path.abspath(event.src_path) == "/home/ec2-user/oi_data.csv":
            ltp = get_ltp(instrument_key="NSE_INDEX|Nifty 50")
            self.process_file_change(ltp)

        # elif event.src_path.endswith(CSV_FILE1):

        #     self.process_file_change1()

        elif os.path.abspath(event.src_path) == "/home/ec2-user/oi_data-copy.csv":
            ltp = get_ltp(instrument_key="NSE_INDEX|Nifty 50")
            self.process_file_change1(ltp)

    def process_file_change(self, ltp):

        with open(CSV_FILE, "r", newline="") as f:

            f.seek(self.file_position)

            reader = csv.reader(f)

            for row in reader:

                if not row:
                    continue

                data[self.row_count]['Timestamp'] = row[0]
            
                data[self.row_count]['data'] = np.array(
                    [0.0 if x == '' else float(x) for x in row[1:]],
                    dtype=np.float64
                )

                self.row_count += 1

                print("New row:", self.row_count)
                # ltp = 22701
                if  1 < self.row_count <= 4:
                    latest_2 = data[self.row_count - 2:self.row_count]['data']
                    call_list = [latest_2[:, i].tolist() for i in range(0, latest_2.shape[1], 2)]
                    put_list = [latest_2[:, i].tolist() for i in range(1, latest_2.shape[1], 2)]
                    k, start, x, startp, startp1 = get_all_indexes(data, self.row_count, 3, ltp)
                    call_array = np.asarray(call_list)[:k + 1]
                    put_array = np.asarray(put_list)[:k + 1]
                    c = latest_2[:, 0::2][:startp]
                    p = latest_2[:, 1::2][:startp1]
                    c_ratio = call_array[:, -1] / call_array[:, -2]
                    p_ratio = put_array[:, -1] / put_array[:, -2]
                    timestamp = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%H:%M:%S")
                    if np.all(c_ratio < 1):
                        with open("trade2.txt", "a") as file:
                            file.write(f"selll Nifty {timestamp}\n")
                    elif np.all(p_ratio > 1):
                        with open("trade2.txt", "a") as file:
                            file.write(f"selll Nifty {timestamp}\n")
                    elif np.all(c_ratio > 1):
                        with open("trade2.txt", "a") as file:
                            file.write(f"Nifty buyy {timestamp}\n")
                    elif np.all(p_ratio < 1):
                        with open("trade2.txt", "a") as file:
                            file.write(f"Nifty buyy {timestamp}\n")
                    cr = c[:, -1] / c[:, -2]
                    pr = p[:, -1] / p[:, -2]

                    # if np.all(cr < 1) or np.all(pr > 1):
                    #     with open("trade2.txt", "a") as file:
                    #         file.write(f"selll Nifty {timestamp}\n")

                    # elif np.all(cr > 1) or np.all(pr < 1):
                    #     with open("trade2.txt", "a") as file:
                    #         file.write(f"Nifty buyy {timestamp}\n")
                    return
                # ltp = 22702
                latest_3 = data[self.row_count - 3:self.row_count]['data']
                call_list = [latest_3[:, i].tolist() for i in range(0, latest_3.shape[1], 2)]
                put_list = [latest_3[:, i].tolist() for i in range(1, latest_3.shape[1], 2)]
                k, start, x, startp, startp1 = get_all_indexes(data, self.row_count, 3, self.ltp)
                # print("s", k, start, startp, startp1)
                call_array = np.asarray(call_list)[1:k + 1]
                put_carray = np.asarray(put_list)[1:k + 1]
                ratios = call_array[:, -1] / call_array[:, -2]
                ratios1 = put_carray[:, -1] / put_carray[:, -2]
                ratios2 = call_array[:, -1] / call_array[:, -3]
                ratios3 = put_carray[:, -1] / put_carray[:, -3]
                # print("ratios", call_array, put_carray)
                timestamp = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%H:%M:%S")
                
                if np.all(ratios < 0.97) or np.all(ratios1 > 1.05):
                    with open("trade.txt", "a") as file:
                        file.write(f"Nifty selll {timestamp}\n")
                elif np.all(ratios > 1.02) or np.all(ratios1 <= 0.95):
                    with open("trade.txt", "a") as file:
                        file.write(f"Nifty buyy {timestamp}\n")
          
                if np.all(ratios2 < 0.965) or np.all(ratios3 > 1.05):
                    signal = self.check_trend(call_array)
                    if signal:
                        with open("trade.txt", "a") as file:
                            file.write(f"Nifty sellll {timestamp}\n")
                elif np.all(ratios2 > 1.05) or np.all(ratios3 < 0.965):
                    signal = self.check_trend(call_array)
                    if signal:
                        with open("trade.txt", "a") as file:
                            file.write(f"Nifty buyyy {timestamp}\n")

                # old_low = intraday_low
                # old_high = intraday_high

                # if ltp < intraday_low:
                #     intraday_low = ltp

                # if ltp > intraday_high:
                #     intraday_high = ltp

                # if ltp == old_low or ltp == old_high:
                #     print("hi", ltp)

            # IMPORTANT: before `with` closes the file
            self.file_position = f.tell()

    def process_file_change1(self, ltp):

        with open(CSV_FILE1, "r", newline="") as f:
        
            f.seek(self.file_position1)
        
            reader = csv.reader(f)
        
            for row in reader:
        
                if not row:
                    continue
        
                data1[self.row_count1]['Timestamp'] = row[0]
                    
                data1[self.row_count1]['data'] = np.array(
                    [0.0 if x == '' else float(x) for x in row[1:]],
                    dtype=np.float64
                )
        
                self.row_count1 += 1
                # ltp = 23280
                # ltp = get_ltp(instrument_key="NSE_INDEX|Nifty 50")
                call_totals_otm, put_totals_itm, call_totals_itm, put_totals_otm = get_call_put_delta_sums(data1, self.row_count1, 5, self.ltp)
                # print("o", call_totals_itm)
                # print("o", call_totals_otm)
                # print("o", put_totals_itm)
                # print("o", put_totals_otm)
                timestamp = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%H:%M:%S")
                if np.all(call_totals_itm[1:] < put_totals_otm[1:]):
                    with open("trade2.txt", "a") as file:
                        file.write(f"nifty buy {timestamp}\n")
                if np.all(call_totals_itm > put_totals_otm):
                    with open("trade2.txt", "a") as file:
                        file.write(f"nifty sell {timestamp}\n")
                if np.all(call_totals_otm > put_totals_itm):
                    with open("trade2.txt", "a") as file:
                        file.write(f"nifty sell {timestamp}\n")
        
                print("New row1:", self.row_count1)
            self.file_position1 = f.tell()


handler = CSVHandler(row_count, file_position, row_count1, file_position1)

observer = Observer()
observer.schedule(
    handler,
    path=".",
    recursive=False
)

observer.start()
IST = ZoneInfo("Asia/Kolkata")

# try:
#     while True:
#         now = datetime.now(IST)
#         print(now)

#         if (
#             now.time() >= time.time(10, 30, 10)
#             and now.minute % 5 == 0
#             and now.second == 10
#         ):

#             candles = fetch_intraday_candles(
#                 instrument_key="NSE_INDEX|Nifty 50"
#             )

#             save_intraday_to_csv(
#                 upstox_response=candles,
#                 file_name="customers-100.csv"
#             )
#             time.sleep(1)
#             closes = dataframe()
#             sum10 = closes[-10:].sum()
#             sum20 = closes[-20:].sum()
#             sma10 = sum10 * 0.1
#             sma20 = sum20 * 0.05
#             sma10 = sma10.round(2)
#             sma20 = sma20.round(2)
#             sma10_list.append(float(sma10))
#             sma20_list.append(float(sma20))
#             if len(sma10_list) >= 3:
#                 print("lennnn")
#                 p10, p20 = sma10_list[-2], sma20_list[-2]
#                 n10, n20 = sma10_list[-1], sma20_list[-1]
#                 prev_diff = p10 - p20
#                 new_diff  = n10 - n20
#                 trends = trend(sma10_list)
#                 print("t", trends)
    
#                 if prev_diff < 0 and new_diff > 0:
#                     with open("trade2.txt", "a") as file:
#                         file.write(f"BUY CROSS {now}\n")
    
#                 elif prev_diff > 0 and new_diff < 0:
#                     with open("trade2.txt", "a") as file:
#                         file.write(f"SELL CROSS {now}\n")
#         time.sleep(1)

# except KeyboardInterrupt:
#     observer.stop()

observer.join()  

# while True:

#     with open(CSV_FILE, "r", newline="") as f:

#         f.seek(file_position)

#         reader = csv.reader(f)

#         for row in reader:

#             if not row:
#                 continue

#             data[row_count]['Timestamp'] = row[0]
            
#             data[row_count]['data'] = np.array(
#                 [0.0 if x == '' else float(x) for x in row[1:]],
#                 dtype=np.float64
#             )

#             row_count += 1

#             print("New row:", row_count)

#             latest = data[row_count - 1]['data']
#             latest_5 = data[row_count - 5:row_count]['data']
#             latest_3 = data[row_count - 3:row_count]['data']
#             # second_last = latest_5[:, -2]
#             # trend(second_last)
#             # call_list = [latest_5[:, i].tolist() for i in range(0, latest_5.shape[1], 2)]
#             # call_list = [latest_3[:, i].tolist() for i in range(0, latest_3.shape[1], 2)]
#             # call_list_oi = [latest_5[:, i].tolist() for i in range(2, latest_5.shape[1], 4)]
#             # print("clist", call_list)
#             # put_list  = [latest_5[:, i].tolist() for i in range(1, latest_5.shape[1], 2)]
#             # trend = (
#             #     np.all(np.diff(call_list, axis=1) > 0, axis=1).astype(np.int8)
#             #     -
#             #     np.all(np.diff(call_list, axis=1) < 0, axis=1).astype(np.int8)
#             # )
#             # print("trend", trend)
#             # ratio = call_list[0][-1] / call_list[0][-2]
#             # print("r", ratio)

#             # if ratio < 0.95:
#             #     with open("trade2.txt", "a") as file:
#             #         file.write(f"nifty sell \n")
#             # if np.all(trend == -1) and ratio < 0.94:
#             #     print("sell")

#             # if ratio > 1.05:
#             #     with open("trade2.txt", "a") as file:
#             #         file.write(f"nifty buy \n")
#             # if np.all(trend == 1) and ratio > 1.05:
#             #     print("buy")
#             # all_puts = [put_list]
#             # all_calls = np.asarray(all_calls)[0]
#             # print("latest", all_calls)
#             # print("latest", all_puts)
#             ltp = 23232
#             # inc, dec = trend(latest_5)
#             # if inc:
#             #     print("increasing buy")
#             # elif dec:
#             #     print("decreasing sell")
#             call_totals_otm, put_totals_itm, call_totals_itm, put_totals_otm = get_call_put_delta_sums(data, row_count, 6, ltp)
#             # k, start, x, startp = get_all_indexes(data, row_count, 3, ltp)
#             # call_price_otm, call_price_itm, put_price_otm, put_price_itm = get_call_put_price_list(k, startp, x)
#             # trends = [trend(price_list) for price_list in call_price_itm]
#             # print("t", trends)
#             # if np.all(np.array(trends)[:, 0]):
#             #     print("hiiiii")
#             # if np.all(np.array(trends)[:, 1]):
#             #     print("bye")
#             # ratio1 = call_price_itm[:, -1] / call_price_itm[:, 0]
#             # ratio2 = put_price_otm[:, -1] / put_price_otm[:, 0]
#             # print("r", ratio1, ratio2)
#             # get_call_put_delta_sum(k ,start, x)
#             diff = put_totals_otm - call_totals_itm
#             print(diff)
#             if np.all(call_totals_itm < put_totals_otm):
#             #     diff = put_totals_otm - call_totals_itm
#                 print("hiii")
#             # call_delta_sum = latest[2::4].sum()
#             # put_delta_sum = latest[3::4].sum()
#             # print(call_delta_sum, put_delta_sum)

#         file_position = f.tell()

#     time.sleep(5)

  //trade2.py
  import numpy as np
from openpyxl.utils import column_index_from_string
from datetime import datetime, time
from zoneinfo import ZoneInfo
import pandas as pd
import csv
import os
from time import monotonic
import time
import threading
from datetime import datetime, time as dt_time
import requests
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from collections import deque
from time import now

ACCESS_TOKEN = ""

FIRST_STRIKE = 22000

BASE_DIR = "/home/ec2-user"

# CSV_FILE = os.path.join(BASE_DIR, "oi_data.csv")

def check_signal(call_total_itm, put_total_otm,
                 call_total_otm, put_total_itm):
    

    c_itm = np.asarray(call_total_itm)
    p_otm = np.asarray(put_total_otm)
    c_otm = np.asarray(call_total_otm[:3])
    p_itm = np.asarray(put_total_itm[:3])
    if (np.all((c_itm < 25000) & (p_otm > c_itm)) and \
           np.all((c_otm < 25000) & (p_itm > c_otm))):
        print("up")
    elif (np.all((c_itm > 0) & (p_otm < 0)) and \
           np.all((c_otm > 25000) & (p_itm < c_otm))):
        print("down")

def get_all_indexes(data, row_count, no_of_rows, ltp):
# Last 5 rows, numeric data only
    x = data[row_count - no_of_rows:row_count]['data']
    # print("x", x)

    # ATM
    atm = int((ltp + 25) // 50 * 50)


    # First strike = 1 step below ATM
    # first = atm + 50
    first = atm
    # print("first", first)


    # Starting strike index
    k = (first - FIRST_STRIKE) // 50
    # print("k", k)


    # Call Delta indexes:
    # k*4+2, k*4+6, k*4+10, ...
    start = k * 4 + 2
    startp = k * 2
    startp1 = k * 2 + 1
    return k, start, x, startp, startp1

def get_call_put_delta_sum(k ,start, x):
    call_indexes = start - np.arange(k) * 4
    # call_indexess = np.append(call_indexes, call_indexes[-1] - 4)
    call_indexess = start - np.arange(k + 1) * 4

    # Five consecutive Call Delta columns
    call_totals_otm = x[:, start:start + 20:4].sum(axis=0)
    call_totals_itm = x[:, call_indexess[1:]].sum(axis=0)

    # Five consecutive Put Delta columns
    put_totals_itm = x[:, start + 1:start + 21:4].sum(axis=0)
    put_indexes = call_indexes - 3
    put_totals_otm = x[:, put_indexes].sum(axis=0)
    print(call_totals_itm, put_totals_otm)
    print(call_totals_otm, put_totals_itm)
    return call_totals_itm, call_totals_otm, put_totals_itm, put_totals_otm

def get_call_put_price_list(k ,startp, x):
    print("hi")
    call_indexes = startp - np.arange(k) * 2
    call_indexess = startp - np.arange(k + 1) * 2
    call_price_otm = x[:, startp:startp + 12:2].T
    call_price_itm = x[:, call_indexess[1:]].T[:6]
    put_price_itm = x[:, startp + 1:startp + 13:2].T
    put_indexes = call_indexes - 1
    put_price_otm = x[:, put_indexes].T
    # print(call_price_otm, call_price_itm)
    # print(put_price_otm, put_price_itm)
    return call_price_otm, call_price_itm, put_price_otm, put_price_itm

def get_call_put_delta_sums(data, row_count, no_of_rows, ltp):

    # Last 5 rows, numeric data only
    x = data[row_count - no_of_rows:row_count]['data']
    # print("x", x)

    # ATM
    atm = int((ltp + 25) // 50 * 50)


    # First strike = 1 step below ATM
    # first = atm + 50
    first = atm
    # print("first", first)


    # Starting strike index
    k = (first - FIRST_STRIKE) // 50


    # Call Delta indexes:
    # k*4+2, k*4+6, k*4+10, ...
    start = k * 4 + 2
    call_indexes = start - np.arange(k) * 4
    call_indexess = np.append(call_indexes, call_indexes[-1] - 4)
    # call_indexess = start - np.arange(k + 1) * 4

    # Five consecutive Call Delta columns
    call_totals_otm = x[:, start:start + 20:4].sum(axis=0)
    call_totals_itm = x[:, call_indexess[1:]].sum(axis=0)

    # Five consecutive Put Delta columns
    put_totals_itm = x[:, start + 1:start + 21:4].sum(axis=0)
    put_indexes = call_indexes - 3

    put_totals_otm = x[:, put_indexes].sum(axis=0)

    # print(call_totals_otm, put_totals_itm)
    # print(call_totals_itm, put_totals_otm)
    # diff = np.asarray(call_totals_itm, dtype=np.int64) - np.asarray(put_totals_otm, dtype=np.int64)
    # diff = np.asarray(call_totals_otm, dtype=np.int64) - np.asarray(put_totals_itm, dtype=np.int64)
    # check_signal(call_totals_itm, put_totals_otm, call_totals_otm, put_totals_itm)
        # print("signal")
    # print(diff)
    # timestamp = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%H:%M:%S")
    # if np.all(diff < 0):
    #     print(f"Buy NIFTY at {timestamp}")
    # first - 50 * np.arange(5)

    return call_totals_otm, put_totals_itm, call_totals_itm, put_totals_otm

# nifty_ltp = 23240

NUM_COLS = 43
NUM_COLS1 = 91
MAX_ROWS = 300
MAX_ROWS1 = 300

CSV_FILE = "Book1.csv"

dtype = [
    ('Timestamp', 'U20'),
    ('data', 'f8', (42))
]

dtype1 = [
    ('Timestamp', 'U20'),
    ('data', 'f8', (90))
]

data = np.empty(MAX_ROWS, dtype=dtype)
data1 = np.empty(MAX_ROWS1, dtype=dtype1)

row_count = 0
row_count1 = 0

with open(CSV_FILE, "r", newline="") as f:

    reader = csv.reader(f)

    # Skip header
    next(reader)

    for row in reader:

        data[row_count]['Timestamp'] = row[0]

        try:
            data[row_count]['data'] = np.asarray(
                row[1:],
                dtype=np.float64
            )

        except ValueError:
            data[row_count]['data'] = np.asarray(
                [0.0 if x == '' else x for x in row[1:]],
                dtype=np.float64
            )

        row_count += 1

    f.seek(0, 2)
    
    file_position = f.tell()

print("Initial rows:", row_count)

# with open(CSV_FILE, "r", newline="") as f:

    # f.seek(0, 2)

    # file_position = f.tell()

def trend(x):
    d = np.diff(x)
    return np.all(d > 0), np.all(d < 0)


sma10_list = deque(maxlen=10)
sma20_list = deque(maxlen=10)
spread_list1 = deque(maxlen=10)
spread_list2 = deque(maxlen=10)
ltp_list = deque(maxlen=10)
sma10 = None
sma20 = None

def get_ltp(instrument_key):
    url = "https://api.upstox.com/v3/market-quote/ltp"

    params = {"instrument_key": instrument_key}

    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Bearer {ACCESS_TOKEN}"
    }

    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()

    data = response.json()
    ltp = data["data"]["NSE_INDEX:Nifty 50"]["last_price"]

    return ltp

def fetch_intraday_candles(instrument_key):
    access_token = ACCESS_TOKEN
    url = f'https://api.upstox.com/v3/historical-candle/intraday/{instrument_key}/minutes/5'

    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Authorization': f'Bearer {access_token}'
    }

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    df = response.json()
    return df

def get_4hour_candles(instrument_key, start_date, end_date, access_token):
    url = (
        f"https://api.upstox.com/v3/historical-candle/"
        f"{instrument_key.replace('|', '%7C')}/hours/4/"
        f"{end_date}/{start_date}"
    )

    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers, timeout=10)

    if response.status_code == 200:
        return response.json()["data"]["candles"]

    print(f"Error: {response.status_code} - {response.text}")
    return None

candles = get_4hour_candles(
    instrument_key="NSE_EQ|INE848E01016",
    start_date="2025-01-01",
    end_date="2025-02-01",
    access_token=ACCESS_TOKEN
)

def upstox_candles_to_df(resp):
    candles = resp['data']['candles']

    df = pd.DataFrame(
        candles,
        columns=['time', 'open', 'high', 'low', 'close', 'volume', 'oi']
    )
    df['time'] = pd.to_datetime(df['time']).dt.tz_localize(None)
    return df

def dataframe():
    df = pd.read_csv("customers-100.csv")

    data = {
        'open': df["open"].tolist(),
        'high': df["high"].tolist(),
        'low': df["low"].tolist(),
        'close': df["close"].tolist()
    }

    df1 = pd.DataFrame(data).tail(20)['close'].to_numpy()
    return df1

def save_intraday_to_csv(upstox_response, file_name):
    new_df = upstox_candles_to_df(upstox_response)
    try:
        old_df = pd.read_csv(file_name)
        old_df['time'] = pd.to_datetime(old_df['time']).dt.tz_localize(None)
    except FileNotFoundError:
        new_df.to_csv(file_name, index=False)
        print("[Excel] First candle file created")
        return

    merged = pd.concat([old_df, new_df])
    merged = merged.drop_duplicates(subset=['time'])
    merged = merged.sort_values('time')
    if len(merged) > len(old_df):
        merged.to_csv(file_name, index=False)
    else:
        print("[Excel] No new candles to append")

def price_towards_sma10(ltp_list, sma10_list):
    distance = np.abs(
        np.asarray(ltp_list[-3:]) -
        np.asarray(sma10_list[-3:])
    )

    if np.all(np.diff(distance) < 0):
        if ltp_list[-1] > sma10_list[-1]:
            return "towards_from_above"
        else:
            return "towards_from_below"

    return "moving_away_or_not_consistent"

class CSVHandler(FileSystemEventHandler):

    def __init__(self, row_count, file_position):
        self.row_count = row_count
        self.file_position = file_position
        self.last_processed_candle = None
        self.long_entry_price = None
        self.long_entry_time = None
        self.short_entry_price = None
        self.short_entry_time = None
        self.position = None
        self.prev_sma10 = None
        self.prev_sma20 = None
        self.prev_ltp = None
        self.last_4hour_run = None
     
    def bollinger_bands(self, data, window_size=20):
        rolling_mean = data['close'].rolling(window=window_size).mean()
        rolling_std = data['close'].rolling(window=window_size).std()
    
        data['Upperband'] = rolling_mean + (2 * rolling_std)
        data['Lowerband'] = rolling_mean - (2 * rolling_std)
    
        last_row = data.iloc[-1]
    
        upperband = round(last_row['Upperband'], 2)
        lowerband = round(last_row['Lowerband'], 2)
    
        return upperband, lowerband

    def get_intraday_candles(self):
        now = datetime.now(IST)
        candle_time = (now.date(), now.hour, now.minute // 5)
        if (dt_time(9, 15, 10) <= now.time() <= dt_time(15, 25, 10)    and now.minute % 5 == 0 and now.second >= 10    and candle_time != self.last_processed_candle):
            candles = fetch_intraday_candles(instrument_key="NSE_INDEX|Nifty 50")
            save_intraday_to_csv(upstox_response=candles, file_name="Book1.csv")
            self.last_processed_candle = candle_time
    
    def get_4_hour_candles(self):
        now = datetime.now(IST)
        
        if now.time() >= dt_time(9, 15) and now.time() < dt_time(9, 16):
            run_key = (now.date(), "09:15")
        
        elif now.time() >= dt_time(13, 15) and now.time() < dt_time(13, 16):
            run_key = (now.date(), "13:15")
        
        else:
            run_key = None
        
        if run_key is not None and run_key != self.last_4hour_run = None:
            data = get_4hour_candles(instrument_key="NSE_INDEX|Nifty 50", start_date="2025-01-01", end_date="2025-02-01", access_token=ACCESS_TOKEN)
            self.last_4hour_run = None = run_key
            self.bollinger_bands(data)

    def run_every_second(self):
        while True:
            now = datetime.now(IST)
            self.get_intraday_candles()
            self.get_4_hour_candles()

            if self.position is not None:

                if dt_time(9, 15) <= now.time() <= dt_time(15, 30):
                    ltp = get_ltp(instrument_key="NSE_INDEX|Nifty 50")
                    self.check_trade_hold(sma10_list[-1], sma20_list[-1], ltp)
                    self.check_long_exit(ltp, now)
                    self.check_short_exit(ltp, now)
                    self.check_trade_exit(ltp, sma10_list, sma20_list)

            time.sleep(1)

    def check_long_exit(self, ltp, now):
        if self.long_entry_price is None:
            return False

        if (
            ltp >= self.long_entry_price + 10
            or monotonic() - self.long_entry_time >= 300
        ):
            with open("trade2.txt", "a") as file:
                file.write(f"SELL {now}\n")

            self.long_entry_price = None
            self.long_entry_time = None
            self.position = None
            return True

        return False

    def check_short_exit(self, ltp, now):
        if self.short_entry_price is None:
            return False

        if (
            ltp <= self.short_entry_price - 10
            or monotonic() - self.short_entry_time >= 300
        ):
            with open("trade2.txt", "a") as file:
                file.write(f"Buy {now}\n")

            self.short_entry_price = None
            self.short_entry_time = None
            self.position = None
            return True

        return False

    def check_trend(self, arr):
        diff = np.diff(arr, axis=1)

        if np.all(diff > 0):
            return True
        elif np.all(diff < 0):
            return True
        
        return False
        
    def on_modified(self, event):
        
        # if event.src_path.endswith(CSV_FILE):

        #     self.process_file_change()

        if os.path.abspath(event.src_path) == "/home/ec2-user/Book1.csv":
            ltp = get_ltp(instrument_key="NSE_INDEX|Nifty 50")
            self.process_file_change(ltp)

    def generally_rising(self, values, min_up=6, min_move=5):
        n = len(values)

        if n < 2:
            return False

        up_count = sum(
            values[i] > values[i - 1]
            for i in range(1, n)
        )

        return (
            values[-1] - values[0] >= min_move
            and up_count >= min_up
        )

    def generally_falling(self, values, min_down=6, min_move=5):
        n = len(values)

        if n < 2:
            return False

        down_count = sum(
            values[i] < values[i - 1]
            for i in range(1, n)
        )

        return (
            values[0] - values[-1] >= min_move
            and down_count >= min_down
        )

    def check_bollinger_trade(self, ltp, upperband, lowerband):
        if ltp >= upperband:
            return -1  # SELL
        elif ltp <= lowerband:
            return 1   # BUY
        return 0      # No trade

    def check_trade_hold(self, sma10, sma20, ltp):

        if self.position == "LONG":

            if (
                sma10 > sma20
                and sma10 > self.prev_sma10
                and sma20 >= self.prev_sma20
                and ltp >= self.prev_ltp
            ):
                print("HOLD LONG")
            else:
                print("LONG HOLD CONDITION FAILED")

        elif self.position == "SHORT":

            if (
                sma10 < sma20
                and sma10 < self.prev_sma10
                and sma20 <= self.prev_sma20
                and ltp <= self.prev_ltp
            ):
                print("HOLD SHORT")
            else:
                print("SHORT HOLD CONDITION FAILED")

        # Update previous candle values
        if self.position is not None:
            self.prev_sma10 = sma10
            self.prev_sma20 = sma20
            self.prev_ltp = ltp

    def check_trade_exit(self, ltp, sma10_list, sma20_list):

        sma10 = sma10_list[-1]
        prev_sma10 = sma10_list[-2]

        sma20 = sma20_list[-1]
        prev_sma20 = sma20_list[-2]

        timestamp = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%H:%M:%S")
        # LONG trade exit
        if self.long_entry_price is not None:

            if (sma10 > sma20 and
                sma10 > prev_sma10 and
                sma20 >= prev_sma20 and
                ltp >= sma10):

                diff = ltp - sma10

                if diff <= 5:
                    with open("trade2.txt", "a") as file:
                        file.write(f"Nifty buyyyy exit {timestamp}\n")

                    self.long_entry_price = None
                    self.long_entry_time = None
                    self.position = None

        # SHORT trade exit
        if self.short_entry_price is not None:

            if (sma10 < sma20 and
                sma10 < prev_sma10 and
                sma20 <= prev_sma20 and
                ltp <= sma10):

                diff = sma10 - ltp

                if diff <= 5:
                    with open("trade2.txt", "a") as file:
                        file.write(f"Nifty sellll exit {timestamp}\n")

                    self.short_entry_price = None
                    self.short_entry_time = None
                    self.position = None

    def process_file_change(self, ltp):

        with open(CSV_FILE, "r", newline="") as f:

            f.seek(self.file_position)

            reader = csv.reader(f)

            for row in reader:

                if not row:
                    continue

                data[self.row_count]['Timestamp'] = row[0]
            
                data[self.row_count]['data'] = np.array(
                    [0.0 if x == '' else float(x) for x in row[1:]],
                    dtype=np.float64
                )

                self.row_count += 1

                print("New row:", self.row_count)
                now = datetime.now(IST)

                # if now.time() >= dt_time(9, 15) and now.time() < dt_time(9, 16):
                #     run_key = (now.date(), "09:15")

                # elif now.time() >= dt_time(13, 15) and now.time() < dt_time(13, 16):
                #     run_key = (now.date(), "13:15")

                # else:
                #     run_key = None

                # if run_key is not None and run_key != last_run:
                #     get_4hour_candles(
                #         instrument_key="NSE_INDEX|Nifty 50",
                #         start_date="2025-01-01",
                #         end_date="2025-02-01",
                #         access_token=ACCESS_TOKEN
                #     )
                #     last_run = run_key
                
                # candle_time = (now.date(), now.hour, now.minute // 5)
                # if (
                #     now.time() >= dt_time(9, 15, 10)
                #     and now.minute % 5 == 0
                #     and now.second >= 10
                #     and candle_time != last_processed_candle
                # ):
                    # candles = fetch_intraday_candles(
                    #     instrument_key="NSE_INDEX|Nifty 50"
                    # )

                    # save_intraday_to_csv(
                    #     upstox_response=candles,
                    #     file_name="customers-100.csv"
                    # )

                time.sleep(1)
                ltp = get_ltp(instrument_key="NSE_INDEX|Nifty 50")
                closes = dataframe()
                sma10 = round(float(closes[-10:].sum()) * 0.1, 2)
                sma20 = round(float(closes[-20:].sum()) * 0.05, 2)
                sma10_list.append(sma10)
                sma20_list.append(sma20)
                spread_list1.append(ltp - sma10)
                spread_list2.append(ltp - sma20)
                ltp_list.append(ltp)
                    
                # last_processed_candle = candle_time

                price_towards_sma10(ltp_list, sma10_list)
                timestamp = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%H:%M:%S")
                if (self.generally_rising(sma10_list) and self.generally_rising(sma20_list)):
                    if sma10_list[-1] > sma20_list[-1] and ltp >= sma10_list[-1] + 5:
                        with open("trade2.txt", "a") as file:
                            file.write(f"Nifty buyyyy {timestamp}\n")
                        self.position = "LONG"
                        self.long_entry_price = ltp
                        self.long_entry_time = time.monotonic()

                elif (self.generally_falling(sma10_list) and self.generally_falling(sma20_list)):
                    if sma10_list[-1] < sma20_list[-1] and ltp <= sma10_list[-1] - 5:
                        with open("trade2.txt", "a") as file:
                            file.write(f"Nifty sellll {timestamp}\n")
                        self.position = "SHORT"
                        self.short_entry_price = ltp
                        self.short_entry_time = time.monotonic()
        
                self.check_trade_hold(sma10_list[-1], sma20_list[-1], ltp)
                self.check_long_exit(ltp, now)
                self.check_short_exit(ltp, now)
                self.check_trade_exit(ltp, sma10_list, sma20_list)

                # old_low = intraday_low
                # old_high = intraday_high

                # if ltp < intraday_low:
                #     intraday_low = ltp

                # if ltp > intraday_high:
                #     intraday_high = ltp

                # if ltp == old_low or ltp == old_high:
                #     print("hi", ltp)
            self.file_position = f.tell()

IST = ZoneInfo("Asia/Kolkata")
handler = CSVHandler(row_count, file_position)

threading.Thread(
    target=handler.run_every_second,
    daemon=True
).start()

observer = Observer()
observer.schedule(
    handler,
    path=".",
    recursive=False
)

observer.start()
IST = ZoneInfo("Asia/Kolkata")

observer.join()  

# while True:

#     with open(CSV_FILE, "r", newline="") as f:

#         f.seek(file_position)

#         reader = csv.reader(f)

#         for row in reader:

#             if not row:
#                 continue

#             data[row_count]['Timestamp'] = row[0]
            
#             data[row_count]['data'] = np.array(
#                 [0.0 if x == '' else float(x) for x in row[1:]],
#                 dtype=np.float64
#             )

#             row_count += 1

#             print("New row:", row_count)

#             latest = data[row_count - 1]['data']
#             latest_5 = data[row_count - 5:row_count]['data']
#             latest_3 = data[row_count - 3:row_count]['data']
#             # second_last = latest_5[:, -2]
#             # trend(second_last)
#             # call_list = [latest_5[:, i].tolist() for i in range(0, latest_5.shape[1], 2)]
#             # call_list = [latest_3[:, i].tolist() for i in range(0, latest_3.shape[1], 2)]
#             # call_list_oi = [latest_5[:, i].tolist() for i in range(2, latest_5.shape[1], 4)]
#             # print("clist", call_list)
#             # put_list  = [latest_5[:, i].tolist() for i in range(1, latest_5.shape[1], 2)]
#             # trend = (
#             #     np.all(np.diff(call_list, axis=1) > 0, axis=1).astype(np.int8)
#             #     -
#             #     np.all(np.diff(call_list, axis=1) < 0, axis=1).astype(np.int8)
#             # )
#             # print("trend", trend)
#             # ratio = call_list[0][-1] / call_list[0][-2]
#             # print("r", ratio)

#             # if ratio < 0.95:
#             #     with open("trade2.txt", "a") as file:
#             #         file.write(f"nifty sell \n")
#             # if np.all(trend == -1) and ratio < 0.94:
#             #     print("sell")

#             # if ratio > 1.05:
#             #     with open("trade2.txt", "a") as file:
#             #         file.write(f"nifty buy \n")
#             # if np.all(trend == 1) and ratio > 1.05:
#             #     print("buy")
#             # all_puts = [put_list]
#             # all_calls = np.asarray(all_calls)[0]
#             # print("latest", all_calls)
#             # print("latest", all_puts)
#             ltp = 23232
#             # inc, dec = trend(latest_5)
#             # if inc:
#             #     print("increasing buy")
#             # elif dec:
#             #     print("decreasing sell")
#             call_totals_otm, put_totals_itm, call_totals_itm, put_totals_otm = get_call_put_delta_sums(data, row_count, 6, ltp)
#             # k, start, x, startp = get_all_indexes(data, row_count, 3, ltp)
#             # call_price_otm, call_price_itm, put_price_otm, put_price_itm = get_call_put_price_list(k, startp, x)
#             # trends = [trend(price_list) for price_list in call_price_itm]
#             # print("t", trends)
#             # if np.all(np.array(trends)[:, 0]):
#             #     print("hiiiii")
#             # if np.all(np.array(trends)[:, 1]):
#             #     print("bye")
#             # ratio1 = call_price_itm[:, -1] / call_price_itm[:, 0]
#             # ratio2 = put_price_otm[:, -1] / put_price_otm[:, 0]
#             # print("r", ratio1, ratio2)
#             # get_call_put_delta_sum(k ,start, x)
#             diff = put_totals_otm - call_totals_itm
#             print(diff)
#             if np.all(call_totals_itm < put_totals_otm):
#             #     diff = put_totals_otm - call_totals_itm
#                 print("hiii")
#             # call_delta_sum = latest[2::4].sum()
#             # put_delta_sum = latest[3::4].sum()
#             # print(call_delta_sum, put_delta_sum)

#         file_position = f.tell()

#     time.sleep(5)

  // justscrap.py
  import time
import numpy as np
import pandas as pd
from collections import deque, defaultdict
import websocket
import threading
import requests
import zipfile
import os
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from test import smsndbdmsfbd
from UpstoxAutoLoginShare import run_5min_scheduler
INPUT_FILE = "intraday_data.xlsx" 
OUTPUT_FILE = "output.xlsx"
MERGED_FILE = "merged_output.xlsx"
access_token = ""
API_KEY = 'ghe3ta94ngx0fjye'

def start_5min_scheduler_thread(instrument_key):
    t = threading.Thread(
        target=run_5min_scheduler,
        args=(instrument_key,),
        daemon=True  # dies when main exits
    )
    t.start()
    return t

def load_access_token():
    with open("access_token.txt", "r") as f:
        return f.read().strip()
    
def fetch_upstox_ltp(access_token, instrument_keys):
    """
    Fetch LTP for one or more Upstox instrument keys.

    Args:
        access_token (str): Upstox OAuth access token
        instrument_keys (str | list[str]): Single key or list of keys

    Returns:
        dict: {instrument_key: ltp}
    """

    if isinstance(instrument_keys, str):
        instrument_keys = [instrument_keys]

    url = "https://api.upstox.com/v3/market-quote/ltp"

    params = {
        "instrument_key": ",".join(instrument_keys)
    }

    headers = {
        "Accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers, params=params, timeout=5)

    if response.status_code == 401:
        raise Exception("401 Unauthorized → Invalid or expired Upstox access token")

    response.raise_for_status()

    data = response.json().get("data", {})

    # Return simplified mapping: {instrument_key: ltp}
    return {
        key: info["last_price"]
        for key, info in data.items()
    }


# ltp_data = fetch_upstox_ltp(
#     access_token=access_token,
#     instrument_keys="NSE_EQ|INE117A01022"
# )

# print(ltp_data)

def normalize_price(price):
    if isinstance(price, dict):
        return float(next(iter(price.values())))
    return float(price)

def run_1sec_tick_loop(orch, instrument_key):
    print("[TICK] 1-second tick loop started")

    while True:
        try:
            price = fetch_upstox_ltp(access_token=access_token, instrument_keys="NSE_EQ|INE356A01018")
            price = normalize_price(price)
            orch.on_tick(price)
            # print("price1", price)
        except Exception as e:
            print("[TICK] Error:", e)

        time.sleep(1)

def start_1sec_tick_thread(orch, instrument_key):
    t = threading.Thread(
        target=run_1sec_tick_loop,
        args=(orch, instrument_key),
        daemon=True
    )
    t.start()
    return t

def tick_direction_score(ticks):
    if len(ticks) < 2:
        return 0.5

    up = 0
    down = 0
    for i in range(1, len(ticks)):
        if ticks[i]["ltp"] > ticks[i-1]["ltp"]:
            up += 1
        elif ticks[i]["ltp"] < ticks[i-1]["ltp"]:
            down += 1

    total = up + down
    if total == 0:
        return 0.5

    return up / total    # BUY → near 1, SELL → near 0

def tick_velocity_score(ticks):
    if len(ticks) < 5:
        return 0.5

    start = ticks[0]["ltp"]
    end = ticks[-1]["ltp"]

    pct = (end - start) / start

    if pct > 0.0003:
        return 1.0
    elif pct < -0.0003:
        return 0.0
    return (pct + 0.0003) / (0.0006)

def micro_structure_score(ticks):
    if len(ticks) < 3:
        return 0.5

    highs, lows = [], []

    for i in range(1, len(ticks)-1):
        p = ticks[i-1]["ltp"]
        c = ticks[i]["ltp"]
        n = ticks[i+1]["ltp"]

        if c > p and c > n:
            highs.append(c)
        if c < p and c < n:
            lows.append(c)

    if len(highs) < 2 or len(lows) < 2:
        return 0.5

    hh = sum(1 for i in range(1, len(highs)) if highs[i] > highs[i-1])
    hl = sum(1 for i in range(1, len(lows)) if lows[i] > lows[i-1])

    lh = sum(1 for i in range(1, len(highs)) if highs[i] < highs[i-1])
    ll = sum(1 for i in range(1, len(lows)) if lows[i] < lows[i-1])

    bull = hh + hl
    bear = lh + ll
    total = bull + bear
    if total == 0:
        return 0.5

    return bull / total

def micro_vwap(ticks):
    total_val = sum(t["ltp"] * t["qty"] for t in ticks)
    total_qty = sum(t["qty"] for t in ticks)
    return total_val / total_qty if total_qty else ticks[-1]["ltp"]

def micro_trend_score(ticks):
    if len(ticks) < 5:
        return {"score": 0.5, "trend": "NEUTRAL"}

    d = tick_direction_score(ticks)
    v = tick_velocity_score(ticks)
    s = micro_structure_score(ticks)

    last = ticks[-1]
    tq = last["buy_qty"] + last["sell_qty"]

    orderflow = last["buy_qty"] / tq if tq else 0.5

    mvwap = micro_vwap(ticks)
    drift_raw = (last["ltp"] - mvwap) / mvwap

    if drift_raw > 0.0002:
        drift = 1.0
    elif drift_raw < -0.0002:
        drift = 0.0
    else:
        drift = (drift_raw + 0.0002) / 0.0004

    score = (
        0.4 * d +
        0.2 * v +
        0.2 * s +
        0.1 * orderflow +
        0.1 * drift
    )

    if score >= 0.65:
        trend = "BUY"
    elif score <= 0.35:
        trend = "SELL"
    else:
        trend = "NEUTRAL"

    return {"score": score, "trend": trend}

def vectorised_signal_engine(df):

    # df = df.copy()
    df = pd.read_excel("merged_output.xlsx")
    bool_cols = [
        'BUY_SIGNAL',
        'SELL_SIGNAL',
        'special_buy',
        'special_sell'
    ]

    for col in bool_cols:
        if col in df.columns:
            df[col] = (
                df[col]
                .fillna(False)
                .astype(bool)
            )
    # ensure columns exist
    if 'BUY_SIGNAL' not in df.columns:
        df['BUY_SIGNAL'] = False
    if 'SELL_SIGNAL' not in df.columns:
        df['SELL_SIGNAL'] = False
    if 'buy_score' not in df.columns:
        df['buy_score'] = 0

    # two consecutive signals
    buy_two = df['BUY_SIGNAL'] & df['BUY_SIGNAL'].shift(1, fill_value=False)
    sell_two = df['SELL_SIGNAL'] & df['SELL_SIGNAL'].shift(1, fill_value=False)

    # special buy rule: buy_score==11 -> enter next candle
    special_buy = (df['BUY_SIGNAL']) & (df['buy_score'] == 11)
    special_buy_entry_idx = special_buy.shift(1).astype("boolean").fillna(False)

    buy_pending = buy_two | special_buy_entry_idx
    sell_entry = sell_two

    df['trade_signal'] = np.where(buy_pending, 'BUY',
                          np.where(sell_entry, 'SELL', None))
    df['entry_price'] = np.where(df['trade_signal'].notna(), df['open'], np.nan)

    return df

def read_excel_when_ready(path, timeout=5.0, check_interval=0.2):
    """
    Wait until Excel file is fully written and readable.
    """
    start = time.time()

    last_size = -1
    while time.time() - start < timeout:
        try:
            size = os.path.getsize(path)

            # size unchanged → likely stable
            if size == last_size:
                return pd.read_excel(path)

            last_size = size
            time.sleep(check_interval)

        except (PermissionError, zipfile.BadZipFile):
            time.sleep(check_interval)

    raise TimeoutError(f"Excel file not stable: {path}")

def safe_concat(old_df, new_df):
    if old_df is None or old_df.empty:
        return new_df.copy()

    # Force consistent dtypes
    new_df = new_df.infer_objects(copy=False)

    return pd.concat(
        [old_df, new_df],
        ignore_index=True,
        copy=False
    )

class ExcelCandleWatcher(FileSystemEventHandler):
    def __init__(self, file_path, orchestrator, debounce_sec=1.0):
        self.file_path = file_path
        self.orch = orchestrator
        self.debounce_sec = debounce_sec
        self._last_event_ts = 0.0

    def on_modified(self, event):
        if not event.src_path.endswith(self.file_path):
            return

        now = time.time()
        if now - self._last_event_ts < self.debounce_sec:
            return
        self._last_event_ts = now

        self.process_file_change()

    def process_file_change(self):
        df = read_excel_when_ready(self.file_path)
        print("[WATCHER] Excel modified")
        self.orch.reload_from_dataframe(df)
        smsndbdmsfbd()

class Trade:
    def __init__(self, direction, entry_price, dynamic_sl=True):
        self.direction = direction          # BUY / SELL
        self.entry_price = float(entry_price)
        self.entry_time = time.time()

        # -----------------------------
        # Dynamic SL based on price
        # -----------------------------
        if dynamic_sl:
            self.sl = self.compute_dynamic_sl(entry_price)
        else:
            self.sl = 3

        self.sl_price = (
            entry_price - self.sl if direction == "BUY"
            else entry_price + self.sl
        )

        self.target_points = 6      # base target
        self.target_price = (
            entry_price + self.target_points if direction == "BUY"
            else entry_price - self.target_points
        )

        # tick buffer
        self.ticks = deque(maxlen=30)

        # micro trend flags
        self.extended_hold = False
        self.extend_start_time = None

        # exit
        self.exited = False
        self.exit_price = None
        self.exit_reason = None

    # ------------------------------------------------------
    # Dynamic Stoploss based on price range
    # ------------------------------------------------------
    @staticmethod
    def compute_dynamic_sl(price):
        if price < 1500:
            return 2.0
        elif price < 2500:
            return 3.0
        elif price < 3500:
            return 5.0
        elif price < 6000:
            return 7.0
        return 9.0      # very high priced stocks

    # ------------------------------------------------------
    # TICK FEEDER – called every websocket tick
    # ------------------------------------------------------
    def on_tick(self, tick):
        if self.exited:
            return True, self.exit_reason, self.exit_price

        self.ticks.append(tick)
        ltp = tick["ltp"]

        # -------------------------------------
        # 1. Check SL
        # -------------------------------------
        if self.direction == "BUY" and ltp <= self.sl_price:
            return self._exit(ltp, "Stoploss Hit")

        if self.direction == "SELL" and ltp >= self.sl_price:
            return self._exit(ltp, "Stoploss Hit")

        # -------------------------------------
        # 2. Check Base Target
        # -------------------------------------
        if self.direction == "BUY" and ltp >= self.target_price:
            return self._exit(ltp, "Base Target Hit")

        if self.direction == "SELL" and ltp <= self.target_price:
            return self._exit(ltp, "Base Target Hit")

        # -------------------------------------
        # 3. Micro Trend Logic
        # -------------------------------------
        trend = micro_trend_score(list(self.ticks))
        print("trend", trend)

        # during trade, if micro-trend is strong → extend hold
        if trend["score"] >= 0.70 and not self.extended_hold:
            self.extended_hold = True
            self.extend_start_time = time.time()

        # extended hold for 30 seconds
        if self.extended_hold:
            if time.time() - self.extend_start_time >= 30:
                return self._exit(ltp, "Extended 30s Exit")

        # -------------------------------------
        # 4. Micro Opposite Trend Exit
        # -------------------------------------
        if self.direction == "BUY" and trend["trend"] == "SELL":
            return self._exit(ltp, "Micro Opp Trend Exit")

        if self.direction == "SELL" and trend["trend"] == "BUY":
            return self._exit(ltp, "Micro Opp Trend Exit")

        return False, None, None

    # ------------------------------------------------------
    def _exit(self, price, reason):
        self.exited = True
        self.exit_price = float(price)
        self.exit_reason = reason
        print("exit", self.exit_price, self.exit_reason)
        return True, reason, price

class TickEntryController:
    """
    Collects ticks (per-second) starting when pending entry is created.
    Sliding window of last 20 tick moves (UP/DOWN/FLAT).
    Entry confirmed when 15/20 UP for BUY or 15/20 DOWN for SELL.
    """
    def __init__(self, direction, planned_entry_price=None):
        self.direction = direction  # 'BUY' or 'SELL'
        self.planned_entry_price = planned_entry_price
        self.tick_window = deque(maxlen=20)
        self.last_price = None
        self.entered_price = None
        self.confirmed = False

    def on_tick(self, price):
        """
        Feed a raw price tick.
        Returns:
          None -> not confirmed yet
          dict -> {'direction','entry_price'} when confirmed
        """
        price = float(price)
        if self.last_price is None:
            self.tick_window.append("FLAT")
        else:
            if price > self.last_price:
                self.tick_window.append("UP")
            elif price < self.last_price:
                self.tick_window.append("DOWN")
            else:
                self.tick_window.append("FLAT")
        self.last_price = price

        if len(self.tick_window) < 20:
            return None

        ups = self.tick_window.count("UP")
        downs = self.tick_window.count("DOWN")

        if self.direction == 'BUY' and ups >= 15:
            self.entered_price = price
            self.confirmed = True
            return {'direction': 'BUY', 'entry_price': price}
        if self.direction == 'SELL' and downs >= 15:
            self.entered_price = price
            self.confirmed = True
            return {'direction': 'SELL', 'entry_price': price}

        return None


class AdvancedExitEngine:
    def __init__(self,
                 direction,
                 entry_price,
                 target_price,
                 sl_price,
                 entry_time=None,
                 base_target_pts=6.0,
                 break_even_trigger=1.0,
                 break_even_buffer=0.3,
                 trailing_step=0.5,
                 trend_ticks=15,
                 trend_min=10,
                 rescue_delay=120,
                 rescue_threshold=10,
                 momentum_hold_seconds=30):
        self.direction = direction.upper()
        assert self.direction in ('BUY','SELL')
        self.entry_price = float(entry_price)
        self.entry_time = float(entry_time)
        self.output_file = "output.xlsx"
        self.target_price = target_price
        self.sl_price = sl_price

        price = fetch_upstox_ltp(access_token=access_token, instrument_keys="NSE_EQ|INE356A01018")
        price = normalize_price(price)

        self.base_target_pts = float(base_target_pts)
        self.trend_ticks = int(trend_ticks)
        self.trend_min = int(trend_min)
        self.ticks = deque(maxlen=self.trend_ticks)
        self.tick_qty = deque(maxlen=self.trend_ticks)  # optional qty for vwap if available

        # configs
        self.break_even_trigger = float(break_even_trigger)
        self.break_even_buffer = float(break_even_buffer)
        self.trailing_step = float(trailing_step)
        self.rescue_delay = float(rescue_delay)
        self.rescue_threshold = int(rescue_threshold)
        self.momentum_hold_seconds = float(momentum_hold_seconds)

        # momentum hold
        self.in_extended_hold = False
        self.hold_until = None
        self.momentum_start = None

        # state
        self.exited = False
        self.exit_price = None
        self.exit_reason = None
        self.last_ts = None
        self.tick_count = 0

    def _exit(self, price, reason):
        self.exited = True
        self.exit_price = float(price)
        self.exit_reason = reason

    def _micro_trend_counts(self):
        if len(self.ticks) < 2:
            return 0,0
        arr = np.array(self.ticks)
        diffs = np.diff(arr)
        up = int(np.sum(diffs > 0))
        down = int(np.sum(diffs < 0))
        return up, down
    
    def _append_trade(self, trade: dict):
        try:
            df = pd.read_excel(self.output_file)
        except FileNotFoundError:
            df = pd.DataFrame()

        df.loc[len(df)] = trade
        df.to_excel(self.output_file, index=False)
    
    def _exit_buy_trade(self, price, profit):
        print("exit_buy")
        exit = {
            "direction": "SELL",
            "exit_price": price,
            "profit": profit,
            "exit_time": datetime.now(),
            "status.1": "EXIT"
        }

        print(">>> EXIT CREATED <<<", exit)

        self.active_trade_meta = exit
        self._append_trade(exit)
    
    def _exit_sell_trade(self, price, profit):
        print("exit_sell")
        exit = {
            "direction": "BUY",
            "exit_price": price,
            "profit": profit,
            "exit_time": datetime.now(),
            "status": "EXIT"
        }

        print(">>> EXIT CREATED <<<", exit)

        self.active_trade_meta = exit
        self._append_trade(exit)
    
    def _check_break_even(self, price):
        if self.direction == 'BUY':
            # profit_pts = price - self.entry_price
            profit_pts = 3
            print("ppoints", profit_pts, self.entry_price, self.break_even_trigger)
            if profit_pts >= self.break_even_trigger:
                new_sl = self.entry_price + self.break_even_buffer
                if new_sl > self.sl_price:
                    self.sl_price = new_sl
                    print("new_sl", new_sl)
        else:
            # profit_pts = self.entry_price - price
            profit_pts = 3
            print("lpoints", profit_pts)
            if profit_pts >= self.break_even_trigger:
                new_sl = self.entry_price - self.break_even_buffer
                if new_sl < self.sl_price:
                    self.sl_price = new_sl
                    print("new_sll", new_sl)
    
    def _check_trailing_sl(self, price):
        if self.direction == 'BUY':
            trail_candidate = price - self.trailing_step
            print("tcan", trail_candidate)
            if trail_candidate > self.sl_price:
                self.sl_price = trail_candidate
        else:
            trail_candidate = price + self.trailing_step
            print("tr", trail_candidate)
            if trail_candidate < self.sl_price:
                self.sl_price = trail_candidate
        return trail_candidate

    def _detect_micro_trend(self, ts=None):
        now = ts if ts is not None else time.time()
        # print("now", now)
        self.last_ts = now
        if len(self.ticks) >= self.trend_ticks:
            up, down = self._micro_trend_counts()
            print("ud", up, down)
            # strong momentum in our favour -> start/extend hold
            if (self.direction == 'BUY' and up >= self.trend_min) or (self.direction == 'SELL' and down >= self.trend_min):
                # start or extend momentum hold window
                self.in_extended_hold = True
                self.momentum_start = self.momentum_start or now
                self.hold_until = now + self.momentum_hold_seconds
                print("hold", self.hold_until)
            else:
                # if no momentum and not inside a hold window, reset momentum_start
                if not (self.in_extended_hold and now < (self.hold_until or 0)):
                    self.momentum_start = None

    # def _rescue(self, price, ts=None):
    #     now = ts if ts is not None else time.time()
    #     # print("now", now)
    #     self.last_ts = now
    #     time_since_entry = now - self.entry_time
    #     in_loss = (price < self.entry_price) if (self.direction=='BUY') else (price > self.entry_price)
    #     if time_since_entry >= self.rescue_delay and in_loss and len(self.ticks) >= self.trend_ticks:
    #         up, down = self._micro_trend_counts()
    #         if (self.direction == 'BUY' and down >= self.rescue_threshold) or (self.direction == 'SELL' and up >= self.rescue_threshold):
    #             self._exit(price, "Tick Rescue Exit (opposite micro-trend)")
    #             return True, {"reason": self.exit_reason, "exit_price": self.exit_price}  
            
    def _on_tick(self, price, qty=None, ts=None):
        self.df = pd.read_excel("output.xlsx")
        last_row = self.df.iloc[-1]
        entry_time = last_row.get("entry_time").timestamp()

        if self.exited:
            return True, {"reason": self.exit_reason, "exit_price": self.exit_price}

        now = ts if ts is not None else time.time()
        self.last_ts = now
        self.tick_count += 1

        self.ticks.append(float(price))
        # print("stick", self.tick_count, self.ticks)
        if qty is None:
            self.tick_qty.append(1)
            print(self.tick_qty)
        else:
            self.tick_qty.append(qty)
        
        time_since_entry = entry_time - now
        print("int", int(time_since_entry))
        if int(time_since_entry) < 19750:
            self._check_break_even(price)
        
        if int(time_since_entry) < 19700:
            if self.direction == 'BUY':
                # time_since_entry = datetime.fromtimestamp(time_since_entry).strftime("%H:%M:%S")
                print("BREAK", self.sl_price)
                if price >= self.target_price:
                    print("more")
                    profit_pts = price - self.entry_price
                    self._exit_buy_trade(price, profit_pts)
                    return True, {"reason": self.exit_reason, "exit_price": self.exit_price}
                if price <= self.sl_price:
                    print("less")
                    profit_pts = price - self.entry_price  
                    self._exit_buy_trade(price, profit_pts)
                    return True, {"reason": self.exit_reason, "exit_price": self.exit_price}
                if self.sl_price < price < self.target_price and int(time_since_entry) < 19600:
                    print("IN RANGE → manage trade")
                    self._exit_buy_trade(price, profit_pts = (price - self.entry_price))
            else:
                print("BREAKK", self.sl_price, self.target_price)
                if price <= self.target_price:
                    print("more1")
                    profit_pts = self.entry_price - price
                    self._exit_sell_trade(price, profit_pts)
                    return True, {"reason": self.exit_reason, "exit_price": self.exit_price}
                if price >= self.sl_price:
                    print("less1")
                    profit_pts = 3
                    self._exit_sell_trade(price, profit_pts)
                    return True, {"reason": self.exit_reason, "exit_price": self.exit_price}
                if self.target_price < price < self.sl_price and int(time_since_entry) < 19600:
                    print("IN RANGE → manage trade")
                    profit_pts = price - self.entry_price
                    self._exit_sell_trade(price, profit_pts)

        # # trailing
        # if self.direction == 'BUY':
        #     trail_candidate = price - self.trailing_step
        #     print("tcan", trail_candidate)
        #     if trail_candidate > self.sl_price:
        #         self.sl_price = trail_candidate
        # else:
        #     trail_candidate = price + self.trailing_step
        #     print("tr", trail_candidate)
        #     if trail_candidate < self.sl_price:
        #         self.sl_price = trail_candidate

        # # micro trend detection
        # if len(self.ticks) >= self.trend_ticks:
        #     up, down = self._micro_trend_counts()
        #     print("ud", up, down)
        #     # strong momentum in our favour -> start/extend hold
        #     if (self.direction == 'BUY' and up >= self.trend_min) or (self.direction == 'SELL' and down >= self.trend_min):
        #         # start or extend momentum hold window
        #         self.in_extended_hold = True
        #         self.momentum_start = self.momentum_start or now
        #         self.hold_until = now + self.momentum_hold_seconds
        #         print("hold", self.hold_until)
        #     else:
        #         # if no momentum and not inside a hold window, reset momentum_start
        #         if not (self.in_extended_hold and now < (self.hold_until or 0)):
        #             self.momentum_start = None

        # ######if in extended hold, exit after hold expires
        # if self.in_extended_hold and (now >= (self.hold_until or 0)):
        #     self._exit(price, "Momentum Hold Expired")
        #     return True, {"reason": self.exit_reason, "exit_price": self.exit_price}

        # rescue logic
        # time_since_entry = now - self.entry_time
        # in_loss = (price < self.entry_price) if (self.direction=='BUY') else (price > self.entry_price)
        # if time_since_entry >= self.rescue_delay and in_loss and len(self.ticks) >= self.trend_ticks:
        #     up, down = self._micro_trend_counts()
        #     if (self.direction == 'BUY' and down >= self.rescue_threshold) or (self.direction == 'SELL' and up >= self.rescue_threshold):
        #         self._exit(price, "Tick Rescue Exit (opposite micro-trend)")
        #         return True, {"reason": self.exit_reason, "exit_price": self.exit_price}

        return False, None

    def info(self):
        return {
            "direction": self.direction,
            "entry_price": self.entry_price,
            "target_price": self.target_price,
            "sl_price": self.sl_price,
            "exited": self.exited,
            "exit_price": self.exit_price,
            "exit_reason": self.exit_reason,
            "ticks_tracked": list(self.ticks)
        }


class TradingOrchestrator:

    def __init__(self, input_file, output_file, base_target_pts=6.0, extend_target=6.0):
        self.input_file = input_file
        self.output_file = output_file
        self.EXTEND_TARGET = extend_target

        # ---------- STATE ----------
        self.df = pd.DataFrame()
        self.pending_entry = None
        self.active_exit_engine = None
        self.active_trade_meta = None
        self.base_target_pts=base_target_pts

        # ---------- INIT ----------
        # self._load_initial_candles()
    @staticmethod
    def _sl_points_by_band(price):
        # print("bdjfbjdfds")
        if price < 2000:
            return 3
        elif price < 3000:
            return 5
        elif price < 5000:
            return 7
        else:
            return 9

    def is_entry_time(self):
        now = datetime.now()
        return (
            now.hour == 9 and
            now.minute >= 26 and
            (now.minute - 26) % 5 == 0
        )

    def _load_initial_candles(self):
        print("hiiii")
        try:
            self.df = pd.read_excel("intraday_data.xlsx")
            self.df = vectorised_signal_engine(self.df)
            print("vect")
            print(f"[ORCH] Loaded {len(self.df)} candles")
            last_row = self.df.iloc[-1]
            signal = last_row.get("trade_signal")
            print("signal", signal)
            if signal in ("BUY", "SELL") and self.active_trade_meta is None:
                self._enter_trade_immediately(last_row)
        except FileNotFoundError:
            self.df = pd.DataFrame()
            print("[ORCH] Waiting for candle file...")

    def reload_from_dataframe(self, new_df: pd.DataFrame):
        if new_df.empty:
            return

        if len(new_df) <= len(self.df):
            return

        appended = new_df.iloc[len(self.df):].copy()
        self.df = new_df.copy()

        for _, row in appended.iterrows():
            self.on_new_candle(row.to_dict())

    def on_new_candle(self, candle: dict):
        self.df = pd.read_excel(self.input_file)
        self.df = vectorised_signal_engine(self.df)
        last_row = self.df.iloc[-1]
        signal = last_row.get("trade_signal")
        print("signal", signal)

        if signal in ("BUY", "SELL") and self.active_trade_meta is None:
            # self.pending_entry = last_row
            # print("Signal stored. Waiting for entry time.")
            self._enter_trade_immediately(last_row)

    def _enter_trade_immediately(self, row):
        price = fetch_upstox_ltp(access_token=access_token, instrument_keys="NSE_EQ|INE356A01018")
        price = normalize_price(price)
        target_price = (price + self.base_target_pts) if row["trade_signal"]=='BUY' else (price - self.base_target_pts)
        sl_points = self._sl_points_by_band(price)
        sl_price = (price - sl_points) if row["trade_signal"]=='BUY' else (price + sl_points)
        entry = {
            "direction": row["trade_signal"],
            "entry_price": price,
            "entry_time": row.get("timestamp", datetime.now()),
            "status": "OPEN",
            "target_price": target_price,
            "sl_price": sl_price
        }
        print(">>> ENTRY CREATED <<<", entry)

        self.active_trade_meta = entry
        entry_price = self.active_trade_meta["entry_price"]
        entry_time = datetime.now().timestamp()
        direction = self.active_trade_meta["direction"]
        entry_price = normalize_price(entry_price)
        target_price = self.active_trade_meta["target_price"]
        sl_price = self.active_trade_meta["sl_price"]

        self.active_exit_engine = AdvancedExitEngine(
            direction=direction,
            entry_price = entry_price,
            entry_time=entry_time,
            base_target_pts=self.EXTEND_TARGET,
            target_price=target_price,
            sl_price=sl_price
        )
        
        self._append_trade(entry)
    
    def on_tick(self, price, qty=None, ts=None):
        if not self.active_trade_meta:
            return
        
        print("[TICK] price =", price)
        now = ts if ts is not None else time.time()

        # if self.pending_entry and self.active_trade_meta is None and self.is_entry_time():
        #     print("ENTRY TIME → executing stored signal")
        #     self._enter_trade_immediately(self.pending_entry)
        #     self.pending_entry = None

        if self.active_exit_engine is not None:
            exited, info = self.active_exit_engine._on_tick(price=price, qty=qty, ts=now)
        
        if exited == True:
            self.active_exit_engine = None
            self.active_trade_meta = None

    def _append_trade(self, trade: dict):
        try:
            df = pd.read_excel(self.output_file)
        except FileNotFoundError:
            df = pd.DataFrame()

        df.loc[len(df)] = trade
        df.to_excel(self.output_file, index=False)

    def status(self):
        return {
            "pending": self.pending_entry is not None,
            "active_trade": self.active_trade_meta,
        }


if __name__ == "__main__":
    # Orchestrator always reads enriched / merged candles
    INPUT_FILE1 = "merged_output.xlsx"
    orch = TradingOrchestrator(INPUT_FILE1, OUTPUT_FILE)

    start_5min_scheduler_thread(
        instrument_key="NSE_EQ|INE356A01018"
    )

    start_1sec_tick_thread(
        orch=orch,
        instrument_key="NSE_EQ|INE356A01018"
    )

    # Watch raw candle input file
    handler = ExcelCandleWatcher(
        file_path=INPUT_FILE,
        orchestrator=orch,
        debounce_sec=1.0
    )

    observer = Observer()
    observer.schedule(handler, ".", recursive=False)
    observer.start()

    print("Pipeline active:")
    print("INPUT_FILE  -> abc() -> MERGED_FILE -> TradingOrchestrator")
    print("Waiting for new candles...")

    observer.join()

# INE003A01024
# INE356A01018

  // test.py
  import pandas as pd

def calculate_dema(df, period):
    ema = df['close'].ewm(span=period, adjust=False).mean()
    dema = 2 * ema - ema.ewm(span=period, adjust=False).mean()
    df[f'{period}_DEMA'] = dema
    return df

def calculate_macd(df, fast=3, slow=10, signal=16):
    df['EMA_fast'] = df['close'].ewm(span=fast, adjust=False).mean()
    df['EMA_slow'] = df['close'].ewm(span=slow, adjust=False).mean()
    df['MACD'] = df['EMA_fast'] - df['EMA_slow']
    df['Signal'] = df['MACD'].ewm(span=signal, adjust=False).mean()
    df['Histogram'] = df['MACD'] - df['Signal']
    return df


# def smsndbdmsfbd():
#     df = pd.read_excel("intraday_data.xlsx")

#     data = {
#         'open': df["Open"].tolist(),
#         'high': df["High"].tolist(),
#         'low': df["Low"].tolist(),
#         'close': df["Close"].tolist()
#     }

#     df1 = pd.DataFrame(data)
#     twenty = calculate_dema(df1, period=10)
#     twenty = pd.DataFrame(twenty)
#     macd = calculate_macd(df1, fast=3, slow=10, signal=16)
#     macd  = pd.DataFrame(macd)
#     macd = macd.round(2)
#     # macd.to_string('output.txt', index=False)
#     combined = pd.concat([twenty, macd.drop(columns=['open', 'high', 'low', 'close', '10_DEMA'])], axis=1)
    
#     cond1 = df1['close'].shift(1) < df1['high']
#     cond2 = (df1['high'] - df1['10_DEMA']) <= 2
#     cond3 = ((df1['MACD'] < df1['MACD'].shift(1)) & 
#              (df1['MACD'].shift(1) < df1['MACD'].shift(2))) | (df1['MACD'] < 0)
#     cond4 = ((df1['Histogram'] < df1['Histogram'].shift(1)) & 
#              (df1['Histogram'].shift(1) < df1['Histogram'].shift(2))) | (df1['Histogram'] < 0)
#     df1['sell_score'] = (1.0*cond1 + 1.0*cond2 + 1.8*cond3 + 1.8*cond4)
#     df1['SELL_SIGNAL'] = df1['sell_score'] >= 3
    
#     combined = pd.concat([twenty, macd.drop(columns=['open', 'high', 'low', 'close', '10_DEMA']), df1.drop(columns=['open', 'high', 'low', 'close', '10_DEMA', 'EMA_fast', 'EMA_slow', 'Signal', 'Histogram'])], axis=1)
#     combined = combined.round(2)
#     columns_to_save = ['open', 'high', 'low', 'close', '10_DEMA', 'EMA_fast', 'EMA_slow', 'MACD', 'Signal', 'Histogram', 'sell_score', 'SELL_SIGNAL']
#     # combined.to_csv("output.txt", sep='\t', index=False)
#     combined = combined.loc[:, ~combined.columns.duplicated()]
#     combined.to_excel("sell.xlsx", columns=columns_to_save, index=False)

#     df1['close_prev'] = df1['close'].shift(1)
#     df1['MACD_prev'] = combined['MACD'].shift(1)
#     df1['Hist_prev'] = combined['Histogram'].shift(1)
#     df1['DEMA_prev'] = combined['10_DEMA'].shift(1)

#     combined = pd.DataFrame(combined)

#     # Vectorized boolean masks
#     cond_price_rising     = (df1['close_prev'] <= df1['close']) * 2
#     cond_above_dema       = ((df1['high'] - df1['10_DEMA']) >= 2) * 2
#     cond_macd_positive    = (combined['MACD'] > 0) * 1
#     cond_macd_increasing  = (combined['MACD'] > df1['MACD_prev']) * 2
#     cond_hist_positive    = (combined['Histogram'] > 0) * 1
#     cond_hist_increasing  = (combined['Histogram'] > df1['Hist_prev']) * 1
#     cond_dema_rising      = (combined['10_DEMA'] > df1['DEMA_prev']) * 1
#     cond_trend_order      = ((df1['close'] > combined['EMA_fast']) & (combined['EMA_fast'] > combined['EMA_slow'])) * 1

#     df1['buy_score'] = (
#         cond_price_rising +
#         cond_above_dema +
#         cond_macd_positive +
#         cond_macd_increasing +
#         cond_hist_positive +
#         cond_hist_increasing +
#         cond_dema_rising +
#         cond_trend_order
#     )

#     # Final buy signal flag
#     df1['BUY_SIGNAL'] = df1['buy_score'] >= 6
#     output_cols = ['open', 'high', 'low', 'close', '10_DEMA', 'MACD', 'Histogram', 'buy_score', 'BUY_SIGNAL']
#     output_df = df1[output_cols].copy()
#     output_df.to_excel("buy_output.xlsx", index=False)

#     df2 = pd.read_excel("sell.xlsx")
#     df3 = pd.read_excel("buy_output.xlsx")
#     combined_df = pd.concat([df2, df3], axis=1)
#     combined_df = combined_df.loc[:, ~combined_df.columns.duplicated()]
#     output_file = "merged_output.xlsx"
#     combined_df.to_excel(output_file, index=False)

    # df = pd.read_excel("buy_output.xlsx")   # replace with your file name
    # df.to_csv("buy_signal_output.txt", sep='\t', index=False)

# while True :
# smsndbdmsfbd()

def smsndbdmsfbd():
    df = pd.read_excel("intraday_data.xlsx")

    data = {
        'open': df["open"].tolist(),
        'high': df["high"].tolist(),
        'low': df["low"].tolist(),
        'close': df["close"].tolist()
    }

    df1 = pd.DataFrame(data)
    # print("df1", df1)
    twenty = calculate_dema(df1, period=10)
    twenty = pd.DataFrame(twenty)
    macd = calculate_macd(df1, fast=3, slow=10, signal=16)
    macd  = pd.DataFrame(macd)
    macd = macd.round(2)
    # print("macd", macd)
    # macd.to_string('output.txt', index=False)
    combined = pd.concat([twenty, macd.drop(columns=['open', 'high', 'low', 'close', '10_DEMA'])], axis=1)
    
    cond1 = df1['close'].shift(1) < df1['high']
    cond2 = (df1['high'] - df1['10_DEMA']) <= 2
    cond3 = ((df1['MACD'] < df1['MACD'].shift(1)) & 
             (df1['MACD'].shift(1) < df1['MACD'].shift(2))) | (df1['MACD'] < 0)
    cond4 = ((df1['Histogram'] < df1['Histogram'].shift(1)) & 
             (df1['Histogram'].shift(1) < df1['Histogram'].shift(2))) | (df1['Histogram'] < 0)
    df1['sell_score'] = (1.0*cond1 + 1.0*cond2 + 1.8*cond3 + 1.8*cond4)
    df1['SELL_SIGNAL'] = df1['sell_score'] >= 3
    
    combined = pd.concat([twenty, macd.drop(columns=['open', 'high', 'low', 'close', '10_DEMA']), df1.drop(columns=['open', 'high', 'low', 'close', '10_DEMA', 'EMA_fast', 'EMA_slow', 'Signal', 'Histogram'])], axis=1)
    combined = combined.round(2)
    # print("combined", combined)
    columns_to_save = ['open', 'high', 'low', 'close', '10_DEMA', 'EMA_fast', 'EMA_slow', 'MACD', 'Signal', 'Histogram', 'sell_score', 'SELL_SIGNAL']
    # combined.to_csv("output.txt", sep='\t', index=False)
    combined = combined.loc[:, ~combined.columns.duplicated()]
    combined.to_excel("sell.xlsx", columns=columns_to_save, index=False)

    df1['close_prev'] = df1['close'].shift(1)
    df1['MACD_prev'] = combined['MACD'].shift(1)
    df1['Hist_prev'] = combined['Histogram'].shift(1)
    df1['DEMA_prev'] = combined['10_DEMA'].shift(1)

    combined = pd.DataFrame(combined)
    # print("Como", combined)

    # Vectorized boolean masks
    cond_price_rising     = (df1['close_prev'] <= df1['close']) * 2
    cond_above_dema       = ((df1['high'] - df1['10_DEMA']) >= 2) * 2
    cond_macd_positive    = (combined['MACD'] > 0) * 1
    cond_macd_increasing  = (combined['MACD'] > df1['MACD_prev']) * 2
    cond_hist_positive    = (combined['Histogram'] > 0) * 1
    cond_hist_increasing  = (combined['Histogram'] > df1['Hist_prev']) * 1
    cond_dema_rising      = (combined['10_DEMA'] > df1['DEMA_prev']) * 1
    cond_trend_order      = ((df1['close'] > combined['EMA_fast']) & (combined['EMA_fast'] > combined['EMA_slow'])) * 1

    df1['buy_score'] = (
        cond_price_rising +
        cond_above_dema +
        cond_macd_positive +
        cond_macd_increasing +
        cond_hist_positive +
        cond_hist_increasing +
        cond_dema_rising +
        cond_trend_order
    )

    # Final buy signal flag
    df1['BUY_SIGNAL'] = df1['buy_score'] >= 6
    output_cols = ['open', 'high', 'low', 'close', '10_DEMA', 'MACD', 'Histogram', 'buy_score', 'BUY_SIGNAL']
    output_df = df1[output_cols].copy()
    output_df.to_excel("buy_output.xlsx", index=False)
    # print("buy", output_df)

    sell_df = pd.read_excel("sell.xlsx")
    buy_df = pd.read_excel("buy_output.xlsx")
    # combined_df = pd.concat([df2, df3], axis=1, ignore_index=True)
    # combined_df = combined_df.loc[:, ~combined_df.columns.duplicated()]
    buy_cols = list(buy_df.columns)
    sell_cols = list(sell_df.columns)

    common_cols = [c for c in buy_cols if c in sell_cols]
    buy_only_cols = [c for c in buy_cols if c not in common_cols]
    sell_only_cols = [c for c in sell_cols if c not in common_cols]

    # Create merged dataframe manually
    merged_df = pd.DataFrame()

    # 1️⃣ Common columns (from buy file)
    for col in common_cols:
        merged_df[col] = buy_df[col]

    # 2️⃣ Buy-only columns
    for col in buy_only_cols:
        merged_df[col] = buy_df[col]

    # 3️⃣ Sell-only columns
    for col in sell_only_cols:
        merged_df[col] = sell_df[col]

    # Write output (with headers)
    output_file = "merged_output.xlsx"
    merged_df.to_excel(output_file, index=False)

smsndbdmsfbd()

    # def _load_initial_candles(self):
    #     print("hiiii")
    #     try:
    #         self.df = pd.read_excel(self.input_file)
    #         self.df = vectorised_signal_engine(self.df)
    #         print("vect")
    #         print(f"[ORCH] Loaded {len(self.df)} candles")
    #         last_row = self.df.iloc[-1]
    #         signal = last_row.get("trade_signal")
    #         if signal in ("BUY", "SELL") and self.active_trade_meta is None:
    #             self._enter_trade_immediately(last_row)
    #     except FileNotFoundError:
    #         self.df = pd.DataFrame()
    #         print("[ORCH] Waiting for candle file...")

    # # ---------------------------------------------------------
    # # FILE / DATAFRAME RELOAD (called by watcher)
    # # ---------------------------------------------------------
    # def reload_from_dataframe(self, new_df: pd.DataFrame):
    #     if new_df.empty:
    #         return

    #     if len(new_df) <= len(self.df):
    #         return

    #     appended = new_df.iloc[len(self.df):].copy()
    #     self.df = new_df.copy()

    #     for _, row in appended.iterrows():
    #         self.on_new_candle(row.to_dict())

# class TradingOrchestrator:
#     """
#     Orchestrates:
#     - vectorised candle detection (loads input excel)
#     - creates pending tick tracker when a new candle has trade_signal
#     - receives per-second ticks via .on_tick(price, qty=None, ts=None)
#     - confirms entry using TickEntryController
#     - once entry confirmed, creates AdvancedExitEngine and monitors ticks until exit
#     - writes result to output excel on trade close
#     """

#     def __init__(self, input_file, output_file, extend_target=6.0):
#         self.input_file = input_file
#         self.output_file = output_file
#         self.EXTEND_TARGET = extend_target

#         # load candles
#         self.df = pd.read_excel(self.input_file)
#         print("xcxc", self.df)
#         # ensure necessary columns
#         for col in ['buy_score','BUY_SIGNAL','sell_score','SELL_SIGNAL','10_DEMA','MACD','Histogram']:
#             if col not in self.df.columns:
#                 self.df[col] = 0
#         # initial vectorised detection
#         self.df = vectorised_signal_engine(self.df)

#         # runtime state
#         self.pending_entry = None   # dict {row_index, direction, planned_entry_price, controller}
#         self.active_exit_engine = None  # AdvancedExitEngine instance
#         self.active_trade_meta = None   # meta info for logging (row index, start ts)
#         self.trades_log = []  # list of trade dicts
#         self.row_ticks = defaultdict(list)

    
    
#     # def reload_from_file(self):

#     #     new_df = pd.read_excel(self.input_file)
#     #     print("cxcxc", new_df.tail(2))

#     #     # First load safety
#     #     if len(new_df) <= len(self.df):
#     #         return

#     #     # Identify appended rows
#     #     appended = new_df.iloc[len(self.df):].copy()

#     #     # Replace df reference
#     #     self.df = new_df.copy()

#     #     # Process each new candle sequentially
#     #     for _, row in appended.iterrows():
#     #         self.on_new_candle(row.to_dict())
    
#     def reload_from_dataframe(self, new_df):
#         print("[ORCH] Reload dataframe")
#         print("Tail:\n", new_df.tail(2))

#         if self.df is not None and len(new_df) <= len(self.df):
#             return

#         # Identify appended rows
#         old_len = 0 if self.df is None else len(self.df)
#         appended = new_df.iloc[old_len:].copy()

#         # Replace df
#         self.df = new_df.copy()
#         # self.df = safe_concat(self.df, appended)

#         # Process each new candle
#         for _, row in appended.iterrows():
#             self.on_new_candle(row.to_dict())



#     # -------------------------
#     # Called when a new candle row is appended (live)
#     # new_row_dict must contain same columns as self.df
#     # -------------------------
#     def on_new_candle(self, new_row_dict):
#         # append new candle
#         # self.df = pd.concat([self.df, pd.DataFrame([new_row_dict])], ignore_index=True)
#         self.df.loc[len(self.df)] = new_row_dict
#         print("self", self.df)
#         # recompute signals (cheap to recompute whole df for simplicity)
#         self.df = vectorised_signal_engine(self.df)

#         last_idx = self.df.index[-1]
#         last_row = self.df.loc[last_idx]
#         print("last", last_row)

#         # if trade signal shown, create pending tick entry tracker
#         if last_row['trade_signal'] in ('BUY','SELL'):
#             # do not overwrite an existing pending entry or active trade
#             if (self.pending_entry is None) and (self.active_exit_engine is None):
#                 controller = TickEntryController(direction=last_row['trade_signal'],
#                                                 planned_entry_price=last_row['entry_price'])
#                 self.pending_entry = {
#                     'row_index': int(last_idx),
#                     'direction': last_row['trade_signal'],
#                     'planned_entry_price': float(last_row['entry_price']),
#                     'controller': controller,
#                     'started_at': time.time()
#                 }
#                 print(f"[ORCH] Pending {last_row['trade_signal']} at row {last_idx}, planned entry price {last_row['entry_price']}")

#             # CREATE EXIT ENGINE IMMEDIATELY
#             # self.active_exit_engine = AdvancedExitEngine(
#             #     direction=direction,
#             #     entry_price=entry_price,
#             #     entry_time=now,
#             #     base_target_pts=self.EXTEND_TARGET)
            
#             # self.active_trade_meta = {
#             #     'row_index': int(last_idx),
#             #     'direction': direction,
#             #     'entry_price': entry_price,
#             #     'entry_time': now,
#             #     'status': 'OPEN'
#             # }

#             # WRITE ENTRY IMMEDIATELY
#             self._append_trade_to_excel({
#                 **self.active_trade_meta,
#                 'exit_price': None,
#                 'exit_time': None,
#                 'profit': None
#             })

#     # -------------------------
#     # Called for every incoming tick (price float). qty and ts optional.
#     # -------------------------
#     def on_tick(self, price, qty=None, ts=None):
#         now = ts if ts is not None else time.time()

#         # store tick for last candle (optional)
#         # self.row_ticks[self.df.index[-1]].append((now, price, qty))

#         # 1) if no pending entry and no active trade: nothing to do
#         if self.pending_entry is None and self.active_exit_engine is None:
#             return

#         # 2) If pending entry exists -> feed tick to its controller
#         if self.pending_entry is not None and self.active_exit_engine is None:
#             controller = self.pending_entry['controller']
#             res = controller.on_tick(price)
#             if res:
#                 # entry confirmed; create exit engine
#                 entry_price = res['entry_price']
#                 direction = res['direction']
#                 print(f"[ORCH] ENTRY CONFIRMED {direction} @ {entry_price} for row {self.pending_entry['row_index']}")

#                 # create AdvancedExitEngine with chosen parameters
#                 engine = AdvancedExitEngine(direction=direction,
#                                             entry_price=entry_price,
#                                             entry_time=now,
#                                             base_target_pts=self.EXTEND_TARGET,
#                                             # you can pass get_sl_points_fn=... if desired
#                                             )

#                 # store active trade meta
#                 self.active_exit_engine = engine
#                 self.active_trade_meta = {
#                     'row_index': self.pending_entry['row_index'],
#                     'direction': direction,
#                     'entry_price': entry_price,
#                     'entry_time': now
#                 }
#                 # clear pending
#                 self.pending_entry = None
#             return

#         # 3) If active exit engine exists -> feed tick to it
#         if self.active_exit_engine is not None:
#             exited, info = self.active_exit_engine.on_tick(price, qty=qty, ts=now)
#             if exited:
#                 # record trade
#                 meta = self.active_trade_meta or {}
#                 trade_record = {
#                     'row_index': meta.get('row_index'),
#                     'direction': meta.get('direction'),
#                     'entry_price': meta.get('entry_price'),
#                     'entry_time': meta.get('entry_time'),
#                     'exit_price': self.active_exit_engine.exit_price,
#                     'exit_reason': self.active_exit_engine.exit_reason,
#                     'exit_time': now,
#                     'profit': (self.active_exit_engine.exit_price - meta.get('entry_price')) if meta.get('direction')=='BUY' else (meta.get('entry_price') - self.active_exit_engine.exit_price)
#                 }
#                 print("[ORCH] TRADE CLOSED:", trade_record)
#                 self.trades_log.append(trade_record)
#                 # write to excel (append)
#                 self._append_trade_to_excel(trade_record)
#                 # clear active engine
#                 self.active_exit_engine = None
#                 self.active_trade_meta = None
#             return

#     # -------------------------
#     # Append trade record to output excel
#     # -------------------------
#     def _append_trade_to_excel(self, trade_record):
#         # try to load existing file
#         try:
#             out_df = pd.read_excel(self.output_file)
#         except FileNotFoundError:
#             out_df = pd.DataFrame()

#         # append_df = pd.DataFrame([trade_record])
#         # out_df = pd.concat([out_df, append_df], ignore_index=True)
#         out_df.loc[len(out_df)] = trade_record
#         print("out", out_df)
#         out_df.to_excel(self.output_file, index=False)
#         print(f"[ORCH] Trade appended to {self.output_file}")

#     # -------------------------
#     # Utility: show active state
#     # -------------------------
#     def status(self):
#         return {
#             'pending_entry': self.pending_entry,
#             'active_trade': None if self.active_exit_engine is None else self.active_trade_meta,
#             'trades_logged': len(self.trades_log)
#         }

  // zero.py
  from playwright.sync_api import Playwright, sync_playwright, expect
from urllib.parse import parse_qs,urlparse,quote
import pyotp
import requests
import json
import time
import re
# from datetime import datetime
import datetime as dt
# from upstox_api.api import Upstox
import pandas as pd
import numpy as np
import os
from urllib.parse import quote
import logging
from kiteconnect import KiteConnect, KiteTicker
# from getLtp import get_ltp
import uncompyle6
from uncompyle6.main import decompile_file
import decompyle3
import finnhub
import http.client
import struct
import websocket
import threading

# from openchart import NSEData
# nse = NSEData()
# from nsetools import Nse

filename = 'TradeToolOptChain_V1.cpython-311.pyc'
output_file = 'myfile_decompiled.py'

# python -m playwright install
# python -m  playwright codegen demo.playwright.dev/todomvc


API_KEY = 'ghe3ta94ngx0fjye'
SECRET_KEY = 's90mkn0609yr4u4s7fbj0xnvlqj62xdj'
RURL = 'https://127.0.0.1:5000/'

TOTP_KEY = 'CO5IOOJ6AM62SGBLMGHNJZVZLCH7TFSC'
MOBILE_NO = ''
PASS = ''
access_token="NUgFtcE5YybTQQP4KFxEpskQHDZvQUvu"
EXCHANGE      = "NSE"
TRADING_SYMBOL = "HDFCBANK" 
ACCESS_TOKEN = "cEakkarrZZsvjAcx8xXW69kin8nyGg0A"
request_token = "9URjlo3FbRjvQyqRNFFbhTg3VRbS4I3n"

file_path = "hdfcbank_intraday_signals.xlsx"
EXCEL_FILE = "zerodha_utbot_stc_intraday.xlsx"
rurlEncode = quote(RURL,safe="")


# symbol = 'NSE_EQ|HDFCBANK'
instrument_token = "341249" 

INTERVAL      = "15minute"                # 1minute, 5minute, 15minute, 30minute, 60minute, day
DAYS_BACK     = 5                         # rolling history window for indicators
SLEEP_SECONDS = 60 


AUTH_URL = f'https://kite.zerodha.com/connect/login?v=3&api_key={API_KEY}'

SYMBOL = "HDFCBANK"  # NSE symbol
FILE_PATH = r"C:\Users\Ramchandra\Desktop\up\hdfcbank_intraday_signals.xlsx"  # change path
FETCH_INTERVAL = 300  # seconds = 5 minutes

# ====== FUNCTION TO FETCH NSE DATA ======
session = requests.Session()

def parse_kite_packet(packet: bytes):
    """
    Parse Zerodha binary tick format.
    Extract:
        - total buy quantity (bytes 20-24)
        - total sell quantity (bytes 24-28)
        - bid entries (64-124) => 5 bids
        - offer entries (124-184) => 5 offers
    """
    ltp = struct.unpack(">I", packet[4:8])[0]
    # -----------------------------
    # 1) Total Buy Quantity (int32)
    # -----------------------------
    total_buy_qty = struct.unpack(">I", packet[20:24])[0]

    # -----------------------------
    # 2) Total Sell Quantity (int32)
    # -----------------------------
    total_sell_qty = struct.unpack(">I", packet[24:28])[0]

    # -----------------------------
    # 3) Bid Book (5 entries × 12 bytes each)
    # Structure per entry:
    #    price  -> int32 (4 bytes)
    #    qty    -> int32 (4 bytes)
    #    orders -> int32 (4 bytes)
    # -----------------------------
    bids = []
    offset = 64
    for i in range(5):
        entry = packet[offset:offset+12]
        price, qty, orders = struct.unpack(">III", entry)
        bids.append({"price": price / 100.0, "qty": qty, "orders": orders})
        offset += 12

    # -----------------------------
    # 4) Offer Book (5 entries × 12 bytes)
    # -----------------------------
    offers = []
    offset = 124
    for i in range(5):
        entry = packet[offset:offset+12]
        price, qty, orders = struct.unpack(">III", entry)
        offers.append({"price": price / 100.0, "qty": qty, "orders": orders})
        offset += 12

    return {
        "ltp": ltp,
        "total_buy_qty": total_buy_qty,
        "total_sell_qty": total_sell_qty,
        "bids": bids,
        "offers": offers
    }

# import onetimepass as otp
# code = otp.get_totp(TOTP_KEY)
# print(code)
# ------------------------------------------------------------
# WEBSOCKET WRAPPER FOR LIVE USAGE
# ------------------------------------------------------------

def on_message(ws, message):
    decoded = parse_kite_packet(message)
    print("ltp:", decoded["ltp"])
    print("Buy Qty:", decoded["total_buy_qty"])
    print("Sell Qty:", decoded["total_sell_qty"])
    print("Top Bid:", decoded["bids"][0])
    print("Top Offer:", decoded["offers"][0])


def on_open(ws):
    print("Connected!")
    # Subscribe example: instrument tokens [token, mode]
    # Mode 3 = full/binary
    payload = struct.pack(">I", 1) + struct.pack(">I", 123456)
    ws.send(payload, opcode=websocket.ABNF.OPCODE_BINARY)


def run_kite_ws(api_key, access_token):
    ws_url = f"wss://ws.kite.trade?api_key={api_key}&access_token={access_token}"
    ws = websocket.WebSocketApp(
        ws_url,
        on_message=on_message
    )
    ws.run_forever()


# Example:
# run_kite_ws("xxxx", "yyyy")


def getAccessToken(code):
    url = 'https://app.kite.trade/session/token'

    headers = {
        'accept': 'application/json',
        'Content-Type': 'application/x-www-form-urlencoded',
        'X-Kite-Version': '3'
    }

    data1 = {
        'api_key': API_KEY,
        'request_token': SECRET_KEY,
        'grant_type': 'authorization_code'
    }

    response = requests.post(url, headers=headers, data=data1)
    json_response = response.json()
    # print("jj",json_response)
    value = json_response['access_token']

    # print(f"access_token:  {json_response['access_token']}")
    return json_response, value

def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://kite.zerodha.com/")
    page.get_by_role("textbox", name="Phone number or User ID").click()
    page.get_by_role("textbox", name="Phone number or User ID").fill(MOBILE_NO)
    page.get_by_role("textbox", name="Password").click()
    page.get_by_role("textbox", name="Password").fill(PASS)
    page.get_by_role("button", name="Login").click()
    page.get_by_role("spinbutton", name="External TOTP").click()
    otp = pyotp.TOTP(TOTP_KEY).now()
    print("otp", otp)
    page.get_by_role("spinbutton", name="External TOTP").fill(otp)
    context.close()
    browser.close()

# with sync_playwright() as playwright:
#     code = run(playwright)
#     print("code", code)

logging.basicConfig(level=logging.DEBUG)

def run2(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    # page.goto(url)
    page.goto("https://kite.zerodha.com/connect/login?api_key=ghe3ta94ngx0fjye")
    page.get_by_role("textbox", name="Phone number or User ID").click()
    page.get_by_role("textbox", name="Phone number or User ID").fill(MOBILE_NO)
    page.get_by_role("textbox", name="Password").click()
    page.get_by_role("textbox", name="Password").fill(PASS)
    page.get_by_role("button", name="Login").click()
    page.get_by_role("spinbutton", name="External TOTP").click()
    otp = pyotp.TOTP(TOTP_KEY).now()
    print("otp", otp)
    page.get_by_role("spinbutton", name="External TOTP").fill(otp)
    # context.close()
    # browser.close()

# with sync_playwright() as playwright:
#     code = run2(playwright)
#     print("code", code)

def fetch_option_chain(symbol="NIFTY"):
    url = f"https://www.nseindia.com/api/option-chain-indices?symbol={symbol}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    session = requests.Session()
    session.get("https://www.nseindia.com", headers=headers)  # Initialize session
    response = session.get(url, headers=headers)
    print(response)
    data = response.json()
    return data


# def history_data():
#     url = f"https://api.kite.trade/instruments/historical/341249/15minute"

#     params = {
#         "from": "2025-08-28 09:15:00",
#         "to": "2025-08-28 15:30:00"
#     }

#     headers = {
#         "X-Kite-Version": "3",
#         "Authorization": "token ghe3ta94ngx0fjye:s90mkn0609yr4u4s7fbj0xnvlqj62xdj"  
#     }
#     print("ddjshdsds")
#     response = requests.get(url, params=params, headers=headers)
#     print(response)

# def fetch_data():
#     today = dt.date.today()
#     five_days_ago = today - dt.timedelta(days=5)
#     instrument_key1 = "NSE_EQ%7CINE347A01017"
#     url = (
#     f"https://api.upstox.com/v3/historical-candle?instrument_token={instrument_key1}&unit=minutes&interval=15&to_date={today}&from_date={five_days_ago}")   
#     header = {
#     'Accept': 'application/json',
#     }
#     response = requests.request("GET", url, headers=header)
#     print("response1", response.text)
#     data2 = response.json()
#     print("data2", data2)
#     candles = data2['data']['candles']
#     df = pd.DataFrame(candles, columns=['timestamp', 'open', 'high', 'low', 'close', 'volume'])
#     df['timestamp'] = pd.to_datetime(df['timestamp'])
#     df.set_index('timestamp', inplace=True)
#     return df
    


# https://www.youtube.com/watch?v=e91yuEE-LDo
def get_data():
    url = "https://www.nseindia.com/api/historicalOR/foCPV?from=25-08-2025&to=28-08-2025&instrumentType=OPTSTK&symbol=GODREJPROP&year=2025&expiryDate=28-AUG-2025&optionType=CE&strikePrice=2000"
    response = requests.get(url)

def fetch_option_chain(symbol="NIFTY"):
    base_url = "https://www.nseindia.com"
    option_chain_url = f"https://www.nseindia.com/api/historicalOR/foCPV?from=01-09-2025&to=09-09-2025&instrumentType=OPTIDX&symbol={symbol}&year=2025&expiryDate=30-SEP-2025&optionType=CE&strikePrice=24800"

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0 Safari/537.36",
        "Accept-Language": "en-IN,en;q=0.9",
        "Accept": "application/json, text/plain, */*",
        "Referer": "https://www.nseindia.com/option-chain",
    }

    session = requests.Session()
    
        # Visit homepage to get cookies
    session.get(base_url, headers=headers, timeout=5)
    # time.sleep(30)  # random delay
        
    response = session.get(option_chain_url, headers=headers, timeout=5)
    df = pd.DataFrame(response)
    print(response.json())

# data = fetch_option_chain("NIFTY")
# def on_ticks(ws, ticks):
#     global current_candle, candle_data
#     ts = int(time.time())
#     bucket = ts - (ts % interval)

#     for tick in ticks:
#         lp = tick["last_price"]
#         if current_candle is None or current_candle["ts"] != bucket:
#             if current_candle:
#                 candle_data.append(current_candle)
#             current_candle = {"open": lp, "high": lp, "low": lp, "close": lp, "ts": bucket}
#         else:
#             current_candle["high"] = max(current_candle["high"], lp)
#             current_candle["low"] = min(current_candle["low"], lp)
#             current_candle["close"] = lp

# def on_connect(ws, response):
#     ws.subscribe([instrument_token])
#     ws.set_mode(ws.MODE_FULL, [instrument_token])


def save_access_token(token):
    with open("access_token.txt", "w") as f:
        f.write(token)

def save_access_token_kite(token):
    with open("access_token_kite.txt", "w") as f:
        f.write(token)

def get_holdings(api_key, access_token):
    url = "https://api.kite.trade/portfolio/holdings"

    headers = {
        "X-Kite-Version": "3",
        "Authorization": f"token {api_key}:{access_token}",
    }

    response = requests.get(url, headers=headers)
    response.raise_for_status()     # raise error if failed

    return response.json()          # return holdings JSON

def start_kite_ws(api_key, access_token, instrument_tokens):
    """
    Connects to Kite WebSocket and subscribes in FULL mode

    instrument_tokens: list of int (e.g. [408065])
    """

    ws_url = f"wss://ws.kite.trade?api_key={api_key}&access_token={access_token}"
    print("hiiii8")
    def on_open(ws):
        print("[KITE WS] Connected")
        print("hiiii")

        # Equivalent of:
        # { a: "mode", v: ["full", [408065]] }
        msg = {
            "a": "mode",
            "v": ["full", instrument_tokens]
        }
        ws.send(json.dumps(msg))
        print("[KITE WS] Subscribed FULL mode:", instrument_tokens)

    def on_message(ws, message):
        # Kite sends binary frames also; leave decoding to you
        print("[KITE WS] Raw message:", message)
        # pass

    def on_error(ws, error):
        print("[KITE WS] Error:", error)

    def on_close(ws, close_status_code, close_msg):
        print("[KITE WS] Closed", close_status_code, close_msg)

    ws = websocket.WebSocketApp(
        ws_url,
        on_open=on_open,
        on_message=on_message,
        on_error=on_error,
        on_close=on_close
    )

    # run websocket in background thread
    # threading.Thread(target=ws.run_forever, daemon=True).start()
    ws.run_forever()

def fetch_ltp(api_key, access_token, instruments):
    """
    instruments: list of strings
        e.g. ["NSE:INFY", "BSE:SENSEX", "NSE:NIFTY 50"]
    """
    url = "https://api.kite.trade/quote/ltp"

    params = {
        "i": instruments
    }

    headers = {
        "X-Kite-Version": "3",
        "Authorization": f"token {api_key}:{access_token}"
    }

    response = requests.get(url, headers=headers, params=params, timeout=5)
    response.raise_for_status()

    return response.json()

def place_kite_order(
    api_key: str,
    access_token: str,
    tradingsymbol: str,
    exchange: str,
    transaction_type: str,
    quantity: int,
    product: str = "MIS",
    order_type: str = "MARKET",
    validity: str = "DAY"
):
    url = "https://api.kite.trade/orders/regular"

    headers = {
        "X-Kite-Version": "3",
        "Authorization": f"token {api_key}:{access_token}"
    }

    payload = {
        "tradingsymbol": tradingsymbol,
        "exchange": exchange,
        "transaction_type": transaction_type,  # BUY / SELL
        "order_type": order_type,
        "quantity": quantity,
        "product": product,
        "validity": validity
    }

    response = requests.post(url, headers=headers, data=payload)
    data = response.json()
    print("Kite response:", data)

    if data.get("status") == "success":
        return data["data"]["order_id"]
    else:
        raise Exception(data)

def login():

    kite = KiteConnect(api_key=API_KEY)
    # url = kite.login_url()
    # print(url)

    # https://www.youtube.com/watch?v=mUzNMbdr3QE&list=PL8s2frOzp0990WjdNHmbEZIeao19FvySQ&index=7

    # data = kite.generate_session("bu3yCpLs6vpW5A96La9hKboHkdNovGrF", SECRET_KEY)
    # kite.set_access_token(data["access_token"])
    # token = data["access_token"]
    # save_access_token_kite(data["access_token"])
    # sdssd = get_holdings(api_key=API_KEY, access_token=ACCESS_TOKEN)
    # print(sdssd)

    # headers = {
    #     "X-Kite-Version": "3",
    #     "Authorization": f"token {API_KEY}:ACCESS_TOKEN"
    # }

    # payload = {
    # "tradingsymbol": "ACC",
    # "exchange": "NSE",
    # "transaction_type": "BUY",
    # "order_type": "MARKET",
    # "quantity": 1,
    # "product": "MIS",
    # "validity": "DAY"
    # }
    # url1 = "https://api.kite.trade/orders/regular"
    # response = requests.post(url1, headers=headers, data=payload)
    # print(response.json())
    # return data["access_token"]
    # order_id = place_kite_order(
    # api_key=API_KEY,
    # access_token=ACCESS_TOKEN,
    # tradingsymbol="ACC",
    # exchange="NSE",
    # transaction_type="BUY",
    # quantity=1
    # )

    # print("Order placed:", order_id)

login()


import hashlib
import requests


def generate_kite_session(api_key, api_secret, request_token):
    """
    Exchange request_token for access_token.
    """

    # checksum = sha256(api_key + request_token + api_secret)
    checksum_str = api_key + request_token + api_secret
    checksum = hashlib.sha256(checksum_str.encode()).hexdigest()

    url = "https://api.kite.trade/session/token"

    headers = {
        "X-Kite-Version": "3"
    }

    data = {
        "api_key": api_key,
        "request_token": request_token,
        "checksum": checksum
    }

    r = requests.post(url, headers=headers, data=data)
    r.raise_for_status()

    return r.json()

# session = generate_kite_session(
#     api_key=API_KEY,
#     api_secret=SECRET_KEY,
#     request_token=request_token
# )

# print(session)
# print("ACCESS TOKEN =", session["data"]["access_token"])
