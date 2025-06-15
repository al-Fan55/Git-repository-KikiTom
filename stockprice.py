import os
import json
import time
import re
import glob
from datetime import datetime
import yfinance as yf
import xlwings as xw

CACHE_DIR = r"C:\yfJSON"

def ensure_cache_dir():
    if not os.path.exists(CACHE_DIR):
        os.makedirs(CACHE_DIR)

def get_today_str():
    return datetime.now().strftime("%Y-%m-%d")

def get_cache_file_path(date_str=None):
    if date_str is None:
        date_str = get_today_str()
    return os.path.join(CACHE_DIR, f"yf_cache_{date_str}.json")

def find_latest_cache_file():
    pattern = os.path.join(CACHE_DIR, "yf_cache_*.json")
    files = glob.glob(pattern)
    if not files:
        return None
    files_with_dates = []
    for file in files:
        basename = os.path.basename(file)
        match = re.search(r"yf_cache_(\d{4}-\d{2}-\d{2})\.json", basename)
        if match:
            date_str = match.group(1)
            try:
                date_obj = datetime.strptime(date_str, "%Y-%m-%d")
                files_with_dates.append((date_obj, file))
            except:
                continue
    if not files_with_dates:
        return None
    latest_file = max(files_with_dates, key=lambda x: x[0])[1]
    return latest_file

def load_latest_cache_if_today_is_empty(today_str):
    path_today = get_cache_file_path(today_str)
    if os.path.exists(path_today):
        with open(path_today, "r", encoding="utf-8") as f:
            data = json.load(f)
            if data:
                return data
    latest_path = get_latest_available_cache()
    if latest_path:
        with open(latest_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def get_latest_available_cache():
    files = [f for f in os.listdir(CACHE_DIR) if f.startswith("yf_cache_") and f.endswith(".json")]
    if not files:
        return None
    files.sort(reverse=True)
    return os.path.join(CACHE_DIR, files[0])

def load_cache(date_str=None):
    if date_str is not None:
        path = get_cache_file_path(date_str)
    else:
        path = find_latest_cache_file()
        if not path:
            return {}
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_cache(cache_data, date_str=None):
    path = get_cache_file_path(date_str)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cache_data, f, ensure_ascii=False, indent=2)

def get_price_and_data(ticker, cache_data, today_str, use_online):
    if ticker in cache_data and today_str in cache_data[ticker]:
        return cache_data[ticker][today_str]

    if use_online:
        try:
            time.sleep(1.5)
            stock = yf.Ticker(ticker)
            hist = stock.history(period='1d')
            if not hist.empty:
                price = hist['Close'].iloc[-1]
                full_data = stock.info

                if ticker not in cache_data:
                    cache_data[ticker] = {}
                cache_data[ticker][today_str] = {
                    "price": price,
                    "info": full_data
                }
                return cache_data[ticker][today_str]
        except Exception as e:
            print(f"{ticker} の取得エラー: {e}")

    return None

def get_tickers_from_excel(wb):
    sheet = wb.sheets.active    # ← アクティブなシートを取得
    tickers = []
    row = 5
    while True:
        code = sheet.range(f"B{row}").value
        if not code:
            break
        try:
            code_int = int(float(code))
            ticker = f"{code_int}.T"
        except:
            ticker = str(code).strip()
        tickers.append(ticker)
        row += 1
    return tickers

def write_prices_to_excel(wb, result_dict):
    # sheet = wb.sheets[0]
    sheet = wb.sheets.active    # ← アクティブなシートを取得
    row = 5
    for ticker, data in result_dict.items():
        price = data.get("price", "N/A")
        sheet.range(f"I{row}").value = price
        print(f"{ticker}: {price} 円")
        row += 1

def main(use_online=True, wb=None):
    wb = xw.Book.caller()  # Excelから呼ばれたときのみ取得

    ensure_cache_dir()
    today_str = get_today_str()
    cache_data = load_cache(today_str)

    tickers = get_tickers_from_excel(wb)
    result_dict = {}

    for ticker in tickers:
        data = get_price_and_data(ticker, cache_data, today_str, use_online)
        if data:
            result_dict[ticker] = data
        else:
            result_dict[ticker] = {"price": "取得失敗", "info": {}}

    write_prices_to_excel(wb, result_dict)

    if use_online:
        save_cache(cache_data, today_str)

if __name__ == "__main__":
    book_path = r"E:\ExcelDocs\API-ポートフォリオ.xlsm"
    app = xw.App(visible=False)
    wb = app.books.open(book_path)

    try:
        while True:
            answer = input("yFinanceから株価データを取得しますか？ (y/n): ").strip().lower()
            if answer in ("y", ""):
                use_online = True
                break
            elif answer == "n":
                print("（注意）古いキャッシュデータを使います。")
                use_online = False
                break
            else:
                print("y または n を入力してください。")

        main(use_online, wb)

    finally:
        wb.save()
        wb.close()
        app.quit()
