import xlwings as xw
import pandas as pd
import time
import requests


wb = xw.Book('cryptocurrency.xlsx')
sheet = wb.sheets['Sheet1']


def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        "vs_currency": "usd",
        "order": "market_cap_desc",
        "per_page": 50,
        "page": 1,
        "sparkline": False
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return pd.DataFrame(response.json())[
            ["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"]
        ]
    else:
        print("Error fetching data:", response.status_code)
        return None


while True:
    data = fetch_crypto_data()
    if data is not None:
        sheet.range('A1').options(pd.DataFrame, index=False).value = data
        print("Updated Excel sheet with live data.")
    time.sleep(20) 
