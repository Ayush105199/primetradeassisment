import xlwings as xw
import pandas as pd
import time
import requests

API_KEY = "2f488ea6-3e97-47f3-84f3-e41617fffcdf" 
headers = {"X-CMC_PRO_API_KEY": API_KEY}


wb = xw.Book('cryptocurrency.xlsx')
sheet = wb.sheets['Sheet1']


def fetch_crypto_data():
    url = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest"
    params = {
        "start": "1",
        "limit": "50",
        "convert": "USD"
    }
    response = requests.get(url, headers=headers, params=params)
    
    if response.status_code == 200:
        data = response.json()["data"]
      
        df = pd.DataFrame(data)
  
        df["current_price"] = df["quote"].apply(lambda x: x["USD"]["price"] if isinstance(x, dict) and "USD" in x else None)
        df["market_cap"] = df["quote"].apply(lambda x: x["USD"]["market_cap"] if isinstance(x, dict) and "USD" in x else None)
        df["total_volume"] = df["quote"].apply(lambda x: x["USD"]["volume_24h"] if isinstance(x, dict) and "USD" in x else None)
        df["price_change_percentage_24h"] = df["quote"].apply(lambda x: x["USD"]["percent_change_24h"] if isinstance(x, dict) and "USD" in x else None)
        
       
        df = df[["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"]]
        
        return df
    else:
        print("Error fetching data:", response.status_code, response.text)
        return None


while True:
    data = fetch_crypto_data()
    if data is not None:
        sheet.range('A1').options(pd.DataFrame, index=False).value = data
        print("Updated Excel sheet with live data.")
    time.sleep(20)  
