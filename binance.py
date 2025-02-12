import xlwings as xw
import pandas as pd


wb = xw.Book('cryptocurrency.xlsx').sheets('Sheet1')

while True:
 
    ticker_data = pd.read_json('https://api.binance.com/api/v1/ticker/24hr')
    
    exchange_info = pd.read_json('https://api.binance.com/api/v3/exchangeInfo')
    
  
    symbols_info = pd.DataFrame(exchange_info['symbols'].tolist())
    symbols_info = symbols_info[['symbol', 'baseAsset', 'quoteAsset']]
    
    
    filtered_data = ticker_data[['symbol', 'lastPrice', 'quoteVolume', 'priceChangePercent']]
    filtered_data.columns = ['Symbol', 'Current Price', '24h Trading Volume', '24h Price Change (%)']
    
    
    merged_data = pd.merge(filtered_data, symbols_info, on='symbol', how='left')
    

    merged_data['Name'] = merged_data['baseAsset']
   
    final_data = merged_data[['Name', 'Symbol', 'Current Price', '24h Trading Volume', '24h Price Change (%)']]
    
   
    wb.range('A1').options(pd.DataFrame).value = final_data