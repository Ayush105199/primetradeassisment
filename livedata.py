import xlwings,pandas
wb=xlwings.Book('cryptocurrency.xlsx').sheets('Sheet1')
while True:
    data= pandas.read_json('https://api.binance.com/api/v1/ticker/24hr')
    wb.range('a1').options(pandas.DataFrame).value=pandas.DataFrame(data)

