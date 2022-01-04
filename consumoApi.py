import xlwings,pandas

wb = xlwings.Book('dadosApi.xlsx').sheets('Sheet1')
while True:
  data = pandas.read_json('https://api.binance.com/api/v1/ticker/allPrices')
  wb.range('a1').options(pandas.DataFrame, index=False).value = pandas.DataFrame(data)


