import numpy as np
import pandas as pd
import requests 
import xlsxwriter
import math
import finnhub
from secrets import FINNHUB_API_TOKEN
import websocket
import time

# There are 2 tables on the Wikipedia page
# we want the first table

payload=pd.read_html('https://en.wikipedia.org/wiki/List_of_S%26P_500_companies')

stocks = payload[0]

finnhub_client = finnhub.Client(api_key=FINNHUB_API_TOKEN)

# def on_message(ws, message):
#     print(message)

# def on_error(ws, error):
#     print(error)

# def on_close(ws):
#     print("### closed ###")

columns = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
final_dataframe = pd.DataFrame(columns = columns)

# getting all S&P 500 stocks in a pandas DataFrame with starting day price and market cap
i = 0
for stock in stocks['Symbol']:
    time.sleep(1.95)
    print(f"{i}     Retrieving data on {stock}")
    price = finnhub_client.quote(symbol=stock)['c']
    market_cap = finnhub_client.company_profile2(symbol=stock)['marketCapitalization']
    final_dataframe = final_dataframe.append(
        pd.Series(
            [
                stock,
                price,
                market_cap,
                'N/A'
            ],
            index = columns
        ),
        ignore_index = True
        )
    i += 1

time_format = time.strftime("%H:%M:%S", time.gmtime(1.95*len(stocks['Symbol'])))
print(f"Elapsed time: {time_format}")
print(final_dataframe)

# def on_open(ws):
#     for stock in stocks['Symbol']:
#         ws.send('{"type":"subscribe","symbol":%s}' % (stock))


# if __name__ == "__main__":
#     websocket.enableTrace(True)
#     ws = websocket.WebSocketApp("wss://ws.finnhub.io?token=c6qlpiqad3i891nj5hcg",
#                               on_message = on_message,
#                               on_error = on_error,
#                               on_close = on_close)
#     ws.on_open = on_open
#     ws.run_forever()

portfolio_size = input('Enter the value of your portfolio:  ')

try:
    val = float(portfolio_size)
    print(val)
except ValueError:
    print("You didn't enter a number \nPlease try again:")
    portfolio_size = input('Enter the value of your portfolio:  ')
    val = float(portfolio_size)

position_size = val/len(final_dataframe.index)
for i in range(0, len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe.loc[i, 'Stock Price'])

writer = pd.ExcelWriter('recomended trades.xlsx', engine = 'xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index = False)

background_color = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

integer_format = writer.book.add_format(
    {
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

writer.sheets['Recommended Trades'].write('A1', 'Ticker', string_format)
writer.sheets['Recommended Trades'].write('B1', 'Stock Price', dollar_format)
writer.sheets['Recommended Trades'].write('C1', 'Market Capitalization', dollar_format)
writer.sheets['Recommended Trades'].write('D1', 'Number of Shares to Buy', integer_format)

column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format]
    'C': ['Market Capitalization', dollar_format]
    'D': ['Number of Shares to Buy', integer_format]
}

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats)
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], column_formats[column][1])

writer.save()