import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math
import yfinance as yf

# Call Yahoo Finance API to obtain data
#symbol = "AAPL"
#stock = yf.Ticker(symbol)
#print(stock.info)

symbols = pd.read_csv(r'C:\Users\Kevin Tang\Desktop\SoftwareProjects\algorithmic-trading-python\Data\sp_500_stocks.csv')

# Parse API Call


#Adding Our Stocks Data to a Pandas DataFrame
columns_name = ['Ticker', 'Stock Price', 'Market Capitalization', 'Number of Shares to Buy']
data_frame = pd.DataFrame(columns = columns_name)

# Iterative API Calls (Very Slow)
for symbol in symbols['Ticker'][:5]:
    stock = yf.Ticker(symbol) # Very slow HTTP Request
    
    current_price = stock.info['currentPrice']
    market_cap = stock.info['marketCap']
    
    data_frame = data_frame.append(
        pd.Series(
        [
            symbol,
            current_price,
            market_cap,
            'N/A'
        ],
        index = columns_name
        ),
        ignore_index=True
    )
    
#Calculate Number of Shares to Buy
portfolio_size = input('Enter the cash value of your portfolio:')
try:
    val = float(portfolio_size)
    print(val)
except ValueError:
    print('Please enter a numerical value \n')
    portfolio_size = input('Enter the value of your portfolio:')
    val = float(portfolio_size)
    
position_size = val/len(data_frame.index)
for i in range(0,len(data_frame.index)):
    data_frame.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/data_frame['Stock Price'][i])
print(data_frame)

#Formatting Excel Output and Export
writer = pd.ExcelWriter('recommended_trades.xlsx', engine = 'xlsxwriter')
data_frame.to_excel(writer, 'Recommneded Trades', index = False)
background_color = '#0a0a23'
font_color = '#ffff'

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

column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format],
    'C': ['Market Capitilization', dollar_format],
    'D': ['Number of Shares to Buy', integer_format]
}

writer.sheets['Recommneded Trades'].write('A1', 'Ticker', string_format)
writer.sheets['Recommneded Trades'].write('B1', 'Stock Price', dollar_format)
writer.sheets['Recommneded Trades'].write('C1', 'Market Captilization', dollar_format)
writer.sheets['Recommneded Trades'].write('D1', 'Number of Shares to Buy', integer_format)

for column in column_formats.keys():
    writer.sheets['Recommneded Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
writer.save()