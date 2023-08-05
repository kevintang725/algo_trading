import math
import pandas as pd
import yfinance as yf
import numpy as np
from scipy import stats #The SciPy stats module
import statistics 

# Functions

def import_symbols_table(path_file):
    symbols = pd.read_csv(path_file)
    return symbols

def parse_api_data(symbols):
    #Adding Our Stocks Data to a Pandas DataFrame
    column_names = [
        'Ticker',
        'Current Price',
        'Day High',
        'Day Low',
        'Previous Close',
        'Trailing PE',
        'Forward PE',
        'PE Percentile',
        'Price-to-Book Ratio',
        'PB Percentile',
        'Price-to-Sales Ratio',
        'PS Percentile',
        'EV/EBITDA',
        'EV/EBITDA Percentile',
        'EV/GP',
        'EV/GP Percentile',
        'RV Score'
    ]
    data_frame = pd.DataFrame(columns = column_names)

    # Iterative API Calls (Very Slow)
    for symbol in symbols['Ticker'][:10]:
        stock = yf.Ticker(symbol) # Very slow HTTP Request
        
        #print(stock.info)
        
        try:
            # Market Cap
            market_cap = stock.info['marketCap']
            
            # Current Day Price Highs and Lows
            current_price = stock.info['currentPrice']
            previous_close = stock.info['previousClose']
            market_open = stock.info['open']
            day_low = stock.info['dayLow']
            day_high = stock.info['dayHigh']
            
            # Trading Volume
            current_volume = stock.info['volume']
            avg_volume = stock.info['averageVolume']
            avg_volume_10days = stock.info['averageVolume10days']
            
            # Price History
            fiftyTwoWeekLow = stock.info['fiftyTwoWeekLow']
            fiftyTwoWeekHigh = stock.info['fiftyTwoWeekHigh']
            fiftyDayAverage = stock.info['fiftyDayAverage']
            twohundredDayAverage = stock.info['twoHundredDayAverage']
            
            # Outstanding Shares, Shorts, and Floats
            sharesOutstanding = stock.info['sharesOutstanding']
            sharesShort = stock.info['sharesShort']
            sharesPercentSharesOut = stock.info['sharesPercentSharesOut']
            sharesShortPriorMonth = stock.info['sharesShortPriorMonth']
            shortRatio = stock.info['shortRatio']
            sharesPercentOfFloat = stock.info['shortPercentOfFloat']
            
            # Earnings Per Share and Profit Margins
            trailing_EPS = stock.info['trailingEps']
            forward_EPS = stock.info['forwardEps']
        
            # Dividend
            dividend_yield = stock.info['dividendYield']
            dividend_rate = stock.info['dividendRate']
            
            # Price-to-Earning Ratios
            trailing_PE = stock.info['trailingPE']
            forward_PE = stock.info['forwardPE']
            
            # Price-to-Book Ratios
            priceToBook = stock.info['priceToBook']
            bookValue = stock.info['bookValue']
            
            # Price-to-Sales Ratios
            priceToSales12Months = stock.info['priceToSalesTrailing12Months']
            
            # Profit and Revnue Margins
            grossProfits = stock.info['grossProfits']
            grossMargins = stock.info['grossMargins']
            profitMargins = stock.info['profitMargins']
            enterpriseValue = stock.info['enterpriseValue']
            ebidtdaMargins = stock.info['ebitdaMargins']
        except:
            pass
        
        # Try to calculate Enterprise Value Ratios
        try:
            ev_to_ebitda = enterpriseValue/ebidtdaMargins
        except TypeError:
            ev_to_ebitda = np.NaN
        
        try:
            ev_to_gross_profit = enterpriseValue/grossProfits
        except TypeError:
            ev_to_gross_profit = np.NaN
        
        # Append Data to Dataframe
        data_frame = data_frame.append(
            pd.Series(
            [
                symbol,
                current_price,
                day_high,
                day_low,
                previous_close,
                trailing_PE,
                forward_PE,
                'N/A',
                priceToBook,
                'N/A',
                priceToSales12Months,
                'N/A',
                ev_to_ebitda,
                'N/A',
                ev_to_gross_profit,
                'N/A',
                'N/A',
            ],
            index = column_names
            ),
            ignore_index=True
        )
        
    # Identify N/A values and fill with average column values
    data_frame = fill_nan_with_mean(data_frame)

    return data_frame

def metrics_parameters():
    metrics = {
            'Forward PE': 'PE Percentile',
            'Price-to-Book Ratio':'PB Percentile',
            'Price-to-Sales Ratio': 'PS Percentile',
            'EV/EBITDA':'EV/EBITDA Percentile',
            'EV/GP':'EV/GP Percentile'
    }
    return metrics

def fill_nan_with_mean(data_frame):
    for column in ['Forward PE', 'Price-to-Book Ratio','Price-to-Sales Ratio',  'EV/EBITDA','EV/GP']:
        data_frame[column].fillna(data_frame[column].mean(), inplace = True)
    return data_frame

def calculate_value_percentiles(data_frame):
    metrics = metrics_parameters()
    for row in data_frame.index:
        for metric in metrics.keys():
            data_frame.loc[row, metrics[metric]] = stats.percentileofscore(data_frame[metric], data_frame.loc[row, metric])/100

    # Print each percentile score to make sure it was calculated properly
    for metric in metrics.values():
        print(data_frame[metric])

    #Print the entire DataFrame    
    data_frame
    return data_frame

def calculate_rv_score(data_frame):
    metrics = metrics_parameters()
    for row in data_frame.index:
        value_percentiles = []
        for metric in metrics.keys():
            value_percentiles.append(data_frame.loc[row, metrics[metric]])
        data_frame.loc[row, 'RV Score'] = statistics.mean(value_percentiles)
    
    return data_frame

def calculate_number_of_shares_to_buy(data_frame):

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

    return data_frame

def export_to_excel(data_frame):
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
    #writer.sheets['Recommneded Trades'].write('B1', 'Stock Price', dollar_format)
    #writer.sheets['Recommneded Trades'].write('C1', 'Market Captilization', dollar_format)
    #writer.sheets['Recommneded Trades'].write('D1', 'Number of Shares to Buy', integer_format)

    for column in column_formats.keys():
        writer.sheets['Recommneded Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
    writer.save()
    
    