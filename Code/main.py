import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math
import yfinance as yf
import functions as my_finance_func

def main():
    # Call Yahoo Finance API to obtain data
    path_file = r'C:\Users\Kevin Tang\Desktop\SoftwareProjects\algo_trading\Data\sp_500_stocks.csv'
    symbols = my_finance_func.import_symbols_table(path_file)

    # Parse API Call
    df = my_finance_func.parse_api_data(symbols)
    
    # Calculate percentiles and fill in to metrics table
    df = my_finance_func.calculate_value_percentiles(df)
    
    # Calculate RV score and fill
    df = my_finance_func.calculate_rv_score(df)

    # Calculate number of shares to buy
    #df = my_finance_func.calculate_number_of_shares_to_buy(df)

    # Export to excel file
    my_finance_func.export_to_excel(df)

if __name__ == "__main__":
    main()
    