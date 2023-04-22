import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math
from secretss import IEX_CLOUD_API_TOKEN

#Read the tickers and ave them in stocks
stocks=pd.read_csv('sp_500_stocks.csv')
my_columns=['Ticker', 'Stock Price', 'Market Capitalization','P/E ratio', 'Number of Shares to Buy']

#Function to divide the ticker list to multiple lists of size n
def chunks(lst,n):
    for i in range(0,len(lst),n):
        yield lst[i:i+n]

symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []

for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))
final_dataframe = pd.DataFrame(columns=my_columns)

for symbol_string in symbol_strings:
     batch_api_call_url = f'https://cloud.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote&token={IEX_CLOUD_API_TOKEN}'
     data = requests.get(batch_api_call_url).json()
     for symbol in symbol_string.split(','):
        final_dataframe = final_dataframe._append(
        pd.Series(
                [
                    symbol,
                    data[symbol]['quote']['latestPrice'],
                    data[symbol]['quote']['marketCap'],
                    data[symbol]['quote']['peRatio'],
                    'N/A'

                ],
                index=my_columns),
            ignore_index=True)
portfolio_size= float(1000000)
position_size=portfolio_size/len(final_dataframe.index)
for i in range(0,len(final_dataframe['Ticker'])-1):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / final_dataframe['Stock Price'][i])
print(final_dataframe)

writer = pd.ExcelWriter('recommended_trades.xlsx', engine='xlsxwriter')
final_dataframe.to_excel(writer, sheet_name='Recommended Trades', index = False)
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
            'num_format':'$0.00',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )

integer_format = writer.book.add_format(
        {
            'num_format':'0',
            'font_color': font_color,
            'bg_color': background_color,
            'border': 1
        }
    )
writer.sheets['Recommended Trades'].write('A1', 'Ticker', string_format)
writer.sheets['Recommended Trades'].write('B1', 'Price', string_format)
writer.sheets['Recommended Trades'].write('C1', 'Market Capitalization', string_format)
writer.sheets['Recommended Trades'].write('D1', 'P/E Ratio', string_format)
writer.sheets['Recommended Trades'].write('E1', 'Number Of Shares to Buy', string_format)

writer.sheets['Recommended Trades'].set_column('A:A', 20, string_format)
writer.sheets['Recommended Trades'].set_column('B:B', 20, dollar_format)
writer.sheets['Recommended Trades'].set_column('C:C', 20, dollar_format)
writer.sheets['Recommended Trades'].set_column('D:D', 20, string_format)
writer.sheets['Recommended Trades'].set_column('E:E', 20, integer_format)

column_formats = {
                    'A': ['Ticker', string_format],
                    'B': ['Price', dollar_format],
                    'C': ['Market Capitalization', dollar_format],
                    'D': ['P/E ratio', string_format],
                    'E': ['Number of Shares to Buy', integer_format]

}

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)
writer.close()