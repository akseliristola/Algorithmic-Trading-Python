import numpy as np
import pandas as pd
import requests
import math
from scipy import stats
from secretss import IEX_CLOUD_API_TOKEN
from constants import chunks,stocks,portfolio_size

def valueStratExcel():

    symbol_groups = list(chunks(stocks['Ticker'], 100))
    symbol_strings = []
    for i in range(0, len(symbol_groups)):
        symbol_strings.append(','.join(symbol_groups[i]))
    rv_columns = [
        'Ticker',
        'Price',
        'Number of Shares to Buy',
        'Price-to-Earnings Ratio',
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

    rv_dataframe = pd.DataFrame(columns=rv_columns)

    for symbol_string in symbol_strings:
        batch_api_call_url = f'https://cloud.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote,advanced-stats&token={IEX_CLOUD_API_TOKEN}'
        data = requests.get(batch_api_call_url).json()
        for symbol in symbol_string.split(','):
            enterprise_value = data[symbol]['advanced-stats']['enterpriseValue']
            ebitda = data[symbol]['advanced-stats']['EBITDA']
            gross_profit = data[symbol]['advanced-stats']['grossProfit']

            try:
                ev_to_ebitda = enterprise_value / ebitda
            except TypeError:
                ev_to_ebitda = np.NaN

            try:
                ev_to_gross_profit = enterprise_value / gross_profit
            except TypeError:
                ev_to_gross_profit = np.NaN

            rv_dataframe = rv_dataframe._append(
                pd.Series([
                    symbol,
                    data[symbol]['quote']['latestPrice'],
                    'N/A',
                    data[symbol]['quote']['peRatio'],
                    'N/A',
                    data[symbol]['advanced-stats']['priceToBook'],
                    'N/A',
                    data[symbol]['advanced-stats']['priceToSales'],
                    'N/A',
                    ev_to_ebitda,
                    'N/A',
                    ev_to_gross_profit,
                    'N/A',
                    'N/A'
                ],
                    index=rv_columns),
                ignore_index=True
            )
    for column in ['Price-to-Earnings Ratio', 'Price-to-Book Ratio','Price-to-Sales Ratio',  'EV/EBITDA','EV/GP']:
        rv_dataframe[column].fillna(rv_dataframe[column].mean(), inplace = True)
    metrics = {
                'Price-to-Earnings Ratio': 'PE Percentile',
                'Price-to-Book Ratio':'PB Percentile',
                'Price-to-Sales Ratio': 'PS Percentile',
                'EV/EBITDA':'EV/EBITDA Percentile',
                'EV/GP':'EV/GP Percentile'
    }

    for row in rv_dataframe.index:
        for metric in metrics.keys():
            rv_dataframe.loc[row, metrics[metric]] = stats.percentileofscore(rv_dataframe[metric], rv_dataframe.loc[row, metric])/100

    # Print each percentile score to make sure it was calculated properly
    for metric in metrics.values():
        print(rv_dataframe[metric])
    from statistics import mean

    for row in rv_dataframe.index:
        value_percentiles = []
        for metric in metrics.keys():
            value_percentiles.append(rv_dataframe.loc[row, metrics[metric]])
        rv_dataframe.loc[row, 'RV Score'] = mean(value_percentiles)
    rv_dataframe.sort_values(by = 'RV Score', inplace = True)
    rv_dataframe = rv_dataframe[:50]
    rv_dataframe.reset_index(drop = True, inplace = True)

    position_size = float(portfolio_size) / len(rv_dataframe.index)
    for i in range(0, len(rv_dataframe['Ticker'])-1):
        rv_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / rv_dataframe['Price'][i])
    writer = pd.ExcelWriter('../excel/value_strategy.xlsx', engine='xlsxwriter')
    rv_dataframe.to_excel(writer, sheet_name='Value Strategy', index = False)
    background_color = '#0a0a23'
    font_color = '#ffffff'

    string_template = writer.book.add_format(
            {
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )

    dollar_template = writer.book.add_format(
            {
                'num_format':'$0.00',
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )

    integer_template = writer.book.add_format(
            {
                'num_format':'0',
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )

    float_template = writer.book.add_format(
            {
                'num_format':'0',
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )

    percent_template = writer.book.add_format(
            {
                'num_format':'0.0%',
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            }
        )
    column_formats = {
                        'A': ['Ticker', string_template],
                        'B': ['Price', dollar_template],
                        'C': ['Number of Shares to Buy', integer_template],
                        'D': ['Price-to-Earnings Ratio', float_template],
                        'E': ['PE Percentile', percent_template],
                        'F': ['Price-to-Book Ratio', float_template],
                        'G': ['PB Percentile',percent_template],
                        'H': ['Price-to-Sales Ratio', float_template],
                        'I': ['PS Percentile', percent_template],
                        'J': ['EV/EBITDA', float_template],
                        'K': ['EV/EBITDA Percentile', percent_template],
                        'L': ['EV/GP', float_template],
                        'M': ['EV/GP Percentile', percent_template],
                        'N': ['RV Score', percent_template]
                     }

    for column in column_formats.keys():
        writer.sheets['Value Strategy'].set_column(f'{column}:{column}', 25, column_formats[column][1])
        writer.sheets['Value Strategy'].write(f'{column}1', column_formats[column][0], column_formats[column][1])
    writer.close()