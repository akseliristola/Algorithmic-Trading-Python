import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math
from scipy import stats
from secretss import IEX_CLOUD_API_TOKEN
from statistics import mean
from constants import  stocks,chunks,portfolio_size

def momentumStratExcel():
    symbol_groups = list(chunks(stocks['Ticker'], 100))
    symbol_strings = []

    for i in range(0, len(symbol_groups)):
        symbol_strings.append(','.join(symbol_groups[i]))

    hqm_columns = [
                    'Ticker',
                    'Price',
                    'Number of Shares to Buy',
                    'One-Year Price Return',
                    'One-Year Return Percentile',
                    'Six-Month Price Return',
                    'Six-Month Return Percentile',
                    'Three-Month Price Return',
                    'Three-Month Return Percentile',
                    'One-Month Price Return',
                    'One-Month Return Percentile',
                    'HQM Score']
    hqm_dataframe = pd.DataFrame(columns = hqm_columns)
    for symbol_string in symbol_strings:
        batch_api_call_url = f'https://cloud.iexapis.com/stable/stock/market/batch/?types=stats,quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
        data = requests.get(batch_api_call_url).json()
        for symbol in symbol_string.split(','):
            hqm_dataframe = hqm_dataframe._append(
                                            pd.Series([symbol,
                                                       data[symbol]['quote']['latestPrice'],
                                                       'N/A',
                                                       data[symbol]['stats']['year1ChangePercent'],
                                                       'N/A',
                                                       data[symbol]['stats']['month6ChangePercent'],
                                                       'N/A',
                                                       data[symbol]['stats']['month3ChangePercent'],
                                                       'N/A',
                                                       data[symbol]['stats']['month1ChangePercent'],
                                                       'N/A',
                                                       'N/A'
                                                       ],
                                                      index = hqm_columns),
                                            ignore_index = True)
    time_periods = [
                    'One-Year',
                    'Six-Month',
                    'Three-Month',
                    'One-Month'
                    ]

    for row in hqm_dataframe.index:
        for time_period in time_periods:
            hqm_dataframe.loc[row, f'{time_period} Return Percentile'] = stats.percentileofscore(np.nan_to_num(hqm_dataframe[f'{time_period} Price Return']), hqm_dataframe.loc[row, f'{time_period} Price Return'])/100

    # Print each percentile score to make sure it was calculated properly
    for time_period in time_periods:
        pass
        #print(hqm_dataframe[f'{time_period} Return Percentile'])

    for row in hqm_dataframe.index:
        momentum_percentiles = []
        for time_period in time_periods:
            momentum_percentiles.append(hqm_dataframe.loc[row, f'{time_period} Return Percentile'])
        hqm_dataframe.loc[row, 'HQM Score'] = mean(momentum_percentiles)
    hqm_dataframe.sort_values('HQM Score',ascending=False,inplace=True)
    hqm_dataframe=hqm_dataframe[:51]
    position_size = float(portfolio_size) / len(hqm_dataframe.index)

    for row in hqm_dataframe.index:
        hqm_dataframe.loc[row, 'Number of Shares to Buy'] = math.floor(position_size / hqm_dataframe['Price'][row])

    hqm_dataframe.reset_index(drop = True, inplace = True)

    print(hqm_dataframe)
    writer = pd.ExcelWriter('../excel/momentum_strategy.xlsx', engine='xlsxwriter')
    hqm_dataframe.to_excel(writer, sheet_name='Momentum Strategy', index = False)

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
                        'D': ['One-Year Price Return', percent_template],
                        'E': ['One-Year Return Percentile', percent_template],
                        'F': ['Six-Month Price Return', percent_template],
                        'G': ['Six-Month Return Percentile', percent_template],
                        'H': ['Three-Month Price Return', percent_template],
                        'I': ['Three-Month Return Percentile', percent_template],
                        'J': ['One-Month Price Return', percent_template],
                        'K': ['One-Month Return Percentile', percent_template],
                        'L': ['HQM Score', string_template]
                        }

    for column in column_formats.keys():
        writer.sheets['Momentum Strategy'].set_column(f'{column}:{column}', 20, column_formats[column][1])
        writer.sheets['Momentum Strategy'].write(f'{column}1', column_formats[column][0], string_template)

    writer.close()