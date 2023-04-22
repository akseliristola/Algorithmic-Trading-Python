import pandas as pd
#Read the tickers and ave them in stocks
stocks=pd.read_csv('../sp_500_stocks.csv')
#Function to divide the ticker list to multiple lists of size n
def chunks(lst,n):
    for i in range(0,len(lst),n):
        yield lst[i:i+n]
#Setting a portfolio size
portfolio_size=float(1000000)