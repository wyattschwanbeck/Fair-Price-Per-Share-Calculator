# -*- coding: utf-8 -*-
"""
Written by Wyatt Schwanbeck via sources from scrapehero

"""

import yfinance as yf

def get_stock_data(ticker):
    # Download historical data for the given ticker
    stock = yf.Ticker(ticker)
    hist = stock.history(period="5d")  # download max available data
    
    # Calculate beta using historical data
    returns = (hist['Close'].pct_change() + 1).cumprod()
    #market_returns = (hist['SPY'].pct_change() + 1).cumprod()  # assume SP500 as the market index
    beta = 1#(returns.cov(market_returns) / market_returns.var()).mean()

    return {"Open":hist["Open"][-1], "Beta":beta}