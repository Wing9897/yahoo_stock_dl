import yfinance as yf
import pandas as pd

# 指定股票代碼和日期範圍（例如 AAPL 從 2020 到現在）
ticker = 'AAPL'
data = yf.download(ticker, start='2020-01-01', end='2025-09-23')

# 儲存為 CSV
data.to_csv(f'{ticker}_historical.csv')
print(data.head())  # 顯示前幾行