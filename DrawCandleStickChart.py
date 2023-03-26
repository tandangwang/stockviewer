import matplotlib.pyplot as plt
from mpl_finance import candlestick_ohlc
import pandas as pd
import matplotlib.dates as mpl_dates
import mplfinance as mpf

# # Read data from CSV file
# data = pd.read_csv('file_result.xlsx.csv')
# data['Date'] = pd.to_datetime(data['Date'])
# data['Date'] = data['Date'].apply(mpl_dates.date2num)
# data = data.loc[:, ['Date', 'Open', 'High', 'Low', 'Close']]

# Read data from CSV file
data = pd.read_excel('file_result.xlsx', sheet_name='Sheet1')
data['日期'] = pd.to_datetime(data['日期'])
data['日期'] = data['日期'].apply(mpl_dates.date2num)
data = data.loc[:, ['日期', '开盘', '最高', '最低', '收盘']]

# Creating Subplots
fig, ax = plt.subplots()

# Creating Candlestick chart
candlestick_ohlc(ax, data.values, width=0.6, colorup='green', colordown='red', alpha=0.8)

# Setting labels & titles
ax.set_xlabel('Date')
ax.set_ylabel('Price')
fig.suptitle('Daily Candlestick Chart of file_result')

# Formatting Date
date_format = mpl_dates.DateFormatter('%d-%m-%Y')
ax.xaxis.set_major_formatter(date_format)
fig.autofmt_xdate()

# Show plot
plt.show()
