import yfinance as yf
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Fetching historical data from Yahoo Finance
def fetch_data(ticker, start_date, end_date):
    data = yf.download(ticker, start=start_date, end=end_date)
    data = data[['Close']]
    return data

# Moving Average calculation
def calculate_moving_average(data, window):
    data[f'MA_{window}'] = data['Close'].rolling(window=window).mean()
    return data

# Bollinger Bands calculation
def calculate_bollinger_bands(data, window=20):
    data['MA_20'] = data['Close'].rolling(window=window).mean()
    rolling_std = data['Close'].rolling(window=window).std().squeeze()  # Ensure it's a Series
    data['BB_upper'] = data['MA_20'] + 2 * rolling_std
    data['BB_lower'] = data['MA_20'] - 2 * rolling_std
    return data

# CCI calculation
def calculate_cci(data, window=20):
    tp = (data['Close'] + data['Close'] + data['Close']) / 3
    ma = tp.rolling(window=window).mean()
    md = tp.rolling(window=window).apply(lambda x: np.mean(np.abs(x - x.mean())))
    data['CCI'] = (tp - ma) / (0.015 * md)
    return data

# Plot Moving Average
def plot_moving_average(data):
    plt.figure(figsize=(12, 6))
    plt.plot(data['Close'], label='Close Price')
    plt.plot(data['MA_1'], label='1-Day MA')
    plt.plot(data['MA_7'], label='1-Week MA')
    plt.title('Moving Average')
    plt.xlabel('Date')
    plt.ylabel('Price')
    plt.legend()
    plt.grid()
    plt.savefig('moving_average_plot.png')
    plt.close()

# Plot Bollinger Bands
def plot_bollinger_bands(data):
    plt.figure(figsize=(12, 6))
    plt.plot(data['Close'], label='Close Price')
    plt.plot(data['MA_20'], label='20-Day MA (Middle Band)')
    plt.plot(data['BB_upper'], label='Upper Band', linestyle='--')
    plt.plot(data['BB_lower'], label='Lower Band', linestyle='--')
    plt.fill_between(data.index, data['BB_lower'], data['BB_upper'], color='lightgrey', alpha=0.5)
    plt.title('Bollinger Bands')
    plt.xlabel('Date')
    plt.ylabel('Price')
    plt.legend()
    plt.grid()
    plt.savefig('bollinger_bands_plot.png')
    plt.close()

# Plot CCI
def plot_cci(data):
    plt.figure(figsize=(12, 6))
    plt.plot(data['CCI'], label='CCI', color='purple')
    plt.axhline(100, color='red', linestyle='--')
    plt.axhline(-100, color='green', linestyle='--')
    plt.title('Commodity Channel Index (CCI)')
    plt.xlabel('Date')
    plt.ylabel('CCI Value')
    plt.legend()
    plt.grid()
    plt.savefig('cci_plot.png')
    plt.close()

# Decision-making based on technical indicators
def make_decision(data):
    decisions = {}

    # Access the last row of the DataFrame for decision making
    last_row = data.iloc[-1]

    # Extract scalar values to avoid Series comparison issues
    close_price = last_row['Close'].item()
    ma_1 = last_row['MA_1'].item()
    bb_upper = last_row['BB_upper'].item()
    bb_lower = last_row['BB_lower'].item()
    cci = last_row['CCI'].item()

    # Decision for Moving Average
    if close_price > ma_1:
        decisions['Moving Average'] = 'BUY'
    elif close_price < ma_1:
        decisions['Moving Average'] = 'SELL'
    else:
        decisions['Moving Average'] = 'NEUTRAL'

    # Decision for Bollinger Bands
    if close_price > bb_upper:
        decisions['Bollinger Bands'] = 'SELL'
    elif close_price < bb_lower:
        decisions['Bollinger Bands'] = 'BUY'
    else:
        decisions['Bollinger Bands'] = 'NEUTRAL'

    # Decision for CCI
    if cci > 100:
        decisions['CCI'] = 'SELL'
    elif cci < -100:
        decisions['CCI'] = 'BUY'
    else:
        decisions['CCI'] = 'NEUTRAL'
        
    return decisions    

# Main function to execute the analysis
def main():
    # Parameters
    ticker = "EURINR=X"
    start_date = "2023-01-01"
    end_date = "2024-09-30"

    # Fetch and prepare data
    data = fetch_data(ticker, start_date, end_date)
    
    # Calculate indicators
    data = calculate_moving_average(data, window=1)   # 1-day MA
    data = calculate_moving_average(data, window=7)   # 1-week MA
    data = calculate_bollinger_bands(data, window=20)
    data = calculate_cci(data, window=20)
    
    # Plot each indicator
    plot_moving_average(data)
    plot_bollinger_bands(data)
    plot_cci(data)
    
    # Display the last row for verification
    print(data.tail(1))
    
    # Make decisions based on the latest data
    decisions = make_decision(data)
    print("Decisions based on the indicators:", decisions)

    # Save the decisions to be used in the PowerPoint report
    data.to_csv("EURINR_analysis_results.csv")

if __name__ == "__main__":
    main()