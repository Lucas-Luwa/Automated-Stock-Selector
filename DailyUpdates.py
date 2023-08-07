import yfinance as yf
import sys

def main(ticker):
    tick = yf.Ticker(ticker)
    tdData = tick.history(period='1d')
    return tdData['Close'][0]
    print(get_current_price('GOOG'))


if __name__ == "__main__":
    main
