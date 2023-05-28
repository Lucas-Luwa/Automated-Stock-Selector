import yfinance as yf

def get_current_price(ticker):
    tick = yf.Ticker(ticker)
    tdData = tick.history(period='1d')
    return tdData['Close'][0]
print(get_current_price('GOOG'))