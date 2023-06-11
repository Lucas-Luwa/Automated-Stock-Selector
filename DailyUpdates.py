import yfinance as yf
import sys

def get_current_price(ticker):
    tick = yf.Ticker(ticker)
    tdData = tick.history(period='1d')
    return tdData['Close'][0]
    print(get_current_price('GOOG'))


if __name__ == "__main__":
    if len(sys.argv) > 1:
        method_name = sys.argv[1]
        if method_name == "":
            result = 
        elif method_name == "":
            result = 
        else:
            result = "Invalid method name"
        print(result)
    else:
        print("No method name provided")
