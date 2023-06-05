import pandas as pd # I really wanted to call this catBear but will stay with normal naming conventions...iykyk ;)
import numpy

def main():
    
    pass

def getErrorCode(input):
    switch = {
        0: "E0: P/E Ratio is missing",
        1: "E1: P/E Ratio is below the cutoff of -100",
        2: "E2: P/E Ratio is above the cutoff of +300",
        3: "E3: Missing Market Cap Value",
        }
    return switch.get(input, "")

if __name__ == "__main__":
    main()