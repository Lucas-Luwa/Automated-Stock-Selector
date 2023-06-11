from scipy import stats
import math
import statistics
import numpy as np

def main():
    #GOOG
    # arr = [9.8, 11.3, 26.1, 25.2, 26.2, 50.5, 25.1, 23.9, 25.0, 21.9, 25.6, 26.8]
    #AMZN - Deal with cases like this 
    arr = [57.1, 45.8, 42.0, 53.5, 141.1, -2574.0, 497.0, -633.8, 375.2, 139.2, 153.6, 78.7, 75.9, 62.7, 50.6, -470.5, 291.5]
    retV, lbda = stats.yeojohnson(arr)
    currV = stats.yeojohnson(20, lbda)
    print("RES", currV)
    print(statistics.stdev(arr), statistics.mean(arr))
    print(statistics.stdev(retV), statistics.mean(retV))
    print(zScore(currV, statistics.mean(retV), statistics.stdev(retV)))
    calc = 1 - stats.norm.cdf(zScore(currV, statistics.mean(retV), statistics.stdev(retV)))
    print(calc)
    return retV

def zScore(val, mean, stdDev):   
    return (val - mean)/stdDev

if __name__ == "__main__":
    main()