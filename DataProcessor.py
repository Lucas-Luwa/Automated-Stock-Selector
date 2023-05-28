import openpyxl
import time
import datetime
import calendar
import os

#Modify Toggles if desired
wipeCurrVerNum = True

def main():
    global sheetNames, writeExcelFileName, rawDataBook, core1, rowIndecies
    start = time.time()
    writeExcelFileName = generateFileName() #We write to this
    rawDataFileName = findRawDataFileName()
    rawDataBook = openpyxl.load_workbook(rawDataFileName)#Pull data from this 
    rawData = rawDataBook.active
    coreName = "CoreExcelFiles/P2MasterTemplate5.28.23.xlsx"
    core1 = openpyxl.load_workbook(coreName)
    sheetNames = core1.sheetnames
    
    rowIndecies = [3] * (len(sheetNames))
    #Iterate through it, but for now we go for Miscellaneous
    excelWriter()

def excelWriter():
    stoppingIndecies = [702, 1220, 1053, 73, 45, 148, 175, 792, 261, 182, 1498, 68, 680]
    counter = 0;
    for row in rawDataBook['Miscellaneous'].iter_rows(3, 45 - 1):
        currSheet = 'Miscellaneous'
        tempWKST = core1['Miscellaneous']
        sheetIndex = sheetNames.index('Miscellaneous')
        #TAGS
        if redFlagsS1(row, currSheet) and redFlagsS2(row) and redFlagsS3(row, currSheet):
            for i in range(1,6):
                tempWKST.cell(row = rowIndecies[sheetIndex], column = i).value = row[i - 1].value
            #SERIES 1
            for i in range(7, 16):
                x = i - 2
                if i >= 8: x += 2
                if i >= 10: x += 1
                if i >= 15: i += 1
                if not i == 14:
                    tempWKST.cell(row = rowIndecies[sheetIndex], column = i).value = row[x].value
                else: 
                    high, low = row[x].value.split('/')
                    tempWKST.cell(row = rowIndecies[sheetIndex], column = i).value = high
                    tempWKST.cell(row = rowIndecies[sheetIndex], column = i + 1).value = low
            #SERIES 3
            counter = 34;
            for i in range (215, 220):
                if i == 218: counter += 2
                tempWKST.cell(row = rowIndecies[sheetIndex], column = i).value = row[counter].value
                counter += 1
            rowIndecies[sheetIndex] += 1
            core1.save("ProcessedSheets\\" + monthYR + "\\" + writeExcelFileName)

def redFlagsS1(row, sheetName):
    #41
    #PE, MKTCAP, Share Short, %Insider, %Institution, 52HighLow, DailyTrade - Series 1
    performanceValues = [5, 8, 10, 11, 12, 15, 16] 
    performanceLength = list()
    errorNum = 0
    for i in performanceValues:
        performanceLength.append(len(removeNonNumeric(row[i].value)))
        if i == 5 and len(removeNonNumeric(row[i].value)) == 0: #Nothing is there PE
            return errorHandler(row, errorNum, sheetName)
        if i == 5 and float(removeNonNumeric(row[i].value)) < -100: #PE Over 300
            return errorHandler(row, errorNum + 1, sheetName)
        if i == 5 and float(removeNonNumeric(row[i].value)) > 300: #PE Under -100
            return errorHandler(row, errorNum + 2, sheetName)
        if i == 8 and len(removeNonNumeric(row[i].value)) == 0: #Nothing is there MKTCAP
            return errorHandler(row, errorNum + 3, sheetName)
        if i == 10 and len(removeNonNumeric(row[i].value)) == 0: #Nothing is there SHARE SHORT
            return errorHandler(row, errorNum + 4, sheetName)
        if i == 10 and float(removeNonNumeric(row[i].value)) > 20.0: #Short percentage is over 20 percent...
            return errorHandler(row, errorNum + 5, sheetName)
        if i == 11 and len(removeNonNumeric(row[i].value)) == 0: #Nothing is there Percent Insider
            return errorHandler(row, errorNum + 6, sheetName)
        if i == 12 and float(removeNonNumeric(row[i].value)) == 0: #Nothing is there Institution %
            return errorHandler(row, errorNum + 7, sheetName)
        if i == 15 and len(removeNonNumeric(row[i].value)) == 0: #Nothing is there high low
            return errorHandler(row, errorNum + 8, sheetName)
        if i == 16 and len(removeNonNumeric(row[i].value)) == 0: #Nothing is there daily trade volume
            return errorHandler(row, errorNum + 9, sheetName)
        if i == 16 and performanceLength.count(0) >= 6: # 6 or more fields missing in Series
            return errorHandler(row, errorNum + 10, sheetName)
    return True

def redFlagsS2(row):
    return True

# if avg all roic or last 5 avg or last 3 above 2 we good. otherwise eliminated
#assets less than 1000 eliminated
def redFlagsS3(row, sheetName):
    performanceValues = [34, 35, 40] 
    #Total Liabilities, Total Assets and Number of Employees
    errorNum = 11
    for i in performanceValues: 
        if i == 35 and len(removeNonNumeric(row[i].value)) == 0: #Nothing is there Total Liabilities
            return errorHandler(row, errorNum, sheetName)
        if i == 36 and len(removeNonNumeric(row[i].value)) == 0: #Nothing is there Total Assets
            return errorHandler(row, errorNum + 1, sheetName)
        # if i == 36 and float(removeNonNumeric(row[i].value)) < 10: #Assets less than $10 Million
        #     return errorHandler(row, errorNum + 2, sheetName)
    return True
def removeNonNumeric(input):
    output = ""
    approvedSet = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
    approvedSet2 = ['-', '.']
    counter = 0
    for i in range(0, len(input) - 1):
        if (input[i] in approvedSet):
            counter += 1
            output+= input[i]
        if (input[i] in approvedSet2):
            output+= input[i]
    if counter == 0: return ""
    return output

def errorHandler(row, flag, industry):
    tempWKST = core1['ELIMINATED']
    sheetIndex = sheetNames.index('ELIMINATED')
    flagComment = getErrorCode(flag)
    for i in range(1,7):
        x = i
        if i == 5: 
            tempWKST.cell(row = rowIndecies[sheetIndex] - 1, column = i).value = industry
        else:
            if i > 5: x -= 1
            tempWKST.cell(row = rowIndecies[sheetIndex] - 1, column = i).value = row[x - 1].value
    tempWKST.cell(row = rowIndecies[sheetIndex] - 1, column = 7).value = flagComment
    rowIndecies[sheetIndex] += 1
    core1.save("ProcessedSheets\\" + monthYR + "\\" + writeExcelFileName)
    return False;

def findRawDataFileName():
    global monthYR
    recoveryFile = open('Recovery.txt')
    recoveryLines = recoveryFile.readlines()
    if (len(recoveryLines) < 3):
        print("ERROR: Please check that your Recovery.txt file has been generated properly.")
        monthYR = "ERROR"
        versionNum = 0
    else:
        monthYR = recoveryLines[4].split(" ")[1].strip()
        versionNum = recoveryLines[5].split(" ")[1].strip()
    return str(monthYR) + "RawDataV" + str(versionNum) + ".xlsx"

def generateFileName():
    recoveryFile = open('Recovery.txt')
    recoveryLines = recoveryFile.readlines()
    recoveryFile.close()
    MonthYear = recoveryLines[4].split(" ")[1].strip()
    filePath = "C://Users//User//Documents//GitHub//Automated-Stock-Selector//ProcessedSheets"
    directoryFiles = os.listdir(filePath)
    if not (MonthYear in directoryFiles): #Create the folder if it isn't already there.
        folderPath = os.path.join(filePath, MonthYear)
        os.mkdir(folderPath)
    filePath = os.path.join(filePath, MonthYear)
    versionNumber = -1
    #Initial Proposed file.
    if len(recoveryLines) == 7: 
        recoveryFile = open("Recovery.txt","a")
        recoveryFile.write("\nDataProcessVersion: " + str(1))
        versionNumber = 1
        recoveryFile.close()
    else: 
        writeToRecovery(0, None, wipeCurrVerNum)
        versionNumber = (int(recoveryLines[7].split(" ")[1].strip()) + 1) if not wipeCurrVerNum else 1
    proposedFileName = "ProcessedDataV" + str(versionNumber) + ".xlsx"

    directoryFiles = os.listdir(filePath)
    #Check this proposed fileName
    contentModified = False
    while proposedFileName in directoryFiles:
        contentModified = True
        versionNumber += 1
        proposedFileName = "ProcessedDataV" + str(versionNumber) + ".xlsx"
    if contentModified:
        writeToRecovery(1, versionNumber, False)
    return proposedFileName

def writeToRecovery(toggle, versionNumber, wipeCurrVerNum):
    recoveryFile = open('Recovery.txt', 'r')
    recoveryLines = recoveryFile.readlines()
    if not wipeCurrVerNum:
        recoveryLines[7] = "DataProcessVersion: " + \
            (str((int(recoveryLines[7].split(" ")[1].strip()) + 1)) if toggle == 0 else str(versionNumber))
    else:
        recoveryLines[7] = "DataProcessVersion: " + str(1)
    recoveryFile = open('Recovery.txt', 'w')
    for i in range(0, len(recoveryLines)):
        recoveryFile.write(str(recoveryLines[i]))
    recoveryFile.close()

def getErrorCode(input):
    switch = {
        0: "E0: P/E Ratio is missing",
        1: "E1: P/E Ratio is below the cutoff of -100",
        2: "E2: P/E Ratio is above the cutoff of +300",
        3: "E3: Missing Market Cap Value",
        4: "E4: Missing Shares Shorted Value",
        5: "E5: Shares Shorted Value over 20 percent",
        6: "E6: Missing Percent of Insiders",
        7: "E7: Percent held by institutions is missing",
        8: "E8: 52 Week High and Low missing",
        9: "E9: Daily trade volume missing",
        10: "E10: Missing 6 or more of Series 1 fields",
        11: "E11: Missing Total Liabilities",
        12: "E12: Missing Total Assets",
        13: "E13: Total assets are below $1000",
        14: "E14: Missing 6 or more of Series 2 fields",
        15: "E15: Series 3 Value is missing",
        16: "E16: ROIC value is below (last 3, last 5 avg and all avg)",
        }

    return switch.get(input, "")
if __name__ == "__main__":
    main()