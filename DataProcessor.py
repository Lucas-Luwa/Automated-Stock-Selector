import openpyxl
import time
import datetime
import calendar
import os

#Modify Toggles if desired
wipeCurrVerNum = True
rawDataFileName = "May2023RawDataV1.xlsx"


def main():
    readExcelFileName = generateFileName() #We write to this
    rawDataBook = openpyxl.load_workbook(rawDataFileName)
    rawData = rawDataBook.active

    print(readExcelFileName)

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

#-100 to +300
#6/12 fields
#>40 insider
#>20 short
# if stock is missing PE, or any series 3
# if avg all roic or last 5 avg or last 3 above 2 we good. otherwise eliminated
#assets less than 1000 eliminated

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
        3: "E3: Missing 6 or more of Series 1 fields",
        4: "E4: Missing 6 or more of Series 2 fields",
        5: "E5: Series 3 Value is missing",
        6: "E6: ROIC value is below (last 3, last 5 avg and all avg)",
        7: "E7: Assets below $1000",
        8: "E8: >40% of company is held by insiders",
        9: "E9: >20% of company shares shorted",
        }

    return switch.get(input, "")
if __name__ == "__main__":
    main()