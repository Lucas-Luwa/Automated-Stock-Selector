import openpyxl
import time
import datetime
import calendar
import os

def main():
    readExcelFileName = generateFileName()
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
        writeToRecovery(0, None)
        versionNumber = int(recoveryLines[7].split(" ")[1].strip()) + 1
    proposedFileName = "ProcessedDataV" + str(versionNumber) + ".xlsx"

    directoryFiles = os.listdir(filePath)
    #Check this proposed fileName
    contentModified = False
    while proposedFileName in directoryFiles:
        contentModified = True
        versionNumber += 1
        proposedFileName = "ProcessedDataV" + str(versionNumber) + ".xlsx"
    if contentModified:
        writeToRecovery(1, versionNumber)
    return proposedFileName

def writeToRecovery(toggle, versionNumber):
    recoveryFile = open('Recovery.txt', 'r')
    recoveryLines = recoveryFile.readlines()
    recoveryLines[7] = "DataProcessVersion: " + \
        (str((int(recoveryLines[7].split(" ")[1].strip()) + 1)) if toggle == 0 else str(versionNumber))
    recoveryFile = open('Recovery.txt', 'w')
    for i in range(0, len(recoveryLines)):
        recoveryFile.write(str(recoveryLines[i]))
    recoveryFile.close()
# def recordVersion

if __name__ == "__main__":
    main()