#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import numpy as np
import os
import re
import sys
import datetime
import brobot
import getpass
import math
import traceback
from tabulate import tabulate

import win32console
_stdin = win32console.GetStdHandle(win32console.STD_INPUT_HANDLE)


# In[ ]:


def _breakString(textString, breakLength = 35):
    textLength = len(textString)
    tokenLocation = breakLength
    while(tokenLocation < textLength or breakLength <= 0):
        for index in range(tokenLocation, textLength):
            if textString[index] == ",":
                textString = textString[:index + 1] + "\n" + textString[index + 1:]
                textLength += 1
                tokenLocation = index + 1 + breakLength # +1 for skiping \n
                break
            elif textString[index] == " ":
                textString = textString[:index] + "\n" + textString[index + 1:]
                textLength += 1
                tokenLocation = index + 1 + breakLength # +1 for skiping \n
                break
        if index == textLength - 1:
            break
    return textString

def _calculator(inputString):
    inputString = inputString.replace(" ", "")
    calculateResult = 0
    pattern = re.compile(r"(?P<operator>[-\+])\s*(?P<number>\d*\.?\d*)")
    matcherList = pattern.findall(inputString)
    if len(matcherList) == 0 and len(inputString) != 0:
        raise ValueError
    try:
        validLen = len(inputString)
        executeLen = 0
        for group in matcherList:
            executeLen += len(group[0] + group[1])
            calculateResult = eval(f"{calculateResult}{group[0]}{group[1]}")
        if(executeLen == validLen): #if some characters of inputString are not fed to calculation, there is invalid input
            return calculateResult
        else:
            raise ValueError
    except:
        raise

def _controller():
    global GlobalVar
    showFlag = True
    action = None
    _showBrief()
    _timeRecorder(0)              
    while(True):
        if GlobalVar.workingIdentity != None:
            workHint = f"[Working on Row {_getIndexByIdentity(GlobalVar.workingIdentity)}]"
        else:
            workHint = ""
        timeHint = f"[{_getNowWithOffset().strftime('%Y/%m/%d %H:%M')}{'' if GlobalVar.minuteOffset == 0 else '(' + str(GlobalVar.minuteOffset) + ')'}]"
        action = input(f"{timeHint}{workHint} To do? ").lower().strip()
        if action in GlobalVar.functionDict:
            try:
                eval(GlobalVar.functionDict[action])
            except Exception as e:
                error_class = e.__class__.__name__ #取得錯誤類型
                detail = e.args[0] #取得詳細內容
                cl, exc, tb = sys.exc_info() #取得Call Stack
                lastCallStack = traceback.extract_tb(tb)[-1] #取得Call Stack的最後一筆資料
                fileName = lastCallStack[0] #取得發生的檔案名稱
                lineNum = lastCallStack[1] #取得發生的行號
                funcName = lastCallStack[2] #取得發生的函數名稱
                errMsg = "File \"{}\", line {}, in {}: [{}] {}".format(fileName, lineNum, funcName, error_class, detail)
                print(errMsg)
        elif action == "?":
            for key in GlobalVar.functionDict:
                print(key)
        elif action == "ex":
            break
        else:
            print("Unknow function")

def _createNewFileNameAndDataFrame():
    global GlobalVar
    tempFileName = f"WeeklyReport-V1.0-{GlobalVar.firstDayOfWorkWeek.replace('/','.')}-{GlobalVar.owner}"
    fileExt = ".xlsx"
    count = 0
    while os.path.isfile(tempFileName + fileExt):
        count +=1
        tempFileName = f"{tempFileName}-{count}"
    GlobalVar.fileName = tempFileName + fileExt
    GlobalVar.reportDf = pd.DataFrame(columns = GlobalVar.metadata + ["identity"])

def _defaultTerminalSize():
    with open("systemCommand", "r") as file:
        for line in file:
            if os.system(line) != 0:
                print(f"Fail to execute \"{line}\"")  
def _firstExecute():
    global GlobalVar
    print("First time execute, setting...")
    ownerName = input("Owner Name? ").strip()
    fileName = _input_def("Import file name?(Keep empty if none) ", "" if GlobalVar.fileName == None else GlobalVar.fileName)
    with open("person.config", "w") as newConfig:
        newConfig.write(f"owner={ownerName}\n")
        newConfig.write(f"fileName={fileName}\n")
        newConfig.write("displayColumns="  + ",".join(list(map(str,GlobalVar.displayColumns))) + "\n")
        newConfig.write(f"substringLength={GlobalVar.substringLength}\n")
        newConfig.write(f"apOption={','.join(list(map(str, GlobalVar.apOption)))}\n")
        newConfig.write(f"skillOption={','.join(list(map(str, GlobalVar.skillOption)))}\n")
        newConfig.write(f"minuteOffset={GlobalVar.minuteOffset}\n")
        newConfig.write(f"firstDayOfWorkWeek={_getFirstDayOfWeek()}\n")
        newConfig.write(f"breakStringLength={GlobalVar.breakStringLength}\n")
    print("Setting Ok")
    return True

def _getConfig(targetKey):
    pattern = re.compile(r"(?P<key>" + targetKey + ")=(?P<value>.*)")
    with open("person.config", "r") as config:
        lines = config.readlines()
        for index, line in enumerate(lines):
            matcher = pattern.match(line.strip())
            if matcher:
                return matcher.group("value")
def _getDataByIdentity(identity, columnName):
    global GlobalVar
    return GlobalVar.reportDf[GlobalVar.reportDf["identity"] == identity][columnName].tolist()[0]
                        
def _getFirstDayOfWeek(gap = "/", typeOfDate = 'str'):
    date = datetime.datetime.today()
    start = date - datetime.timedelta(days=date.weekday())
    if typeOfDate == 'str':
        return start.strftime(f"%Y{gap}%m{gap}%d")
    elif typeOfDate == 'datetime':
        return start.replace(hour=0, minute=0, second=0, microsecond=0)
                        
def _getIndexByIdentity(identity):
    global GlobalVar
    return GlobalVar.reportDf[GlobalVar.reportDf["identity"] == identity].index.tolist()[0]
          
def _getNowWithOffset():
    global GlobalVar
    return datetime.datetime.now() + datetime.timedelta(minutes = GlobalVar.minuteOffset)

def _initializeTerminal():
    print("Initializing terminal...")
    width = 140
    height = 30
    if not os.path.isfile("systemCommand"):
        with open("systemCommand", "w") as file:
            #set window size
            file.write(f"mode {width},{height}\n")
            #set window buffer size
            file.write(f'powershell -command "&{{$H=get-host;$W=$H.ui.rawui;$B=$W.buffersize;$B.width={width};$B.height=600;$W.buffersize=$B;}}"\n')
    with open("systemCommand", "r") as file:
        for index, line in enumerate(file):
            print(f"Executing command {index}", end = "\r")
            if os.system(line) != 0:
                print(f"Fail to execute \"{line}\"")                        

def _input_def(prompt, default=''):
    keys = []
    for c in str(default):
        evt = win32console.PyINPUT_RECORDType(win32console.KEY_EVENT)
        evt.Char = c
        evt.RepeatCount = 1
        evt.KeyDown = True
        keys.append(evt)

    _stdin.WriteConsoleInput(keys)
    return input(prompt)


def _readConfig():
    global GlobalVar
    pattern = re.compile(r"(?P<key>[a-zA-Z1-9]*)=(?P<value>.*)")
    with open("person.config", "r") as config:
        try:
            while True:
                content = next(config)
                matcher = pattern.match(content.strip())
                if matcher:
                    key = matcher.group("key")
                    value = matcher.group("value")
                    if key == "owner":
                        GlobalVar.owner = value
                    elif key == "fileName":
                        GlobalVar.fileName = value
                    elif key == "displayColumns":
                        GlobalVar.displayColumns = list(map(int, list(map(str.strip, value.split(",")))))
                    elif key == "substringLength":
                        GlobalVar.substringLength = value
                    elif key == "apOption":
                        GlobalVar.apOption = list(map(str, list(map(str.strip, value.split(",")))))
                    elif key == "skillOption":
                        GlobalVar.skillOption = list(map(str, list(map(str.strip, value.split(",")))))
                    elif key == "minuteOffset":
                        GlobalVar.minuteOffset = value
                    elif key == "firstDayOfWorkWeek":
                        GlobalVar.firstDayOfWorkWeek = value
                    elif key == "breakStringLength":
                        GlobalVar.breakStringLength = value
        except StopIteration: # EOF
            pass

def _readReport():
    global GlobalVar
    GlobalVar.reportDf = pd.read_excel(GlobalVar.fileName, header = 0, dtype = str)
    GlobalVar.reportDf = GlobalVar.reportDf.replace(np.nan, "")
    
    #get each row a identity
    GlobalVar.reportDf = GlobalVar.reportDf.assign(identity = pd.Series(np.arange(GlobalVar.reportDf.shape[0])).array).astype(str)
        
def _reorder():
    global GlobalVar
    checkPart = GlobalVar.reportDf.loc[GlobalVar.reportDf['SKILL'] == "Check"].sort_values(by=["AP"])
    complementaryPart = GlobalVar.reportDf.loc[GlobalVar.reportDf['SKILL'] != "Check"].sort_values(by=["OA_NO", "AP", "OA_DESC", "SKILL"])
    if checkPart.size > 0:
        GlobalVar.reportDf = pd.concat([checkPart, complementaryPart]).reset_index(drop=True)
    else:
        GlobalVar.reportDf = complementaryPart.reset_index(drop=True)
    GlobalVar.reportDf = GlobalVar.reportDf.astype(str)

def _saveXlsx():
    global GlobalVar
    GlobalVar.reportDf["A_DATE"] = GlobalVar.firstDayOfWorkWeek.replace("/", "")
    GlobalVar.reportDf["ITEM"] = np.arange(len(GlobalVar.reportDf)) + 1
    GlobalVar.reportDf["OWNER"] = GlobalVar.owner
    if "identity" in GlobalVar.reportDf.columns:
        tempDataFrame = GlobalVar.reportDf.drop(["identity"], axis = 1)
    else:
        tempDataFrame = GlobalVar.reportDf
    tempDataFrame.to_excel(GlobalVar.fileName, index = False)

def _selfCheck():
    global GlobalVar
    
    firstTime = False
    if not os.path.isfile("person.config"):
        firstTime = _firstExecute()
        
    _readConfig()
    
    #check week
    newWeek = False
    if not firstTime:
        if GlobalVar.firstDayOfWorkWeek == None:
            GlobalVar.firstDayOfWorkWeek = _getFirstDayOfWeek()
            _updateConfig("firstDayOfWorkWeek", GlobalVar.firstDayOfWorkWeek)
            newWeek = True
        else:
            try:
                workWeekInConfig = datetime.datetime.strptime(GlobalVar.firstDayOfWorkWeek, "%Y/%m/%d")
                if workWeekInConfig < _getFirstDayOfWeek(typeOfDate="datetime"):
                    answer = input("New week arrive, press enter to create new week report or input any character to stay in old week. ")
                    if answer == "":
                        GlobalVar.firstDayOfWorkWeek = _getFirstDayOfWeek()
                        _updateConfig("firstDayOfWorkWeek", GlobalVar.firstDayOfWorkWeek)
                        newWeek = True
            except ValueError:
                GlobalVar.firstDayOfWorkWeek = _getFirstDayOfWeek()
                _updateConfig("firstDayOfWorkWeek", GlobalVar.firstDayOfWorkWeek)
                newWeek = True
                print("Get error when read firstDayOfWorkWeek, set it to new week")

                        
    #check owner
    if GlobalVar.owner == None:
        GlobalVar.owner = "NoName"
        print("no owner name, default owner name to NoName")
    
    #check file        
    if GlobalVar.fileName != None and os.path.isfile(GlobalVar.fileName) and not newWeek:
        try:
            _readReport()
            _reorder()
        except FileNotFoundError:
            _createNewFileNameAndDataFrame()
            _updateConfig("fileName", GlobalVar.fileName)
            _saveXlsx()                               
            print(f"Weekly report not found, creating one named {GlobalVar.fileName}")
    else:
        _createNewFileNameAndDataFrame()
        _updateConfig("fileName", GlobalVar.fileName)
        _saveXlsx()                         

                      
    
    #check displayColumns
    try:
        for columnIndex in GlobalVar.displayColumns:
            GlobalVar.metadata[columnIndex]
    except (IndexError, TypeError):
        print(f"load displayColumns error, default displayColumns to {','.join(list(map(str,GlobalVar.displayColumns)))}")

    #check substringLength
    try:
        GlobalVar.substringLength = int(GlobalVar.substringLength)
    except (ValueError, TypeError):
        print("load substringLength error, default substringLength to 0 (no effected)")
        GlobalVar.substringLength = 0

    #check apOption
    if GlobalVar.apOption == None:
        print(f"loading apOption error, default apOption to {','.join(list(map(str, GlobalVar.apOption)))}")
    
    #check skillOption
    if GlobalVar.skillOption == None:
        print(f"loading skillOption error, default skillOption to {','.join(list(map(str, GlobalVar.skillOption)))}")
    
    #check minuteOffset
    try:
        GlobalVar.minuteOffset = int(GlobalVar.minuteOffset)
    except ValueError:
        print(f"loading minuteOffset error, default minuteOffset to 0")
        GlobalVar.minuteOffset = 0
    
    #check breakStringLength
    try:
        GlobalVar.breakStringLength = int(GlobalVar.breakStringLength)
    except (ValueError, TypeError):
        print("load breakStringLength error, default breakStringLength to 35")
        GlobalVar.breakStringLength = 35
              
    return True
              
def _settleWork():
    global GlobalVar
    floorMinute = math.floor(int(_getNowWithOffset().strftime("%M")) / 15) * 15
    nowDate = _getNowWithOffset().replace(minute=floorMinute, second=0, microsecond=0)
                      
    timeDelta = nowDate - datetime.datetime.strptime(GlobalVar.workingTime, "%Y/%m/%d %H:%M:%S")
    usingHours = math.ceil((timeDelta.total_seconds() / 3600.0) / 0.25) * 0.25
    print(f"Work settle: Row {_getIndexByIdentity(GlobalVar.workingIdentity)} {_getDataByIdentity(GlobalVar.workingIdentity, 'OA_DESC')}, working time {usingHours} hours")
    GlobalVar.reportDf.at[_getIndexByIdentity(GlobalVar.workingIdentity), "W_HOUR"] = GlobalVar.filter("W_HOUR", 
                                                                             _getIndexByIdentity(GlobalVar.workingIdentity), 
                                                                             defaultString = GlobalVar.reportDf.at[_getIndexByIdentity(GlobalVar.workingIdentity), "W_HOUR"],
                                                                             valueString = f"+{usingHours}")
    GlobalVar.workingIdentity = None
    GlobalVar.workingTime = None
    
def _showBrief():
    def autoSubstring(targetString, n = None):
        n = 0 if n == None else int(n)
        if n == 0 or n >= len(targetString):
            return targetString
        else:
            return targetString[:n] + "..."
    global GlobalVar
    tempDataFrame = GlobalVar.reportDf.copy().applymap(lambda x: autoSubstring(str(x), GlobalVar.substringLength))
    displayMetadata = list()
    for i in range(len(GlobalVar.displayColumns)):
        displayMetadata.append(GlobalVar.metadata[GlobalVar.displayColumns[i]])
    print(tabulate(tempDataFrame[displayMetadata], headers='keys', tablefmt="grid"))

def _resizeTerminal(width, height):
    height = 30 if height < 30 else height
    width = 120 if width < 120 else width
    #set window buffer size
    os.system(f"mode con: cols={width} lines={height}")
    #set window buffer size
    #{{ to escape {
    os.system(f'powershell -command "&{{$H=get-host;$W=$H.ui.rawui;$B=$W.buffersize;$B.width={width};$B.height={height};$W.buffersize=$B;}}"')

          
def _timeRecorder(number):
    defaultFileName = "timeRecorder"
    if os.path.isfile(defaultFileName):
        with open(defaultFileName, "r+") as file:
            try:
                lastTime, usedHour = tuple(file.read().split(","))
            except ValueError:
                print(f"No data in {defaultFileName}, reset to now")
                lastTime, usedHour = datetime.datetime.today().strftime("%Y/%m/%d"), "0"
            lastDate = datetime.datetime.strptime(lastTime, "%Y/%m/%d").date()
            if lastDate == datetime.date.today():
                usedHour = eval(f"{float(usedHour)}+{number}")
            else:
                usedHour = number;
                print(f"Another new day, set calculator of hours to {usedHour}")
            file.seek(0)
            file.truncate()
            file.write(f"{datetime.date.today().strftime('%Y/%m/%d')},{usedHour}")
    else:
        print(f"Can not find {defaultFileName}, set calculator of hours to {number}")
        usedHour = number
        with open(defaultFileName, "w") as file:
            file.write(datetime.datetime.today().strftime("%Y/%m/%d") + f",{usedHour}")
    print(f"Record {usedHour} hours today")

def _updateConfig(targetKey, targetValue):
    pattern = re.compile(r"(?P<key>" + targetKey + ")=(?P<value>.*)")
    find = False
    with open("person.config", "r+") as config:
        lines = config.readlines()
        for index, line in enumerate(lines):
            matcher = pattern.match(line.strip())
            if matcher:
                lines[index] = f"{targetKey}={targetValue}\n"
                find = True
                break
        config.seek(0)
        config.writelines(lines)
    if not find:
        with open("person.config", "a") as config:
            config.write(f"{targetKey}={targetValue}\n")
                       
def addHours():
    global GlobalVar
    while(True):
        try:
            index = int(input("Which row? ").strip())
            old = float(GlobalVar.reportDf.at[index, "W_HOUR"])  
            GlobalVar.reportDf.at[index, "W_HOUR"] = GlobalVar.filter("W_HOUR", index, "W_HOUR: ", old)
            break                                                                              
        except (KeyError, ValueError):
            print(f"input is not valid")
    _showBrief()
                       
def addNewRow(assignColumnDict = [], defaultColumnDict = dict()):
    global GlobalVar
    newDataDict = dict()
    if len(GlobalVar.reportDf["identity"]) != 0:
        newRowIdentity = int(GlobalVar.reportDf["identity"].max()) + 1
    else:
        newRowIdentity = 0
    
    for key in GlobalVar.metadata:
        if key not in GlobalVar.constMetadata:
            if len(assignColumnDict) == 0:
#                 newDataDict[key] = input(key + ": ")
                newDataDict[key] = GlobalVar.filter(key, showMessage = f"{key}: ")
            elif key in assignColumnDict:
                newDataDict[key] = GlobalVar.filter(key, showMessage = f"{key}: ")
            elif defaultColumnDict.get(key) != None:
                newDataDict[key] = GlobalVar.filter(key, valueString = defaultColumnDict.get(key))
            else:
                newDataDict[key] = ""
        else:
            newDataDict[key] = ""
    newDataDict["identity"] = str(newRowIdentity)
    GlobalVar.reportDf = GlobalVar.reportDf.append(newDataDict, ignore_index=True)
    _reorder()
    _showBrief()
    return newRowIdentity

def calcHours():
    global GlobalVar
    print(sum(list(map(float, GlobalVar.reportDf["W_HOUR"]))))

def displayAll():
    def getWidthAndHeight(string):
        width = string.find("\n") + 1
        height = len(string) // width + 1
        return width, height
    global GlobalVar
    if "identity" in GlobalVar.reportDf.columns:
        outputString = tabulate(GlobalVar.reportDf.drop(["identity"], axis = 1), headers='keys', tablefmt="grid")
    else:
        outputString = tabulate(GlobalVar.reportDf, headers='keys', tablefmt="grid")
    width, height = getWidthAndHeight(outputString)
    _resizeTerminal(width=width, height=height)
    print(outputString)
    input("Press any key to continue...")
    _defaultTerminalSize()
    _showBrief()
                       
def displaySimple():
    _showBrief()

def doneOa():
    global GlobalVar
    while(True):
        try:
            index = int(input("Which row? ").strip())
            temp = GlobalVar.filter("OA_STATUS", index, valueString = "done") #temp for try and except
            GlobalVar.reportDf.at[index, "OA_STATUS"] = temp
            break                                                                              
        except (KeyError, ValueError):
            print(f"input is not valid")
    _showBrief()

def editRow():
    global GlobalVar
    while(True):
        try:
            index = int(input("Which row? ").strip())
            GlobalVar.reportDf.at[int(index), "A_DATE"] #"A_DATE" is only for testing that if index exists 

            count = 0
            metadataPair = {"all": "all"}
            for column in GlobalVar.metadata:
                if column not in GlobalVar.constMetadata:
                    count += 1
                    print(f"{count}.{column}", end = "｜")
                    metadataPair[str(count)] = column            
            editTarget = str(input("Which column? Or all?").lower().strip())
            if editTarget != "all":
                GlobalVar.reportDf.at[index, metadataPair[editTarget]]
            break
        except (KeyError, ValueError):
            print(f"input is not valid")

    
    if editTarget != "all":
        GlobalVar.reportDf.at[index, metadataPair[editTarget]] = GlobalVar.filter(metadataPair[editTarget], index,
                                                                    f"{metadataPair[editTarget]}: ", 
                                                                    GlobalVar.reportDf.at[int(index), metadataPair[editTarget]])
    else:
        for metadata in GlobalVar.metadata:
            if metadata not in GlobalVar.constMetadata:
                GlobalVar.reportDf.at[index, metadata] = GlobalVar.filter(metadata, index,
                                                                    f"{metadata}: ",
                                                                    GlobalVar.reportDf.at[index, metadata])
    _reorder()
    _showBrief()

def removeRow():
    global GlobalVar
    while(True):
        try:
            index = int(input("Which row? "))
            if input(f"Confirm to delete Row {index} {GlobalVar.reportDf.at[index, 'OA_DESC']}? Enter n to stop ") != "n":
                if GlobalVar.workingIdentity != None:
                    if _getIndexByIdentity(GlobalVar.workingIdentity) == index:
                        GlobalVar.workingIdentity = None
                        GlobalVar.workingTime = None
                GlobalVar.reportDf = GlobalVar.reportDf.drop(GlobalVar.reportDf.index[index])
                _reorder()
                _showBrief()
            else:
                print("canceled")
            break
        except (KeyError, ValueError):
             print(f"input is not valid")
                       
def resetPersonConfig():
    global GlobalVar
    takeBreak()
    _firstExecute()

    _readConfig()

    #check file
    if GlobalVar.fileName != None and os.path.isfile(GlobalVar.fileName):
        _readReport()
        _reorder()
        _showBrief()
    else:
        _createNewFileNameAndDataFrame()
        _updateConfig("fileName", GlobalVar.fileName)
        _saveXlsx()     

def resetTimer():
    defaultFileName = "timeRecorder"
    while(True):
        try:
            resetNum = int(input("Reset to? "))
            with open(defaultFileName, "w") as file:
                file.write(datetime.datetime.today().strftime("%Y/%m/%d") + "," + str(resetNum))
            print("Reset successfully")
            break
        except ValueError:
            print(f"input is not valid")

                     
def saveExcel():
    _saveXlsx()

def setting():
    global GlobalVar
    for index, value in enumerate(GlobalVar.settingOption):
        print(f"{index}. {value} |" , end=" ")
    print()
    while(True):
        try:
            inputValue = int(input(f"Pick one? "))
            settingWord = GlobalVar.settingOption[inputValue]
            break
        except (TypeError, IndexError, ValueError):
            print(f"input is not valid")
    
    newSetting = _input_def(f"change {settingWord}: ", _getConfig(settingWord))
    try:
        newSetting = int(newSetting)
        exec(f"GlobalVar.{settingWord} = {newSetting}")
    except:
        exec(f"GlobalVar.{settingWord} = {newSetting.split(',')}")
    _updateConfig(settingWord, newSetting)
                     
def showUsingHour():
    _timeRecorder(0)
    
def takeBreak():
    global GlobalVar
    if GlobalVar.workingTime != None:
        _settleWork()      
        _showBrief()
    else:
        print("no processing work")

def transferHour():
    global GlobalVar
    while(True):
        try:
            fromIndex = int(input("From row: ").strip())
            toIndex = int(input("To row: ").strip())
            transHours = float(input("How many hours? ")) 
            GlobalVar.reportDf.at[fromIndex, "W_HOUR"] = str(float(GlobalVar.reportDf.at[fromIndex, "W_HOUR"]) - transHours)
            GlobalVar.reportDf.at[toIndex, "W_HOUR"] = str(float(GlobalVar.reportDf.at[toIndex, "W_HOUR"]) + transHours)
            break                                                                              
        except (KeyError, ValueError):
            print(f"input is not valid")
    _showBrief()

                     
def updateOaInfo():
    global GlobalVar
    oaList = list(GlobalVar.reportDf.loc[GlobalVar.reportDf["OA_NO"] != "", "OA_NO"])
    userId = input("user id: ")
    password = getpass.getpass("enter password: ")
    print(brobot.fillInOaInfo(GlobalVar.reportDf, oaList, userId, password))
    _reorder()
    _showBrief()
                     
def working():
    global GlobalVar
    if GlobalVar.workingTime != None:
        _settleWork()
        _showBrief()
    while(True):
        dictPattern = re.compile(r"addNewRow")
        candidateFunctionList = []
        validFunction = dict()
        for key, value in GlobalVar.functionDict.items():
            if dictPattern.match(value) != None:
                candidateFunctionList.append(key)
                validFunction[key] = value
        tempInput = input(f"Which row? Or create one? ({', '.join(candidateFunctionList)}) ").strip()
        try:     
            tempInput = int(tempInput)
            GlobalVar.workingIdentity = str(GlobalVar.reportDf.loc[tempInput, ["identity"]].tolist()[0])
            break
        except ValueError:
            if validFunction.get(tempInput) == None:
                print(f"input is not valid")
            else:
                GlobalVar.workingIdentity = str(eval(validFunction.get(tempInput)))
                break
        except KeyError:
            print(f"input is not valid")

            
    floorMinute = math.floor(int(_getNowWithOffset().strftime("%M")) / 15) * 15  
    GlobalVar.workingTime = _getNowWithOffset().replace(minute=floorMinute, second=0, microsecond=0).strftime("%Y/%m/%d %H:%M:%S")
    print(f"Start working: Row {_getIndexByIdentity(GlobalVar.workingIdentity)} {_getDataByIdentity(GlobalVar.workingIdentity, 'OA_DESC')}")


          
          
#Global
class GlobalVar():
    reportDf = None
    metadata =  ['A_DATE', 'ITEM', 'OA_DESC', 'AP', 'SKILL', 'SITE', 'DUE_DATE', 'COMPLET_D', 'OWNER', 'IT_STATUS', 
                 'OA_NO', 'PROGRAM', 'W_HOUR', 'REMARK', 'PROG_CNT', 'OA_STATUS']
    constMetadata = ['A_DATE', 'ITEM', 'OWNER']
    dateMetadata = ['DUE_DATE', 'COMPLET_D']
    settingOption = ["displayColumns", "substringLength", "apOption", "skillOption", "minuteOffset", "breakStringLength"]
    functionDict = {"new": "addNewRow()", 
                    "oa": "addNewRow(['OA_NO', 'AP', 'SKILL', 'IT_STATUS', 'PROGRAM', 'W_HOUR','REMARK', 'PROG_CNT'])",
                    "meeting": "addNewRow(['OA_DESC', 'DUE_DATE', 'COMPLET_D', 'W_HOUR', 'REMARK'], {'AP': 'Meeting', 'SKILL': 'Check', 'PROG_CNT': '0'})",
                    "all": "displayAll()", "simple": "displaySimple()", "save": "saveExcel()", "weekhour": "calcHours()", 
                    "edit": "editRow()", "remove": "removeRow()", "update": "updateOaInfo()",
                    "addhour": "addHours()", "dayhour": "showUsingHour()", "transfer": "transferHour()",
                    "done": "doneOa()", 
                    "work": "working()", "break": "takeBreak()",
                    "setting": "setting()", "resettimer": "resetTimer()"}
                    #"reset": "resetPersonConfig()"
    #temp work area
    workingIdentity = None
    workingTime = None
    
    #custom settings
    owner = None
    fileName = None
    displayColumns = [2, 12, 13, 10, 15,]
    substringLength = 0
    apOption = ["SAP", "Meeting", "Training", "User Support", "文件製作", "Others"]
    skillOption = ["ABAP", "Check", "java", "jsp"]
    minuteOffset = 0
    firstDayOfWorkWeek = None
    breakStringLength = 35
                     
    @staticmethod
    def filter(targetColumn, targetIndex = -1, showMessage = "", defaultString = "", valueString = ""):
        global GlobalVar
        def checkInput(candidateList, showMessage, defaultString = "",restrictive = True):
            for index, value in enumerate(candidateList):
                print(f"{index}. {value} |" , end=" ")
            print()
            while(True):
                try:
                    if restrictive:
                        inputValue = input(f"Pick one? {showMessage + defaultString} -> ")
                    else:
                        inputValue = _input_def(f"Pick one or input string? {showMessage}", defaultString)
                    
                    if inputValue == "":
                        return inputValue
                    else:
                        return candidateList[int(inputValue)]
                except ValueError:
                    if restrictive:
                        print("please input index")
                    else:
                        return inputValue
                except IndexError:
                    if restrictive:
                        print("please input valid index")
                    else:
                        return inputValue
        if valueString != "":
            if targetColumn == "OA_STATUS":
                if valueString == "done" and targetIndex != -1:
                    dueDate = GlobalVar.reportDf.at[targetIndex, "DUE_DATE"]
                    todayDate = datetime.datetime.today()
                    if dueDate == "":
                        GlobalVar.reportDf.at[targetIndex, "COMPLET_D"] = todayDate.strftime("%Y/%m/%d")
                        GlobalVar.reportDf.at[targetIndex, "DUE_DATE"] = todayDate.strftime("%Y/%m/%d")
                        print(f"DUE_DATE and COMPLET_D are filled in {todayDate.strftime('%Y/%m/%d')} automatically")
                    else:
                        if datetime.datetime.strptime(dueDate, "%Y/%m/%d") > todayDate:
                            GlobalVar.reportDf.at[targetIndex, "COMPLET_D"] = todayDate.strftime("%Y/%m/%d")
                            print(f"COMPLET_D is filled in {todayDate.strftime('%Y/%m/%d')} automatically")
                        else:
                            GlobalVar.reportDf.at[targetIndex, "COMPLET_D"] = dueDate
                            print(f"DUE_DATE is before today, so COMPLET_D is filled in DUE_DATE({dueDate}) automatically")
            elif targetColumn == "W_HOUR":
                calculateResult = _calculator(valueString)
                valueString = str(eval(f"{float(defaultString)}+{calculateResult}"))
                _timeRecorder(calculateResult)                                  
        elif targetColumn == "AP":
            valueString = checkInput(GlobalVar.apOption, showMessage, defaultString, restrictive=False)                
        elif targetColumn == "SKILL":
            valueString = checkInput(GlobalVar.skillOption, showMessage, defaultString, restrictive=False)
        elif targetColumn == "DUE_DATE":
            while(True):
                valueString = _input_def(f"(t for today) {showMessage}", defaultString)
                try:
                    if valueString == "": #to keep valueString empty
                        break                    
                    elif valueString.lower() == "t":
                        valueString = datetime.datetime.today().strftime(f"%Y/%m/%d")
                    else:
                        datetime.datetime.strptime(valueString, '%Y/%m/%d')
                    break
                except ValueError:
                    print(f"please input string which be like {_getFirstDayOfWeek()}")

        elif targetColumn == "COMPLET_D":
            monday = datetime.datetime.today() - datetime.timedelta(days=datetime.datetime.today().weekday())
            while(True):
                valueString = checkInput(list(map(lambda x : (monday + datetime.timedelta(x)).strftime("%Y/%m/%d"), list(range(5)))),
                                        f"(t for today) {showMessage}", defaultString = defaultString, restrictive = False)
                try:
                    if valueString == "": #to keep valueString empty
                        break                    
                    elif valueString.lower() == "t":
                        valueString = datetime.datetime.today().strftime(f"%Y/%m/%d")
                    else:
                        datetime.datetime.strptime(valueString, '%Y/%m/%d')
                    break
                except ValueError:
                    print(f"please input string which be like {_getFirstDayOfWeek()}, or input index")

        elif targetColumn == "IT_STATUS":
            valueString = checkInput(["設計完成", "設計中"], showMessage = showMessage, defaultString = defaultString)
        elif targetColumn == "OA_NO":
            pattern = re.compile(r"^(SAI){1,2}(\d{6})$")
            while(True):
                if defaultString == "":
                    valueString = _input_def(showMessage, "SAI").strip()
                else:
                    valueString = _input_def(showMessage, defaultString).strip()
                if valueString == "SAI" or valueString == "":
                    valueString = ""
                    break
                matcher = pattern.match(valueString)
                try:
                    valueString = matcher.group(1) + matcher.group(2)
                    break
                except AttributeError:
                    print("invalid OA number, please input SAI with 6 digits")
        elif targetColumn == "W_HOUR":
            if defaultString == "":
                defaultString = "0"
            while(True):
                try:
                    tempInput = input(f"Type operation(ex. +2.25) {showMessage}: {defaultString} ")
                    if tempInput == "":
                        tempInput = "+0"
                    calculateResult = _calculator(tempInput)
                    valueString = str(float(defaultString) + calculateResult)
                    _timeRecorder(calculateResult)
                    break                                                                              
                except (AttributeError, SyntaxError):
                    valueString = ""
                    print("Please input operator + or - with number ex. +2.5")                
        elif targetColumn == "PROG_CNT":
            valueString = checkInput(["0", "1"], showMessage = showMessage, defaultString = defaultString)
        elif targetColumn == "OA_STATUS":
            valueString = checkInput(["done"], showMessage = showMessage, defaultString = defaultString)
        elif targetColumn == "REMARK" or targetColumn == "PROGRAM":
            valueString = _breakString(_input_def(showMessage, defaultString), GlobalVar.breakStringLength)
        else:
            valueString = _input_def(showMessage, defaultString)
        return valueString

def initializeApp():
    try:
        _initializeTerminal()
        if(_selfCheck()):
            _controller()
    except:
        traceback.print_exc()
        os.system("pause")


# In[ ]:


initializeApp()

