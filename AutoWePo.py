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
from tabulate import tabulate


import win32console
_stdin = win32console.GetStdHandle(win32console.STD_INPUT_HANDLE)


# In[ ]:


def _readReport():
    global GlobalVar
    GlobalVar.reportDf = pd.read_excel(GlobalVar.fileName, header = 0, dtype = str)
    GlobalVar.reportDf = GlobalVar.reportDf.replace(np.nan, "")
    _reorder()
    
def _reorder():
    global GlobalVar
    checkPart = GlobalVar.reportDf.loc[GlobalVar.reportDf['SKILL'] == "Check"].sort_values(by=["AP"])
    complementaryPart = GlobalVar.reportDf.loc[GlobalVar.reportDf['SKILL'] != "Check"].sort_values(by=["OA_NO", "AP", "OA_DESC", "SKILL"])
    if checkPart.size > 0:
        GlobalVar.reportDf = pd.concat([checkPart, complementaryPart]).reset_index(drop=True)
    else:
        GlobalVar.reportDf = complementaryPart.reset_index(drop=True)
#         complementaryPart = GlobalVar.reportDf.loc[GlobalVar.reportDf['SKILL'] != "Check"]
#         complementaryPart = complementaryPart.sort_values(by=["OA_NO", "AP", "SKILL"])
    

def _showBrief():
    def autoSubstring(targetString, n = None):
        n = 0 if n == None else n
        if n == 0 or n >= len(targetString):
            return targetString
        else:
            return targetString[:n] + "..."
    global GlobalVar
    tempDataFrame = GlobalVar.reportDf.copy().applymap(lambda x: autoSubstring(x, GlobalVar.substringLength))
    displayMetadata = list()
    for i in range(len(GlobalVar.displayColumns)):
        displayMetadata.append(GlobalVar.metadata[GlobalVar.displayColumns[i]])
    print(tabulate(tempDataFrame[displayMetadata], headers='keys', tablefmt="grid"))

def _saveXlsx():
    global GlobalVar
    GlobalVar.reportDf["A_DATE"] = _getFirstDayOfWeek("")
    GlobalVar.reportDf["ITEM"] = np.arange(len(GlobalVar.reportDf)) + 1
    GlobalVar.reportDf["OWNER"] = GlobalVar.owner
    GlobalVar.reportDf.to_excel(GlobalVar.fileName, index = False)
    

            
def _getFirstDayOfWeek(gap = "/"):
    date = datetime.datetime.today()
    start = date - datetime.timedelta(days=date.weekday())
    return start.strftime(f"%Y{gap}%m{gap}%d")

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
                            
def displayAll():
    global GlobalVar
    print(tabulate(GlobalVar.reportDf, headers='keys', tablefmt="grid"))
    
def addNewRow(assignDict = []):
    global GlobalVar
    newDataDict = dict()
    for key in GlobalVar.metadata:
        if key not in GlobalVar.constMetadata:
            if len(assignDict) == 0:
#                 newDataDict[key] = input(key + ": ")
                newDataDict[key] = GlobalVar.filter(key, showMessage = f"{key}: ")
            elif key in assignDict:
                newDataDict[key] = GlobalVar.filter(key, showMessage = f"{key}: ")
            else:
                newDataDict[key] = ""
        else:
            newDataDict[key] = ""
    GlobalVar.reportDf = GlobalVar.reportDf.append(newDataDict, ignore_index=True)
    _reorder()
    _showBrief()

def saveExcel():
    _saveXlsx()

def showUsingHour():
    _timeRecorder("+",0)
    
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
    
def calcHours():
    global GlobalVar
    print(sum(list(map(float, GlobalVar.reportDf["W_HOUR"]))))

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

def removeRow():
    global GlobalVar
    index = int(input("Which row?"))
    if input(f"Confirm to delete {index} row? enter n to stop ") != "n":
        GlobalVar.reportDf = GlobalVar.reportDf.drop(GlobalVar.reportDf.index[index])
        _reorder()
        _showBrief()
    else:
        print("canceled")

def updateOaInfo():
    global GlobalVar
    oaList = list(GlobalVar.reportDf.loc[GlobalVar.reportDf["OA_NO"] != "", "OA_NO"])
    userId = input("user id: ")
    password = getpass.getpass("enter password: ")
    print(brobot.fillInOaInfo(GlobalVar.reportDf, oaList, userId, password))
    _reorder()
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
        

                            
                                                              

    



# In[ ]:


class GlobalVar():
    reportDf = None
    metadata =  ['A_DATE', 'ITEM', 'OA_DESC', 'AP', 'SKILL', 'SITE', 'DUE_DATE', 'COMPLET_D', 'OWNER', 'IT_STATUS', 
                 'OA_NO', 'PROGRAM', 'W_HOUR', 'REMARK', 'PROG_CNT', 'OA_STATUS']
    constMetadata = ['A_DATE', 'ITEM', 'OWNER']
    dateMetadata = ['DUE_DATE', 'COMPLET_D']
    displayColumns = [2, 12, 13, 10, 15,]
    functionDict = {"new": "addNewRow()", 
                    "newoa": "addNewRow(['OA_NO', 'AP', 'SKILL', 'IT_STATUS', 'PROGRAM', 'W_HOUR','REMARK', 'PROG_CNT'])", 
                    "all": "displayAll()", "save": "saveExcel()", "calchour": "calcHours()", 
                    "edit": "editRow()", "remove": "removeRow()", "update": "updateOaInfo()",
                    "addhour": "addHours()", "hour": "showUsingHour()",
                    "done": "doneOa()"}
    owner = None
    fileName = None
    substringLength = 30
    
    @staticmethod
    def filter(targetColumn, targetIndex = -1, showMessage = "", defaultString = "", valueString = ""):
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
        if targetColumn == "AP":
            valueString = checkInput(["SAP", "Meeting", "Training", "User Support", "文件製作", "Others"], 
                                     showMessage, defaultString, restrictive=False)
        elif targetColumn == "SKILL":
            valueString = checkInput(["ABAP", "Check", "java", "jsp"], 
                                     showMessage, defaultString, restrictive=False)
        elif targetColumn == "DUE_DATE":
            while(True):
                valueString = input(f"(t for today) {showMessage}")
                try:
                    if valueString.lower() == "t":
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
                                        f"(t for today) {showMessage}", restrictive = False)
                try:
                    if valueString.lower() == "t":
                        valueString = datetime.datetime.today().strftime(f"%Y/%m/%d")
                    else:
                        datetime.datetime.strptime(valueString, '%Y/%m/%d')
                    break
                except ValueError:
                    print(f"please input string which be like {_getFirstDayOfWeek()}, or input index")

        elif targetColumn == "IT_STATUS":
            valueString = checkInput(["設計完成", "設計中"], showMessage)
        elif targetColumn == "OA_NO":
            pattern = re.compile(r"(^SAI\d{6}$)")
            while(True):
                if defaultString == "":
                    valueString = input(showMessage + "SAI").strip()
                else:
                    valueString = _input_def(showMessage, defaultString).strip()
                if valueString == "":
                    break
                valueString = "SAI" + valueString
                matcher = pattern.match(valueString)
                try:
                    matcher.group(0)
                    break
                except AttributeError:
                    print("invalid OA number, please input 6 digits")
        elif targetColumn == "PROGRAM":
            valueString = _input_def(showMessage, defaultString)
        elif targetColumn == "W_HOUR":
            pattern = re.compile(r"(?P<operator>[\+-])\s*(?P<number>\d*\.?\d*)")
            if defaultString == "":
                defaultString = "0"
            while(True):
                try:
                    matcher = pattern.match(input(f"Type operation(ex. +2.25): {defaultString} "))
                    valueString = str(eval(f"{float(defaultString)}{matcher.group('operator')}{matcher.group('number')}"))
                    _timeRecorder(matcher.group('operator'), float(matcher.group('number')))
                    break                                                                              
                except AttributeError:
                    print("Please input operator + or - with number ex. +2.5")                
        elif targetColumn == "REMARK":
            valueString = _input_def(showMessage, defaultString)
        elif targetColumn == "PROG_CNT":
            valueString = checkInput(["0", "1"], showMessage)
        elif targetColumn == "OA_STATUS":
            if valueString == "":
                valueString = checkInput(["done"], showMessage)
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



        else:
            valueString = _input_def(showMessage, defaultString)
        return valueString

def _timeRecorder(operator, number):
    defaultFileName = "timeRecorder"
    if os.path.isfile(defaultFileName):
        with open(defaultFileName, "r+") as file:
            lastTime, usedHour = tuple(file.read().split(","))
            lastDate = datetime.datetime.strptime(lastTime, "%Y/%m/%d").date()
            if lastDate == datetime.date.today():
                usedHour = eval(f"{float(usedHour)}{operator}{number}")
            else:
                usedHour = number;
                print(f"Another new day, set calculator of hours to {usedHour}")
            file.seek(0)
            file.write(f"{datetime.date.today().strftime('%Y/%m/%d')},{usedHour}")
    else:
        print(f"Can not find {defaultFileName}, set calculator of hours to {number}")
        usedHour = number
        with open(defaultFileName, "w") as file:
            file.write(datetime.datetime.today().strftime("%Y/%m/%d") + f",{usedHour}")
    print(f"Record {usedHour} hours today")
        
def _firstExecute():
    global GlobalVar
    print("First time execute, setting...")
    with open("person.config", "w") as newConfig:
        newConfig.write("owner=" + input("Owner Name? ").strip() + "\n")

        newConfig.write("fileName=" + input("import weekly report? (keep empty if no) file name: ") + "\n")

        newConfig.write("displayColumns="  + ",".join(list(map(str,GlobalVar.displayColumns))) + "\n")
        
        newConfig.write(f"substringLength={GlobalVar.substringLength}\n")
    print("Setting Ok")

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
                    if key == "fileName":
                        GlobalVar.fileName = value
                    if key == "displayColumns":
                        GlobalVar.displayColumns = list(map(int, list(map(str.strip, value.split(",")))))
                    if key == "substringLength":
                        GlobalVar.substringLength = value
        except StopIteration: # EOF
            pass
        
        
def _selfCheck():
    global GlobalVar
    if not os.path.isfile("person.config"):
        _firstExecute()
        
    _readConfig()
    
    #check owner
    if GlobalVar.owner == None:
        print("ERROR: do not get owner name")
        return False
    
    #check file
    if GlobalVar.fileName != None and os.path.isfile(GlobalVar.fileName):
        _readReport()
    else:
        GlobalVar.fileName = f"WeeklyReport-V1.0-{_getFirstDayOfWeek('.')}-{GlobalVar.owner}.xlsx"
        GlobalVar.reportDf = pd.DataFrame(columns = GlobalVar.metadata)
        with open("person.config", "r+") as config:
            pattern = re.compile(r"(?P<key>[a-zA-Z1-9]*)=(?P<value>.*)")
            with open("person.config", "r+") as config:
                lines = config.readlines()
                for index, line in enumerate(lines):
                    matcher = pattern.match(line.strip())
                    if matcher:
                        key = matcher.group("key")
                        value = matcher.group("value")
                        if key == "fileName":
                            lines[index] = f"fileName={GlobalVar.fileName}\n"
                config.seek(0)
                config.writelines(lines)
        _saveXlsx()
    
    #check displayColumns
    try:
        for columnIndex in GlobalVar.displayColumns:
            GlobalVar.metadata[columnIndex]
    except (IndexError, TypeError):
        print(f"load displayColumns error, default displayColumns to {','.join(list(map(str,GlobalVar.displayColumns)))}")

                       
    #check substringLength
    try:
        int(GlobalVar.substringLength)
    except (ValueError, TypeError):
        print("load substringLength error, default substringLength to 30")
        GlobalVar.substringLength = 30
                       
    return True

def _controller():
    global GlobalVar
    showFlag = True
    action = None
    _showBrief()
    while(True):
        action = input(f"[{datetime.datetime.now().strftime('%Y/%m/%d %H:%M')}] To do? ").lower().strip()
        if action in GlobalVar.functionDict:
#             try:
                eval(GlobalVar.functionDict[action])
#             except:
#                 print("Unexpected error:", sys.exc_info()[0])
        elif action == "?":
            for key in GlobalVar.functionDict:
                print(key)
        elif action == "ex":
            break
        else:
            print("Unknow function")

def initializeApp():
    if(_selfCheck()):
        _timeRecorder("+",0)
        _controller()


# In[ ]:


initializeApp()

