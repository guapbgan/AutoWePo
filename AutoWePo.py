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
    global GlobalVar
    displayMetadata = list()
    for i in range(len(GlobalVar.displayColumns)):
        displayMetadata.append(GlobalVar.metadata[GlobalVar.displayColumns[i]])
    print(tabulate(GlobalVar.reportDf[displayMetadata], headers='keys', tablefmt='psql'))

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
            
def displayAll():
    global GlobalVar
    print(tabulate(GlobalVar.reportDf, headers='keys', tablefmt='psql'))
    
def addNewRow(assignDict = []):
    global GlobalVar
    newDataDict = dict()
    for key in GlobalVar.metadata:
        if key not in GlobalVar.constMetadata:
            if len(assignDict) == 0:
#                 newDataDict[key] = input(key + ": ")
                newDataDict[key] = GlobalVar.filter(key, f"{key}: ")
            elif key in assignDict:
                newDataDict[key] = GlobalVar.filter(key, f"{key}: ")
            else:
                newDataDict[key] = ""
        else:
            newDataDict[key] = ""
    GlobalVar.reportDf = GlobalVar.reportDf.append(newDataDict, ignore_index=True)
    _reorder()
    _showBrief()

def saveExcel():
    _saveXlsx()

def editRow():
    global GlobalVar
    index = input("Which row? ").strip()
    count = 0
    metadataPair = {"all": "all"}
    for column in GlobalVar.metadata:
        if column not in GlobalVar.constMetadata:
            count += 1
            print(f"{count}.{column}", end = "｜")
            metadataPair[str(count)] = column
    
    while(True):
        editTarget = input("Which column? Or all?").lower().strip()
        if editTarget in metadataPair:
            break
        else:
            print("Not correct. Try again.")
    
    if editTarget != "all":
        GlobalVar.reportDf.at[int(index), metadataPair[editTarget]] = GlobalVar.filter(metadataPair[editTarget], 
                                                                                       f"{metadataPair[editTarget]}: {GlobalVar.reportDf.at[int(index), metadataPair[editTarget]]} ->")
    else:
        for metadata in GlobalVar.metadata:
            if metadata not in GlobalVar.constMetadata:
                GlobalVar.reportDf.at[int(index), metadata] = GlobalVar.filter(metadata ,
                                                                               f"{metadata}: {GlobalVar.reportDf.at[int(index), metadata]} ->")
    _reorder()
    _showBrief()
    
def calcHours():
    global GlobalVar
    print(sum(list(map(float, GlobalVar.reportDf["W_HOUR"]))))
    
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
        

def firstExecute():
    global GlobalVar
    print("First time execute, setting...")
    with open("person.config", "w") as newConfig:
        newConfig.write("owner=" + input("Owner Name? ").strip() + "\n")

        newConfig.write("fileName=" + input("import weekly report? (keep empty if no) file name: ") + "\n")

        newConfig.write("simpleDisplayColumn="  + ", ".join(list(map(str,GlobalVar.displayColumns))) + "\n")
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
                    if key == "simpleDisplayColumn":
                        GlobalVar.displayColumns = list(map(int, list(map(str.strip, value.split(",")))))
        except StopIteration: # EOF
            pass
        
    



# In[ ]:


class GlobalVar():
    reportDf = None
    metadata =  ['A_DATE', 'ITEM', 'OA_DESC', 'AP', 'SKILL', 'SITE', 'DUE_DATE', 'COMPLET_D', 'OWNER', 'IT_STATUS', 
                 'OA_NO', 'PROGRAM', 'W_HOUR', 'REMARK', 'PROG_CNT', 'OA_STATUS']
    constMetadata = ['A_DATE', 'ITEM', 'OWNER']
    dateMetadata = ['DUE_DATE', 'COMPLET_D']
    displayColumns = [2, 12, 13, 10, 15,]
    functionDict = {"new": "addNewRow()", "newoa": "addNewRow(['OA_NO', 'AP', 'SKILL', 'IT_STATUS', 'PROGRAM', 'REMARK', 'PROG_CNT'])", 
                    "all": "displayAll()", "save": "saveExcel()", "calchour": "calcHours()", 
                    "edit": "editRow()", "remove": "removeRow()", "update": "updateOaInfo()"}
    owner = None
    fileName = None
    
    @staticmethod
    def filter(targetColumn, showMessage = ""):
        def checkInput(candidateList, showMessage, restrictive = True):
            for index, value in enumerate(candidateList):
                print(f"{index}. {value} |", end=" ")
            if restrictive:
                print()
            else:
                print("or other string")
            while(True):
                try:
                    inputValue = input(f"Which one? {showMessage}")
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
            valueString = checkInput(["SAP", "Meeting", "Training", "User Support", "文件製作", "Others"], showMessage, restrictive=False)
        elif targetColumn == "SKILL":
            valueString = checkInput(["ABAP", "Check", "java", "jsp"], showMessage, restrictive=False)
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
                valueString = "SAI" + input(showMessage + "SAI").strip()
                matcher = pattern.match(valueString)
                try:
                    matcher.group(0)
                    break
                except AttributeError:
                    print("invalid OA number, please input 6 digits")
        elif targetColumn == "PROGRAM":
            valueString = input(showMessage)
        elif targetColumn == "W_HOUR":
            while(True):
                try:
                    inputValue = input(showMessage)
                    if inputValue == "":
                        inputValue = 0
                    temp = float(inputValue)
                    break
                except ValueError:
                    print("please input digital")
            valueString = str(temp)
        elif targetColumn == "REMARK":
            valueString = input(showMessage)
        elif targetColumn == "PROG_CNT":
            valueString = checkInput(["0", "1"], showMessage)
        elif targetColumn == "OA_STATUS":
            valueString = checkInput(["done"], showMessage)
        else:
            valueString = input(showMessage)
        return valueString

    
def _selfCheck():
    global GlobalVar
    if not os.path.isfile("person.config"):
        firstExecute()
        
    _readConfig()
    
    if GlobalVar.owner == None:
        print("ERROR: do not get owner name")
        return False
    
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
    return True

def _controller():
    global GlobalVar
    showFlag = True
    action = None
    _showBrief()
    while(True):
        action = input("To do? ").lower().strip()
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
        _controller()


# In[ ]:


initializeApp()

