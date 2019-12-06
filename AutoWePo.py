#!/usr/bin/env python
# coding: utf-8

# In[12]:


import pandas as pd
import numpy as np
import os
from tabulate import tabulate


# In[25]:


def _readReport():
    global GlobalVar
    GlobalVar.reportDf = pd.read_excel(GlobalVar.fileName, header = 0)
    #replace datetime to only date 
    for dateColumn in GlobalVar.dateMetadata:
        GlobalVar.reportDf[dateColumn] = GlobalVar.reportDf[dateColumn].dt.date
    #replace nan to empty string
    GlobalVar.reportDf = GlobalVar.reportDf.replace(np.nan, "")
    _reorder()
    
def _reorder():
    global GlobalVar
    checkPart = GlobalVar.reportDf.loc[GlobalVar.reportDf['SKILL'] == "Check"].sort_values(by=["AP"])
    complementaryPart = GlobalVar.reportDf.loc[GlobalVar.reportDf['SKILL'] != "Check"].sort_values(by=["OA_NO", "AP", "OA_DESC", "SKILL"])
    if checkPart.size > 0:
        GlobalVar.reportDf = pd.concat([checkPart, complementaryPart]).reset_index()
    else:
        GlobalVar.reportDf = complementaryPart.reset_index()
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
    GlobalVar.reportDf.to_excel(GlobalVar.fileName, index = False)
    
def _controller():
    global GlobalVar
    showFlag = True
    action = None
    _showBrief()
    while(True):
        action = input("To do? ").lower().strip()
        if action in GlobalVar.functionDict:
#             try:
                GlobalVar.functionDict[action]()
#             except:
#                 print("Unexpected error:", sys.exc_info()[0])
        elif action == "?":
            for key in GlobalVar.functionDict:
                print(key)
        elif action == "ex":
            break
        else:
            print("Unknow function")
            

def displayAll():
    global GlobalVar
    print(tabulate(GlobalVar.reportDf, headers='keys', tablefmt='psql'))
    
def addNewRow():
    global GlobalVar
    newDataDict = dict()
    for key in GlobalVar.metadata:
        if key not in GlobalVar.constMetadata:
            newDataDict[key] = input(key + ": ")
        else:
            newDataDict[key] = ""
    GlobalVar.reportDf = GlobalVar.reportDf.append(newDataDict, ignore_index=True)
    _reorder()
    _showBrief()

def saveExcel():
    _saveXlsx()

def editRow():
    global GlobalVar
    index = input("Which row?")
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
        GlobalVar.reportDf.at[int(index), metadataPair[editTarget]] = input(f"{GlobalVar.reportDf.at[int(index), metadataPair[editTarget]]} ->")
    else:
        for metadata in GlobalVar.metadata:
            if metadata not in GlobalVar.constMetadata:
                GlobalVar.reportDf.at[int(index), metadata] = input(f"{metadata}:{GlobalVar.reportDf.iloc[int(index)][metadata]} ->")
    _showBrief()

class GlobalVar():
    reportDf = None
    metadata =  ['A_DATE', 'ITEM', 'OA_DESC', 'AP', 'SKILL', 'SITE', 'DUE_DATE', 'COMPLET_D', 'OWNER', 'IT_STATUS', 'OA_NO', 'PROGRAM', 'W_HOUR', 'REMARK', 'PROG_CNT', 'OA_STATUS']
    constMetadata = ['A_DATE', 'ITEM', 'OWNER']
    dateMetadata = ['DUE_DATE', 'COMPLET_D']
    displayColumns = [2, 12, 13]
    functionDict = {"new": addNewRow, "all": displayAll, "save": saveExcel, "edit": editRow}
#     fileName = "WeeklyReport-V1.0-2019.09.30-TonyOu.xlsx"
    fileName = "new.xlsx"
    
def initializeApp():
    _readReport()
    _controller()


# In[26]:


initializeApp()

