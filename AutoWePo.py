#!/usr/bin/env python
# coding: utf-8

# In[27]:


import pandas as pd
import sys
import os
from tabulate import tabulate


# In[31]:


def _readReport():
    global GlobalVar
    GlobalVar.reportDf = pd.read_excel("WeeklyReport-V1.0-2019.09.30-TonyOu.xlsx", header = 0)
    for dateColumn in GlobalVar.dateMetadata:
        GlobalVar.reportDf[dateColumn] = GlobalVar.reportDf[dateColumn].dt.date

def _showBrief():
    global GlobalVar
    displayMetadata = list()
    for i in range(len(GlobalVar.displayColumns)):
        displayMetadata.append(GlobalVar.metadata[GlobalVar.displayColumns[i]])
    print(tabulate(GlobalVar.reportDf[displayMetadata], headers='keys', tablefmt='psql'))

def _saveXlsx():
    global GlobalVar
    GlobalVar.reportDf.to_excel(GlobalVar.fileName, index = False)
    
def saveExcel():
    _saveXlsx()
    
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
    _showBrief()
    
def _controller():
    global GlobalVar
    showFlag = True
    action = None
    _showBrief()
    while(True):
        action = input("To do? ").lower().strip()
        if action in GlobalVar.functionDict:
            try:
                GlobalVar.functionDict[action]()
            except:
                print("Unexpected error:", sys.exc_info()[0])
        else:
            print("Unknow function")
def initializeApp():
    _readReport()
    _controller()

class GlobalVar():
    reportDf = None
    metadata =  ['A_DATE', 'ITEM', 'OA_DESC', 'AP', 'SKILL', 'SITE', 'DUE_DATE', 'COMPLET_D', 'OWNER', 'IT_STATUS', 'OA_NO', 'PROGRAM', 'W_HOUR', 'REMARK', 'PROG_CNT', 'OA_STATUS']
    constMetadata = ['A_DATE', 'ITEM']
    dateMetadata = ['DUE_DATE', 'COMPLET_D']
    displayColumns = [2, 12, 13]
    functionDict = {"new": addNewRow, "all": displayAll, "save": saveExcel}
    fileName = "WeeklyReport-V1.0-2019.09.30-TonyOu.xlsx"


# In[ ]:


initializeApp()

