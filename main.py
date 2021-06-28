import requests
import sys
import traceback
import pandas as pd
import time
import numpy as np
import openpyxl
import requests
from openpyxl import load_workbook
from datetime import datetime
import matplotlib.pyplot as plt
import docx
import os

def updatefile():
    dls = "https://bdu.fstec.ru/files/documents/vullist.xlsx"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.106 Safari/537.36',
        'Referrer': 'https://bdu.fstec.ru/vul%27%7D'}
    resp = requests.get(dls, headers=headers)

    output = open('test.xlsx', 'wb')
    output.write(resp.content)
    output.close()

workbook = load_workbook('test.xlsx')
column = workbook.active

Everything = {}
Vulner = []
Description = []
Soft = []
SoftType = []
VulnerClass = []
Date = []
VulnerLVL = []
Status = []
Exploit = []
ProtInfo = []
CWEType = []
Measures = []
ID = []

def initial_file():
    Counter = 0
    for row in range(1, column.max_row):
        if 'Flash Player' in str(column.cell(row=row, column=5).value):
            Vulner.append(column.cell(row=row, column=2).value)
            Description.append(column.cell(row=row, column=3).value)
            Soft.append(column.cell(row=row, column=5).value)
            SoftType.append(column.cell(row=row, column=7).value)
            VulnerClass.append(column.cell(row=row, column=9).value)
            Date.append(column.cell(row=row, column=10).value)
            VulnerLVL.append(column.cell(row=row, column=13).value)
            Status.append(column.cell(row=row, column=15).value)
            Exploit.append(column.cell(row=row, column=16).value)
            ProtInfo.append(column.cell(row=row, column=17).value)
            CWEType.append(column.cell(row=row, column=21).value)
            Measures.append(column.cell(row=row, column=14).value)
            ID.append(column.cell(row=row, column=2).value)
            Counter += 1
    for i in range(Counter):
        Everything[i] = (
            Vulner[i], Description[i], Soft[i], SoftType[i], VulnerClass[i], Date[i],
            VulnerLVL[i],
            Status[i], Exploit[i], ProtInfo[i], CWEType[i], Measures[i], ID[i])


def CountCheker(Count):
    Count = 0
    for t in Count:
        Count += 1
    return Count

Length = CountCheker(Everything)

def TranslateToDate(Date):
    Translate = "%d.%m.%Y"
    Dated= datetime.strptime(Date, Translate)
    return Dated

def TranslatorTDate(Date):
    TranslatorTDate = "%m/%d/%y"
    DateTranslator = datetime.strptime(Date, TranslatorTDate)
    return DateTranslator

def FilterByDate(Date1,Date2,Length):
    for i in Length:
        try:
            if Everything[i][5] is None:
                del Everything[i]
                continue
            if (TranslateToDate(Everything[i][5]) < Date1) | (TranslateToDate(Everything[i][5] > Date2)):
                del Everything[i]
                continue
        except KeyError:
            continue

def FilterByExploit():
    for i in Length:
        try:
            if 'Существует' in str(Everything[i][8]):
                continue
            elif 'уточняются' in str(Everything[i][8]):
                del Everything[i]
        except KeyError:
            continue

def FilterByStatus():
    for i in Length:
        try:
            if 'Подтверждена' in str(Everything[i][7]):
                continue
            else:
                del Everything[i]
        except KeyError:
            continue

def UyazCounter():
    CritCounter = 0
    HighCounter = 0
    MediumCounter = 0
    LowCounter = 0
    for i in Length:
        try:
            if 'Критический' in str(Everything[i][6]):
                CritCounter += 1
                continue
            elif 'Высокий' in str(Everything[i][6]):
                HighCounter += 1
                continue
            elif 'Средний' in str(Everything[i][6]):
                MediumCounter += 1
                continue
            elif 'Низкий' in str(Everything[i][6]):
                LowCounter += 1
                continue
            else:
                continue
        except KeyError:
            continue
    return [CritCounter, HighCounter, MediumCounter, LowCounter]

def diagram():
    index = ['Критический','Высокий','Средний','Низкий']
    values = UyazCounter()
    plt.bar(index, values)
    plt.savefig('graph.png', bbox_inches='tight')
    plt.show()