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

Date1 = 0
Date2 = 0
Everything = {}
Vulner = []
Description = []
Soft = []
SoftType = []
VulnerClass = []
DateFound = []
VulnerLVL = []
Status = []
Exploit = []
ProtInfo = []
CWEType = []
Measures = []
ID = []
Length = 0

def initial_file():
    Counter = 0
    for row in range(1, column.max_row):
        if 'Thunderbird' in str(column.cell(row=row, column=5).value):
            Vulner.append(column.cell(row=row, column=2).value)
            Description.append(column.cell(row=row, column=3).value)
            Soft.append(column.cell(row=row, column=5).value)
            SoftType.append(column.cell(row=row, column=7).value)
            VulnerClass.append(column.cell(row=row, column=9).value)
            DateFound.append(column.cell(row=row, column=10).value)
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
            Vulner[i], Description[i], Soft[i], SoftType[i], VulnerClass[i], DateFound[i],
            VulnerLVL[i],
            Status[i], Exploit[i], ProtInfo[i], CWEType[i], Measures[i], ID[i])


def CountCheker(Count):
    Counter = 0
    for t in Count:
        Counter += 1
    return Counter


def TranslateToDate(Date):
    Translate = "%d.%m.%Y"
    Dated = datetime.strptime(Date, Translate)
    return Dated
'''
def TranslatorTDate(Date):
    TranslatorTDate = "%d/%m/%Y"
    DateTranslator = datetime.strptime(Date, TranslatorTDate)
    return DateTranslator
    '''
def FilterByClass(Length, Class):
    for i in range(Length):
        if (Class != 1) & ('Уязвимость архитектуры' in Everything[i][5]):
            del Everything[i]
            continue
        elif (Class != 2) & ('Уязвимость кода' in Everything[i][5]):
            del Everything[i]
            continue
        elif (Class != 3) & ('Уязвимость многофакторная' in Everything[i][5]):
            del Everything[i]
            continue


def FilterByStatus(Length):
    for i in range(Length):
        try:
            if 'Подтверждена' in str(Everything[i][7]):
                continue
            elif 'Подтверждена' not in str(Everything[i][7]):
                del Everything[i]
                continue
        except KeyError:
            continue


def FilterByUSTRINFO():
    for i in Everything:
        try:
            if 'Переменная' in str(Everything[i][9]):
                continue
            else:
                del Everything[i]
                continue
        except KeyError:
            continue

def FilterByDate(Date1, Date2, Length):
    for i in range(Length):
        try:
            if Everything[i][5] is None:
                del Everything[i]
                continue
            if (TranslateToDate(Everything[i][5]) < TranslateToDate(Date1)) | (
                    TranslateToDate(Everything[i][5]) > TranslateToDate(Date2)):
                del Everything[i]
                continue
        except KeyError:
            continue


def FilterByExploit(Length):
    for i in range(Length):
        try:
            if 'Существует' in str(Everything[i][8]):
                continue
            elif 'уточняются' in str(Everything[i][8]):
                del Everything[i]
        except KeyError:
            continue

def FilterByDangerLVL():
    for i in Everything:
        try:
            if 'Переменная' in str(Everything[i][6]):
                continue
            else:
                del Everything[i]
                continue
        except KeyError:
            continue

def UyazCounter():
    CritCounter = 0
    HighCounter = 0
    MediumCounter = 0
    LowCounter = 0
    for i in Everything:
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
        except KeyError:
            continue
    return [CritCounter, HighCounter, MediumCounter, LowCounter]

def diagram():
    index = ['Критический', 'Высокий', 'Средний', 'Низкий']
    values = UyazCounter()
    print(values)
    plt.bar(index, values)
    plt.savefig('graph.png', bbox_inches='tight')
    plt.show()


initial_file()
Length = CountCheker(Everything)
FilterByDate(str('12.03.2016'), str('21.05.2021'), Length)
diagram()



