#!/usr/bin/env python
# -*- coding: utf-8 -*-

# РїРѕРґРєР»СЋС‡Р°РµРј РЅСѓР¶РЅС‹Рµ Р±РёР±Р»РёРѕС‚РµРєРё
import win32com.client as com_client # РґР»СЏ Р·Р°РїРёСЃРё РІ excel С„Р°Р№Р»
import xlrd # РґР»СЏ С‡С‚РµРЅРёСЏ РёР· excel С„Р°Р№Р»Р°
import shutil # РґР»СЏ СЃРѕР·РґР°РЅРёСЏ excel С„Р°Р№Р»Р° СЃ РѕРїРѕР·РґР°РІС€РёРјРё
import datetime # РґР»СЏ СЂР°Р±РѕС‚С‹ СЃ С‚РµРєСѓС‰РµР№ РґР°С‚РѕР№

#СЃРѕР·РґР°РµРј РЅСѓР¶РЅС‹Рµ РїРµСЂРµРјРµРЅРЅС‹Рµ
################
a = []
headline1 = ""
headline2 = ""
headline3 = []
toname = {1 : 'янв.', 2 : 'фев.', 3 : 'март', 4 : 'апр.', 5 : 'май.', 6 : 'июнь', 7 : 'июль', 8 : 'авг.', 9 : 'сен.', 10 : 'окт.', 11 : 'ноя.', 12 : 'дек.'}
names = ['Омаров', 'Зуев', 'Свинолупов', 'Елюбаев', 'Салаватов', 'Нурмухамбетов', 'Иманмәлік']
###################

#СЃРѕР·РґР°РЅРёРµ РєР»Р°СЃСЃР° РґР»СЏ РєР°Р¶РґРѕРіРѕ СѓС‡РµРЅРёРєР°
############
class id:
    def __init__(self, ind, name, clas, hour, minute):
        self.ind = ind
        self.name = name
        self.clas = clas
        self.hour = hour
        self.minute = minute
        self.mt = self.minute + self.hour * 60
        if (len(self.clas) == 2):
            self.clasnum = int(self.clas[0])
        elif (len(self.clas) == 3):
            self.clasnum = int(self.clas[0] + self.clas[1])
##########



#С‡С‚РµРЅРёРµ РґР°РЅРЅС‹С… РёР· excel С„Р°Р№Р»Р°
#################################
def ReadData(Path):
    global headline1, headline2, headline3
    wb = xlrd.open_workbook(r'C:\Users\ilyas\OneDrive\Desktop\Tourniquet 2019' + Path, on_demand = True)    
    ##wb = xlrd.open_workbook(r'C:\Users\FizMat\Desktop\projecy' + Path, on_demand = True)
    ws = wb.sheet_by_name("Лист1")
    headline1 = ws.cell(1, 1).value
    headline2 = ws.cell(2, 1).value
    for i in range(8):
        headline3.append(ws.cell(3, i).value)
    stnum = 5
    while (True):
        ind = ws.cell(stnum, 0).value
        name = ws.cell(stnum, 2).value
        clas = ws.cell(stnum, 5).value
        b = True
        DateEnter = xlrd.xldate_as_tuple(ws.cell(stnum, 4).value, wb.datemode)
        if (DateEnter[3] >= 9):
            break
        if (DateEnter[3] > 7 or (DateEnter[3] == 7 and DateEnter[4] > 45)):
            if len(name) > 0:
                for j in names:
                    if j in name:
                        b = False
                if b:      
                    a.append(id(ind, name, clas, DateEnter[3], DateEnter[4]))
        stnum += 1
#############################


#РєРѕРЅРІРµСЂС‚РёСЂРѕРІР°РЅРёРµ РІСЂРµРјРµРЅРё РІ СЃС‚СЂРѕРєРѕРІС‹Р№ С„РѕСЂРјР°С‚ РїРѕ С‚РµРєСѓС‰РёРј С‡Р°СЃСѓ Рё РјРёРЅСѓС‚РѕР№
#############################
def convert(x, y):
    x = str(x)
    y = str(y)
    if (len(y) == 1):
        return x + ":0" + y
    return x + ":" + y
##################


#Р·Р°РїРёСЃСЊ РѕРїРѕР·РґР°РІС€РёС… РІ РѕС‚РґРµР»СЊРЅС‹Р№ excel С„Р°Р№Р»
##############################
def WriteToCurrentSheet(sheetname, a, wb):
    global headline1, headline2, headline3
    ws = wb.Worksheets(sheetname)
    ws.Cells(1, 2).Value = headline1
    ws.Cells(2, 2).Value = headline2
    for i in range(8):
        ws.Cells(4, i + 1).Value = headline3[i]
    a.sort(key=lambda x: x.clas)
    f = open("Late List.txt", "w", encoding="utf-8")
    b = a[:]
    c = []
    ind = 0
    for i in range(len(b)):
        if i + 1 == len(b) or b[i].clas != b[i + 1].clas:
            f.write(str(b[i].clas) + ' - ')
            for j in range(ind, i + 1):
                if b[i].name == '' or (i > 0 and b[i].name == b[i - 1].name):
                    continue
                if(j != ind):
                    f.write(', ' + b[i].name)
                else:
                    f.write(b[i].name)
            ind = i + 1
            f.write('\n')
    f.close()
    for i in range(len(a)):
        ws.Cells(i + 5, 1).Value = a[i].ind
        ws.Cells(i + 5, 2).Value = a[i].name
        ws.Cells(i + 5, 3).Value = a[i].clas
        ws.Cells(i + 5, 5).Value = convert(a[i].hour, a[i].minute)
    num = len(a) + 5
    while (ws.Cells(num, 1).Value):
        for i in range(1, 9):
            ws.Cells(num, i).Value = ""
        num += 1
######################


#РґР»СЏ РєР°Р¶РґРѕРіРѕ РєР»Р°СЃСЃР° Р·Р°РїРёСЃСЊ РІ РѕС‚РґРµР»СЊРЅС‹Р№ Р»РёСЃС‚
##############################
def WriteData(LateName):
    global a
    excel = com_client.Dispatch('Excel.Application')
    excel.visible = False
    print('Writing data...')
    wb = excel.Workbooks.Open(r'C:\Users\ilyas\OneDrive\Desktop\Tourniquet 2019' + LateName)    
    ##wb = excel.Workbooks.Open(r'C:\Users\FizMat\Desktop\projecy' + LateName)
    WriteToCurrentSheet("123", a, wb)
    wb.Save()
    excel.Application.Quit()
    print("Done!")
    print("programm maded by:")
    for i in names:
        print("           " + i)
######################################


#СЃРѕР·РґР°РЅРёРµ С„Р°Р№Р»Р° СЃ РѕРїРѕР·РґР°РІС€РёРјРё РїРѕ Р·Р°РґР°РЅРЅРѕРјСѓ РѕР±СЂР°Р·С†Сѓ
######################
def Create(Latename):
    shutil.copy(r'C:\Users\ilyas\OneDrive\Desktop\Tourniquet 2019\Образец.xlsx', r'C:\Users\ilyas\OneDrive\Desktop\Tourniquet 2019\tmp')
    shutil.move(r'C:\Users\ilyas\OneDrive\Desktop\Tourniquet 2019\tmp\Образец.xlsx', r'C:\Users\ilyas\OneDrive\Desktop\Tourniquet 2019' + Latename)
####################


#РѕРїСЂРµРґРµР»РµРЅРёРµ РёРјРµРЅРё РёСЃС…РѕРґРЅРѕРіРѕ С„Р°Р№Р»Р° РІ Р·Р°РІРёСЃРёРјРѕСЃС‚Рё РѕС‚ С‚РµРєСѓС‰РµР№ РґР°С‚С‹
########################
def MakePath():
    now = datetime.date.today()
    st = '\Время прихода - время ухода за '
    st += str(now.day) + ' ' + toname[now.month] + ' ' + str(now.year) + '.xlsx'
    return st
#########################


#РѕРїСЂРµРґРµР»РµРЅРёРµ РёРјРµРЅРё С„Р°Р№Р»Р° СЃ РѕРїРѕР·РґР°РІС€РёРјРё РІ Р·Р°РІРёСЃРёРјРѕСЃС‚Рё РѕС‚ С‚РµРєСѓС‰РµР№ РґР°С‚С‹
############################
def MakeLateName():
    now = datetime.date.today()
    st = '\Опоздавшие за '
    st += str(now.day) + ' ' + toname[now.month] + ' ' + str(now.year) + '.xlsx'
    return st
########################


# РіР»Р°РІРЅР°СЏ С„СѓРЅРєС†РёСЏ РєРѕС‚РѕСЂР°СЏ РІС‹Р·РІР°РµС‚ РґСЂСѓРіРёРµ
######################
def main(Path, LateName):
    Create(LateName) # СЃРѕР·РґР°РЅРёРµ С„Р°Р№Р»Р° СЃ РѕРїРѕР·РґР°РІС€РёРјРё
    ReadData(Path) # С‡С‚РµРЅРёРµ РґР°РЅРЅС‹С… РёР· С„Р°Р№Р»Р° Рё Р·Р°РїРёСЃСЊ РІ РѕС‚РґРµР»СЊРЅС‹Р№ СЃРїРёСЃРѕРє
    WriteData(LateName) # Р·Р°РїРёСЃСЊ РѕРїРѕР·РґР°РІС€РёС… СѓС‡РµРЅРёРєРѕРІ РІ РѕС‚РґРµР»СЊРЅС‹Р№ excel С„Р°Р№Р»
#####################


main(MakePath(), MakeLateName()) # РІС‹Р·РѕРІ РіР»Р°РІРЅРѕР№ С„СѓРЅРєС†РёРё