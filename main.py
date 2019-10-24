#!/usr/bin/env python
# -*- coding: utf-8 -*-

# подключаем нужные библиотеки
import win32com.client as com_client # для записи в excel файл
import xlrd # для чтения из excel файла
import shutil # для создания excel файла с опоздавшими
import datetime # для работы с текущей датой

#создаем нужные переменные

################
a = []
headline1 = ""
headline2 = ""
headline3 = []
toname = {1 : 'янв.', 2 : 'фев.', 3 : 'мар.', 4 : 'апр.', 
          5 : 'май.', 6 : 'июн.', 7 : 'июл.', 8 : 'авг.', 
          9 : 'сен.', 10 : 'окт.', 11 : 'ноя.', 12 : 'дек.'}
###################

#создание класса для каждого ученика

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

#чтение данных из excel файла

#################################
def ReadData(Path):
    global headline1, headline2, headline3
    wb = xlrd.open_workbook(r'C:\Users\ilyas\OneDrive\Desktop\Tourniquet 2019' + Path, on_demand = True)    
    ##wb = xlrd.open_workbook(r'C:\Users\grimson\Desktop\Tourniquet 2019' + Path, on_demand = True)    
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
        DateEnter = xlrd.xldate_as_tuple(ws.cell(stnum, 4).value, wb.datemode)
        if (DateEnter[3] >= 9):
            break
        if (DateEnter[3] > 7 or (DateEnter[3] == 7 and DateEnter[4] > 45)):
            a.append(id(ind, name, clas, DateEnter[3], DateEnter[4]))
        stnum += 1
#############################
        
#конвертирование времени в строковый формат по текущим часу и минутой

#############################
def convert(x, y):
    x = str(x)
    y = str(y)
    if (len(y) == 1):
        return x + ":0" + y
    return x + ":" + y
##################

#запись опоздавших в отдельный excel файл

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
        
#для каждого класса запись в отдельный лист

##############################
def WriteData(LateName):
    global a
    excel = com_client.Dispatch('Excel.Application')
    excel.visible = False
    print('Writing data...')
    wb = excel.Workbooks.Open(r'C:\Users\ilyas\OneDrive\Desktop\Tourniquet 2019' + LateName)    
    ##wb = excel.Workbooks.Open(r'C:\Users\grimson\Desktop\Tourniquet 2019' + LateName)    
    ##wb = excel.Workbooks.Open(r'C:\Users\FizMat\Desktop\projecy' + LateName)
    WriteToCurrentSheet("123", a, wb)
    wb.Save()
    excel.Application.Quit()
    print("Done!")
######################################

#создание файла с опоздавшими по заданному образцу

#C:\Users\ilyas\OneDrive\Desktop\Tourniquet 2019
######################
def Create(Latename):
    shutil.copy(r'C:\Users\ilyas\OneDrive\Desktop\Tourniquet 2019\Образец.xlsx', r'C:\Users\ilyas\OneDrive\Desktop\Tourniquet 2019\tmp')
    shutil.move(r'C:\Users\ilyas\OneDrive\Desktop\Tourniquet 2019\tmp\Образец.xlsx', r'C:\Users\ilyas\OneDrive\Desktop\Tourniquet 2019' + Latename)    
    ##shutil.copy(r'C:\Users\grimson\Desktop\Tourniquet 2019\Образец.xlsx', r'C:\Users\grimson\Desktop\Tourniquet 2019\tmp')
    ##shutil.move(r'C:\Users\grimson\Desktop\Tourniquet 2019\tmp\Образец.xlsx', r'C:\Users\grimson\Desktop\Tourniquet 2019' + Latename)    
####################

#определение имени исходного файла в зависимости от текущей даты

########################
def MakePath():
    now = datetime.date.today()
    st = '\Время прихода - время ухода за '
    st += str(now.day) + ' ' + toname[now.month] + ' ' + str(now.year) + '.xlsx'
    return st
#########################

#определение имени файла с опоздавшими в зависимости от текущей даты

############################
def MakeLateName():
    now = datetime.date.today()
    st = '\Опоздавшие за '
    st += str(now.day) + ' ' + toname[now.month] + ' ' + str(now.year) + '.xlsx'
    return st
########################

# главная функция которая вызвает другие

######################
def main(Path, LateName):
    Create(LateName) # создание файла с опоздавшими
    ReadData(Path) # чтение данных из файла и запись в отдельный список
    WriteData(LateName) # запись опоздавших учеников в отдельный excel файл
#####################


main(MakePath(), MakeLateName()) # вызов главной функции
