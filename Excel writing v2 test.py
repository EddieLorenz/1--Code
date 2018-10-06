import xlwt
import xlrd
from tkinter import filedialog
from tkinter import *


#В скобках номер для питона, без скобок - Excel
dates = [] # 1 (0) столбец
statement = [] # 2 (1) столбец
content = [] # 3 (2) столбец
total = [] # 6 (5) столбец
amount = [] # 7 (6) столбец
nds = [] # 8 (7) столбец
summ = [] # 9 (8) столбец
operationType = [] #10(9) столбец
companyName = [] #11(10) столбец


root = Tk()
root.filename = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("all files","*.*"),("jpeg files","*.jpg")))
workbook = xlrd.open_workbook(root.filename)
sheet = workbook.sheet_by_index(0)

for sheet in workbook.sheets():
    for rowidx in range(sheet.nrows):
        row = sheet.row(rowidx)
        for colidx, cell in enumerate(row):
            if cell.value == "Содержание" :
                startNumber = rowidx + 1
                print('Начало таблицы: ' + str(startNumber))
                break

for sheet in workbook.sheets():
    for rowidx in range(sheet.nrows):
        row = sheet.row(rowidx)
        for colidx, cell in enumerate(row):
            if cell.value == "Итого" :
                finishNumber = rowidx - 1
                print('Конец таблицы: ' + str(finishNumber))
                break


number_of_operations = finishNumber - startNumber + 1
typeInput = input("Тип операции 60 или 62? ")
if typeInput == '60':
    numberSplit = 1
else:
    numberSplit = 3
    

def readbook():
    currentNumber = startNumber
    for i in range(number_of_operations):
        takeNumber = currentNumber - startNumber
        dates.append(str(sheet.cell_value(currentNumber, 0))) # 1 (0) столбец
        statement.append(str(sheet.cell_value(currentNumber, 1))) # 2 (1) столбец
        content.append(str(sheet.cell_value(currentNumber, 2))) # 3 (2) столбец
        total.append(round(sheet.cell_value(currentNumber, 5), 2)) # 6 (5) столбец
        amount.append(str(round(total[takeNumber] / 18.10))) # 7 (6) столбец
        summ.append(round(sheet.cell_value(currentNumber, 8) ,2)) # 9 (8) столбец
        nds.append(round(total[takeNumber]-summ[takeNumber], 2)) # 8 (7) столбец
        operationType.append(sheet.cell_value(currentNumber, 9))
        companyName.append(sheet.cell_value(currentNumber, 2).split('\n')[int(numberSplit)])
        currentNumber += 1


def writebook():
    wb = xlwt.Workbook()
    ws = wb.add_sheet("A test Sheet", cell_overwrite_ok = True)
    startNW = 0 #startNumberWrite
    currentNW = startNW #currentNumberWrite
    for i in range(number_of_operations):
        ws.write(currentNW, 0, dates[currentNW])
        ws.write(currentNW, 1, statement[currentNW])
        ws.write(currentNW, 2, content[currentNW])
        ws.write(currentNW, 3, total[currentNW])
        ws.write(currentNW, 4, amount[currentNW])
        ws.write(currentNW, 5, summ[currentNW])
        ws.write(currentNW, 6, nds[currentNW])
        ws.write(currentNW, 7, operationType[currentNW])
        ws.write(currentNW, 8, companyName[currentNW])
        currentNW += 1
    root = Tk()
    root.filename =  filedialog.asksaveasfilename(initialdir = "/",title = "Select file",filetypes = (("xls Excel","*.xls"),("all files","*.*")))
    wb.save(root.filename)

readbook()
writebook()

print("Work is done")
