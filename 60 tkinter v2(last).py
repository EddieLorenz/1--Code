import xlrd
import clipboard
import time
import keyboard
import win32api
import win32com.client
import win32gui
import win32process
import win32con
import ctypes
from tkinter import filedialog
from tkinter import *


#функциия клика
def click(x, y):
    win32api.SetCursorPos((x, y))
    time.sleep(.3)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
    time.sleep(.2)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)


def closeLastWindow():
    time.sleep(timerb)
    click(298, 123)  # закрыть окно


def writeCompanyName():
    time.sleep(timer)
    keyboard.press_and_release('alt + shift')
    keyboard.press_and_release('right')
    time.sleep(timers)
    keyboard.write(companyName[currentNumber])  # пишем имя фирмы
    time.sleep(timers)
    keyboard.press_and_release('enter')
    time.sleep(timers)
    keyboard.press_and_release('alt + shift')



##Goods
def clickFirstJournal():
    time.sleep(timer)
    click(110, 694)  # Первый журнал, реализация
    time.sleep(timers)


def windowGoods():
    keyboard.press_and_release('ins')
    time.sleep(timers)
    click(619, 365)
    keyboard.press_and_release('enter')
    time.sleep(timer)


def dateForGoods():
    keyboard.press_and_release('tab')
    time.sleep(timers)
    clipboard.copy(dates[currentNumber])
    keyboard.write(dates[currentNumber])
    time.sleep(timers)
    keyboard.press_and_release('tab')
    keyboard.press_and_release('tab')
    time.sleep(timers)
    keyboard.write(dates[currentNumber])
    for i in range(2):
        keyboard.press_and_release('tab')
        time.sleep(.3)
    keyboard.press_and_release('shift + /')
    time.sleep(timers)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    keyboard.write(dates[currentNumber])
    for i in range(4):
        keyboard.press_and_release('tab')
        time.sleep(.3)


def dataEntryGoods():
    keyboard.press_and_release('ins')
    time.sleep(timers)
    keyboard.press_and_release('enter')
    time.sleep(timer)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    keyboard.write(amount[currentNumber])
    time.sleep(timers)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    keyboard.press_and_release('backspace')
    time.sleep(timers)
    keyboard.write(summ[currentNumber])
    keyboard.press_and_release('enter')
    time.sleep(timers)
    clipboard.copy(nds[currentNumber])
    keyboard.press_and_release('backspace')
    time.sleep(timers)
    keyboard.write(nds[currentNumber])
    keyboard.press_and_release('enter')
    time.sleep(timers)
    keyboard.press_and_release('ctrl + enter')
    time.sleep(timer)
    keyboard.press_and_release('enter')


##Services
def clickSecondJournal():
    time.sleep(timer)
    click(334, 694)  # второй журнал с услугами для 60
    time.sleep(timers)

def windowGeneralServices():
    keyboard.press_and_release('ins')
    time.sleep(timer)
    click(607, 414)
    keyboard.press_and_release('enter')
    time.sleep(timerb)


def dateForServices():
    keyboard.press_and_release('tab')
    keyboard.press_and_release('tab')
    time.sleep(timers)
    clipboard.copy(dates[currentNumber])
    keyboard.write(dates[currentNumber])
    time.sleep(timers)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    keyboard.write(dates[currentNumber])
    time.sleep(timers)
    click(436, 188)
    keyboard.press_and_release('shift + /')
    time.sleep(timers)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    keyboard.write(dates[currentNumber])
    time.sleep(timers)


def dataEntryServices():
    keyboard.press_and_release('ins')  #новая строка
    time.sleep(timers)
    keyboard.write('20')  #корр.счёт
    keyboard.press_and_release('enter')
    time.sleep(timer)
    keyboard.press_and_release('enter')
    time.sleep(timer)
    click(741, 378)
    time.sleep(timers)
    clipboard.copy(summ[currentNumber])  #сумма
    keyboard.write(summ[currentNumber])
    time.sleep(timers)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    clipboard.copy(nds[currentNumber])  #ндс
    keyboard.write(nds[currentNumber])
    time.sleep(timers)
    keyboard.press_and_release('ctrl + enter')
    time.sleep(timer)
    keyboard.press_and_release('enter')


##Materials
def clickThirdJournal():
    time.sleep(timer)
    click(554, 694)  # третий журнал с материалами для 60
    time.sleep(timers)


def windowMaterials():
    keyboard.press_and_release('ins')
    time.sleep(timers)
    click(645, 337)
    time.sleep(timers)
    keyboard.press_and_release('enter')
    time.sleep(timer)


def moveMaterials():
    time.sleep(timers)
    keyboard.press_and_release('ins')  # перемещение материалов
    time.sleep(timer)
    click(616, 363)
    keyboard.press_and_release('enter')
    time.sleep(timer)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    clipboard.copy(dates[currentNumber])
    keyboard.write(dates[currentNumber])
    time.sleep(timers)
    for i in range(3):
        keyboard.press_and_release('tab')
    time.sleep(timers)
    keyboard.write('20')
    time.sleep(timers)
    keyboard.press_and_release('ins')
    time.sleep(timers)
    keyboard.press_and_release('enter')
    time.sleep(timer)
    keyboard.press_and_release('enter')
    time.sleep(timers)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    clipboard.copy(amount[currentNumber])
    keyboard.write(amount[currentNumber])
    time.sleep(timers)
    keyboard.press_and_release('ctrl + enter')
    time.sleep(timer)
    keyboard.press_and_release('enter')


def dateForMaterials():
    keyboard.press_and_release('tab')
    time.sleep(timers)
    clipboard.copy(dates[currentNumber])
    keyboard.write(dates[currentNumber])
    time.sleep(timers)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    keyboard.write(dates[currentNumber])
    time.sleep(timers)
    click(400, 197)
    time.sleep(timers)
    keyboard.press_and_release('shift + /')
    time.sleep(timers)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    keyboard.write(dates[currentNumber])
    time.sleep(timers)


def dateEntryMaterials():
    keyboard.press_and_release('ins')  # Создаём новую строку
    time.sleep(timers)
    keyboard.press_and_release('enter')
    time.sleep(timer)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    clipboard.copy(amount[currentNumber])
    keyboard.write(amount[currentNumber])
    time.sleep(timers)
    keyboard.press_and_release('tab')
    time.sleep(timers)
    keyboard.press_and_release('tab')
    clipboard.copy(summ[currentNumber])
    keyboard.write(summ[currentNumber])
    time.sleep(timers)
    keyboard.press_and_release('tab')
    clipboard.copy(nds[currentNumber])
    keyboard.write(nds[currentNumber])
    time.sleep(timers)
    keyboard.press_and_release('ctrl + enter')
    time.sleep(timer)
    keyboard.press_and_release('enter')


def appends():
    dates.append(str(sheet.cell_value(currentNumber, 0)))
    typeOperation.append(str(sheet.cell_value(currentNumber, 7)))
    amount.append(str(sheet.cell_value(currentNumber, 4)))
    summ.append(str(sheet.cell_value(currentNumber, 5)))
    total.append(str(sheet.cell_value(currentNumber, 3)))
    companyName.append(str(sheet.cell_value(currentNumber, 8)))
    nds.append(str(sheet.cell_value(currentNumber, 6)))

# открываем Excel файл и разбиваем столбцы
root = Tk()
root.filename = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("all files","*.*"),("jpeg files","*.jpg")))
workbook = xlrd.open_workbook(root.filename)
sheet = workbook.sheet_by_index(0)

#задаём столбцы
dates = []  #0 столбец, дата
typeOperation = []  #2 столбец, тип операции
amount = []  #6 столбец, количество
summ = []  #8 столбец, сумма
total = []  #5 столбец, общая
companyName = []  #8 столбец, имя фирмы
nds = []  #7 столбец, НДС
currentNumber = 0
timers = 0.5  #Задержка в 0.5секунд
timer = 1.5  #Выбираем задержку
timerb = 3.5

#задаём откуда искать услуги и материалы
services = 'усл'
materials = 'мат'

#Разрешение 1366*768
time.sleep(5)

u = ctypes.windll.LoadLibrary("user32.dll")
pf = getattr(u, "GetKeyboardLayout")

if hex(pf(0)) == '0x4090409':
    for i in range(sheet.nrows):
        appends()
        if services in typeOperation[currentNumber].split(' '):
            clickSecondJournal()
            windowGeneralServices()
            dateForServices()
            click(286, 216)  # исполнитель услуга 60
            writeCompanyName()
            dataEntryServices()
            closeLastWindow()
            currentNumber = currentNumber + 1
        elif materials in typeOperation[currentNumber].split(' '):
            clickThirdJournal()
            windowMaterials()
            dateForMaterials()
            click(270, 245)  # клик по исполнителю в материалах
            writeCompanyName()
            dateEntryMaterials()
            closeLastWindow()
            clickThirdJournal()
            moveMaterials()
            closeLastWindow()
            currentNumber = currentNumber + 1
        else:
            clickFirstJournal()
            windowGoods()
            dateForGoods()
            writeCompanyName()
            dataEntryGoods()
            closeLastWindow()
            currentNumber = currentNumber + 1
else:
    print("Переключите на английский язык")
print('Завершено! ')
