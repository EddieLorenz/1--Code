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
def click(x,y):
    win32api.SetCursorPos((x,y))
    time.sleep(.3)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    time.sleep(.2)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)



#открываем Excel файл и разбиваем столбцы
root = Tk()
root.filename = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("all files","*.*"),("jpeg files","*.jpg")))
workbook = xlrd.open_workbook(root.filename)
sheet = workbook.sheet_by_index(0)

#задаём столбцы
dates = [] # 0 столбец, дата
typeOperation = [] #2 столбец, тип операции
amount = [] #6 столбец, количество
summ = [] # 8 столбец, сумма
total = [] # 5 столбец, общая
companyName = [] # 8 столбец, имя фирмы
nds = [] # 7 столбец, НДС
currentNumber = 0

timers = 0.5 # Задержка в 0.5секунд
timer = 1.5 #Выбираем задержку
timerb = 3.5

#задаём откуда искать услуги и материалы
services = 'усл'
print("Проверить номера стройматериалов и услуг")

time.sleep(5)

u = ctypes.windll.LoadLibrary("user32.dll")
pf = getattr(u, "GetKeyboardLayout")

if hex(pf(0)) == '0x4090409':
    for i in range(sheet.nrows):
        dates.append(str(sheet.cell_value(currentNumber, 0)))
        typeOperation.append(str(sheet.cell_value(currentNumber, 7)))
        amount.append(str(sheet.cell_value(currentNumber, 4)))
        summ.append(str(sheet.cell_value(currentNumber, 5)))
        total.append(str((sheet.cell_value(currentNumber, 3))))
        companyName.append(str(sheet.cell_value(currentNumber, 8)))
        nds.append(str(sheet.cell_value(currentNumber, 6))) 
        if services in typeOperation[currentNumber].split(' '):
            click(110, 694) #Первый журнал, реализация
            time.sleep(timers)
            keyboard.press_and_release('ins')
            time.sleep(timers)
            click(603, 325) # Клик по отгрузке
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timer)
            keyboard.press_and_release('tab')
            keyboard.write(dates[currentNumber])
            time.sleep(timers)
            keyboard.press_and_release('tab')
            keyboard.press_and_release('tab')#Открывается окно с исполнителем
            time.sleep(timer)
            keyboard.press_and_release('alt + shift')
            time.sleep(timers)
            keyboard.press_and_release('right')
            keyboard.write(companyName[currentNumber])
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timers)
            keyboard.press_and_release('alt + shift')
            time.sleep(timers)
            keyboard.press_and_release('ins')
            time.sleep(timer)
            keyboard.write('000000002')
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timers)
            keyboard.press_and_release('tab')
            keyboard.write('1')
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timers)
            keyboard.write(summ[currentNumber])
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timers)
            keyboard.press_and_release('tab')
            time.sleep(timers)
            keyboard.write(nds[currentNumber])
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timers)
            click(381, 664) # счёт-фактура
            time.sleep(timers)
            click(407, 636)
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timer)
            click(615, 278)
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timer)
            keyboard.press_and_release('tab')
            time.sleep(timers)
            keyboard.write(dates[currentNumber])
            keyboard.press_and_release('enter')
            time.sleep(timers)
            keyboard.press_and_release('ctrl + enter')
            time.sleep(timer)
            keyboard.press_and_release('enter')
            time.sleep(timerb)
            click(292, 123)
            time.sleep(timer)
            click(1000, 693)
            time.sleep(timers)
            keyboard.press_and_release('ctrl + enter')
            time.sleep(timer)
            keyboard.press_and_release('enter')
            time.sleep(timerb)
            click(293, 121)
            time.sleep(timer)
            currentNumber = currentNumber + 1
        else:
            click(110, 694) #Первый журнал, реализация
            time.sleep(timers)
            keyboard.press_and_release('ins')
            time.sleep(timers)
            click(618, 389)
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timer)
            keyboard.press_and_release('tab')
            time.sleep(timers)
            keyboard.write(dates[currentNumber])
            time.sleep(timers)
            keyboard.press_and_release('tab')
            keyboard.press_and_release('enter')
            keyboard.press_and_release('tab')
            time.sleep(timer)
            keyboard.press_and_release('alt + shift')
            keyboard.press_and_release('right')
            time.sleep(timers)
            keyboard.write(companyName[currentNumber])
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timer)
            keyboard.press_and_release('alt + shift')
            time.sleep(timers)
            keyboard.press_and_release('ins')
            time.sleep(timer)
            keyboard.write('000000001')
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timers)
            keyboard.press_and_release('tab')
            time.sleep(timers)
            keyboard.write(amount[currentNumber])
            time.sleep(timers)
            keyboard.press_and_release('tab')
            time.sleep(timers)
            keyboard.press_and_release('tab')
            time.sleep(timers)
            keyboard.write(summ[currentNumber])
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timers)
            keyboard.write(nds[currentNumber])
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timers)
            click(380, 667)
            time.sleep(timers)
            click(407, 636)
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timer)
            click(607, 305)
            time.sleep(timers)
            keyboard.press_and_release('enter')
            time.sleep(timer)
            keyboard.press_and_release('tab')
            time.sleep(timers)
            keyboard.write(dates[currentNumber])
            keyboard.press_and_release('enter')
            time.sleep(timers)
            keyboard.press_and_release('ctrl + enter')
            time.sleep(timer)
            keyboard.press_and_release('enter')
            time.sleep(timerb)
            click(292, 124)
            time.sleep(timer)
            click(1000, 695)
            time.sleep(timers)
            keyboard.press_and_release('ctrl + enter')
            time.sleep(timer)
            keyboard.press_and_release('enter')
            time.sleep(timerb)
            click(293, 121)
            time.sleep(timer)
            currentNumber = currentNumber + 1
else:
    print("Переключите на английский язык")
print('Завершено! ')
