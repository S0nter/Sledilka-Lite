import time
import csv
import os
import datetime
import winshell
from win32com.client import Dispatch
from tkinter import messagebox


sid = 0
data = []
startover = 0
start = datetime.date.today()
delta = datetime.date.today()


def dataload():
    print('dataload')
    global sid, data, startover, start
    data = []
    with open('dataL.csv') as f:
        reader = csv.reader(f)
        for row in reader:
            sid = int(row[0])
            startover = int(row[1])
            start = datetime.datetime.strptime(row[2], '%Y-%m-%d').date()
            data.append(row)
    print('data', data)


def datasave():
    print('datasave')
    global sid, startover, start, data
    if len(data) > 0 and delta.days > 0:    # Если заход в прогу был не сегодня:
        print('if len(data) > 0 and delta.days > 0:')
        sid = 0
        data.append([sid, startover, str(start)])
    else:
        print('else:        data[-1] = [sid, startover, str(start)]')
        data[-1] = [sid, startover, str(start)]
    with open('dataL.csv', 'w', newline='') as file:
        writer = csv.writer(file, delimiter=",")
        if startover == 1:
            if delta.days > 0 and not len(data) > 1:
                '''Если сегодня старта проги не было и кол-во строк в документе не превышает одну:'''
                print('if delta.days > 0 and not len(data) > 1:')
                writer.writerow(data[-1])
                writer.writerow([sid, startover, start])
                data.append([sid, startover, str(start)])
            elif delta.days > 0 and len(data) > 1:
                '''Если сегодня старта проги не было и кол-во строк в документе превышает одну:'''
                print('elif delta.days > 0 and len(data) > 1:')
                writer.writerows(data[:-1])
                writer.writerow([sid, startover, start])
                data[-1] = [sid, startover, str(start)]
            elif not (delta.days > 0) and not len(data) > 1:
                '''Если сегодня старт проги был и кол-во строк в документе превышает одну:'''
                print('elif not (delta.days > 0) and not len(data) > 1:')
                writer.writerow(data[-1])
                data[-1] = [sid, startover, str(start)]
            elif not (delta.days > 0) and len(data) > 1:
                '''Если сегодня старт проги был и кол-во строк в документе превышает одну:'''
                print('elif not (delta.days > 0) and len(data) > 1:')
                writer.writerows(data[:-1])
                writer.writerow([sid, startover, start])
                data[-1] = [sid, startover, str(start)]
        elif startover == 0:
            print('elif startover == 0:')
            make_file()
            writer.writerow([sid, startover, start])
        print(data)


def make_file():
    print('makefile')
    with open('dataL.csv', 'w', newline='') as file:
        writer = csv.writer(file, delimiter=",")
        writer.writerow([sid, startover, str(start)])


def remake_file():
    with open('dataL(restored).csv', 'w', newline='') as file:
        writer = csv.writer(file, delimiter=",")
        writer.writerow([sid, startover, str(start)])


def make_shortcut(name, target, path_to_save, w_dir='default', icon='default'):
    print('make_shortcut')
    if path_to_save == 'desktop':
        '''Saving on desktop'''
        # Соединяем пути, с учётом разных операционок.
        path = os.path.join(winshell.desktop(), str(name) + '.lnk')
    elif path_to_save == 'startup':
        '''Adding to startup (windows)'''
        user = os.path.expanduser('~')
        path_to_save = os.path.join(r"%s/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Startup/" % user,
                                    str(name) + '.lnk')
    else:
        path_to_save = os.path.join(path_to_save, str(name) + '.lnk')
    if icon == 'default':
        icon = target
    if w_dir == 'default':
        w_dir = os.path.dirname(target)
    # С помощью метода Dispatch, обьявляем работу с Wscript
    # (работа с ярлыками, реестром и прочей системной информацией в windows)
    shell = Dispatch('WScript.Shell')
    # Создаём ярлык.
    shortcut = shell.CreateShortCut(path_to_save)
    # Путь к файлу, к которому делаем ярлык.
    shortcut.Targetpath = target
    # Путь к рабочей папке.
    shortcut.WorkingDirectory = w_dir
    # Тырим иконку.
    shortcut.IconLocation = icon
    # Обязательное действо, сохраняем ярлык.
    shortcut.save()


try:
    print('try dataload')
    dataload()
except FileNotFoundError:
    print('except FileNotFoundError:')
    make_file()
except ValueError:
    print('ValueError')
    messagebox.showwarning('Внимание!', 'Не трогать файлы!')
    make_file()
except IndexError:
    print('IndexError')
    make_file()
dataload()

if startover == 0:
    start = datetime.date.today()
    now = datetime.date.today()
    delta = now - start
    datasave()
    startover = 1
elif startover == 1:
    now = datetime.date.today()
    delta = now - start
    if delta.days > 0:
        messagebox.showinfo("Информация", f"Время проведённое за компьютером в прошлый раз (минуты): {sid}")
        start = datetime.date.today()
        datasave()
        now = datetime.date.today()
        delta = now - start
make_shortcut(name='Sledilka Lite',
              target=os.path.abspath('Sledilka Lite.exe'),
              path_to_save='startup')
while True:
    for _ in range(30):
        time.sleep(60)
        sid += 1
        datasave()
        print('end')
    os.popen('shutdown -t 10 -s -c "Требуется отдых от монитора"')
