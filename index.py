import win32api
import win32print
import time


def printdoc(filename):
    # status 状态为1表示打印中 状态为0表示空闲
    print(win32print.GetDefaultPrinter())
    handle = win32print.OpenPrinter('HP LaserJet Professional P1106')
    status1 = 1024  # 初始化为1进入语句
    status2 = 2
    while (status1 == 1024) and (status2 >= 2):
        statusqueue = (win32print.GetPrinter(handle))
        print(statusqueue)
        status1 = statusqueue[-3]
        status2 = statusqueue[-2]
        print("打印" + filename + "等待中")
        time.sleep(1)
        continue
    print("开始打印" + filename)
    win32api.ShellExecute(0, 'print', filename, win32print.GetDefaultPrinterW(), ".", 0)
    time.sleep(2)


for i in range(1, 10):
    printdoc(str(i) + ".docx")
