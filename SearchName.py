
import win32gui
import win32com
import win32con
import pandas as pd
import os
import time
import xlwings as xw
from PIL import ImageGrab

hwnd_title = {}

def get_all_hwnd(hwnd, mouse):
    if (win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd)):
        hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})

def open_xlsx(dst_path,sname):
    file = os.path.basename(dst_path)
    df = pd.read_excel(dst_path)

    app = xw.App(visible=True, add_book=False)
    # 1. 使用 xlwings 的 读取 path 文件 启动
    wb = app.books.open(dst_path)
    sht = wb.sheets[0]

    i = df[df['姓名'].isin([sname])].index[0]
    sht.api.Rows("1:"+str(i+1)).EntireRow.Hidden = True  
    time.sleep(2)

    # 2.获取打开excel句柄并置到最前端截图保存
    win32gui.EnumWindows(get_all_hwnd, 0)

    for h,t in hwnd_title.items():
        if t:
            if t == file+' - Excel': 
                shell = win32com.client.Dispatch("WScript.Shell")
                shell.SendKeys('%')
                win32gui.SetForegroundWindow(h)
                win32gui.ShowWindow(h,win32con.SW_SHOWMAXIMIZED)
                time.sleep(1)
                left, top, right, bottom = win32gui.GetWindowRect(h)     #得到窗口矩形框
                rect = (left, top, right, bottom)
                img = ImageGrab.grab(rect)
                ImgName = os.path.splitext(dst_path)[0]
                ImgName = str(ImgName).replace(".","")
                img.save(ImgName+'.jpg')
    wb.close()  # 不保存，直接关闭
    app.quit()  # 退出
  
def find_text_xlsx():
    dst_path = input("输入目标地址文件夹")
    sname = input("输入查找的名字")
    ext_name = 'xlsx'
    for file in os.listdir(dst_path):
        if file.endswith(ext_name):
            path = dst_path+'\\'+file
            print(path)
            open_xlsx(path,sname)


find_text_xlsx()


