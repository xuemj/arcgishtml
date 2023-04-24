# coding=utf-8
import Tkinter as tk
import os
from tkFileDialog import askopenfilename



file_dir1 = "C:/Users/Administrator/Desktop/123/222"
file_dir2 = "C:/Users/Administrator/Desktop/123/44451211"
lists1 = []
lists2 = []

listspath=[]
def selectPath():
    path = askopenfilename()  #使用askdirectory()方法返回文件夹的路径
    listspath.append(path)
    lb1.config(text=os.path.basename(path))
    return listspath

def getfilelist(filepath):
    filelist = os.listdir(filepath)
    files = []
    for i in range(len(filelist)):
        child = os.path.join('%s/%s' % (filepath, filelist[i]))
        if os.path.isdir(child):
            files.extend(getfilelist(child))
        else:
            files.append(child)
    return files


def appedlists(f):
    lists = []
    for filepath in getfilelist(f):
        if filepath.endswith("shp") and not filepath.startswith('~$'):
            lists.append(filepath)
    return lists


lists1 = appedlists(file_dir1)
lists2 = appedlists(file_dir2)
top = tk.Tk()
top.title("shape相交分析工具")  # 设置窗口
top.maxsize(720, 950)
top.minsize(720, 950)


F = tk.Frame(top, bg='white', width=150, height=900,
             highlightbackground='black',
             highlightcolor='green',
             highlightthickness=1, takefocus=1
             )
F.place(x=70, y=50)
tk.Button(top, text='选取范围：',command=selectPath).place(x=10, y=10)
lb1 = tk.Label(top, text='').place(x=90, y=13)

tk.Label(top, text='基础资料:', bd=6).place(x=0, y=50)

checkvar1 = tk.IntVar()
checkvar2 = tk.IntVar()
checkvar3 = tk.IntVar()
checkvar4 = tk.IntVar()
checkvar5 = tk.IntVar()
checkvar6 = tk.IntVar()
checkvar7 = tk.IntVar()
checkvar8 = tk.IntVar()
checkvar9 = tk.IntVar()
checkvar10 = tk.IntVar()
checkvar11 = tk.IntVar()
checkvar12 = tk.IntVar()
checkvar13 = tk.IntVar()
checkvar14 = tk.IntVar()
checkvar15 = tk.IntVar()
checkvar16 = tk.IntVar()
checkvar17 = tk.IntVar()
checkvar18 = tk.IntVar()
checkvar19 = tk.IntVar()
checkvar20 = tk.IntVar()
checkvar21 = tk.IntVar()
checkvar22 = tk.IntVar()
checkvar23 = tk.IntVar()
checkvar24 = tk.IntVar()
checkvar25 = tk.IntVar()
checkvar26 = tk.IntVar()
checkvar27 = tk.IntVar()
checkvar28 = tk.IntVar()
checkvar29 = tk.IntVar()
checkvar30 = tk.IntVar()
checkvar_arr = [checkvar1, checkvar2, checkvar3, checkvar4, checkvar5, checkvar6, checkvar7, checkvar8,
                checkvar9, checkvar10, checkvar11, checkvar12, checkvar13, checkvar14, checkvar15, checkvar16,
                checkvar17, checkvar18, checkvar19, checkvar20, checkvar21, checkvar22, checkvar23, checkvar24,
                checkvar25, checkvar26, checkvar27, checkvar28, checkvar29, checkvar30]

for x in range(len(lists1)):
    check1 = tk.Checkbutton(F, text=os.path.basename(lists1[x]), variable=checkvar_arr[x], onvalue=1, offvalue=0,
                            )
    check1.pack(anchor=tk.W, pady=1)

F = tk.Frame(top, bg='white', width=150, height=900,
             highlightbackground='black',
             highlightcolor='green',
             highlightthickness=1, takefocus=1)
F.place(x=360, y=50)
tk.Label(top, text='现状图:', bd=6).place(x=290, y=50)

check1 = tk.IntVar()
check2 = tk.IntVar()
check3 = tk.IntVar()
check4 = tk.IntVar()
check5 = tk.IntVar()
check6 = tk.IntVar()
check7 = tk.IntVar()
check8 = tk.IntVar()
check9 = tk.IntVar()
check10 = tk.IntVar()
check11 = tk.IntVar()
check12 = tk.IntVar()
check13 = tk.IntVar()
check14 = tk.IntVar()
check15 = tk.IntVar()
check16 = tk.IntVar()
check17 = tk.IntVar()
check18 = tk.IntVar()
check19 = tk.IntVar()
check20 = tk.IntVar()
check21 = tk.IntVar()
check22 = tk.IntVar()
check23 = tk.IntVar()
check24 = tk.IntVar()
check25 = tk.IntVar()
check26 = tk.IntVar()
check27 = tk.IntVar()
check28 = tk.IntVar()
check29 = tk.IntVar()
check30 = tk.IntVar()
check_arr = [check1, check2, check3, check4, check5, check6, check7, check8, check9, check10, check11, check12, check13,
             check14, check15, check16, check17, check18, check19, check20, check21, check22, check23, check24, check25,
             check26, check27, check28, check29, check30]

for x in range(len(lists2)):
    check1 = tk.Checkbutton(F, text=os.path.basename(lists2[x]), variable=check_arr[x], onvalue=1,
                            offvalue=0)
    check1.pack(anchor=tk.W, pady=1)

tk.Button(top, text='开始分析', bd=12).place(relx=0.89, rely=0.95, width=80, height=50)

top.mainloop()
