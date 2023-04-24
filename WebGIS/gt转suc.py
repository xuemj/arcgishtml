# -*- coding: utf-8 -*-
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

root = tk.Tk()
root.withdraw()

messagebox.showinfo('提示','请选择文件！')
datapath = filedialog.askopenfilename()
#读取文档
ss_arr = []
sd_arr = []
file1 = open(datapath,'r')
for line in file1.readlines():
    lin = line.strip()
    if 'SS' in lin:
        li = lin.replace(" ","").strip('SS')
        ss_arr.append(li)
    if 'SD' in lin:
        li = lin.replace(" ","").strip('SD')
        sd_arr.append(li)

print('---------',ss_arr)
print('======',sd_arr)
file1.close()

str1 = os.path.basename(datapath)
str2 = str1.split('.')[0]
f1 = open(os.path.dirname(datapath) + "/" + str2+".SUC","w+",encoding="utf-8")
for i in range(0,len(ss_arr)):
    ss = ss_arr[i]
    sd = sd_arr[i]
    
    s1 = ss.split(',')[0]
    s2 = ss.split(',')[1]
    f1.writelines(s1+','+sd+','+s2+','+'0\n')
f1.close
messagebox.showinfo('提示','SUC文件转换完成！')

