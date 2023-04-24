# -*- coding: utf-8 -*-
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import xlwt
import xlrd

root = tk.Tk()
root.withdraw()

# 获取filepath文件夹下的所有的文件
def getfilelist(filepath):
    filelist =  os.listdir(filepath)
    files = []
    for i in range(len(filelist)):
        child = os.path.join('%s/%s'%(filepath, filelist[i]))
        index = child.find('~$')
        if index == -1:
            if os.path.isdir(child):
                files.extend(getfilelist(child))
            else:
                files.append(child)
    return files

messagebox.showinfo('提示','请选择文件夹！')
Folderpath = filedialog.askdirectory() #获得选择好的文件夹
filelists = getfilelist(Folderpath)
writeexcel = xlwt.Workbook()
sheet = writeexcel.add_sheet('Sheet1')
i = 0
for filepath in filelists:
    if filepath.endswith('.xls') and not filepath.startswith('~$'):
        workbook = xlrd.open_workbook(filepath)
        sheet1 = workbook.sheet_by_name("Sheet1")
        #J1
        sheet.write(i,0,sheet1.cell(8,2).value)
        sheet.write(i,1,sheet1.cell(8,3).value)
        sheet.write(i,2,sheet1.cell(8,4).value)
        sheet.write(i,3,0)
        #j2
        sheet.write(i+1,0,sheet1.cell(10,2).value)
        sheet.write(i+1,1,sheet1.cell(10,3).value)
        sheet.write(i+1,2,sheet1.cell(10,4).value)
        sheet.write(i+1,3,0)
        #J3
        sheet.write(i+2,0,sheet1.cell(12,2).value)
        sheet.write(i+2,1,sheet1.cell(12,3).value)
        sheet.write(i+2,2,sheet1.cell(12,4).value)
        sheet.write(i+2,3,0)
        #j4
        sheet.write(i+3,0,sheet1.cell(14,2).value)
        sheet.write(i+3,1,sheet1.cell(14,3).value)
        sheet.write(i+3,2,sheet1.cell(14,4).value)
        sheet.write(i+3,3,0)

        i = i + 4
writeexcel.save(Folderpath+"/"+"界址点坐标.xls")
messagebox.showinfo('提示','抽取界址点坐标完成!')