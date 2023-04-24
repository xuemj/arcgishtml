# coding=utf-8
from csv import list_dialects
from re import U
from site import addsitedir, check_enableusersite  
from sys import executable  
from os import path
from tabnanny import check
interpreter = executable  
sitepkg = path.dirname(interpreter) + "\\site-packages" 
addsitedir(sitepkg)

import os
from Tkinter import *
import Tkinter as tk
import tkFileDialog
import arcpy
import xlwt
import shutil

root = tk.Tk()
root.title('shape处理')
root.maxsize(600, 950)
root.minsize(600, 500)

lists_arr = []
datapath_list = []

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
check31 = tk.IntVar()
check32 = tk.IntVar()
check33 = tk.IntVar()
check34 = tk.IntVar()
check35 = tk.IntVar()
check36 = tk.IntVar()
check37 = tk.IntVar()
check38 = tk.IntVar()
check39 = tk.IntVar()
check40 = tk.IntVar()
check_arr = [check1, check2, check3, check4, check5, check6, check7, check8, check9, check10, check11, check12, check13,
             check14, check15, check16, check17, check18, check19, check20, check21, check22, check23, check24, check25,
             check26, check27, check28, check29, check30,check31,check32,check33,check34,check35,check36,check36,check37,check38,check39,check40]

def to_unicode(unicode_or_str):
    if isinstance(unicode_or_str, str):
        value = unicode_or_str.decode('utf-8')
    else:
        value = unicode_or_str
    return value

def copyshp(shp_path):
    path_str1 = os.path.dirname(shp_path)
    path_str2 = os.path.basename (shp_path)
    path_str3 = path_str2.split('.')[0]
    path_str = path_str1 + "/" + to_unicode(path_str3)
    shp = path_str+'1'+'.shp'
    cpg = path_str+'.cpg'
    dbf = path_str+'.dbf'
    sbn = path_str+'.sbn'
    sbx = path_str+'.sbx'
    shp_xml = path_str+'.shp.xml'
    shx = path_str+'shx'
    try:
        shutil.copy(path_str + '.cpg',path_str +'1'+'.cpg')
    except:
        print('缺少cpg')
    try:
        shutil.copy(path_str + '.dbf',path_str +'1'+'.dbf')
    except:
        print('缺少dbf')
    try:
        shutil.copy(path_str + '.prj',path_str +'1'+'.prj')
    except:
        print('缺少prj')
    try:
        shutil.copy(path_str + '.sbn',path_str +'1'+'.sbn')
    except:
        print('缺少sbn')
    try:
        shutil.copy(path_str + '.sbx',path_str +'1'+'.sbx')
    except:
        print('缺少sbx')
    try:
        shutil.copy(path_str + '.shp.xml',path_str +'1'+'.shp.xml')
    except:
        print('缺少shp.xml')
    try:
        shutil.copy(path_str + '.shx',path_str +'1'+'.shx')
    except:
        print('缺少shx')
    shutil.copy(path_str + '.shp',shp)
    return shp

def select():
    global lists_arr
    if len(lists_arr) > 0:
        del lists_arr[:]
    datapath = tkFileDialog.askopenfilename()
    lb1.config(text=os.path.basename(datapath))
    data_copyfile = copyshp(datapath)
    if data_copyfile.endswith("shp") and not data_copyfile.startswith('~$'):
        datapath_list.append(data_copyfile)
        shpfields = ['DLMC']
        shprows = arcpy.SearchCursor(data_copyfile,shpfields)
        while True:
            shprow = shprows.next()
            if not shprow:
                break
            if shprow.DLMC not in lists_arr:
                lists_arr.append(shprow.DLMC)
        clear()
        for x in range(len(lists_arr)):
            check1 = tk.Checkbutton(F2, text=os.path.basename(lists_arr[x]), variable=check_arr[x], onvalue=1,
                                offvalue=0)
            check1.pack(anchor=tk.W, pady=1)
    return datapath_list
def run():
    analyseArr = []
    for i in range(len(lists_arr)):
        if check_arr[i].get() == 1:
            analyseArr.append(lists_arr[i])
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Sheet1')
    index = 1
    sheet.write(0,0,u'地块类型')
    sheet.write(0,1,u'面积')
    dir_path = os.path.dirname(datapath_list[-1])
    if len(analyseArr) > 0:
        with arcpy.da.UpdateCursor(datapath_list[-1],["DLMC"]) as cursor:
            for row in cursor:
                strDLMC = row[0]
                if strDLMC in analyseArr:
                    print('---')
                else:
                    cursor.deleteRow()
    arcpy.CalculateField_management(datapath_list[-1], 'area', "!shape.area!", 'PYTHON_9.3')
    shpfields = ['area','DLMC']
    shp_area = []
    dlmc_arr = []
    shprows = arcpy.SearchCursor(datapath_list[-1],shpfields)
    while True:
        shprow = shprows.next()
        if not shprow:
            break
        shp_area.append(shprow.area)
        dlmc_arr.append(shprow.DLMC)
    arr = set(dlmc_arr)
    for a in arr :
        shparea = 0.0000
        for i in range(len(shp_area)) :
            if dlmc_arr[i] == a :
                shparea = shparea + shp_area[i]
        sheet.write(index,0,a)
        sheet.write(index,1,shparea)
        index = index + 1
    workbook.save(dir_path + "/"+u"面积统计.xls")
def clear():
    for btn in F2.winfo_children():
        btn.destroy()
    for ch in check_arr:
        ch.set(0)
tk.Button(root, text='选取shp',command=select,bd=6).place(x=100, y=10)
lb1 = tk.Label(root,text='请选择shp...')
lb1.place(x=170, y=20)

F2 = tk.Frame(root, bg='white',
             highlightbackground='black',
             highlightcolor='green',
             highlightthickness=1, takefocus=1)
F2.place(x=200, y=60)

tk.Button(root, text='开始处理',command=run, bd=12).place(relx=0.88, rely=0.9, width=80, height=50)

root.mainloop()
