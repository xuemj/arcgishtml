# coding=utf-8
from array import array
from re import U
from site import addsitedir  
from sys import executable  
from os import path
from turtle import pd  
interpreter = executable  
sitepkg = path.dirname(interpreter) + "\\site-packages" 
addsitedir(sitepkg)

import os
from Tkinter import *
import Tkinter
import tkFileDialog
import arcpy
import xlwt

root = Tkinter.Tk()
root.title('shape相交分析工具')
root.geometry('500x500')
lists = []
analyseArr = []
datapath_list = []
Folderpath_list = []

checkvar1 = IntVar()
checkvar2 = IntVar()
checkvar3 = IntVar()
checkvar4 = IntVar()
checkvar5 = IntVar()
checkvar6 = IntVar()
checkvar7 = IntVar()
checkvar8 = IntVar()
checkvar9 = IntVar()
checkvar10 = IntVar()
checkvar11 = IntVar()
checkvar12 = IntVar()
checkvar13 = IntVar()
checkvar14 = IntVar()
checkvar15 = IntVar()
checkvar16 = IntVar()
checkvar17 = IntVar()
checkvar18 = IntVar()
checkvar19 = IntVar()
checkvar20 = IntVar()
checkvar21 = IntVar()
checkvar22 = IntVar()
checkvar23 = IntVar()
checkvar24 = IntVar()
checkvar25 = IntVar()
checkvar26 = IntVar()
checkvar27 = IntVar()
checkvar28 = IntVar()
checkvar29 = IntVar()
checkvar30 = IntVar()


checkvar_arr = [checkvar1,checkvar2,checkvar3,checkvar4,checkvar5,checkvar6,checkvar7,checkvar8,checkvar9,checkvar10,checkvar11,checkvar12,checkvar13,checkvar14,checkvar15,checkvar16,checkvar17,checkvar18,checkvar19,checkvar20,checkvar21,checkvar22,checkvar23,checkvar24,checkvar25,checkvar26,checkvar27,checkvar28,checkvar29,checkvar30]
def run():
    analyseArr = []
    for i in range(len(lists)):
        if checkvar_arr[i].get() == 1:
            analyseArr.append(lists[i])

    if analyseArr.count == 0 or datapath_list.count == 0:
        lb2.config(text='你还没有选择任何shape')
    else :
        # workbook = xlwt.Workbook()
        # sheet = workbook.add_sheet('shape')
        # row = 0
        workbook2 = xlwt.Workbook()
        sheet1 = workbook2.add_sheet('Sheet1')
        sheet1.write(0,0,u'项目名称')
        sheet1.write(0,1,u'面积')
        sheet1.write(0,2,u'地类编码')
        sheet1.write(0,3,u'地类名称')
        sheet1.write(0,4,u'权属性质')
        sheet1.write(0,5,u'镇（办）')
        sheet1.write(0,6,u'权属单位名称')
        sheet1.write(0,7,u'坐落单位名称')
        sheet1.write(0,8,u'坡度级别')
        row1 = 1
        for analyse_file in analyseArr: 
            str1 = os.path.basename(analyse_file)
            str2 = str1.split('.')[0]
            str3 = str2.replace('(','').replace(')','')
            outpath = u"C:/Users/Administrator/Desktop/相交分析结果/"+str3+"intersect.shp"
            arcpy.Intersect_analysis([datapath_list[-1],analyse_file],outpath,"ALL", "", "INPUT")
            arcpy.CalculateField_management(outpath, 'area', "!shape.area!", 'PYTHON_9.3')
            try :
                shpfields = ['area','DLBM','DLMC','QSXZ','XZQ','QSDWMC','ZLDWMC','PDJB']
                shp_area = []
                dlbm_arr = []
                dlmc_arr = []
                qsxz_arr = []
                xzq_arr = []
                qsdwmc_arr = []
                zldwmc_arr = []
                pdjb_arr = []
                index_arr = []
                shprows = arcpy.SearchCursor(outpath,shpfields)
                while True:
                    shprow = shprows.next()
                    if not shprow:
                        break
                    shp_area.append(shprow.area)
                    dlbm_arr.append(shprow.DLBM)
                    dlmc_arr.append(shprow.DLMC)
                    qsxz_arr.append(shprow.QSXZ)
                    xzq_arr.append(shprow.XZQ)
                    qsdwmc_arr.append(shprow.QSDWMC)
                    zldwmc_arr.append(shprow.ZLDWMC)
                    pdjb_arr.append(shprow.PDJB)
            except: 
                continue
            # ans = sum(shp_area)
            # sheet.write(row,0,str3)
            # sheet.write(row,1,str(ans))

            # row = row + 1
            print('---------------------',xzq_arr)
            while (len(xzq_arr) != len(index_arr)) :
                aarea = 0
                dlbm_str = ''
                dlmc_str = ''
                qsxz_str = ''
                xzq_str = ''
                qsdwmc_str = ''
                zldwmc_str = ''
                pdjb = ''
                mark = False
                for i in range(len(xzq_arr)) :
                    if i in index_arr :
                        continue 
                    else :
                        if mark == False :
                            dlbm_str = dlbm_arr[i]
                            dlmc_str = dlmc_arr[i]
                            qsxz_str = qsxz_arr[i]
                            xzq_str = xzq_arr[i]
                            qsdwmc_str = qsdwmc_arr[i]
                            zldwmc_str = zldwmc_arr[i]
                            pdjb = pdjb_arr[i]
                            mark = True
                        if dlbm_arr[i] == dlbm_str and dlmc_arr[i] == dlmc_str and qsxz_arr[i] == qsxz_str and xzq_arr[i] == xzq_str and qsdwmc_arr[i] == qsdwmc_str and zldwmc_arr[i] == zldwmc_str and pdjb_arr[i] == pdjb :
                            aarea = aarea + shp_area[i]
                            index_arr.append(i)
                sheet1.write(row1,0,str3)
                sheet1.write(row1,1,str(aarea))
                sheet1.write(row1,2,dlbm_str)
                sheet1.write(row1,3,dlmc_str)
                sheet1.write(row1,4,qsxz_str)
                sheet1.write(row1,5,xzq_str)
                sheet1.write(row1,6,qsdwmc_str)
                sheet1.write(row1,7,zldwmc_str)
                sheet1.write(row1,8,str(pdjb))
                row1 = row1 + 1
        # workbook.save(u"C:/Users/Administrator/Desktop/相交分析结果/"+u"shape相交结果.xls")
        workbook2.save(u"C:/Users/Administrator/Desktop/相交分析结果/"+u"shape统计.xls")

def newpath():
    if os.path.exists(u'C:/Users/Administrator/Desktop/相交分析结果'):
        print('文件夹已经存在')
    else:
        os.makedirs(u'C:/Users/Administrator/Desktop/相交分析结果')

def select():
    datapath = tkFileDialog.askopenfilename()
    lb1.config(text=os.path.basename(datapath))
    datapath_list.append(datapath)
    return datapath_list

def getfilelist(filepath):
    filelist =  os.listdir(filepath)
    files = []
    for i in range(len(filelist)):
        child = os.path.join('%s/%s'%(filepath, filelist[i]))
        if os.path.isdir(child):
            files.extend(getfilelist(child))
        else:
            files.append(child)
    return files

def selectPath():
    Folderpath = tkFileDialog.askdirectory() #获得选择好的文件夹
    Folderpath_list.append(Folderpath)
    for filepath in getfilelist(Folderpath):
        if filepath.endswith("shp") and not filepath.startswith('~$'):
            lists.append(filepath)
    setUI()
    return Folderpath_list
def setUI():
    for i in range(len(lists)):
        ch = Checkbutton(root,text=os.path.basename(lists[i]),variable = checkvar_arr[i],onvalue=1,offvalue=0)
        ch.pack()


newpath()
lb = Label(root,text='')
lb.pack()

btn1= Button(root,text='请选择范围shape',command=select)
btn1.pack()

lb1 = Label(root,text='')
lb1.pack()

btn2 = Button(root,text='请选择相交shape文件夹',command=selectPath)
btn2.pack()

btn = Button(root,text="开始分析",fg="red",relief=GROOVE,command=run)
btn.place(relx=0.8,rely=0.9,width=100,height=50)

lb2 = Label(root,text='')
lb2.pack()
root.mainloop()
