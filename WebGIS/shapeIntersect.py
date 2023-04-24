# coding=utf-8
from re import U
from site import addsitedir  
from sys import executable  
from os import path
interpreter = executable  
sitepkg = path.dirname(interpreter) + "\\site-packages" 
addsitedir(sitepkg)

import os
import io
import shutil
from Tkinter import *
import Tkinter as tk
import tkFileDialog
import arcpy
import xlwt

datapath_list = []
outpath_list1 = []
outpath_list2 = []
lists1 = []
lists2 = []
B2_status = True
B1_status = True

root = tk.Tk()
root.title('shape相交分析工具')
root.maxsize(800, 950)
root.minsize(800, 500)

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
checkvar_arr = [checkvar1,checkvar2,checkvar3,checkvar4,checkvar5,checkvar6,checkvar7,checkvar8,checkvar9,checkvar10,checkvar11,checkvar12,checkvar13,checkvar14,checkvar15,checkvar16,checkvar17,checkvar18,checkvar19,checkvar20,checkvar21,checkvar22,checkvar23,checkvar24,checkvar25,checkvar26,checkvar27,checkvar28,checkvar29,checkvar30]

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
def run():
    analyseArr1 = []
    analyseArr2 = []
    str4 = os.path.basename(datapath_list[-1])
    str5 = str4.split('.')[0]
    fpath = str5.replace('(','').replace(')','')
    for i in range(len(lists1)):
        if checkvar_arr[i].get() == 1:
            analyseArr1.append(lists1[i])
    for i in range(len(lists2)):
        if check_arr[i].get() == 1:
            analyseArr2.append(lists2[i])
    
    if len(analyseArr1) > 0 and len(outpath_list1) > 0:
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('Sheet1')
        row = 1
        sheet.write(0,0,u'项目名称')
        sheet.write(0,1,u'面积')
        sheet.write(0,2,'GJLYD')
        out1 = outpath_list1[-1]
        for analyse_file in analyseArr1:
            str1 = os.path.basename(analyse_file)
            str2 = str1.split('.')[0]
            str3 = str2.replace('(','').replace(')','')
            outpath = out1 + "/"+fpath+' X '+str3+".shp"
            if os.path.exists(outpath):
                print('已存在')
            else :
                arcpy.Intersect_analysis([datapath_list[-1],analyse_file],outpath,"ALL", "", "INPUT")
                arcpy.CalculateField_management(outpath, 'area', "!shape.area!", 'PYTHON_9.3')
            shpfields = ['area','GJLYD']
            shp_area = []
            gjlyd_arr = []
            shprows = arcpy.SearchCursor(outpath,shpfields)
            while True:
                shprow = shprows.next()
                if not shprow:
                    break
                shp_area.append(shprow.area)
                try :
                    gjlyd_arr.append(shprow.GJLYD)
                except: 
                    continue
            if len(gjlyd_arr) > 0 :
                arr = set(gjlyd_arr)
                for a in arr :
                    shparea = 0.0000
                    for i in range(len(shp_area)) :
                        if gjlyd_arr[i] == a :
                            shparea = shparea + shp_area[i]
                    sheet.write(row,0,str3)
                    sheet.write(row,1,shparea)
                    sheet.write(row,2,a)
                    row = row + 1
            else :
                ans = sum(shp_area)  
                sheet.write(row,0,str3)
                sheet.write(row,1,ans)
                row = row + 1
        workbook.save(out1 + "/"+u"intersect.xls")
        os.startfile(out1 + "/"+u"intersect.xls")
    if len(analyseArr2) > 0 and len(outpath_list2) > 0:
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
        workbook3 = xlwt.Workbook()+-
        sheet2 = workbook3.add_sheet(u'地类统计')
        sheet2.write(0,0,u'项目名称')
        sheet2.write(0,1,u'面积')
        sheet2.write(0,2,u'地类编码')   
        sheet2.write(0,3,u'地类名称')
        workbook4 = xlwt.Workbook()
        sheet3 = workbook4.add_sheet(u'行政区统计')
        sheet3.write(0,0,u'项目名称')
        sheet3.write(0,1,u'面积')
        sheet3.write(0,2,u'镇（办）')
        sheet3.write(0,3,u'权属单位名称')
        sheet3.write(0,4,u'坐落单位名称')
        row1 = 1
        row2 = 1
        row3 = 1
        out2 = outpath_list2[-1]
        for analyse_file in analyseArr2: 
            str1 = os.path.basename(analyse_file)
            str2 = str1.split('.')[0]
            str3 = str2.replace('(','').replace(')','')
            outpath = out2 + "/" +fpath+' X '+str3+".shp"
            if os.path.exists(outpath):
                print('已存在')
            else :
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
                index_arr1 = []
                index_arr2 = []
                index_arr3 = []
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
            while (len(xzq_arr) != len(index_arr1)) :
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
                    if i in index_arr1 :
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
                            index_arr1.append(i)
                sheet1.write(row1,0,str3)
                sheet1.write(row1,1,aarea)
                sheet1.write(row1,2,dlbm_str)
                sheet1.write(row1,3,dlmc_str)
                sheet1.write(row1,4,qsxz_str)
                sheet1.write(row1,5,xzq_str)
                sheet1.write(row1,6,qsdwmc_str)
                sheet1.write(row1,7,zldwmc_str)
                sheet1.write(row1,8,str(pdjb))
                row1 = row1 + 1
            while (len(xzq_arr) != len(index_arr2)) :
                aarea = 0
                dlbm_str = ''
                dlmc_str = ''
                mark = False
                for i in range(len(xzq_arr)) :
                    if i in index_arr2 :
                        continue 
                    else :
                        if mark == False :
                            dlbm_str = dlbm_arr[i]
                            dlmc_str = dlmc_arr[i]
                            mark = True
                        if dlbm_arr[i] == dlbm_str and dlmc_arr[i] == dlmc_str :
                            aarea = aarea + shp_area[i]
                            index_arr2.append(i)
                sheet2.write(row2,0,str3)
                sheet2.write(row2,1,aarea)
                sheet2.write(row2,2,dlbm_str)
                sheet2.write(row2,3,dlmc_str)
                row2 = row2 + 1
            while (len(xzq_arr) != len(index_arr3)) :
                aarea = 0
                xzq_str = ''
                qsdwmc_str = ''
                zldwmc_str = ''
                mark = False
                for i in range(len(xzq_arr)) :
                    if i in index_arr3 :
                        continue 
                    else :
                        if mark == False :
                            xzq_str = xzq_arr[i]
                            qsdwmc_str = qsdwmc_arr[i]
                            zldwmc_str = zldwmc_arr[i]
                            mark = True
                        if xzq_arr[i] == xzq_str and qsdwmc_arr[i] == qsdwmc_str and zldwmc_arr[i] == zldwmc_str:
                            aarea = aarea + shp_area[i]
                            index_arr3.append(i)
                sheet3.write(row3,0,str3)
                sheet3.write(row3,1,aarea)
                sheet3.write(row3,2,xzq_str)
                sheet3.write(row3,3,qsdwmc_str)
                sheet3.write(row3,4,zldwmc_str)
                row3 = row3 + 1
        workbook2.save(out2 + "/"+u"intersect.xls")
        workbook3.save(out2 + "/"+u"地类统计.xls")
        workbook4.save(out2 + "/"+u"行政区统计.xls")

def del_file(path):
    ls = os.listdir(path)
    for i in ls:
        c_path = os.path.join(path, i)
        if os.path.isdir(c_path):
            del_file(c_path)
        else:
            os.remove(c_path)
def outpath():
    outfolderpath = tkFileDialog.askdirectory()
    lb2.config(text=outfolderpath)
    path1 = outfolderpath + u'/相交分析结果/其他资料'
    path2 = outfolderpath + u'/相交分析结果/现状图'
    if os.path.exists(path1) :
        print()
    else :
        os.makedirs(path1)
    outpath_list1.append(path1)
    if os.path.exists(path2) :
        print()
    else :
        os.makedirs(path2)
    outpath_list2.append(path2)
    
    return outpath_list1,outpath_list2

def readpath1():
    if os.path.exists('dataPath1.txt'):
        file1 = open('dataPath1.txt','r')
        content1 = file1.readline()
        if content1 == '':
            return []
        else :
            content = to_unicode(content1)
            arr = appedlists(content)
            file1.close
            return arr
    else:
        f1 = io.open('dataPath1.txt', 'w+', encoding='utf8')
        f1.close
        return []
def readpath2():
    if os.path.exists('dataPath2.txt'):
        file2 = open('dataPath2.txt','r')
        content2 = file2.readline()
        if content2 == '':
            return []
        else :    
            content = to_unicode(content2)
            arr = appedlists(content)
            file2.close
            return arr
    else:
        f2 = io.open('dataPath2.txt', 'w+', encoding='utf8')
        f2.close
        return []

def to_unicode(unicode_or_str):
    if isinstance(unicode_or_str, str):
        value = unicode_or_str.decode('utf-8')
    else:
        value = unicode_or_str
    return value
def select():
    datapath = tkFileDialog.askopenfilename()
    lb1.config(text=os.path.basename(datapath))
    if datapath.endswith("dwg") and not datapath.startswith('~$'):
        defaultpath = os.path.dirname(datapath)
        strname = os.path.basename(datapath)
        CADname = strname.split('.')[0]
        # 定义工作空间
        arcpy.env.workspace =defaultpath
        # CAD文件路径
        input_cad_dataset =os.path.join(defaultpath,CADname+'.dwg')
        # gdb文件路径
        out_gdb_path = os.path.join(defaultpath,CADname+'.gdb')
        # 要素集文件名称
        out_dataset_name = CADname
        # CAD转shp坐标比例
        reference_scale = "1"
        # 先创建一个gdb地理数据库
        arcpy.CreateFileGDB_management(defaultpath, CADname+'.gdb')
        # 将CAD文件导入到gdb地理数据库，新建一个要素集
        arcpy.CADToGeodatabase_conversion(input_cad_dataset, out_gdb_path, out_dataset_name, reference_scale)
        
        # 切换工作空间到gdb中
        arcpy.env.workspace = out_gdb_path
        # 获取gdb中的文件列表
        datasets = arcpy.ListDatasets(feature_type='feature')
        # 输入shp文件的保存路径
        output_shp_path=defaultpath
        datasets = [''] + datasets if datasets is not None else []
        # 获取每个地理数据库中的要素集
        for ds in datasets:
            for fc in arcpy.ListFeatureClasses(feature_dataset=ds):
                path = os.path.join(arcpy.env.workspace, ds, fc)
                outfc = arcpy.ValidateTableName(fc)
                # 将要素集里的要素转为shp文件
                if outfc == 'Polygon':
                    arcpy.FeatureClassToShapefile_conversion(outfc, output_shp_path)
        newpath = defaultpath + '/' + 'Polygon.shp'
        shutil.copy(u'Polygon.prj',defaultpath+'/Polygon.prj')
        datapath_list.append(newpath)
    elif datapath.endswith("shp") and not datapath.startswith('~$'):
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

def appedlists(f):
    lists = []
    for filepath in getfilelist(f):
        if filepath.endswith("shp") and not filepath.startswith('~$'):
            lists.append(filepath)
    return lists

def selectlist1(): 
    global lists1
    if len(lists1) > 0:
        del lists1[:]
    Folderpath = tkFileDialog.askdirectory() #获得选择好的文件夹
    lists1.extend(appedlists(Folderpath))
    clear1()
    for x in range(len(lists1)):
        check1 = tk.Checkbutton(F1, text=os.path.basename(lists1[x]), variable=checkvar_arr[x], onvalue=1, offvalue=0,
                            )
        check1.pack(anchor=tk.W, pady=1)
    f1 = io.open(u'dataPath1.txt', 'w+', encoding='utf8')
    f1.write(Folderpath)
    f1.close

def selectlist2():
    global lists2
    if len(lists2) > 0 :
        del lists2[:]
    Folderpath = tkFileDialog.askdirectory() #获得选择好的文件夹
    lists2.extend(appedlists(Folderpath))
    clear2()
    for x in range(len(lists2)):
        check1 = tk.Checkbutton(F2, text=os.path.basename(lists2[x]), variable=check_arr[x], onvalue=1,
                            offvalue=0)
        check1.pack(anchor=tk.W, pady=1)
    f2 = io.open(u'dataPath2.txt', 'w+', encoding='utf8')
    f2.write(Folderpath)
    f2.close

def clear1():
    for btn in F1.winfo_children():
        btn.destroy()
    for ch in checkvar_arr:
        ch.set(0)
def clear2():
    for btn in F2.winfo_children():
        btn.destroy()
    for ch in check_arr:
        ch.set(0)
def allselect1():
    global B1_status
    if B1_status == True:
        B1.config(text='取消全选')
        for ch in checkvar_arr:
            ch.set(1)
        B1_status = False
    else :
        B1.config(text='全选')
        for ch in checkvar_arr:
            ch.set(0)
        B1_status = True
def allselect2():
    global B2_status
    if B2_status == True:
        B2.config(text='取消全选')
        for ch in check_arr:
            ch.set(1)
        B2_status = False
    else :
        B2.config(text='全选')
        for ch in check_arr:
            ch.set(0)
        B2_status = True

lists1.extend(readpath1())
lists2.extend(readpath2())

tk.Button(root, text='选取范围',command=select,bd=6).place(x=100, y=10)
lb1 = tk.Label(root,text='请选择shp范围...')
lb1.place(x=170, y=20)

tk.Button(root, text='导出位置',command=outpath,bd=6).place(x=350, y=10)
lb2 = tk.Label(root,text='请选择导出位置...')
lb2.place(x=420, y=20)
F1 = tk.Frame(root, bg='white',
             highlightbackground='black',
             highlightcolor='green',
             highlightthickness=1, takefocus=1
             )

F1.place(x=80, y=60)
tk.Button(root, text='其他资料:',command=selectlist1,bd=6).place(x=9, y=60)
B1 = tk.Button(root,text='全选',command=allselect1,bd=6)
B1.place(x=9,y=120)


F2 = tk.Frame(root, bg='white',
             highlightbackground='black',
             highlightcolor='green',
             highlightthickness=1, takefocus=1)
F2.place(x=460, y=60)
tk.Button(root, text='现状图:',command=selectlist2,bd=6).place(x=390, y=60)

B2 = tk.Button(root,text='全选',command=allselect2,bd=6)
B2.place(x=390,y=120)


tk.Button(root, text='开始分析',command=run, bd=12).place(relx=0.9, rely=0.9, width=80, height=50)

for x in range(len(lists1)):
        check1 = tk.Checkbutton(F1, text=os.path.basename(lists1[x]), variable=checkvar_arr[x], onvalue=1, offvalue=0,
                            )
        check1.pack(anchor=tk.W, pady=1)
for x in range(len(lists2)):
        check1 = tk.Checkbutton(F2, text=os.path.basename(lists2[x]), variable=check_arr[x], onvalue=1,
                            offvalue=0)
        check1.pack(anchor=tk.W, pady=1)

root.mainloop()
