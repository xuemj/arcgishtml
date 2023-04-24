# -*- coding: utf-8 -*-
from operator import ge
import os
from tkinter import filedialog
import tkinter as tk
from tkinter import messagebox
import shapefile
import xlwt

root = tk.Tk()
root.withdraw()

geometry_arr = []
record_arr = []
messagebox.showinfo('提示','请选择shape文件！')
filepath = filedialog.askopenfilename() 
shfile = shapefile.Reader(filepath,encoding="gbk")
shapes = shfile.shapes()

writeexcel = xlwt.Workbook()
sheet = writeexcel.add_sheet('Sheet1')
row = 0
for i in range(len(shapes)):
    geometry = shapes[i]
    print('-----------------',len(geometry.points))
    for j in range(len(geometry.points)-1):
        point = str(geometry.points[j])
        point_m = point.replace('(','').replace(' ','').replace(')','')
        point_arr = point_m.split(',')
        # point_x = format(float(point_arr[1]),'.4f')
        # point_y = format(float(point_arr[0]),'.4f')
        point_x = float(point_arr[1])
        point_y = float(point_arr[0])
        count = j+1
        title = 'J'+str(count)
        sheet.write(row+j,0,title)
        sheet.write(row+j,1,point_x)
        sheet.write(row+j,2,point_y)
        sheet.write(row+j,3,0)
    row = row + len(geometry.points)-1
writeexcel.save(os.path.dirname(filepath)+"/"+"界址点坐标.xls")
messagebox.showinfo('提示','界址点坐标提取完成！')