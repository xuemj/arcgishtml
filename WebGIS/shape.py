# -*- coding: utf-8 -*-
# coding: utf-8
from operator import ge
import os
from tkinter import filedialog
import tkinter as tk
from tkinter import messagebox
import shapefile

root = tk.Tk()
root.withdraw()

geometry_arr = []
record_arr = []
sort_arr = []
messagebox.showinfo('提示','请选择shape文件！')
filepath = filedialog.askopenfilename() 
shfile = shapefile.Reader(filepath,encoding="utf-8")
shapes = shfile.shapes()
dir,ext=os.path.splitext(filepath)
f1 = open(dir + ".txt","w+",encoding="utf-8")
f1.writelines('[属性描述]\n格式版本号=\n数据产生单位=国土资源部\n数据产生日期=2021-6-28\n坐标系=2000国家大地坐标系\n几度分带=3\n投影类型=高斯克吕格\n计量单位=米\n带号=37\n精度=0.0001\n转换坐标=0,0,0,0,0,0\n[地块坐标]\n') 
# 按照地块序号排列数据
for j in range(len(shapes)):
    geometry = shapes[j]
    geometry_arr.append(geometry)
    record = shfile.record(j)
    record_arr.append(record)
    sort_arr.append(record[1])
print('-----------',sort_arr)
l = len(sort_arr)
for m in range(l-1) :
    minindex = m
    for n in range(m+1,l,1) :
        if sort_arr[n] < sort_arr[minindex] :
            minindex = n  
    sort_arr[m],sort_arr[minindex] = sort_arr[minindex],sort_arr[m]
    geometry_arr[m],geometry_arr[minindex] = geometry_arr[minindex],geometry_arr[m]
    record_arr[m],record_arr[minindex] = record_arr[minindex],record_arr[m]
print('=================',sort_arr)
for index in range(len(record_arr)) :
    geometry = geometry_arr[index]
    list_arr = record_arr[index]
    point_count = len(geometry.points)
    f1.writelines(str(point_count)+','+str(list_arr[3])+','+str(list_arr[1])+','+str(list_arr[0])+','+'面'+','+str(list_arr[2])+','+str(list_arr[4])+','+str(list_arr[5])+','+str(list_arr[6])+','+str(list_arr[7])+','+str(list_arr[8])+','+str(list_arr[9])+','+'@'+'\n')
    head_count = 1
    head_count_m= 1
    a = 2 
    list_x = []
    for i in range(len(geometry.points)):
        point = str(geometry.points[i])
        point_m = point.replace('(','').replace(' ','').replace(')','')
        point_arr = point_m.split(',')
        point_x = format(float(point_arr[1]),'.4f')
        point_y = format(float(point_arr[0]),'.4f')
        head = 'J'+str(i+1-head_count_m+1)
        if point in list_x:
            a = a-1
            f1.writelines('J'+str(head_count)+','+str(head_count_m)+','+str(point_x)+','+str(point_y)+'\n')
            head_count_m = head_count_m+1
            head_count = i+a
        else :
            f1.writelines(head+','+str(head_count_m)+','+str(point_x)+','+str(point_y)+'\n')
            list_x.append(point)
    f1.close
messagebox.showinfo('提示','shape坐标txt完成！')
