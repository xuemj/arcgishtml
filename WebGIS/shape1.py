# -*- coding: utf-8 -*-
# coding: utf-8

import os
from tkinter import filedialog
import tkinter as tk
from tkinter import messagebox
import shapefile

root = tk.Tk()
root.withdraw()

messagebox.showinfo('提示','请选择shape文件！')
filepath = filedialog.askopenfilename() 
shfile = shapefile.Reader(filepath,encoding="gbk")
shapes = shfile.shapes()
dir,ext=os.path.splitext(filepath)
f1 = open(dir + ".txt","w+",encoding="utf-8")
f1.writelines('[属性描述]\n坐标系=2000国家大地坐标系\n几度分带=3\n投影类型=高斯克吕格\n计量单位=米\n带号=37\n精度=0.0001\n转换坐标=1,2,3,4,5,6,7\n[地块坐标]\n') 
for index in range(len(shapes)):
    geometry = shapes[index]
    list_arr = shfile.record(index)
    point_count = len(set(geometry.points))
    f1.writelines(str(point_count)+','+str(list_arr[3])+',,'+str(list_arr[0])+',,,,,'+'@'+'\n')
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
messagebox.showinfo('提示','shape补划坐标txt完成！')
