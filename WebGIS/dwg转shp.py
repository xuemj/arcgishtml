# coding=utf-8
import os
from Tkinter import *
import Tkinter as tk
import tkFileDialog
import arcpy

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
        if filepath.endswith("dwg") and not filepath.startswith('~$'):
            lists.append(filepath)
    return lists


Folderpath = tkFileDialog.askdirectory() #获得选择好的文件夹
dwglists = appedlists(Folderpath)

# 自定义输入文件路径
defaultpath= Folderpath
# 输入CAD文件名称
CADname='test'
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
arcpy.CreateFileGDB_management(defaultpath, 'test.gdb')
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

