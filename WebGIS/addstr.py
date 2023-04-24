

import arcpy
f=open("C:/Users/Administrator/Desktop/sxb.txt",'r')
line=f.readline()
while line:
    lineList=line.split(',')
    arcpy.AddField_management('lyr',field_name=lineList[0],field_alias=lineList[1], field_type=lineList[2],field_length=lineList[3],field_precision=lineList[4].replace("\n",""))
    line=f.readline()
f.close()

