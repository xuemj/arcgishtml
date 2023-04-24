import os
import arcpy
  
fileDir=arcpy.GetParameterAsText(0)
out =arcpy.GetParameterAsText(1)
ver=arcpy.GetParameterAsText(2)
  
for r,dirs,files in os.walk(fileDir):
    for mxdFile in files:
        if mxdFile[-3:].lower()=="mxd":
            mxd=arcpy.mapping.MapDocument(os.path.join(r,mxdFile))
            mxd.saveACopy(os.path.join(out,mxdFile),ver)
