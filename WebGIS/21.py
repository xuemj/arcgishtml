import arcpy
import pythonaddins

class button1(object):
    """Implementation for TIF2021_addin.button (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        MXD = arcpy.mapping.MapDocument("current")
        layers = arcpy.mapping.ListLayers(MXD)
        layer_1 = layers[0]
        paths = []
        with arcpy.da.SearchCursor(layer_1,["RefName"]) as cursor:
            for row in cursor:
                print ("---------------",row[0])
                paths.append("L:/2021Q4/BIGDOM/"+row[0]+".tif")
        pylevel = "-1"
        skipfirst = "NONE"
        resample = "NEAREST"
        compress = "NONE"
        quality = "75"
        skipexist = "SKIP_EXISTING"
        for tiff in paths:
            print ("=========",tiff)
            arcpy.BuildPyramids_management(tiff,pylevel, skipfirst, resample, 
                               compress, quality, skipexist)
            