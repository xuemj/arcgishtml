# coding=utf-8
from distutils.core import setup
import py2exe
options = {"py2exe": {"excludes": ["arcpy"]}}  
setup(windows=['shp.py'], options=options) 
