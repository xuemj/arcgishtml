# -*- coding: utf-8 -*-
import os
from Tkinter import *
import Tkinter
import tkFileDialog

lists = []
root = Tkinter.Tk()
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
    print('===============',Folderpath)
    for filepath in getfilelist(Folderpath):
        if filepath.endswith("docx") and not filepath.startswith('~$'):
            print('----------------',filepath)
            lists.append(filepath)


def select():
    datapath = tkFileDialog.askopenfilename()
    print('++++++++++++++++++++',datapath)

select()