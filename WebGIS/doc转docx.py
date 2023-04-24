import os
import tkinter as tk
from tkinter import filedialog
from win32com import client as wc
from tkinter import messagebox

root = tk.Tk()
root.withdraw()

messagebox.showinfo('提示','请选择文件夹！')

Folderpath = filedialog.askdirectory() #获得选择好的文件夹

# 获取filepath文件夹下的所有的文件
def getfilelist(filepath):
    filelist =  os.listdir(filepath)
    files = []
    for i in range(len(filelist)):
        child = os.path.join('%s/%s'%(filepath, filelist[i]))
        index = child.find('~$')
        if index == -1:
            if os.path.isdir(child):
                files.extend(getfilelist(child))
            else:
                files.append(child)
    return files
word = wc.Dispatch("kwps.Application")
filelist = getfilelist(Folderpath)
for file in filelist:
    # 找出文件中以.doc结尾并且不以~$开头的文件（~$是为了排除临时文件的）
    if file.endswith('.doc') and not file.startswith('~$'):
        # try
        # 打开文件
        try:
            doc = word.Documents.Open(file)
            doc.SaveAs("{}x".format(file), 12)  # 另存为后缀为".docx"的文件，其中参数12指docx文件
            doc.Close()  # 关闭原来word文件
            print('变更完毕--------------------------------',file)
            os.remove(file)
        except:
            print('doc----',file)
            continue
word.Quit()
messagebox.showinfo('提示','Word文档变更完成')
