import os
from tkinter import filedialog


txtLists = []
def bLi(rootDir):
    for root,dirs,files in os.walk(rootDir):
        for file in files:
            txtLists.append(os.path.join(root,file))
        for dir in dirs:
            bLi(dir)
            
Folderpath = filedialog.askdirectory() #获得选择好的文件夹
#获取当前文件夹中的文件名称列表  
bLi(Folderpath)

f1 = open(Folderpath + "/" + "result_cam1_pos.txt","w+",encoding="utf-8")
f2 = open(Folderpath + "/" + "result_cam2_pos.txt","w+",encoding="utf-8")
f3 = open(Folderpath + "/" + "result_cam3_pos.txt","w+",encoding="utf-8")
f4 = open(Folderpath + "/" + "result_cam4_pos.txt","w+",encoding="utf-8")
f5 = open(Folderpath + "/" + "result_cam5_pos.txt","w+",encoding="utf-8")

for file in txtLists:
    if "cam1_pos" in file:
        for line in open(file,encoding="utf-8"):
            f1.writelines(line)
        f1.write('\n')
    if "cam2_pos" in file:
        for line in open(file,encoding="utf-8"):
            f2.writelines(line)
        f2.write('\n')
    if "cam3_pos" in file:
        for line in open(file,encoding="utf-8"):
            f3.writelines(line)
        f3.write('\n')
    if "cam4_pos" in file:
        for line in open(file,encoding="utf-8"):
            f4.writelines(line)
        f4.write('\n')
    if "cam5_pos" in file:
        for line in open(file,encoding="utf-8"):
            f5.writelines(line)
        f5.write('\n')
        
f1.close()
f2.close()
f3.close()
f4.close() 
f5.close() 