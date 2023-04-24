
# lists = []

# def getfilelist(filepath):
#     filelist =  os.listdir(filepath)
#     files = []
#     for i in range(len(filelist)):
#         child = os.path.join('%s/%s'%(filepath, filelist[i]))
#         if os.path.isdir(child):
#             files.extend(getfilelist(child))
#         else:
#             files.append(child)
#     return files

# def appedlists(f):
#     for filepath in getfilelist(f):
#         if filepath.endswith("dwg") and not filepath.startswith('~$'):
#             lists.append(filepath)
#     return lists


# Folderpath = tkFileDialog.askdirectory() #获得选择好的文件夹
# appedlists(Folderpath)