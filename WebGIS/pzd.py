# coding=utf-8
from Tkinter import *
root=Tk()
LB2=Listbox(root,selectmode=EXTENDED)
Label(root,text='多选：你会几种编程语言').pack()
for item in ['python','C++','C','Java','Php']:
    LB2.insert(END,item)
LB2.insert(1,'JS','Go','R')
LB2.delete(5,6)
LB2.select_set(0,3)
LB2.select_clear(0,1)

LB2.pack()
 
root.mainloop()
