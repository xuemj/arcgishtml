# -*- coding: utf-8 -*-
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd

root = tk.Tk()
root.withdraw()

messagebox.showinfo('提示','请选择需要整合的Excel表！')
datapath = filedialog.askopenfilename() 
#读取文档
df1 = pd.read_excel(datapath, sheet_name=0)
