# -*- coding: utf-8 -*-
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill,Border, Side, Alignment, Font
alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)

root = tk.Tk()
root.withdraw()

messagebox.showinfo('提示','请选择需要整合的Excel表！')
datapath = filedialog.askopenfilename() 
#读取文档
df1 = pd.read_excel(datapath, sheet_name=0)
xzqdm_list = df1['XZQDM'].tolist()
xzqmc_list = df1['XZQMC'].tolist()
dlmc_list = df1['DLMC'].tolist()
zldwdm_list = df1['ZLDWDM'].tolist()
pjjg_list = df1['PJJG'].tolist()
# dlbm_list = df1['DLBM'].tolist()
bz_list = df1['BZ'].tolist()
zldwmc_list = df1['ZLDWMC'].tolist()
tbdlmj_list = df1['TBDLMJ'].tolist()


df1.to_excel(datapath)