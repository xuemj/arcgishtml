# -*- coding: utf-8 -*-
import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

import pandas as pd
import shutil

root = tk.Tk()
root.withdraw()

messagebox.showinfo('提示','请选择需要比对的Excel总表！')
datapath = filedialog.askopenfilename() #获得选择好的总excel文件做比对
#读取文档
