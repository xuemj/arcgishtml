# coding=utf-8
import os
from turtle import pd
from Tkinter import *
import Tkinter
import tkFileDialog
import xlwt
import arcpy

datapath = tkFileDialog.askopenfilename()

fileds = ['TBDLMJ','PDJB','TCHDJB','TRZDJB','TRYJZHLJB','TRPHZJB','SWDYXJB','TRZJSWRJB','SZJB','GDEJDLJB','XZQDM']
pdmj_1 = 0.000000
pdmj_2 = 0.000000
pdmj_3 = 0.000000
pdmj_4 = 0.000000
pdmj_5 = 0.000000
tchdmj_1 = 0.000000
tchdmj_2 = 0.000000
tchdmj_3 = 0.000000
trzdmj_1 = 0.000000
trzdmj_2 = 0.000000
trzdmj_3 = 0.000000
tryjzhlmj_1 = 0.000000
tryjzhlmj_2 = 0.000000
tryjzhlmj_3 = 0.000000
trphzmj_10 = 0.000000
trphzmj_2a = 0.000000
trphzmj_2b = 0.000000
trphzmj_3a = 0.000000
trphzmj_3b = 0.000000
swdyxmj_1 = 0.000000
swdyxmj_2 = 0.000000
swdyxmj_3 = 0.000000
trzjswrmj_1 = 0.000000
trzjswrmj_2 = 0.000000
trzjswrmj_3 = 0.000000
szmj_1 = 0.000000
szmj_2 = 0.000000
szmj_3 = 0.000000
gdejdlmj_j = 0.000000
gdejdlmj_g = 0.000000


pdmj_1_1001 = 0.000000
pdmj_2_1001 = 0.000000
pdmj_3_1001 = 0.000000
pdmj_4_1001 = 0.000000
pdmj_5_1001 = 0.000000
tchdmj_1_1001 = 0.000000
tchdmj_2_1001 = 0.000000
tchdmj_3_1001 = 0.000000
trzdmj_1_1001 = 0.000000
trzdmj_2_1001 = 0.000000
trzdmj_3_1001 = 0.000000
tryjzhlmj_1_1001 = 0.000000
tryjzhlmj_2_1001 = 0.000000
tryjzhlmj_3_1001 = 0.000000
trphzmj_10_1001 = 0.000000
trphzmj_2a_1001 = 0.000000
trphzmj_2b_1001 = 0.000000
trphzmj_3a_1001 = 0.000000
trphzmj_3b_1001 = 0.000000
swdyxmj_1_1001 = 0.000000
swdyxmj_2_1001 = 0.000000
swdyxmj_3_1001 = 0.000000
trzjswrmj_1_1001 = 0.000000
trzjswrmj_2_1001 = 0.000000
trzjswrmj_3_1001 = 0.000000
szmj_1_1001 = 0.000000
szmj_2_1001 = 0.000000
szmj_3_1001 = 0.000000
gdejdlmj_j_1001 = 0.000000
gdejdlmj_g_1001 = 0.000000


pdmj_1_1002 = 0.000000
pdmj_2_1002 = 0.000000
pdmj_3_1002 = 0.000000
pdmj_4_1002 = 0.000000
pdmj_5_1002 = 0.000000
tchdmj_1_1002 = 0.000000
tchdmj_2_1002 = 0.000000
tchdmj_3_1002 = 0.000000
trzdmj_1_1002 = 0.000000
trzdmj_2_1002 = 0.000000
trzdmj_3_1002 = 0.000000
tryjzhlmj_1_1002 = 0.000000
tryjzhlmj_2_1002 = 0.000000
tryjzhlmj_3_1002 = 0.000000
trphzmj_10_1002 = 0.000000
trphzmj_2a_1002 = 0.000000
trphzmj_2b_1002 = 0.000000
trphzmj_3a_1002 = 0.000000
trphzmj_3b_1002 = 0.000000
swdyxmj_1_1002 = 0.000000
swdyxmj_2_1002 = 0.000000
swdyxmj_3_1002 = 0.000000
trzjswrmj_1_1002 = 0.000000
trzjswrmj_2_1002 = 0.000000
trzjswrmj_3_1002 = 0.000000
szmj_1_1002 = 0.000000
szmj_2_1002 = 0.000000
szmj_3_1002 = 0.000000
gdejdlmj_j_1002 = 0.000000
gdejdlmj_g_1002 = 0.000000


pdmj_1_1003 = 0.000000
pdmj_2_1003 = 0.000000
pdmj_3_1003 = 0.000000
pdmj_4_1003 = 0.000000
pdmj_5_1003 = 0.000000
tchdmj_1_1003 = 0.000000
tchdmj_2_1003 = 0.000000
tchdmj_3_1003 = 0.000000
trzdmj_1_1003 = 0.000000
trzdmj_2_1003 = 0.000000
trzdmj_3_1003 = 0.000000
tryjzhlmj_1_1003 = 0.000000
tryjzhlmj_2_1003 = 0.000000
tryjzhlmj_3_1003 = 0.000000
trphzmj_10_1003 = 0.000000
trphzmj_2a_1003 = 0.000000
trphzmj_2b_1003 = 0.000000
trphzmj_3a_1003 = 0.000000
trphzmj_3b_1003 = 0.000000
swdyxmj_1_1003 = 0.000000
swdyxmj_2_1003 = 0.000000
swdyxmj_3_1003 = 0.000000
trzjswrmj_1_1003 = 0.000000
trzjswrmj_2_1003 = 0.000000
trzjswrmj_3_1003 = 0.000000
szmj_1_1003 = 0.000000
szmj_2_1003 = 0.000000
szmj_3_1003 = 0.000000
gdejdlmj_j_1003 = 0.000000
gdejdlmj_g_1003 = 0.000000

pdmj_1_1004 = 0.000000
pdmj_2_1004 = 0.000000
pdmj_3_1004 = 0.000000
pdmj_4_1004 = 0.000000
pdmj_5_1004 = 0.000000
tchdmj_1_1004 = 0.000000
tchdmj_2_1004 = 0.000000
tchdmj_3_1004 = 0.000000
trzdmj_1_1004 = 0.000000
trzdmj_2_1004 = 0.000000
trzdmj_3_1004 = 0.000000
tryjzhlmj_1_1004 = 0.000000
tryjzhlmj_2_1004 = 0.000000
tryjzhlmj_3_1004 = 0.000000
trphzmj_10_1004 = 0.000000
trphzmj_2a_1004 = 0.000000
trphzmj_2b_1004 = 0.000000
trphzmj_3a_1004 = 0.000000
trphzmj_3b_1004 = 0.000000
swdyxmj_1_1004 = 0.000000
swdyxmj_2_1004 = 0.000000
swdyxmj_3_1004 = 0.000000
trzjswrmj_1_1004 = 0.000000
trzjswrmj_2_1004 = 0.000000
trzjswrmj_3_1004 = 0.000000
szmj_1_1004 = 0.000000
szmj_2_1004 = 0.000000
szmj_3_1004 = 0.000000
gdejdlmj_j_1004 = 0.000000
gdejdlmj_g_1004 = 0.000000

pdmj_1_1005 = 0.000000
pdmj_2_1005 = 0.000000
pdmj_3_1005 = 0.000000
pdmj_4_1005 = 0.000000
pdmj_5_1005 = 0.000000
tchdmj_1_1005 = 0.000000
tchdmj_2_1005 = 0.000000
tchdmj_3_1005 = 0.000000
trzdmj_1_1005 = 0.000000
trzdmj_2_1005 = 0.000000
trzdmj_3_1005 = 0.000000
tryjzhlmj_1_1005 = 0.000000
tryjzhlmj_2_1005 = 0.000000
tryjzhlmj_3_1005 = 0.000000
trphzmj_10_1005 = 0.000000
trphzmj_2a_1005 = 0.000000
trphzmj_2b_1005 = 0.000000
trphzmj_3a_1005 = 0.000000
trphzmj_3b_1005 = 0.000000
swdyxmj_1_1005 = 0.000000
swdyxmj_2_1005 = 0.000000
swdyxmj_3_1005 = 0.000000
trzjswrmj_1_1005 = 0.000000
trzjswrmj_2_1005 = 0.000000
trzjswrmj_3_1005 = 0.000000
szmj_1_1005 = 0.000000
szmj_2_1005 = 0.000000
szmj_3_1005 = 0.000000
gdejdlmj_j_1005 = 0.000000
gdejdlmj_g_1005 = 0.000000

pdmj_1_1100 = 0.000000
pdmj_2_1100 = 0.000000
pdmj_3_1100 = 0.000000
pdmj_4_1100 = 0.000000
pdmj_5_1100 = 0.000000
tchdmj_1_1100 = 0.000000
tchdmj_2_1100 = 0.000000
tchdmj_3_1100 = 0.000000
trzdmj_1_1100 = 0.000000
trzdmj_2_1100 = 0.000000
trzdmj_3_1100 = 0.000000
tryjzhlmj_1_1100 = 0.000000
tryjzhlmj_2_1100 = 0.000000
tryjzhlmj_3_1100 = 0.000000
trphzmj_10_1100 = 0.000000
trphzmj_2a_1100 = 0.000000
trphzmj_2b_1100 = 0.000000
trphzmj_3a_1100 = 0.000000
trphzmj_3b_1100 = 0.000000
swdyxmj_1_1100 = 0.000000
swdyxmj_2_1100 = 0.000000
swdyxmj_3_1100 = 0.000000
trzjswrmj_1_1100 = 0.000000
trzjswrmj_2_1100 = 0.000000
trzjswrmj_3_1100 = 0.000000
szmj_1_1100 = 0.000000
szmj_2_1100 = 0.000000
szmj_3_1100 = 0.000000
gdejdlmj_j_1100 = 0.000000
gdejdlmj_g_1100 = 0.000000

pdmj_1_1101 = 0.000000
pdmj_2_1101 = 0.000000
pdmj_3_1101 = 0.000000
pdmj_4_1101 = 0.000000
pdmj_5_1101 = 0.000000
tchdmj_1_1101 = 0.000000
tchdmj_2_1101 = 0.000000
tchdmj_3_1101 = 0.000000
trzdmj_1_1101 = 0.000000
trzdmj_2_1101 = 0.000000
trzdmj_3_1101 = 0.000000
tryjzhlmj_1_1101 = 0.000000
tryjzhlmj_2_1101 = 0.000000
tryjzhlmj_3_1101 = 0.000000
trphzmj_10_1101 = 0.000000
trphzmj_2a_1101 = 0.000000
trphzmj_2b_1101 = 0.000000
trphzmj_3a_1101 = 0.000000
trphzmj_3b_1101 = 0.000000
swdyxmj_1_1101 = 0.000000
swdyxmj_2_1101 = 0.000000
swdyxmj_3_1101 = 0.000000
trzjswrmj_1_1101 = 0.000000
trzjswrmj_2_1101 = 0.000000
trzjswrmj_3_1101 = 0.000000
szmj_1_1101 = 0.000000
szmj_2_1101 = 0.000000
szmj_3_1101 = 0.000000
gdejdlmj_j_1101 = 0.000000
gdejdlmj_g_1101 = 0.000000

pdmj_1_1102 = 0.000000
pdmj_2_1102 = 0.000000
pdmj_3_1102 = 0.000000
pdmj_4_1102 = 0.000000
pdmj_5_1102 = 0.000000
tchdmj_1_1102 = 0.000000
tchdmj_2_1102 = 0.000000
tchdmj_3_1102 = 0.000000
trzdmj_1_1102 = 0.000000
trzdmj_2_1102 = 0.000000
trzdmj_3_1102 = 0.000000
tryjzhlmj_1_1102 = 0.000000
tryjzhlmj_2_1102 = 0.000000
tryjzhlmj_3_1102 = 0.000000
trphzmj_10_1102 = 0.000000
trphzmj_2a_1102 = 0.000000
trphzmj_2b_1102 = 0.000000
trphzmj_3a_1102 = 0.000000
trphzmj_3b_1102 = 0.000000
swdyxmj_1_1102 = 0.000000
swdyxmj_2_1102 = 0.000000
swdyxmj_3_1102 = 0.000000
trzjswrmj_1_1102 = 0.000000
trzjswrmj_2_1102 = 0.000000
trzjswrmj_3_1102 = 0.000000
szmj_1_1102 = 0.000000
szmj_2_1102 = 0.000000
szmj_3_1102 = 0.000000
gdejdlmj_j_1102 = 0.000000
gdejdlmj_g_1102 = 0.000000

pdmj_1_1103 = 0.000000
pdmj_2_1103 = 0.000000
pdmj_3_1103 = 0.000000
pdmj_4_1103 = 0.000000
pdmj_5_1103 = 0.000000
tchdmj_1_1103 = 0.000000
tchdmj_2_1103 = 0.000000
tchdmj_3_1103 = 0.000000
trzdmj_1_1103 = 0.000000
trzdmj_2_1103 = 0.000000
trzdmj_3_1103 = 0.000000
tryjzhlmj_1_1103 = 0.000000
tryjzhlmj_2_1103 = 0.000000
tryjzhlmj_3_1103 = 0.000000
trphzmj_10_1103 = 0.000000
trphzmj_2a_1103 = 0.000000
trphzmj_2b_1103 = 0.000000
trphzmj_3a_1103 = 0.000000
trphzmj_3b_1103 = 0.000000
swdyxmj_1_1103 = 0.000000
swdyxmj_2_1103 = 0.000000
swdyxmj_3_1103 = 0.000000
trzjswrmj_1_1103 = 0.000000
trzjswrmj_2_1103 = 0.000000
trzjswrmj_3_1103 = 0.000000
szmj_1_1103 = 0.000000
szmj_2_1103 = 0.000000
szmj_3_1103 = 0.000000
gdejdlmj_j_1103 = 0.000000
gdejdlmj_g_1103 = 0.000000

pdmj_1_1104 = 0.000000
pdmj_2_1104 = 0.000000
pdmj_3_1104 = 0.000000
pdmj_4_1104 = 0.000000
pdmj_5_1104 = 0.000000
tchdmj_1_1104 = 0.000000
tchdmj_2_1104 = 0.000000
tchdmj_3_1104 = 0.000000
trzdmj_1_1104 = 0.000000
trzdmj_2_1104 = 0.000000
trzdmj_3_1104 = 0.000000
tryjzhlmj_1_1104 = 0.000000
tryjzhlmj_2_1104 = 0.000000
tryjzhlmj_3_1104 = 0.000000
trphzmj_10_1104 = 0.000000
trphzmj_2a_1104 = 0.000000
trphzmj_2b_1104 = 0.000000
trphzmj_3a_1104 = 0.000000
trphzmj_3b_1104 = 0.000000
swdyxmj_1_1104 = 0.000000
swdyxmj_2_1104 = 0.000000
swdyxmj_3_1104 = 0.000000
trzjswrmj_1_1104 = 0.000000
trzjswrmj_2_1104 = 0.000000
trzjswrmj_3_1104 = 0.000000
szmj_1_1104 = 0.000000
szmj_2_1104 = 0.000000
szmj_3_1104 = 0.000000
gdejdlmj_j_1104 = 0.000000
gdejdlmj_g_1104 = 0.000000

pdmj_1_1105 = 0.000000
pdmj_2_1105 = 0.000000
pdmj_3_1105 = 0.000000
pdmj_4_1105 = 0.000000
pdmj_5_1105 = 0.000000
tchdmj_1_1105 = 0.000000
tchdmj_2_1105 = 0.000000
tchdmj_3_1105 = 0.000000
trzdmj_1_1105 = 0.000000
trzdmj_2_1105 = 0.000000
trzdmj_3_1105 = 0.000000
tryjzhlmj_1_1105 = 0.000000
tryjzhlmj_2_1105 = 0.000000
tryjzhlmj_3_1105 = 0.000000
trphzmj_10_1105 = 0.000000
trphzmj_2a_1105 = 0.000000
trphzmj_2b_1105 = 0.000000
trphzmj_3a_1105 = 0.000000
trphzmj_3b_1105 = 0.000000
swdyxmj_1_1105 = 0.000000
swdyxmj_2_1105 = 0.000000
swdyxmj_3_1105 = 0.000000
trzjswrmj_1_1105 = 0.000000
trzjswrmj_2_1105 = 0.000000
trzjswrmj_3_1105 = 0.000000
szmj_1_1105 = 0.000000
szmj_2_1105 = 0.000000
szmj_3_1105 = 0.000000
gdejdlmj_j_1105 = 0.000000
gdejdlmj_g_1105 = 0.000000

pdmj_1_1106 = 0.000000
pdmj_2_1106 = 0.000000
pdmj_3_1106 = 0.000000
pdmj_4_1106 = 0.000000
pdmj_5_1106 = 0.000000
tchdmj_1_1106 = 0.000000
tchdmj_2_1106 = 0.000000
tchdmj_3_1106 = 0.000000
trzdmj_1_1106 = 0.000000
trzdmj_2_1106 = 0.000000
trzdmj_3_1106 = 0.000000
tryjzhlmj_1_1106 = 0.000000
tryjzhlmj_2_1106 = 0.000000
tryjzhlmj_3_1106 = 0.000000
trphzmj_10_1106 = 0.000000
trphzmj_2a_1106 = 0.000000
trphzmj_2b_1106 = 0.000000
trphzmj_3a_1106 = 0.000000
trphzmj_3b_1106 = 0.000000
swdyxmj_1_1106 = 0.000000
swdyxmj_2_1106 = 0.000000
swdyxmj_3_1106 = 0.000000
trzjswrmj_1_1106 = 0.000000
trzjswrmj_2_1106 = 0.000000
trzjswrmj_3_1106 = 0.000000
szmj_1_1106 = 0.000000
szmj_2_1106 = 0.000000
szmj_3_1106 = 0.000000
gdejdlmj_j_1106 = 0.000000
gdejdlmj_g_1106 = 0.000000

pdmj_1_1107 = 0.000000
pdmj_2_1107 = 0.000000
pdmj_3_1107 = 0.000000
pdmj_4_1107 = 0.000000
pdmj_5_1107 = 0.000000
tchdmj_1_1107 = 0.000000
tchdmj_2_1107 = 0.000000
tchdmj_3_1107 = 0.000000
trzdmj_1_1107 = 0.000000
trzdmj_2_1107 = 0.000000
trzdmj_3_1107 = 0.000000
tryjzhlmj_1_1107 = 0.000000
tryjzhlmj_2_1107 = 0.000000
tryjzhlmj_3_1107 = 0.000000
trphzmj_10_1107 = 0.000000
trphzmj_2a_1107 = 0.000000
trphzmj_2b_1107 = 0.000000
trphzmj_3a_1107 = 0.000000
trphzmj_3b_1107 = 0.000000
swdyxmj_1_1107 = 0.000000
swdyxmj_2_1107 = 0.000000
swdyxmj_3_1107 = 0.000000
trzjswrmj_1_1107 = 0.000000
trzjswrmj_2_1107 = 0.000000
trzjswrmj_3_1107 = 0.000000
szmj_1_1107 = 0.000000
szmj_2_1107 = 0.000000
szmj_3_1107 = 0.000000
gdejdlmj_j_1107 = 0.000000
gdejdlmj_g_1107 = 0.000000

pdmj_1_1108 = 0.000000
pdmj_2_1108 = 0.000000
pdmj_3_1108 = 0.000000
pdmj_4_1108 = 0.000000
pdmj_5_1108 = 0.000000
tchdmj_1_1108 = 0.000000
tchdmj_2_1108 = 0.000000
tchdmj_3_1108 = 0.000000
trzdmj_1_1108 = 0.000000
trzdmj_2_1108 = 0.000000
trzdmj_3_1108 = 0.000000
tryjzhlmj_1_1108 = 0.000000
tryjzhlmj_2_1108 = 0.000000
tryjzhlmj_3_1108 = 0.000000
trphzmj_10_1108 = 0.000000
trphzmj_2a_1108 = 0.000000
trphzmj_2b_1108 = 0.000000
trphzmj_3a_1108 = 0.000000
trphzmj_3b_1108 = 0.000000
swdyxmj_1_1108 = 0.000000
swdyxmj_2_1108 = 0.000000
swdyxmj_3_1108 = 0.000000
trzjswrmj_1_1108 = 0.000000
trzjswrmj_2_1108 = 0.000000
trzjswrmj_3_1108 = 0.000000
szmj_1_1108 = 0.000000
szmj_2_1108 = 0.000000
szmj_3_1108 = 0.000000
gdejdlmj_j_1108 = 0.000000
gdejdlmj_g_1108 = 0.000000

pdmj_1_1109 = 0.000000
pdmj_2_1109 = 0.000000
pdmj_3_1109 = 0.000000
pdmj_4_1109 = 0.000000
pdmj_5_1109 = 0.000000
tchdmj_1_1109 = 0.000000
tchdmj_2_1109 = 0.000000
tchdmj_3_1109 = 0.000000
trzdmj_1_1109 = 0.000000
trzdmj_2_1109 = 0.000000
trzdmj_3_1109 = 0.000000
tryjzhlmj_1_1109 = 0.000000
tryjzhlmj_2_1109 = 0.000000
tryjzhlmj_3_1109 = 0.000000
trphzmj_10_1109 = 0.000000
trphzmj_2a_1109 = 0.000000
trphzmj_2b_1109 = 0.000000
trphzmj_3a_1109 = 0.000000
trphzmj_3b_1109 = 0.000000
swdyxmj_1_1109 = 0.000000
swdyxmj_2_1109 = 0.000000
swdyxmj_3_1109 = 0.000000
trzjswrmj_1_1109 = 0.000000
trzjswrmj_2_1109 = 0.000000
trzjswrmj_3_1109 = 0.000000
szmj_1_1109 = 0.000000
szmj_2_1109 = 0.000000
szmj_3_1109 = 0.000000
gdejdlmj_j_1109 = 0.000000
gdejdlmj_g_1109 = 0.000000

pdmj_1_1110 = 0.000000
pdmj_2_1110 = 0.000000
pdmj_3_1110 = 0.000000
pdmj_4_1110 = 0.000000
pdmj_5_1110 = 0.000000
tchdmj_1_1110 = 0.000000
tchdmj_2_1110 = 0.000000
tchdmj_3_1110 = 0.000000
trzdmj_1_1110 = 0.000000
trzdmj_2_1110 = 0.000000
trzdmj_3_1110 = 0.000000
tryjzhlmj_1_1110 = 0.000000
tryjzhlmj_2_1110 = 0.000000
tryjzhlmj_3_1110 = 0.000000
trphzmj_10_1110 = 0.000000
trphzmj_2a_1110 = 0.000000
trphzmj_2b_1110 = 0.000000
trphzmj_3a_1110 = 0.000000
trphzmj_3b_1110 = 0.000000
swdyxmj_1_1110 = 0.000000
swdyxmj_2_1110 = 0.000000
swdyxmj_3_1110 = 0.000000
trzjswrmj_1_1110 = 0.000000
trzjswrmj_2_1110 = 0.000000
trzjswrmj_3_1110 = 0.000000
szmj_1_1110 = 0.000000
szmj_2_1110 = 0.000000
szmj_3_1110 = 0.000000
gdejdlmj_j_1110 = 0.000000
gdejdlmj_g_1110 = 0.000000
with arcpy.da.SearchCursor(datapath,fileds) as cursor:
    for row in cursor:
        if row[1] == '1' :
            pdmj_1 = pdmj_1+row[0]
        elif row[1] == '2' :
            pdmj_2 = pdmj_2+row[0]
        elif row[1] == '3' :
            pdmj_3 = pdmj_3+row[0]
        elif row[1] == '4' :
            pdmj_4 = pdmj_4+row[0]
        elif row[1] == '5' :
            pdmj_5 = pdmj_5+row[0]

        if row[2] == '1' :
            tchdmj_1 = tchdmj_1+row[0]
        elif row[2] == '2' :
            tchdmj_2 = tchdmj_2+row[0]
        elif row[2] == '3' :
            tchdmj_3 = tchdmj_3+row[0]

        if row[3] == '1' :
            trzdmj_1 = trzdmj_1+row[0]
        elif row[3] == '2' :
            trzdmj_2 = trzdmj_2+row[0]
        elif row[3] == '3' :
            trzdmj_3 = trzdmj_3+row[0]

        if row[4] == '1' :
            tryjzhlmj_1 = tryjzhlmj_1+row[0]
        elif row[4] == '2' :
            tryjzhlmj_2 = tryjzhlmj_2+row[0]
        elif row[4] == '3' :
            tryjzhlmj_3 = tryjzhlmj_3+row[0]

        if row[5] == '10' :
            trphzmj_10 = trphzmj_10+row[0]
        elif row[5] == '2a' :
            trphzmj_2a = trphzmj_2a+row[0]
        elif row[5] == '2b' :
            trphzmj_2b = trphzmj_2b+row[0]
        elif row[5] == '3a' :
            trphzmj_3a = trphzmj_3a+row[0]
        elif row[5] == '3b' :
            trphzmj_3b = trphzmj_3b+row[0]

        if row[6] == '1' :
            swdyxmj_1 = swdyxmj_1+row[0]
        elif row[6] == '2' :
            swdyxmj_2 = swdyxmj_2+row[0]
        elif row[6] == '3' :
            swdyxmj_3 = swdyxmj_3+row[0]

        if row[7] == '1' :
            trzjswrmj_1 = trzjswrmj_1+row[0]
        elif row[7] == '2' :
            trzjswrmj_2 = trzjswrmj_2+row[0]
        elif row[7] == '3' :
            trzjswrmj_3 = trzjswrmj_3+row[0]

        if row[8] == '1' :
            szmj_1 = szmj_1+row[0]
        elif row[8] == '2' :
            szmj_2 = szmj_2+row[0]
        elif row[8] == '3' :
            szmj_3 = szmj_3+row[0]

        if row[9] == 'j' :
            gdejdlmj_j = gdejdlmj_j+row[0]
        elif row[9] == 'g' :
            gdejdlmj_g = gdejdlmj_g+row[0]
        if row[10] == '419001001' :
            if row[1] == '1' :
                pdmj_1_1001 = pdmj_1_1001+row[0]
            elif row[1] == '2' :
                pdmj_2_1001 = pdmj_2_1001+row[0]
            elif row[1] == '3' :
                pdmj_3_1001 = pdmj_3_1001+row[0]
            elif row[1] == '4' :
                pdmj_4_1001 = pdmj_4_1001+row[0]
            elif row[1] == '5' :
                pdmj_5_1001 = pdmj_5_1001+row[0]

            if row[2] == '1' :
                tchdmj_1_1001 = tchdmj_1_1001+row[0]
            elif row[2] == '2' :
                tchdmj_2_1001 = tchdmj_2_1001+row[0]
            elif row[2] == '3' :
                tchdmj_3_1001 = tchdmj_3_1001+row[0]

            if row[3] == '1' :
                trzdmj_1_1001 = trzdmj_1_1001+row[0]
            elif row[3] == '2' :
                trzdmj_2_1001 = trzdmj_2_1001+row[0]
            elif row[3] == '3' :
                trzdmj_3_1001 = trzdmj_3_1001+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1001 = tryjzhlmj_1_1001+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1001 = tryjzhlmj_2_1001+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1001 = tryjzhlmj_3_1001+row[0]

            if row[5] == '10' :
                trphzmj_10_1001 = trphzmj_10_1001+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1001 = trphzmj_2a_1001+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1001 = trphzmj_2b_1001+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1001 = trphzmj_3a_1001+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1001 = trphzmj_3b_1001+row[0]

            if row[6] == '1' :
                swdyxmj_1_1001 = swdyxmj_1_1001+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1001 = swdyxmj_2_1001+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1001 = swdyxmj_3_1001+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1001 = trzjswrmj_1_1001+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1001 = trzjswrmj_2_1001+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1001 = trzjswrmj_3_1001+row[0]

            if row[8] == '1' :
                szmj_1_1001 = szmj_1_1001+row[0]
            elif row[8] == '2' :
                szmj_2_1001 = szmj_2_1001+row[0]
            elif row[8] == '3' :
                szmj_3_1001 = szmj_3_1001+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1001 = gdejdlmj_j_1001+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1001 = gdejdlmj_g_1001+row[0]
        elif row[10] == '419001002' :
            if row[1] == '1' :
                pdmj_1_1002 = pdmj_1_1002+row[0]
            elif row[1] == '2' :
                pdmj_2_1002 = pdmj_2_1002+row[0]
            elif row[1] == '3' :
                pdmj_3_1002 = pdmj_3_1002+row[0]
            elif row[1] == '4' :
                pdmj_4_1002 = pdmj_4_1002+row[0]
            elif row[1] == '5' :
                pdmj_5_1002 = pdmj_5_1002+row[0]

            if row[2] == '1' :
                tchdmj_1_1002 = tchdmj_1_1002+row[0]
            elif row[2] == '2' :
                tchdmj_2_1002 = tchdmj_2_1002+row[0]
            elif row[2] == '3' :
                tchdmj_3_1002 = tchdmj_3_1002+row[0]

            if row[3] == '1' :
                trzdmj_1_1002 = trzdmj_1_1002+row[0]
            elif row[3] == '2' :
                trzdmj_2_1002 = trzdmj_2_1002+row[0]
            elif row[3] == '3' :
                trzdmj_3_1002 = trzdmj_3_1002+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1002 = tryjzhlmj_1_1002+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1002 = tryjzhlmj_2_1002+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1002 = tryjzhlmj_3_1002+row[0]

            if row[5] == '10' :
                trphzmj_10_1002 = trphzmj_10_1002+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1002 = trphzmj_2a_1002+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1002 = trphzmj_2b_1002+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1002 = trphzmj_3a_1002+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1002 = trphzmj_3b_1002+row[0]

            if row[6] == '1' :
                swdyxmj_1_1002 = swdyxmj_1_1002+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1002 = swdyxmj_2_1002+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1002 = swdyxmj_3_1002+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1002 = trzjswrmj_1_1002+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1002 = trzjswrmj_2_1002+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1002 = trzjswrmj_3_1002+row[0]

            if row[8] == '1' :
                szmj_1_1002 = szmj_1_1002+row[0]
            elif row[8] == '2' :
                szmj_2_1002 = szmj_2_1002+row[0]
            elif row[8] == '3' :
                szmj_3_1002 = szmj_3_1002+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1002 = gdejdlmj_j_1002+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1002 = gdejdlmj_g_1002+row[0]
        elif row[10] == '419001003' :
            if row[1] == '1' :
                pdmj_1_1003 = pdmj_1_1003+row[0]
            elif row[1] == '2' :
                pdmj_2_1003 = pdmj_2_1003+row[0]
            elif row[1] == '3' :
                pdmj_3_1003 = pdmj_3_1003+row[0]
            elif row[1] == '4' :
                pdmj_4_1003 = pdmj_4_1003+row[0]
            elif row[1] == '5' :
                pdmj_5_1003 = pdmj_5_1003+row[0]

            if row[2] == '1' :
                tchdmj_1_1003 = tchdmj_1_1003+row[0]
            elif row[2] == '2' :
                tchdmj_2_1003 = tchdmj_2_1003+row[0]
            elif row[2] == '3' :
                tchdmj_3_1003 = tchdmj_3_1003+row[0]

            if row[3] == '1' :
                trzdmj_1_1003 = trzdmj_1_1003+row[0]
            elif row[3] == '2' :
                trzdmj_2_1003 = trzdmj_2_1003+row[0]
            elif row[3] == '3' :
                trzdmj_3_1003 = trzdmj_3_1003+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1003 = tryjzhlmj_1_1003+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1003 = tryjzhlmj_2_1003+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1003 = tryjzhlmj_3_1003+row[0]

            if row[5] == '10' :
                trphzmj_10_1003 = trphzmj_10_1003+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1003 = trphzmj_2a_1003+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1003 = trphzmj_2b_1003+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1003 = trphzmj_3a_1003+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1003 = trphzmj_3b_1003+row[0]

            if row[6] == '1' :
                swdyxmj_1_1003 = swdyxmj_1_1003+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1003 = swdyxmj_2_1003+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1003 = swdyxmj_3_1003+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1003 = trzjswrmj_1_1003+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1003 = trzjswrmj_2_1003+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1003 = trzjswrmj_3_1003+row[0]

            if row[8] == '1' :
                szmj_1_1003 = szmj_1_1003+row[0]
            elif row[8] == '2' :
                szmj_2_1003 = szmj_2_1003+row[0]
            elif row[8] == '3' :
                szmj_3_1003 = szmj_3_1003+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1003 = gdejdlmj_j_1003+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1003 = gdejdlmj_g_1003+row[0]
        elif row[10] == '419001004' :
            if row[1] == '1' :
                pdmj_1_1004 = pdmj_1_1004+row[0]
            elif row[1] == '2' :
                pdmj_2_1004 = pdmj_2_1004+row[0]
            elif row[1] == '3' :
                pdmj_3_1004 = pdmj_3_1004+row[0]
            elif row[1] == '4' :
                pdmj_4_1004 = pdmj_4_1004+row[0]
            elif row[1] == '5' :
                pdmj_5_1004 = pdmj_5_1004+row[0]

            if row[2] == '1' :
                tchdmj_1_1004 = tchdmj_1_1004+row[0]
            elif row[2] == '2' :
                tchdmj_2_1004 = tchdmj_2_1004+row[0]
            elif row[2] == '3' :
                tchdmj_3_1004 = tchdmj_3_1004+row[0]

            if row[3] == '1' :
                trzdmj_1_1004 = trzdmj_1_1004+row[0]
            elif row[3] == '2' :
                trzdmj_2_1004 = trzdmj_2_1004+row[0]
            elif row[3] == '3' :
                trzdmj_3_1004 = trzdmj_3_1004+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1004 = tryjzhlmj_1_1004+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1004 = tryjzhlmj_2_1004+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1004 = tryjzhlmj_3_1004+row[0]

            if row[5] == '10' :
                trphzmj_10_1004 = trphzmj_10_1004+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1004 = trphzmj_2a_1004+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1004 = trphzmj_2b_1004+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1004 = trphzmj_3a_1004+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1004 = trphzmj_3b_1004+row[0]

            if row[6] == '1' :
                swdyxmj_1_1004 = swdyxmj_1_1004+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1004 = swdyxmj_2_1004+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1004 = swdyxmj_3_1004+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1004 = trzjswrmj_1_1004+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1004 = trzjswrmj_2_1004+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1004 = trzjswrmj_3_1004+row[0]

            if row[8] == '1' :
                szmj_1_1004 = szmj_1_1004+row[0]
            elif row[8] == '2' :
                szmj_2_1004 = szmj_2_1004+row[0]
            elif row[8] == '3' :
                szmj_3_1004 = szmj_3_1004+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1004 = gdejdlmj_j_1004+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1004 = gdejdlmj_g_1004+row[0]
        elif row[10] == '419001005' :
            if row[1] == '1' :
                pdmj_1_1005 = pdmj_1_1005+row[0]
            elif row[1] == '2' :
                pdmj_2_1005 = pdmj_2_1005+row[0]
            elif row[1] == '3' :
                pdmj_3_1005 = pdmj_3_1005+row[0]
            elif row[1] == '4' :
                pdmj_4_1005 = pdmj_4_1005+row[0]
            elif row[1] == '5' :
                pdmj_5_1005 = pdmj_5_1005+row[0]

            if row[2] == '1' :
                tchdmj_1_1005 = tchdmj_1_1005+row[0]
            elif row[2] == '2' :
                tchdmj_2_1005 = tchdmj_2_1005+row[0]
            elif row[2] == '3' :
                tchdmj_3_1005 = tchdmj_3_1005+row[0]

            if row[3] == '1' :
                trzdmj_1_1005 = trzdmj_1_1005+row[0]
            elif row[3] == '2' :
                trzdmj_2_1005 = trzdmj_2_1005+row[0]
            elif row[3] == '3' :
                trzdmj_3_1005 = trzdmj_3_1005+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1005 = tryjzhlmj_1_1005+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1005 = tryjzhlmj_2_1005+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1005 = tryjzhlmj_3_1005+row[0]

            if row[5] == '10' :
                trphzmj_10_1005 = trphzmj_10_1005+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1005 = trphzmj_2a_1005+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1005 = trphzmj_2b_1005+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1005 = trphzmj_3a_1005+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1005 = trphzmj_3b_1005+row[0]

            if row[6] == '1' :
                swdyxmj_1_1005 = swdyxmj_1_1005+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1005 = swdyxmj_2_1005+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1005 = swdyxmj_3_1005+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1005 = trzjswrmj_1_1005+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1005 = trzjswrmj_2_1005+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1005 = trzjswrmj_3_1005+row[0]

            if row[8] == '1' :
                szmj_1_1005 = szmj_1_1005+row[0]
            elif row[8] == '2' :
                szmj_2_1005 = szmj_2_1005+row[0]
            elif row[8] == '3' :
                szmj_3_1005 = szmj_3_1005+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1005 = gdejdlmj_j_1005+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1005 = gdejdlmj_g_1005+row[0]
        elif row[10] == '419001100' :
            if row[1] == '1' :
                pdmj_1_1100 = pdmj_1_1100+row[0]
            elif row[1] == '2' :
                pdmj_2_1100 = pdmj_2_1100+row[0]
            elif row[1] == '3' :
                pdmj_3_1100 = pdmj_3_1100+row[0]
            elif row[1] == '4' :
                pdmj_4_1100 = pdmj_4_1100+row[0]
            elif row[1] == '5' :
                pdmj_5_1100 = pdmj_5_1100+row[0]

            if row[2] == '1' :
                tchdmj_1_1100 = tchdmj_1_1100+row[0]
            elif row[2] == '2' :
                tchdmj_2_1100 = tchdmj_2_1100+row[0]
            elif row[2] == '3' :
                tchdmj_3_1100 = tchdmj_3_1100+row[0]

            if row[3] == '1' :
                trzdmj_1_1100 = trzdmj_1_1100+row[0]
            elif row[3] == '2' :
                trzdmj_2_1100 = trzdmj_2_1100+row[0]
            elif row[3] == '3' :
                trzdmj_3_1100 = trzdmj_3_1100+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1100 = tryjzhlmj_1_1100+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1100 = tryjzhlmj_2_1100+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1100 = tryjzhlmj_3_1100+row[0]

            if row[5] == '10' :
                trphzmj_10_1100 = trphzmj_10_1100+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1100 = trphzmj_2a_1100+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1100 = trphzmj_2b_1100+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1100 = trphzmj_3a_1100+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1100 = trphzmj_3b_1100+row[0]

            if row[6] == '1' :
                swdyxmj_1_1100 = swdyxmj_1_1100+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1100 = swdyxmj_2_1100+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1100 = swdyxmj_3_1100+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1100 = trzjswrmj_1_1100+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1100 = trzjswrmj_2_1100+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1100 = trzjswrmj_3_1100+row[0]

            if row[8] == '1' :
                szmj_1_1100 = szmj_1_1100+row[0]
            elif row[8] == '2' :
                szmj_2_1100 = szmj_2_1100+row[0]
            elif row[8] == '3' :
                szmj_3_1100 = szmj_3_1100+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1100 = gdejdlmj_j_1100+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1100 = gdejdlmj_g_1100+row[0]
        elif row[10] == '419001101' :
            if row[1] == '1' :
                pdmj_1_1101 = pdmj_1_1101+row[0]
            elif row[1] == '2' :
                pdmj_2_1101 = pdmj_2_1101+row[0]
            elif row[1] == '3' :
                pdmj_3_1101 = pdmj_3_1101+row[0]
            elif row[1] == '4' :
                pdmj_4_1101 = pdmj_4_1101+row[0]
            elif row[1] == '5' :
                pdmj_5_1101 = pdmj_5_1101+row[0]

            if row[2] == '1' :
                tchdmj_1_1101 = tchdmj_1_1101+row[0]
            elif row[2] == '2' :
                tchdmj_2_1101 = tchdmj_2_1101+row[0]
            elif row[2] == '3' :
                tchdmj_3_1101 = tchdmj_3_1101+row[0]

            if row[3] == '1' :
                trzdmj_1_1101 = trzdmj_1_1101+row[0]
            elif row[3] == '2' :
                trzdmj_2_1101 = trzdmj_2_1101+row[0]
            elif row[3] == '3' :
                trzdmj_3_1101 = trzdmj_3_1101+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1101 = tryjzhlmj_1_1101+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1101 = tryjzhlmj_2_1101+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1101 = tryjzhlmj_3_1101+row[0]

            if row[5] == '10' :
                trphzmj_10_1101 = trphzmj_10_1101+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1101 = trphzmj_2a_1101+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1101 = trphzmj_2b_1101+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1101 = trphzmj_3a_1101+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1101 = trphzmj_3b_1101+row[0]

            if row[6] == '1' :
                swdyxmj_1_1101 = swdyxmj_1_1101+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1101 = swdyxmj_2_1101+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1101 = swdyxmj_3_1101+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1101 = trzjswrmj_1_1101+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1101 = trzjswrmj_2_1101+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1101 = trzjswrmj_3_1101+row[0]

            if row[8] == '1' :
                szmj_1_1101 = szmj_1_1101+row[0]
            elif row[8] == '2' :
                szmj_2_1101 = szmj_2_1101+row[0]
            elif row[8] == '3' :
                szmj_3_1101 = szmj_3_1101+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1101 = gdejdlmj_j_1101+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1101 = gdejdlmj_g_1101+row[0]
        elif row[10] == '419001102' :
            if row[1] == '1' :
                pdmj_1_1102 = pdmj_1_1102+row[0]
            elif row[1] == '2' :
                pdmj_2_1102 = pdmj_2_1102+row[0]
            elif row[1] == '3' :
                pdmj_3_1102 = pdmj_3_1102+row[0]
            elif row[1] == '4' :
                pdmj_4_1102 = pdmj_4_1102+row[0]
            elif row[1] == '5' :
                pdmj_5_1102 = pdmj_5_1102+row[0]

            if row[2] == '1' :
                tchdmj_1_1102 = tchdmj_1_1102+row[0]
            elif row[2] == '2' :
                tchdmj_2_1102 = tchdmj_2_1102+row[0]
            elif row[2] == '3' :
                tchdmj_3_1102 = tchdmj_3_1102+row[0]

            if row[3] == '1' :
                trzdmj_1_1102 = trzdmj_1_1102+row[0]
            elif row[3] == '2' :
                trzdmj_2_1102 = trzdmj_2_1102+row[0]
            elif row[3] == '3' :
                trzdmj_3_1102 = trzdmj_3_1102+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1102 = tryjzhlmj_1_1102+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1102 = tryjzhlmj_2_1102+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1102 = tryjzhlmj_3_1102+row[0]

            if row[5] == '10' :
                trphzmj_10_1102 = trphzmj_10_1102+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1102 = trphzmj_2a_1102+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1102 = trphzmj_2b_1102+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1102 = trphzmj_3a_1102+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1102 = trphzmj_3b_1102+row[0]

            if row[6] == '1' :
                swdyxmj_1_1102 = swdyxmj_1_1102+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1102 = swdyxmj_2_1102+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1102 = swdyxmj_3_1102+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1102 = trzjswrmj_1_1102+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1102 = trzjswrmj_2_1102+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1102 = trzjswrmj_3_1102+row[0]

            if row[8] == '1' :
                szmj_1_1102 = szmj_1_1102+row[0]
            elif row[8] == '2' :
                szmj_2_1102 = szmj_2_1102+row[0]
            elif row[8] == '3' :
                szmj_3_1102 = szmj_3_1102+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1102 = gdejdlmj_j_1102+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1102 = gdejdlmj_g_1102+row[0]

        elif row[10] == '419001103' :
            if row[1] == '1' :
                pdmj_1_1103 = pdmj_1_1103+row[0]
            elif row[1] == '2' :
                pdmj_2_1103 = pdmj_2_1103+row[0]
            elif row[1] == '3' :
                pdmj_3_1103 = pdmj_3_1103+row[0]
            elif row[1] == '4' :
                pdmj_4_1103 = pdmj_4_1103+row[0]
            elif row[1] == '5' :
                pdmj_5_1103 = pdmj_5_1103+row[0]

            if row[2] == '1' :
                tchdmj_1_1103 = tchdmj_1_1103+row[0]
            elif row[2] == '2' :
                tchdmj_2_1103 = tchdmj_2_1103+row[0]
            elif row[2] == '3' :
                tchdmj_3_1103 = tchdmj_3_1103+row[0]

            if row[3] == '1' :
                trzdmj_1_1103 = trzdmj_1_1103+row[0]
            elif row[3] == '2' :
                trzdmj_2_1103 = trzdmj_2_1103+row[0]
            elif row[3] == '3' :
                trzdmj_3_1103 = trzdmj_3_1103+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1103 = tryjzhlmj_1_1103+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1103 = tryjzhlmj_2_1103+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1103 = tryjzhlmj_3_1103+row[0]

            if row[5] == '10' :
                trphzmj_10_1103 = trphzmj_10_1103+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1103 = trphzmj_2a_1103+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1103 = trphzmj_2b_1103+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1103 = trphzmj_3a_1103+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1103 = trphzmj_3b_1103+row[0]

            if row[6] == '1' :
                swdyxmj_1_1103 = swdyxmj_1_1103+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1103 = swdyxmj_2_1103+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1103 = swdyxmj_3_1103+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1103 = trzjswrmj_1_1103+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1103 = trzjswrmj_2_1103+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1103 = trzjswrmj_3_1103+row[0]

            if row[8] == '1' :
                szmj_1_1103 = szmj_1_1103+row[0]
            elif row[8] == '2' :
                szmj_2_1103 = szmj_2_1103+row[0]
            elif row[8] == '3' :
                szmj_3_1103 = szmj_3_1103+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1103 = gdejdlmj_j_1103+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1103 = gdejdlmj_g_1103+row[0]

        elif row[10] == '419001104' :
            if row[1] == '1' :
                pdmj_1_1104 = pdmj_1_1104+row[0]
            elif row[1] == '2' :
                pdmj_2_1104 = pdmj_2_1104+row[0]
            elif row[1] == '3' :
                pdmj_3_1104 = pdmj_3_1104+row[0]
            elif row[1] == '4' :
                pdmj_4_1104 = pdmj_4_1104+row[0]
            elif row[1] == '5' :
                pdmj_5_1104 = pdmj_5_1104+row[0]

            if row[2] == '1' :
                tchdmj_1_1104 = tchdmj_1_1104+row[0]
            elif row[2] == '2' :
                tchdmj_2_1104 = tchdmj_2_1104+row[0]
            elif row[2] == '3' :
                tchdmj_3_1104 = tchdmj_3_1104+row[0]

            if row[3] == '1' :
                trzdmj_1_1104 = trzdmj_1_1104+row[0]
            elif row[3] == '2' :
                trzdmj_2_1104 = trzdmj_2_1104+row[0]
            elif row[3] == '3' :
                trzdmj_3_1104 = trzdmj_3_1104+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1104 = tryjzhlmj_1_1104+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1104 = tryjzhlmj_2_1104+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1104 = tryjzhlmj_3_1104+row[0]

            if row[5] == '10' :
                trphzmj_10_1104 = trphzmj_10_1104+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1104 = trphzmj_2a_1104+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1104 = trphzmj_2b_1104+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1104 = trphzmj_3a_1104+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1104 = trphzmj_3b_1104+row[0]

            if row[6] == '1' :
                swdyxmj_1_1104 = swdyxmj_1_1104+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1104 = swdyxmj_2_1104+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1104 = swdyxmj_3_1104+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1104 = trzjswrmj_1_1104+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1104 = trzjswrmj_2_1104+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1104 = trzjswrmj_3_1104+row[0]

            if row[8] == '1' :
                szmj_1_1104 = szmj_1_1104+row[0]
            elif row[8] == '2' :
                szmj_2_1104 = szmj_2_1104+row[0]
            elif row[8] == '3' :
                szmj_3_1104 = szmj_3_1104+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1104 = gdejdlmj_j_1104+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1104 = gdejdlmj_g_1104+row[0]

        elif row[10] == '419001105' :
            if row[1] == '1' :
                pdmj_1_1105 = pdmj_1_1105+row[0]
            elif row[1] == '2' :
                pdmj_2_1105 = pdmj_2_1105+row[0]
            elif row[1] == '3' :
                pdmj_3_1105 = pdmj_3_1105+row[0]
            elif row[1] == '4' :
                pdmj_4_1105 = pdmj_4_1105+row[0]
            elif row[1] == '5' :
                pdmj_5_1105 = pdmj_5_1105+row[0]

            if row[2] == '1' :
                tchdmj_1_1105 = tchdmj_1_1105+row[0]
            elif row[2] == '2' :
                tchdmj_2_1105 = tchdmj_2_1105+row[0]
            elif row[2] == '3' :
                tchdmj_3_1105 = tchdmj_3_1105+row[0]

            if row[3] == '1' :
                trzdmj_1_1105 = trzdmj_1_1105+row[0]
            elif row[3] == '2' :
                trzdmj_2_1105 = trzdmj_2_1105+row[0]
            elif row[3] == '3' :
                trzdmj_3_1105 = trzdmj_3_1105+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1105 = tryjzhlmj_1_1105+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1105 = tryjzhlmj_2_1105+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1105 = tryjzhlmj_3_1105+row[0]

            if row[5] == '10' :
                trphzmj_10_1105 = trphzmj_10_1105+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1105 = trphzmj_2a_1105+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1105 = trphzmj_2b_1105+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1105 = trphzmj_3a_1105+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1105 = trphzmj_3b_1105+row[0]

            if row[6] == '1' :
                swdyxmj_1_1105 = swdyxmj_1_1105+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1105 = swdyxmj_2_1105+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1105 = swdyxmj_3_1105+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1105 = trzjswrmj_1_1105+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1105 = trzjswrmj_2_1105+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1105 = trzjswrmj_3_1105+row[0]

            if row[8] == '1' :
                szmj_1_1105 = szmj_1_1105+row[0]
            elif row[8] == '2' :
                szmj_2_1105 = szmj_2_1105+row[0]
            elif row[8] == '3' :
                szmj_3_1105 = szmj_3_1105+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1105 = gdejdlmj_j_1105+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1105 = gdejdlmj_g_1105+row[0]
        elif row[10] == '419001106' :
            if row[1] == '1' :
                pdmj_1_1106 = pdmj_1_1106+row[0]
            elif row[1] == '2' :
                pdmj_2_1106 = pdmj_2_1106+row[0]
            elif row[1] == '3' :
                pdmj_3_1106 = pdmj_3_1106+row[0]
            elif row[1] == '4' :
                pdmj_4_1106 = pdmj_4_1106+row[0]
            elif row[1] == '5' :
                pdmj_5_1106 = pdmj_5_1106+row[0]

            if row[2] == '1' :
                tchdmj_1_1106 = tchdmj_1_1106+row[0]
            elif row[2] == '2' :
                tchdmj_2_1106 = tchdmj_2_1106+row[0]
            elif row[2] == '3' :
                tchdmj_3_1106 = tchdmj_3_1106+row[0]

            if row[3] == '1' :
                trzdmj_1_1106 = trzdmj_1_1106+row[0]
            elif row[3] == '2' :
                trzdmj_2_1106 = trzdmj_2_1106+row[0]
            elif row[3] == '3' :
                trzdmj_3_1106 = trzdmj_3_1106+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1106 = tryjzhlmj_1_1106+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1106 = tryjzhlmj_2_1106+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1106 = tryjzhlmj_3_1106+row[0]

            if row[5] == '10' :
                trphzmj_10_1106 = trphzmj_10_1106+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1106 = trphzmj_2a_1106+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1106 = trphzmj_2b_1106+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1106 = trphzmj_3a_1106+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1106 = trphzmj_3b_1106+row[0]

            if row[6] == '1' :
                swdyxmj_1_1106 = swdyxmj_1_1106+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1106 = swdyxmj_2_1106+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1106 = swdyxmj_3_1106+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1106 = trzjswrmj_1_1106+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1106 = trzjswrmj_2_1106+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1106 = trzjswrmj_3_1106+row[0]

            if row[8] == '1' :
                szmj_1_1106 = szmj_1_1106+row[0]
            elif row[8] == '2' :
                szmj_2_1106 = szmj_2_1106+row[0]
            elif row[8] == '3' :
                szmj_3_1106 = szmj_3_1106+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1106 = gdejdlmj_j_1106+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1106 = gdejdlmj_g_1106+row[0]

        elif row[10] == '419001107' :
            if row[1] == '1' :
                pdmj_1_1107 = pdmj_1_1107+row[0]
            elif row[1] == '2' :
                pdmj_2_1107 = pdmj_2_1107+row[0]
            elif row[1] == '3' :
                pdmj_3_1107 = pdmj_3_1107+row[0]
            elif row[1] == '4' :
                pdmj_4_1107 = pdmj_4_1107+row[0]
            elif row[1] == '5' :
                pdmj_5_1107 = pdmj_5_1107+row[0]

            if row[2] == '1' :
                tchdmj_1_1107 = tchdmj_1_1107+row[0]
            elif row[2] == '2' :
                tchdmj_2_1107 = tchdmj_2_1107+row[0]
            elif row[2] == '3' :
                tchdmj_3_1107 = tchdmj_3_1107+row[0]

            if row[3] == '1' :
                trzdmj_1_1107 = trzdmj_1_1107+row[0]
            elif row[3] == '2' :
                trzdmj_2_1107 = trzdmj_2_1107+row[0]
            elif row[3] == '3' :
                trzdmj_3_1107 = trzdmj_3_1107+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1107 = tryjzhlmj_1_1107+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1107 = tryjzhlmj_2_1107+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1107 = tryjzhlmj_3_1107+row[0]

            if row[5] == '10' :
                trphzmj_10_1107 = trphzmj_10_1107+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1107 = trphzmj_2a_1107+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1107 = trphzmj_2b_1107+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1107 = trphzmj_3a_1107+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1107 = trphzmj_3b_1107+row[0]

            if row[6] == '1' :
                swdyxmj_1_1107 = swdyxmj_1_1107+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1107 = swdyxmj_2_1107+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1107 = swdyxmj_3_1107+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1107 = trzjswrmj_1_1107+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1107 = trzjswrmj_2_1107+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1107 = trzjswrmj_3_1107+row[0]

            if row[8] == '1' :
                szmj_1_1107 = szmj_1_1107+row[0]
            elif row[8] == '2' :
                szmj_2_1107 = szmj_2_1107+row[0]
            elif row[8] == '3' :
                szmj_3_1107 = szmj_3_1107+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1107 = gdejdlmj_j_1107+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1107 = gdejdlmj_g_1107+row[0]

        elif row[10] == '419001108' :
            if row[1] == '1' :
                pdmj_1_1108 = pdmj_1_1108+row[0]
            elif row[1] == '2' :
                pdmj_2_1108 = pdmj_2_1108+row[0]
            elif row[1] == '3' :
                pdmj_3_1108 = pdmj_3_1108+row[0]
            elif row[1] == '4' :
                pdmj_4_1108 = pdmj_4_1108+row[0]
            elif row[1] == '5' :
                pdmj_5_1108 = pdmj_5_1108+row[0]

            if row[2] == '1' :
                tchdmj_1_1108 = tchdmj_1_1108+row[0]
            elif row[2] == '2' :
                tchdmj_2_1108 = tchdmj_2_1108+row[0]
            elif row[2] == '3' :
                tchdmj_3_1108 = tchdmj_3_1108+row[0]

            if row[3] == '1' :
                trzdmj_1_1108 = trzdmj_1_1108+row[0]
            elif row[3] == '2' :
                trzdmj_2_1108 = trzdmj_2_1108+row[0]
            elif row[3] == '3' :
                trzdmj_3_1108 = trzdmj_3_1108+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1108 = tryjzhlmj_1_1108+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1108 = tryjzhlmj_2_1108+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1108 = tryjzhlmj_3_1108+row[0]

            if row[5] == '10' :
                trphzmj_10_1108 = trphzmj_10_1108+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1108 = trphzmj_2a_1108+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1108 = trphzmj_2b_1108+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1108 = trphzmj_3a_1108+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1108 = trphzmj_3b_1108+row[0]

            if row[6] == '1' :
                swdyxmj_1_1108 = swdyxmj_1_1108+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1108 = swdyxmj_2_1108+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1108 = swdyxmj_3_1108+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1108 = trzjswrmj_1_1108+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1108 = trzjswrmj_2_1108+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1108 = trzjswrmj_3_1108+row[0]

            if row[8] == '1' :
                szmj_1_1108 = szmj_1_1108+row[0]
            elif row[8] == '2' :
                szmj_2_1108 = szmj_2_1108+row[0]
            elif row[8] == '3' :
                szmj_3_1108 = szmj_3_1108+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1108 = gdejdlmj_j_1108+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1108 = gdejdlmj_g_1108+row[0]

        elif row[10] == '419001109' :
            if row[1] == '1' :
                pdmj_1_1109 = pdmj_1_1109+row[0]
            elif row[1] == '2' :
                pdmj_2_1109 = pdmj_2_1109+row[0]
            elif row[1] == '3' :
                pdmj_3_1109 = pdmj_3_1109+row[0]
            elif row[1] == '4' :
                pdmj_4_1109 = pdmj_4_1109+row[0]
            elif row[1] == '5' :
                pdmj_5_1109 = pdmj_5_1109+row[0]

            if row[2] == '1' :
                tchdmj_1_1109 = tchdmj_1_1109+row[0]
            elif row[2] == '2' :
                tchdmj_2_1109 = tchdmj_2_1109+row[0]
            elif row[2] == '3' :
                tchdmj_3_1109 = tchdmj_3_1109+row[0]

            if row[3] == '1' :
                trzdmj_1_1109 = trzdmj_1_1109+row[0]
            elif row[3] == '2' :
                trzdmj_2_1109 = trzdmj_2_1109+row[0]
            elif row[3] == '3' :
                trzdmj_3_1109 = trzdmj_3_1109+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1109 = tryjzhlmj_1_1109+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1109 = tryjzhlmj_2_1109+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1109 = tryjzhlmj_3_1109+row[0]

            if row[5] == '10' :
                trphzmj_10_1109 = trphzmj_10_1109+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1109 = trphzmj_2a_1109+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1109 = trphzmj_2b_1109+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1109 = trphzmj_3a_1109+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1109 = trphzmj_3b_1109+row[0]

            if row[6] == '1' :
                swdyxmj_1_1109 = swdyxmj_1_1109+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1109 = swdyxmj_2_1109+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1109 = swdyxmj_3_1109+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1109 = trzjswrmj_1_1109+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1109 = trzjswrmj_2_1109+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1109 = trzjswrmj_3_1109+row[0]

            if row[8] == '1' :
                szmj_1_1109 = szmj_1_1109+row[0]
            elif row[8] == '2' :
                szmj_2_1109 = szmj_2_1109+row[0]
            elif row[8] == '3' :
                szmj_3_1109 = szmj_3_1109+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1109 = gdejdlmj_j_1109+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1109 = gdejdlmj_g_1109+row[0]

        elif row[10] == '419001110' :
            if row[1] == '1' :
                pdmj_1_1110 = pdmj_1_1110+row[0]
            elif row[1] == '2' :
                pdmj_2_1110 = pdmj_2_1110+row[0]
            elif row[1] == '3' :
                pdmj_3_1110 = pdmj_3_1110+row[0]
            elif row[1] == '4' :
                pdmj_4_1110 = pdmj_4_1110+row[0]
            elif row[1] == '5' :
                pdmj_5_1110 = pdmj_5_1110+row[0]

            if row[2] == '1' :
                tchdmj_1_1110 = tchdmj_1_1110+row[0]
            elif row[2] == '2' :
                tchdmj_2_1110 = tchdmj_2_1110+row[0]
            elif row[2] == '3' :
                tchdmj_3_1110 = tchdmj_3_1110+row[0]

            if row[3] == '1' :
                trzdmj_1_1110 = trzdmj_1_1110+row[0]
            elif row[3] == '2' :
                trzdmj_2_1110 = trzdmj_2_1110+row[0]
            elif row[3] == '3' :
                trzdmj_3_1110 = trzdmj_3_1110+row[0]

            if row[4] == '1' :
                tryjzhlmj_1_1110 = tryjzhlmj_1_1110+row[0]
            elif row[4] == '2' :
                tryjzhlmj_2_1110 = tryjzhlmj_2_1110+row[0]
            elif row[4] == '3' :
                tryjzhlmj_3_1110 = tryjzhlmj_3_1110+row[0]

            if row[5] == '10' :
                trphzmj_10_1110 = trphzmj_10_1110+row[0]
            elif row[5] == '2a' :
                trphzmj_2a_1110 = trphzmj_2a_1110+row[0]
            elif row[5] == '2b' :
                trphzmj_2b_1110 = trphzmj_2b_1110+row[0]
            elif row[5] == '3a' :
                trphzmj_3a_1110 = trphzmj_3a_1110+row[0]
            elif row[5] == '3b' :
                trphzmj_3b_1110 = trphzmj_3b_1110+row[0]

            if row[6] == '1' :
                swdyxmj_1_1110 = swdyxmj_1_1110+row[0]
            elif row[6] == '2' :
                swdyxmj_2_1110 = swdyxmj_2_1110+row[0]
            elif row[6] == '3' :
                swdyxmj_3_1110 = swdyxmj_3_1110+row[0]

            if row[7] == '1' :
                trzjswrmj_1_1110 = trzjswrmj_1_1110+row[0]
            elif row[7] == '2' :
                trzjswrmj_2_1110 = trzjswrmj_2_1110+row[0]
            elif row[7] == '3' :
                trzjswrmj_3_1110 = trzjswrmj_3_1110+row[0]

            if row[8] == '1' :
                szmj_1_1110 = szmj_1_1110+row[0]
            elif row[8] == '2' :
                szmj_2_1110 = szmj_2_1110+row[0]
            elif row[8] == '3' :
                szmj_3_1110 = szmj_3_1110+row[0]

            if row[9] == 'j' :
                gdejdlmj_j_1110 = gdejdlmj_j_1110+row[0]
            elif row[9] == 'g' :
                gdejdlmj_g_1110 = gdejdlmj_g_1110+row[0]

all_jy = pdmj_1+pdmj_2+pdmj_3+pdmj_4+pdmj_5
all_1001 = pdmj_1_1001+pdmj_2_1001+pdmj_3_1001+pdmj_4_1001+pdmj_5_1001
all_1002 = pdmj_1_1002+pdmj_2_1002+pdmj_3_1002+pdmj_4_1002+pdmj_5_1002
all_1003 = pdmj_1_1003+pdmj_2_1003+pdmj_3_1003+pdmj_4_1003+pdmj_5_1003
all_1004 = pdmj_1_1004+pdmj_2_1004+pdmj_3_1004+pdmj_4_1004+pdmj_5_1004
all_1005 = pdmj_1_1005+pdmj_2_1005+pdmj_3_1005+pdmj_4_1005+pdmj_5_1005
all_1100 = pdmj_1_1100+pdmj_2_1100+pdmj_3_1100+pdmj_4_1100+pdmj_5_1100
all_1101 = pdmj_1_1101+pdmj_2_1101+pdmj_3_1101+pdmj_4_1101+pdmj_5_1101
all_1102 = pdmj_1_1102+pdmj_2_1102+pdmj_3_1102+pdmj_4_1102+pdmj_5_1102
all_1103 = pdmj_1_1103+pdmj_2_1103+pdmj_3_1103+pdmj_4_1103+pdmj_5_1103
all_1104 = pdmj_1_1104+pdmj_2_1104+pdmj_3_1104+pdmj_4_1104+pdmj_5_1104
all_1105 = pdmj_1_1105+pdmj_2_1105+pdmj_3_1105+pdmj_4_1105+pdmj_5_1105
all_1106 = pdmj_1_1106+pdmj_2_1106+pdmj_3_1106+pdmj_4_1106+pdmj_5_1106
all_1107 = pdmj_1_1107+pdmj_2_1107+pdmj_3_1107+pdmj_4_1107+pdmj_5_1107
all_1108 = pdmj_1_1108+pdmj_2_1108+pdmj_3_1108+pdmj_4_1108+pdmj_5_1108
all_1109 = pdmj_1_1109+pdmj_2_1109+pdmj_3_1109+pdmj_4_1109+pdmj_5_1109
all_1110 = pdmj_1_1110+pdmj_2_1110+pdmj_3_1110+pdmj_4_1110+pdmj_5_1110

workbook = xlwt.Workbook()
sheet = workbook.add_sheet('GD')
sheet.write(0,0,'419001')
sheet.write(0,1,u'济源市')
sheet.write(0,2,all_jy)
sheet.write(0,3,pdmj_1)
sheet.write(0,4,pdmj_2)
sheet.write(0,5,pdmj_3)
sheet.write(0,6,pdmj_4)
sheet.write(0,7,pdmj_5)
sheet.write(0,8,tchdmj_1)
sheet.write(0,9,tchdmj_2)
sheet.write(0,10,tchdmj_3)
sheet.write(0,11,trzdmj_1)
sheet.write(0,12,trzdmj_2)
sheet.write(0,13,trzdmj_3)
sheet.write(0,14,tryjzhlmj_1)
sheet.write(0,15,tryjzhlmj_2)
sheet.write(0,16,tryjzhlmj_3)
sheet.write(0,17,trphzmj_10)
sheet.write(0,18,trphzmj_2a)
sheet.write(0,19,trphzmj_2b)
sheet.write(0,20,trphzmj_3a)
sheet.write(0,21,trphzmj_3b)
sheet.write(0,22,swdyxmj_1)
sheet.write(0,23,swdyxmj_2)
sheet.write(0,24,swdyxmj_3)
sheet.write(0,25,trzjswrmj_1)
sheet.write(0,26,trzjswrmj_2)
sheet.write(0,27,trzjswrmj_3)
sheet.write(0,28,szmj_1)
sheet.write(0,29,szmj_2)
sheet.write(0,30,szmj_3)
sheet.write(0,31,gdejdlmj_j)
sheet.write(0,32,gdejdlmj_g)
sheet.write(1,0,'419001001')
sheet.write(1,1,u'沁园街道直属')
sheet.write(1,2,all_1001)
sheet.write(1,3,pdmj_1_1001)
sheet.write(1,4,pdmj_2_1001)
sheet.write(1,5,pdmj_3_1001)
sheet.write(1,6,pdmj_4_1001)
sheet.write(1,7,pdmj_5_1001)
sheet.write(1,8,tchdmj_1_1001)
sheet.write(1,9,tchdmj_2_1001)
sheet.write(1,10,tchdmj_3_1001)
sheet.write(1,11,trzdmj_1_1001)
sheet.write(1,12,trzdmj_2_1001)
sheet.write(1,13,trzdmj_3_1001)
sheet.write(1,14,tryjzhlmj_1_1001)
sheet.write(1,15,tryjzhlmj_2_1001)
sheet.write(1,16,tryjzhlmj_3_1001)
sheet.write(1,17,trphzmj_10_1001)
sheet.write(1,18,trphzmj_2a_1001)
sheet.write(1,19,trphzmj_2b_1001)
sheet.write(1,20,trphzmj_3a_1001)
sheet.write(1,21,trphzmj_3b_1001)
sheet.write(1,22,swdyxmj_1_1001)
sheet.write(1,23,swdyxmj_2_1001)
sheet.write(1,24,swdyxmj_3_1001)
sheet.write(1,25,trzjswrmj_1_1001)
sheet.write(1,26,trzjswrmj_2_1001)
sheet.write(1,27,trzjswrmj_3_1001)
sheet.write(1,28,szmj_1_1001)
sheet.write(1,29,szmj_2_1001)
sheet.write(1,30,szmj_3_1001)
sheet.write(1,31,gdejdlmj_j_1001)
sheet.write(1,32,gdejdlmj_g_1001)

sheet.write(2,0,'419001002')
sheet.write(2,1,u'济水街道直属')
sheet.write(2,2,all_1002)
sheet.write(2,3,pdmj_1_1002)
sheet.write(2,4,pdmj_2_1002)
sheet.write(2,5,pdmj_3_1002)
sheet.write(2,6,pdmj_4_1002)
sheet.write(2,7,pdmj_5_1002)
sheet.write(2,8,tchdmj_1_1002)
sheet.write(2,9,tchdmj_2_1002)
sheet.write(2,10,tchdmj_3_1002)
sheet.write(2,11,trzdmj_1_1002)
sheet.write(2,12,trzdmj_2_1002)
sheet.write(2,13,trzdmj_3_1002)
sheet.write(2,14,tryjzhlmj_1_1002)
sheet.write(2,15,tryjzhlmj_2_1002)
sheet.write(2,16,tryjzhlmj_3_1002)
sheet.write(2,17,trphzmj_10_1002)
sheet.write(2,18,trphzmj_2a_1002)
sheet.write(2,19,trphzmj_2b_1002)
sheet.write(2,20,trphzmj_3a_1002)
sheet.write(2,21,trphzmj_3b_1002)
sheet.write(2,22,swdyxmj_1_1002)
sheet.write(2,23,swdyxmj_2_1002)
sheet.write(2,24,swdyxmj_3_1002)
sheet.write(2,25,trzjswrmj_1_1002)
sheet.write(2,26,trzjswrmj_2_1002)
sheet.write(2,27,trzjswrmj_3_1002)
sheet.write(2,28,szmj_1_1002)
sheet.write(2,29,szmj_2_1002)
sheet.write(2,30,szmj_3_1002)
sheet.write(2,31,gdejdlmj_j_1002)
sheet.write(2,32,gdejdlmj_g_1002)

sheet.write(3,0,'419001003')
sheet.write(3,1,u'北海街道直属')
sheet.write(3,2,all_1003)
sheet.write(3,3,pdmj_1_1003)
sheet.write(3,4,pdmj_2_1003)
sheet.write(3,5,pdmj_3_1003)
sheet.write(3,6,pdmj_4_1003)
sheet.write(3,7,pdmj_5_1003)
sheet.write(3,8,tchdmj_1_1003)
sheet.write(3,9,tchdmj_2_1003)
sheet.write(3,10,tchdmj_3_1003)
sheet.write(3,11,trzdmj_1_1003)
sheet.write(3,12,trzdmj_2_1003)
sheet.write(3,13,trzdmj_3_1003)
sheet.write(3,14,tryjzhlmj_1_1003)
sheet.write(3,15,tryjzhlmj_2_1003)
sheet.write(3,16,tryjzhlmj_3_1003)
sheet.write(3,17,trphzmj_10_1003)
sheet.write(3,18,trphzmj_2a_1003)
sheet.write(3,19,trphzmj_2b_1003)
sheet.write(3,20,trphzmj_3a_1003)
sheet.write(3,21,trphzmj_3b_1003)
sheet.write(3,22,swdyxmj_1_1003)
sheet.write(3,23,swdyxmj_2_1003)
sheet.write(3,24,swdyxmj_3_1003)
sheet.write(3,25,trzjswrmj_1_1003)
sheet.write(3,26,trzjswrmj_2_1003)
sheet.write(3,27,trzjswrmj_3_1003)
sheet.write(3,28,szmj_1_1003)
sheet.write(3,29,szmj_2_1003)
sheet.write(3,30,szmj_3_1003)
sheet.write(3,31,gdejdlmj_j_1003)
sheet.write(3,32,gdejdlmj_g_1003)

sheet.write(4,0,'419001004')
sheet.write(4,1,u'天坛街道直属')
sheet.write(4,2,all_1004)
sheet.write(4,3,pdmj_1_1004)
sheet.write(4,4,pdmj_2_1004)
sheet.write(4,5,pdmj_3_1004)
sheet.write(4,6,pdmj_4_1004)
sheet.write(4,7,pdmj_5_1004)
sheet.write(4,8,tchdmj_1_1004)
sheet.write(4,9,tchdmj_2_1004)
sheet.write(4,10,tchdmj_3_1004)
sheet.write(4,11,trzdmj_1_1004)
sheet.write(4,12,trzdmj_2_1004)
sheet.write(4,13,trzdmj_3_1004)
sheet.write(4,14,tryjzhlmj_1_1004)
sheet.write(4,15,tryjzhlmj_2_1004)
sheet.write(4,16,tryjzhlmj_3_1004)
sheet.write(4,17,trphzmj_10_1004)
sheet.write(4,18,trphzmj_2a_1004)
sheet.write(4,19,trphzmj_2b_1004)
sheet.write(4,20,trphzmj_3a_1004)
sheet.write(4,21,trphzmj_3b_1004)
sheet.write(4,22,swdyxmj_1_1004)
sheet.write(4,23,swdyxmj_2_1004)
sheet.write(4,24,swdyxmj_3_1004)
sheet.write(4,25,trzjswrmj_1_1004)
sheet.write(4,26,trzjswrmj_2_1004)
sheet.write(4,27,trzjswrmj_3_1004)
sheet.write(4,28,szmj_1_1004)
sheet.write(4,29,szmj_2_1004)
sheet.write(4,30,szmj_3_1004)
sheet.write(4,31,gdejdlmj_j_1004)
sheet.write(4,32,gdejdlmj_g_1004)

sheet.write(5,0,'419001005')
sheet.write(5,1,u'玉泉街道直属')
sheet.write(5,2,all_1005)
sheet.write(5,3,pdmj_1_1005)
sheet.write(5,4,pdmj_2_1005)
sheet.write(5,5,pdmj_3_1005)
sheet.write(5,6,pdmj_4_1005)
sheet.write(5,7,pdmj_5_1005)
sheet.write(5,8,tchdmj_1_1005)
sheet.write(5,9,tchdmj_2_1005)
sheet.write(5,10,tchdmj_3_1005)
sheet.write(5,11,trzdmj_1_1005)
sheet.write(5,12,trzdmj_2_1005)
sheet.write(5,13,trzdmj_3_1005)
sheet.write(5,14,tryjzhlmj_1_1005)
sheet.write(5,15,tryjzhlmj_2_1005)
sheet.write(5,16,tryjzhlmj_3_1005)
sheet.write(5,17,trphzmj_10_1005)
sheet.write(5,18,trphzmj_2a_1005)
sheet.write(5,19,trphzmj_2b_1005)
sheet.write(5,20,trphzmj_3a_1005)
sheet.write(5,21,trphzmj_3b_1005)
sheet.write(5,22,swdyxmj_1_1005)
sheet.write(5,23,swdyxmj_2_1005)
sheet.write(5,24,swdyxmj_3_1005)
sheet.write(5,25,trzjswrmj_1_1005)
sheet.write(5,26,trzjswrmj_2_1005)
sheet.write(5,27,trzjswrmj_3_1005)
sheet.write(5,28,szmj_1_1005)
sheet.write(5,29,szmj_2_1005)
sheet.write(5,30,szmj_3_1005)
sheet.write(5,31,gdejdlmj_j_1005)
sheet.write(5,32,gdejdlmj_g_1005)

sheet.write(6,0,'419001100')
sheet.write(6,1,u'克井镇')
sheet.write(6,2,all_1100)
sheet.write(6,3,pdmj_1_1100)
sheet.write(6,4,pdmj_2_1100)
sheet.write(6,5,pdmj_3_1100)
sheet.write(6,6,pdmj_4_1100)
sheet.write(6,7,pdmj_5_1100)
sheet.write(6,8,tchdmj_1_1100)
sheet.write(6,9,tchdmj_2_1100)
sheet.write(6,10,tchdmj_3_1100)
sheet.write(6,11,trzdmj_1_1100)
sheet.write(6,12,trzdmj_2_1100)
sheet.write(6,13,trzdmj_3_1100)
sheet.write(6,14,tryjzhlmj_1_1100)
sheet.write(6,15,tryjzhlmj_2_1100)
sheet.write(6,16,tryjzhlmj_3_1100)
sheet.write(6,17,trphzmj_10_1100)
sheet.write(6,18,trphzmj_2a_1100)
sheet.write(6,19,trphzmj_2b_1100)
sheet.write(6,20,trphzmj_3a_1100)
sheet.write(6,21,trphzmj_3b_1100)
sheet.write(6,22,swdyxmj_1_1100)
sheet.write(6,23,swdyxmj_2_1100)
sheet.write(6,24,swdyxmj_3_1100)
sheet.write(6,25,trzjswrmj_1_1100)
sheet.write(6,26,trzjswrmj_2_1100)
sheet.write(6,27,trzjswrmj_3_1100)
sheet.write(6,28,szmj_1_1100)
sheet.write(6,29,szmj_2_1100)
sheet.write(6,30,szmj_3_1100)
sheet.write(6,31,gdejdlmj_j_1100)
sheet.write(6,32,gdejdlmj_g_1100)

sheet.write(7,0,'419001101')
sheet.write(7,1,u'五龙口镇')
sheet.write(7,2,all_1101)
sheet.write(7,3,pdmj_1_1101)
sheet.write(7,4,pdmj_2_1101)
sheet.write(7,5,pdmj_3_1101)
sheet.write(7,6,pdmj_4_1101)
sheet.write(7,7,pdmj_5_1101)
sheet.write(7,8,tchdmj_1_1101)
sheet.write(7,9,tchdmj_2_1101)
sheet.write(7,10,tchdmj_3_1101)
sheet.write(7,11,trzdmj_1_1101)
sheet.write(7,12,trzdmj_2_1101)
sheet.write(7,13,trzdmj_3_1101)
sheet.write(7,14,tryjzhlmj_1_1101)
sheet.write(7,15,tryjzhlmj_2_1101)
sheet.write(7,16,tryjzhlmj_3_1101)
sheet.write(7,17,trphzmj_10_1101)
sheet.write(7,18,trphzmj_2a_1101)
sheet.write(7,19,trphzmj_2b_1101)
sheet.write(7,20,trphzmj_3a_1101)
sheet.write(7,21,trphzmj_3b_1101)
sheet.write(7,22,swdyxmj_1_1101)
sheet.write(7,23,swdyxmj_2_1101)
sheet.write(7,24,swdyxmj_3_1101)
sheet.write(7,25,trzjswrmj_1_1101)
sheet.write(7,26,trzjswrmj_2_1101)
sheet.write(7,27,trzjswrmj_3_1101)
sheet.write(7,28,szmj_1_1101)
sheet.write(7,29,szmj_2_1101)
sheet.write(7,30,szmj_3_1101)
sheet.write(7,31,gdejdlmj_j_1101)
sheet.write(7,32,gdejdlmj_g_1101)

sheet.write(8,0,'419001102')
sheet.write(8,1,u'只城镇')
sheet.write(8,2,all_1102)
sheet.write(8,3,pdmj_1_1102)
sheet.write(8,4,pdmj_2_1102)
sheet.write(8,5,pdmj_3_1102)
sheet.write(8,6,pdmj_4_1102)
sheet.write(8,7,pdmj_5_1102)
sheet.write(8,8,tchdmj_1_1102)
sheet.write(8,9,tchdmj_2_1102)
sheet.write(8,10,tchdmj_3_1102)
sheet.write(8,11,trzdmj_1_1102)
sheet.write(8,12,trzdmj_2_1102)
sheet.write(8,13,trzdmj_3_1102)
sheet.write(8,14,tryjzhlmj_1_1102)
sheet.write(8,15,tryjzhlmj_2_1102)
sheet.write(8,16,tryjzhlmj_3_1102)
sheet.write(8,17,trphzmj_10_1102)
sheet.write(8,18,trphzmj_2a_1102)
sheet.write(8,19,trphzmj_2b_1102)
sheet.write(8,20,trphzmj_3a_1102)
sheet.write(8,21,trphzmj_3b_1102)
sheet.write(8,22,swdyxmj_1_1102)
sheet.write(8,23,swdyxmj_2_1102)
sheet.write(8,24,swdyxmj_3_1102)
sheet.write(8,25,trzjswrmj_1_1102)
sheet.write(8,26,trzjswrmj_2_1102)
sheet.write(8,27,trzjswrmj_3_1102)
sheet.write(8,28,szmj_1_1102)
sheet.write(8,29,szmj_2_1102)
sheet.write(8,30,szmj_3_1102)
sheet.write(8,31,gdejdlmj_j_1102)
sheet.write(8,32,gdejdlmj_g_1102)

sheet.write(9,0,'419001103')
sheet.write(9,1,u'承留镇')
sheet.write(9,2,all_1103)
sheet.write(9,3,pdmj_1_1103)
sheet.write(9,4,pdmj_2_1103)
sheet.write(9,5,pdmj_3_1103)
sheet.write(9,6,pdmj_4_1103)
sheet.write(9,7,pdmj_5_1103)
sheet.write(9,8,tchdmj_1_1103)
sheet.write(9,9,tchdmj_2_1103)
sheet.write(9,10,tchdmj_3_1103)
sheet.write(9,11,trzdmj_1_1103)
sheet.write(9,12,trzdmj_2_1103)
sheet.write(9,13,trzdmj_3_1103)
sheet.write(9,14,tryjzhlmj_1_1103)
sheet.write(9,15,tryjzhlmj_2_1103)
sheet.write(9,16,tryjzhlmj_3_1103)
sheet.write(9,17,trphzmj_10_1103)
sheet.write(9,18,trphzmj_2a_1103)
sheet.write(9,19,trphzmj_2b_1103)
sheet.write(9,20,trphzmj_3a_1103)
sheet.write(9,21,trphzmj_3b_1103)
sheet.write(9,22,swdyxmj_1_1103)
sheet.write(9,23,swdyxmj_2_1103)
sheet.write(9,24,swdyxmj_3_1103)
sheet.write(9,25,trzjswrmj_1_1103)
sheet.write(9,26,trzjswrmj_2_1103)
sheet.write(9,27,trzjswrmj_3_1103)
sheet.write(9,28,szmj_1_1103)
sheet.write(9,29,szmj_2_1103)
sheet.write(9,30,szmj_3_1103)
sheet.write(9,31,gdejdlmj_j_1103)
sheet.write(9,32,gdejdlmj_g_1103)

sheet.write(10,0,'419001104')
sheet.write(10,1,u'邵原镇')
sheet.write(10,2,all_1104)
sheet.write(10,3,pdmj_1_1104)
sheet.write(10,4,pdmj_2_1104)
sheet.write(10,5,pdmj_3_1104)
sheet.write(10,6,pdmj_4_1104)
sheet.write(10,7,pdmj_5_1104)
sheet.write(10,8,tchdmj_1_1104)
sheet.write(10,9,tchdmj_2_1104)
sheet.write(10,10,tchdmj_3_1104)
sheet.write(10,11,trzdmj_1_1104)
sheet.write(10,12,trzdmj_2_1104)
sheet.write(10,13,trzdmj_3_1104)
sheet.write(10,14,tryjzhlmj_1_1104)
sheet.write(10,15,tryjzhlmj_2_1104)
sheet.write(10,16,tryjzhlmj_3_1104)
sheet.write(10,17,trphzmj_10_1104)
sheet.write(10,18,trphzmj_2a_1104)
sheet.write(10,19,trphzmj_2b_1104)
sheet.write(10,20,trphzmj_3a_1104)
sheet.write(10,21,trphzmj_3b_1104)
sheet.write(10,22,swdyxmj_1_1104)
sheet.write(10,23,swdyxmj_2_1104)
sheet.write(10,24,swdyxmj_3_1104)
sheet.write(10,25,trzjswrmj_1_1104)
sheet.write(10,26,trzjswrmj_2_1104)
sheet.write(10,27,trzjswrmj_3_1104)
sheet.write(10,28,szmj_1_1104)
sheet.write(10,29,szmj_2_1104)
sheet.write(10,30,szmj_3_1104)
sheet.write(10,31,gdejdlmj_j_1104)
sheet.write(10,32,gdejdlmj_g_1104)

sheet.write(11,0,'419001105')
sheet.write(11,1,u'坡头镇')
sheet.write(11,2,all_1105)
sheet.write(11,3,pdmj_1_1105)
sheet.write(11,4,pdmj_2_1105)
sheet.write(11,5,pdmj_3_1105)
sheet.write(11,6,pdmj_4_1105)
sheet.write(11,7,pdmj_5_1105)
sheet.write(11,8,tchdmj_1_1105)
sheet.write(11,9,tchdmj_2_1105)
sheet.write(11,10,tchdmj_3_1105)
sheet.write(11,11,trzdmj_1_1105)
sheet.write(11,12,trzdmj_2_1105)
sheet.write(11,13,trzdmj_3_1105)
sheet.write(11,14,tryjzhlmj_1_1105)
sheet.write(11,15,tryjzhlmj_2_1105)
sheet.write(11,16,tryjzhlmj_3_1105)
sheet.write(11,17,trphzmj_10_1105)
sheet.write(11,18,trphzmj_2a_1105)
sheet.write(11,19,trphzmj_2b_1105)
sheet.write(11,20,trphzmj_3a_1105)
sheet.write(11,21,trphzmj_3b_1105)
sheet.write(11,22,swdyxmj_1_1105)
sheet.write(11,23,swdyxmj_2_1105)
sheet.write(11,24,swdyxmj_3_1105)
sheet.write(11,25,trzjswrmj_1_1105)
sheet.write(11,26,trzjswrmj_2_1105)
sheet.write(11,27,trzjswrmj_3_1105)
sheet.write(11,28,szmj_1_1105)
sheet.write(11,29,szmj_2_1105)
sheet.write(11,30,szmj_3_1105)
sheet.write(11,31,gdejdlmj_j_1105)
sheet.write(11,32,gdejdlmj_g_1105)

sheet.write(12,0,'419001106')
sheet.write(12,1,u'梨林镇')
sheet.write(12,2,all_1106)
sheet.write(12,3,pdmj_1_1106)
sheet.write(12,4,pdmj_2_1106)
sheet.write(12,5,pdmj_3_1106)
sheet.write(12,6,pdmj_4_1106)
sheet.write(12,7,pdmj_5_1106)
sheet.write(12,8,tchdmj_1_1106)
sheet.write(12,9,tchdmj_2_1106)
sheet.write(12,10,tchdmj_3_1106)
sheet.write(12,11,trzdmj_1_1106)
sheet.write(12,12,trzdmj_2_1106)
sheet.write(12,13,trzdmj_3_1106)
sheet.write(12,14,tryjzhlmj_1_1106)
sheet.write(12,15,tryjzhlmj_2_1106)
sheet.write(12,16,tryjzhlmj_3_1106)
sheet.write(12,17,trphzmj_10_1106)
sheet.write(12,18,trphzmj_2a_1106)
sheet.write(12,19,trphzmj_2b_1106)
sheet.write(12,20,trphzmj_3a_1106)
sheet.write(12,21,trphzmj_3b_1106)
sheet.write(12,22,swdyxmj_1_1106)
sheet.write(12,23,swdyxmj_2_1106)
sheet.write(12,24,swdyxmj_3_1106)
sheet.write(12,25,trzjswrmj_1_1106)
sheet.write(12,26,trzjswrmj_2_1106)
sheet.write(12,27,trzjswrmj_3_1106)
sheet.write(12,28,szmj_1_1106)
sheet.write(12,29,szmj_2_1106)
sheet.write(12,30,szmj_3_1106)
sheet.write(12,31,gdejdlmj_j_1106)
sheet.write(12,32,gdejdlmj_g_1106)

sheet.write(13,0,'419001107')
sheet.write(13,1,u'大峪镇')
sheet.write(13,2,all_1107)
sheet.write(13,3,pdmj_1_1107)
sheet.write(13,4,pdmj_2_1107)
sheet.write(13,5,pdmj_3_1107)
sheet.write(13,6,pdmj_4_1107)
sheet.write(13,7,pdmj_5_1107)
sheet.write(13,8,tchdmj_1_1107)
sheet.write(13,9,tchdmj_2_1107)
sheet.write(13,10,tchdmj_3_1107)
sheet.write(13,11,trzdmj_1_1107)
sheet.write(13,12,trzdmj_2_1107)
sheet.write(13,13,trzdmj_3_1107)
sheet.write(13,14,tryjzhlmj_1_1107)
sheet.write(13,15,tryjzhlmj_2_1107)
sheet.write(13,16,tryjzhlmj_3_1107)
sheet.write(13,17,trphzmj_10_1107)
sheet.write(13,18,trphzmj_2a_1107)
sheet.write(13,19,trphzmj_2b_1107)
sheet.write(13,20,trphzmj_3a_1107)
sheet.write(13,21,trphzmj_3b_1107)
sheet.write(13,22,swdyxmj_1_1107)
sheet.write(13,23,swdyxmj_2_1107)
sheet.write(13,24,swdyxmj_3_1107)
sheet.write(13,25,trzjswrmj_1_1107)
sheet.write(13,26,trzjswrmj_2_1107)
sheet.write(13,27,trzjswrmj_3_1107)
sheet.write(13,28,szmj_1_1107)
sheet.write(13,29,szmj_2_1107)
sheet.write(13,30,szmj_3_1107)
sheet.write(13,31,gdejdlmj_j_1107)
sheet.write(13,32,gdejdlmj_g_1107)

sheet.write(14,0,'419001108')
sheet.write(14,1,u'思礼镇')
sheet.write(14,2,all_1108)
sheet.write(14,3,pdmj_1_1108)
sheet.write(14,4,pdmj_2_1108)
sheet.write(14,5,pdmj_3_1108)
sheet.write(14,6,pdmj_4_1108)
sheet.write(14,7,pdmj_5_1108)
sheet.write(14,8,tchdmj_1_1108)
sheet.write(14,9,tchdmj_2_1108)
sheet.write(14,10,tchdmj_3_1108)
sheet.write(14,11,trzdmj_1_1108)
sheet.write(14,12,trzdmj_2_1108)
sheet.write(14,13,trzdmj_3_1108)
sheet.write(14,14,tryjzhlmj_1_1108)
sheet.write(14,15,tryjzhlmj_2_1108)
sheet.write(14,16,tryjzhlmj_3_1108)
sheet.write(14,17,trphzmj_10_1108)
sheet.write(14,18,trphzmj_2a_1108)
sheet.write(14,19,trphzmj_2b_1108)
sheet.write(14,20,trphzmj_3a_1108)
sheet.write(14,21,trphzmj_3b_1108)
sheet.write(14,22,swdyxmj_1_1108)
sheet.write(14,23,swdyxmj_2_1108)
sheet.write(14,24,swdyxmj_3_1108)
sheet.write(14,25,trzjswrmj_1_1108)
sheet.write(14,26,trzjswrmj_2_1108)
sheet.write(14,27,trzjswrmj_3_1108)
sheet.write(14,28,szmj_1_1108)
sheet.write(14,29,szmj_2_1108)
sheet.write(14,30,szmj_3_1108)
sheet.write(14,31,gdejdlmj_j_1108)
sheet.write(14,32,gdejdlmj_g_1108)

sheet.write(15,0,'419001109')
sheet.write(15,1,u'王屋镇')
sheet.write(15,2,all_1109)
sheet.write(15,3,pdmj_1_1109)
sheet.write(15,4,pdmj_2_1109)
sheet.write(15,5,pdmj_3_1109)
sheet.write(15,6,pdmj_4_1109)
sheet.write(15,7,pdmj_5_1109)
sheet.write(15,8,tchdmj_1_1109)
sheet.write(15,9,tchdmj_2_1109)
sheet.write(15,10,tchdmj_3_1109)
sheet.write(15,11,trzdmj_1_1109)
sheet.write(15,12,trzdmj_2_1109)
sheet.write(15,13,trzdmj_3_1109)
sheet.write(15,14,tryjzhlmj_1_1109)
sheet.write(15,15,tryjzhlmj_2_1109)
sheet.write(15,16,tryjzhlmj_3_1109)
sheet.write(15,17,trphzmj_10_1109)
sheet.write(15,18,trphzmj_2a_1109)
sheet.write(15,19,trphzmj_2b_1109)
sheet.write(15,20,trphzmj_3a_1109)
sheet.write(15,21,trphzmj_3b_1109)
sheet.write(15,22,swdyxmj_1_1109)
sheet.write(15,23,swdyxmj_2_1109)
sheet.write(15,24,swdyxmj_3_1109)
sheet.write(15,25,trzjswrmj_1_1109)
sheet.write(15,26,trzjswrmj_2_1109)
sheet.write(15,27,trzjswrmj_3_1109)
sheet.write(15,28,szmj_1_1109)
sheet.write(15,29,szmj_2_1109)
sheet.write(15,30,szmj_3_1109)
sheet.write(15,31,gdejdlmj_j_1109)
sheet.write(15,32,gdejdlmj_g_1109)

sheet.write(16,0,'419001110')
sheet.write(16,1,u'下冶镇')
sheet.write(16,2,all_1110)
sheet.write(16,3,pdmj_1_1110)
sheet.write(16,4,pdmj_2_1110)
sheet.write(16,5,pdmj_3_1110)
sheet.write(16,6,pdmj_4_1110)
sheet.write(16,7,pdmj_5_1110)
sheet.write(16,8,tchdmj_1_1110)
sheet.write(16,9,tchdmj_2_1110)
sheet.write(16,10,tchdmj_3_1110)
sheet.write(16,11,trzdmj_1_1110)
sheet.write(16,12,trzdmj_2_1110)
sheet.write(16,13,trzdmj_3_1110)
sheet.write(16,14,tryjzhlmj_1_1110)
sheet.write(16,15,tryjzhlmj_2_1110)
sheet.write(16,16,tryjzhlmj_3_1110)
sheet.write(16,17,trphzmj_10_1110)
sheet.write(16,18,trphzmj_2a_1110)
sheet.write(16,19,trphzmj_2b_1110)
sheet.write(16,20,trphzmj_3a_1110)
sheet.write(16,21,trphzmj_3b_1110)
sheet.write(16,22,swdyxmj_1_1110)
sheet.write(16,23,swdyxmj_2_1110)
sheet.write(16,24,swdyxmj_3_1110)
sheet.write(16,25,trzjswrmj_1_1110)
sheet.write(16,26,trzjswrmj_2_1110)
sheet.write(16,27,trzjswrmj_3_1110)
sheet.write(16,28,szmj_1_1110)
sheet.write(16,29,szmj_2_1110)
sheet.write(16,30,szmj_3_1110)
sheet.write(16,31,gdejdlmj_j_1110)
sheet.write(16,32,gdejdlmj_g_1110)

workbook.save(os.path.dirname(datapath)+"/"+"TJ.xls")