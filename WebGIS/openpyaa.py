import pandas as pd
import os
import datetime
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

CONTENT_ROW = 4

root = tk.Tk()
root.withdraw()

Folderpath = filedialog.askdirectory() #获得选择好的文件夹
Filepath = filedialog.askopenfilename() #获得选择好的文件

df1 = pd.read_excel(Filepath, sheet_name=0)
nodeDataList = list(df1.groupby(['FBFMC']))

for nodeData in nodeDataList:
    # 将数据按照承包方归类排序
    nodeData[1].sort_values(by = 'CBFBM',inplace=True,ascending=True)
    select_list = []
    beizhu_list = nodeData[1]['BZ'].tolist() #备注
    cbfbm_list = nodeData[1]['CBFBM'].tolist() #承包方
    for row in range(0, len(beizhu_list)):
        if beizhu_list[row] > 0 :
            select_list.append(cbfbm_list[row])
    #筛查后的数据源
    sourceData = nodeData[1][nodeData[1].CBFBM.isin(list(set(select_list)))]
    if len(sourceData) == 0:
        continue  
    # 创建一个workbook 
    workbook = Workbook()
    # 创建一个worksheet
    worksheet = workbook.active
    #设置列宽、行高
    worksheet.column_dimensions['A'].width = 9.8
    worksheet.column_dimensions['B'].width = 22.3
    worksheet.column_dimensions['C'].width = 21
    worksheet.column_dimensions['D'].width = 9.2
    worksheet.column_dimensions['E'].width = 13
    worksheet.column_dimensions['F'].width = 11
    worksheet.row_dimensions[1].height = 20
    worksheet.row_dimensions[2].height = 30
    worksheet.row_dimensions[3].height = 35.25
# 参数对应 行, 列, 值
    # 获取当前时间戳
    year = datetime.datetime.now().year
    month = datetime.datetime.now().month
    day = datetime.datetime.now().day
    timeStr = '公示日期：' + str(year)+'年'+str(month)+'月'+str(day)+'日'+' -- '+ str(month)+'月'+ str(day+7)+'日'
    worksheet.merge_cells(start_row=1,end_row=1,start_column=1,end_column=6)
    worksheet.cell(1,1).value = '农村土地承包经营权证“基本农田”信息变更公示表'

  

    worksheet.merge_cells(start_row=2,end_row=2,start_column=1,end_column=2)
    worksheet.cell(2,1).value = str(nodeData[0])
    
   
    worksheet.merge_cells(start_row=2,end_row=2,start_column=3,end_column=6)
    worksheet.cell(2,3).value = timeStr
    worksheet.cell(2, 3).alignment = Alignment(horizontal='center', vertical='center')
    
    worksheet.cell(3,1).value = '承包方'
    worksheet.cell(3, 1).alignment = Alignment(horizontal='center', vertical='center')
    
    worksheet.cell(3, 2).value = '地块代码'
    
    
    worksheet.cell(3, 3).value = '坐落（四至）'

    worksheet.cell(3, 4).value = '面积（亩）'

    #提取df列的值
    cbflist = sourceData['ZJRXM'].tolist() #承包方
    dkbmlist = list(pd.Series(sourceData['DKBM'].tolist()).astype(str)) #地块代码
    dkdzlist = sourceData['DKDZ'].tolist() #东至
    dkxzlist = sourceData['DKXZ'].tolist() #西至
    dknzlist = sourceData['DKNZ'].tolist() #南至
    dkbzlist = sourceData['DKBZ'].tolist() #北至
    mjlist = sourceData['SCMJM'].tolist() #实测面积
    jbntlist = sourceData['SFJBNT'].tolist() #是否基本农田
    dkbzxxlist = sourceData['DKBZXX'].tolist() #备注
    zhen = sourceData['ZHEN'].tolist()
    cun = sourceData['CUN'].tolist()
    zu = sourceData['ZU'].tolist()
    CBFStr = ''
    record_row = 0
    next_page_horizon, next_page_vertical = worksheet.page_breaks
    for row in range(0, len(sourceData.index)):
        count_row = row*4
        for i in range(0,4) :
            #设置行高
            worksheet.row_dimensions[CONTENT_ROW + count_row + i].height = 14.5
            if i==0:
                worksheet.cell(CONTENT_ROW +count_row+i,3).value = '东：' + dkdzlist[row]

            elif i==1:
                worksheet.cell(CONTENT_ROW + count_row + i, 3).value = '西：' + dkxzlist[row]

            elif i==2:
                worksheet.cell(CONTENT_ROW + count_row + i, 3).value = '南：' + dknzlist[row]

            else:
                worksheet.cell(CONTENT_ROW + count_row + i, 3).value = '北：' + dkbzlist[row]


        #合并单元格
        isbottom = False
        if row > 0 and row % 10 == 0:
            isbottom = True
        if (CBFStr != cbflist[row] or isbottom) and row > 0:
            worksheet.merge_cells(start_row=CONTENT_ROW+record_row,end_row=CONTENT_ROW+count_row-1,start_column=1,end_column=1)
            worksheet.cell(CONTENT_ROW + record_row,1).value = CBFStr

            if row == len(sourceData.index)-1:
                worksheet.merge_cells(start_row=CONTENT_ROW + count_row, end_row=CONTENT_ROW + count_row + 3,start_column=1, end_column=1)
                worksheet.cell(CONTENT_ROW + count_row, 1).value = cbflist[row]
                worksheet.cell(CONTENT_ROW + count_row, 1).alignment = Alignment(horizontal='center',vertical='center')
            record_row = count_row
        elif row == len(sourceData.index)-1:
            worksheet.merge_cells(start_row=CONTENT_ROW+record_row,end_row=CONTENT_ROW+count_row+3,start_column=1,end_column=1)
            worksheet.cell(CONTENT_ROW + record_row, 1).value = cbflist[row]
            worksheet.cell(CONTENT_ROW + record_row, 1).font = font1
            worksheet.cell(CONTENT_ROW + record_row, 1).border = border1
            worksheet.cell(CONTENT_ROW + record_row, 1).alignment = Alignment(horizontal='center', vertical='center')
        CBFStr = cbflist[row]

        worksheet.merge_cells(start_row=CONTENT_ROW+count_row,end_row=CONTENT_ROW+count_row+3,start_column=2,end_column=2)
        worksheet.cell(CONTENT_ROW + count_row, 2).value = dkbmlist[row]
        worksheet.cell(CONTENT_ROW + count_row, 2).font = font1
        worksheet.cell(CONTENT_ROW + count_row, 2).alignment = Alignment(horizontal='center', vertical='center')

        worksheet.merge_cells(start_row=CONTENT_ROW+count_row, end_row=CONTENT_ROW+count_row+3,start_column=4,end_column=4)
        worksheet.cell(CONTENT_ROW + count_row, 4).value = mjlist[row]
        worksheet.cell(CONTENT_ROW + count_row, 4).font = font1
        worksheet.cell(CONTENT_ROW + count_row, 4).alignment = Alignment(horizontal='center', vertical='center')

        jbntStr = ''
        if jbntlist[row] == 1:
            jbntStr = '是'
        else:
            jbntStr = '否'
        worksheet.merge_cells(start_row=CONTENT_ROW+count_row, end_row=CONTENT_ROW+count_row+3,start_column=5,end_column=5)
        worksheet.cell(CONTENT_ROW + count_row, 5).value = jbntStr
        worksheet.cell(CONTENT_ROW + count_row, 5).font = font1
        worksheet.cell(CONTENT_ROW + count_row, 5).alignment = Alignment(horizontal='center', vertical='center')

        worksheet.merge_cells(start_row=CONTENT_ROW+count_row,end_row=CONTENT_ROW+count_row+3, start_column=6,end_column=6)
        worksheet.cell(CONTENT_ROW + count_row,6).value = dkbzxxlist[row]
        worksheet.cell(CONTENT_ROW + count_row, 6).font = font1
        worksheet.cell(CONTENT_ROW + count_row, 6).alignment = Alignment(horizontal='center', vertical='center',wrapText=True)
        # 插入水平分页符
        if row > 0 and row % 10 == 0:
            next_page_horizon.append(Break(count_row+3))

    worksheet.print_title_rows = '$1:$3'

    # 保存文件和Excel
    zhenPath = str(zhen[0])
    cunPath = str(cun[0])
    zuPath = str(zu[0])
    path = Folderpath + '/' + zhenPath + '/' + cunPath + '/' + zuPath
    # path = Folderpath + '/' + str(nodeData[0])
    isExists = os.path.exists(path)
    if not isExists:
        os.makedirs(path)
    workbook.save(path + '/' + str(nodeData[0]) + '.xlsx')

messagebox.showinfo('提示','Excel表格已经生成完毕！')

