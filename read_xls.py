#!/usr/bin/env python
#coding=utf-8
#-------------------------------------------------------------------------------
# Name:        模块1
# Purpose:
#
# Author:      Administrator
#
# Created:     30-06-2017
# Copyright:   (c) Administrator 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import xlrd,xlwt

#定义字典KV映射.
LNames=dict()
for i,lName in enumerate(wh1.row_values(0)):
    LNames.setdefault(lName,i)

Labels=['申报日期','运抵国','报关单号','物料名称','集装箱号','毛重','箱数','数量','单位','币别','金额','净重'
]
class ORDER_FORM(object): #订单

    def __init__(self):
        super(ORDER_FORM,self).__init__()
        self.fDATE='' #早报日期
        self.fCOUNTRY=''#运抵国
        self.fFORM_MARK=''#报关单号
        self.fBOX_MARK=''#集装箱号
        self.fITEMs=[]#多项::'物料名称,毛重,箱数,数量,单位,币别,金额,净重'


class fITEMs(object):

    def __init__(self):
        super(fITEMs,self).__init__()
        self.fNAME=''   #物料名称
        self.fKG=''     #毛重
        self.fCTNS=''   #箱数
        self.fAMOUNT='' #数量
        self.fUNIT=''   #单位
        self.fCURRENCY=''#币别
        self.fMONEY=''  #金额
        self.fNET=''    #净重


workbook = xlrd.open_workbook('G:\\test222\\LOAD_xls\\test.xls')
worksheets = workbook.sheet_names()             #获得所有工作表名
print('worksheets is %s' %worksheets)           #打印表名
wh1 = workbook.sheet_by_name(worksheets[0])     #定位到工作表
num_rows = wh1.nrows
LF=[]
LFORM = []
ITEM = []

#读入数据.
for curr_row in range(1,num_rows):
    if wh1.cell_value(curr_row,0)=='合计': #当单元格的值为合计是退出
        print(wh1.cell_value(curr_row,0))
        break
    if wh1.cell_value(curr_row,LNames['报关单号']) != '':
        for i in Labels[:5]:
            if i=='物料名称': #忽略 物料 列。
                continue

            LFORM.append(wh1.cell_value(curr_row, LNames[i]))
        items=[]
        for i in Labels[3:]:
            if i=='集装箱号': #忽略 柜号 列。
                continue
            items.append(wh1.cell_value(curr_row,LNames[i]))
        ITEM.append(items.copy())

    elif wh1.cell_value(curr_row,LNames['报关单号']) == '':
        items=[]
        for i in Labels[3:]:
            if i=='集装箱号':
                continue
            items.append(wh1.cell_value(curr_row,LNames[i]))
        ITEM.append(items.copy())


    if (wh1.cell_value(curr_row + 1, LNames['报关单号']) != '' and wh1.cell_value(curr_row, LNames['报关单号']) == '')\
            or (wh1.cell_value(curr_row,LNames['报关单号']) != '' and  wh1.cell_value(curr_row+1,LNames['报关单号']))\
            or wh1.cell_value(curr_row +1,LNames['审核标志'])=='合计':
        LFORM.append(ITEM.copy())
        LF.append(LFORM.copy())
        LFORM.clear()
        ITEM.clear()


#创建workbook和sheet对象
workbook = xlwt.Workbook() #注意Workbook的开头W要大写
sheet1 = workbook.add_sheet('sheet1',cell_overwrite_ok=True)

#向sheet页中写入数据
Y=0
val1=0
val2=0
for row in LF:
    Y+=1
    fstr=row[0].split('-')
    sheet1.write(Y, 0, fstr[-2]+'.'+fstr[-2])#日期
##    sheet1.write(Y, 0, row[0])#日期
    sheet1.write(Y, 4, row[1])#运抵国
    sheet1.row(Y).set_cell_text(5, row[2])
    # sheet1.write(Y, 5, row[2])#报关单号
    #sheet1.write(Y, 6, row[-1])#品名：1
    sheet1.write(Y, 7, row[3])#柜号
    for rowd in row[-1]:
        val1+=rowd[1]
        val2+=rowd[2]

    sheet1.write(Y, 8, val1)  #毛重
    sheet1.write(Y, 9, val2)  #箱数
    val1=0
    val2=0
    for rowd in row[-1]:
        sheet1.write(Y, 6, rowd[0])  # 品名
        sheet1.write(Y, 10, rowd[3]) # 数量
        sheet1.write(Y, 11, rowd[4]) # 单位
        sheet1.write(Y, 12, rowd[5]) # 币别
        sheet1.write(Y, 13, rowd[6])  #金额
        sheet1.write(Y, 14, rowd[7]) #净重
        if len(row)>1:
            Y+=1
    if len(row) > 1:
            Y-=1






# sheet1.write(0,0,'this should overwrite1')
# sheet1.write(0,1,'aaaaaaaaaaaa')
# #保存该excel文件,有同名文件时直接覆盖
workbook.save('G:\\test222\\LOAD_xls\\test_out.xls')


##num_rows = wh1.nrows
##for curr_row in range(1,num_rows):
##    row = wh1.row_values(curr_row)
##    print('row%s is %s' %(curr_row,row))

def main():
    pass

if __name__ == '__main__':
    main()
