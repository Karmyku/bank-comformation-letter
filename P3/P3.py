#用python批量生成银行询证函（excel格式）
import openpyxl
from copy import copy

wb=openpyxl.load_workbook('example1.xlsx')
模板=wb['模板']
银行存款=wb['银行存款']
银行借款=wb['银行借款']
def 存款行复制函数(k,i):           #两个sheet间复制值
    新建工作表[f'A{k}']=银行存款[f'C{i}'].value
    新建工作表[f'C{k}']=银行存款[f'E{i}'].value
    新建工作表[f'E{k}']=银行存款[f'G{i}'].value
    新建工作表[f'F{k}']=银行存款[f'H{i}'].value
    新建工作表[f'G{k}']=银行存款[f'I{i}'].value
    新建工作表[f'H{k}']=银行存款[f'J{i}'].value
    新建工作表[f'J{k}']=银行存款[f'L{i}'].value
    新建工作表[f'K{k}']=银行存款[f'M{i}'].value
    新建工作表[f'L{k}']=银行存款[f'N{i}'].value
    新建工作表[f'M{k}']=银行存款[f'O{i}'].value
    新建工作表[f'N{k}']=银行存款[f'P{i}'].value

def copy_rows(sheet, row_idx):    #插入行并复制上面一行格式
    row = sheet[row_idx]
    sheet.insert_rows(row_idx)      #在哪一行上面插入一行
    for cell in row:                #复制上面一行格式，将新一行的单元格格式全部更改一致
        new_cell = sheet.cell(row=row_idx, column=cell.col_idx)
        if cell.has_style:
            new_cell.font = copy(cell.font)
            new_cell.border = copy(cell.border)
            new_cell.fill = copy(cell.fill)
            new_cell.number_format = copy(cell.number_format)
            new_cell.protection = copy(cell.protection)
            new_cell.alignment = copy(cell.alignment)

i=3

while 银行存款[f'B{i}'].value!=0:       #当有清单的银行名称那列有银行名称时
    新建工作表=wb.copy_worksheet(模板)     #新建sheet
    新建工作表.title=银行存款[f'B{i}'].value     #将新建sheet的名字改成银行名称
    新建工作表['A4']=银行存款[f'B{i}'].value
    新建工作表['M2']=str(新建工作表['M2'].value)+str(银行存款[f'A{i}'].value)     #银行将询证函编号继续插入
    k=20
    存款行复制函数(k,i)        #运行上面复制询证函银行存款内容的函数
    while 银行存款[f'B{i}'].value==银行存款[f'B{i+1}'].value:   #当清单上银行名称与下面的银行名称一致时
        k=k+1
        copy_rows(新建工作表, k)     #插入新的一行，并与上面一行的格式一致
        i=i+1
        存款行复制函数(k,i)        #运行上面复制询证函银行存款内容的函数
    i=i+1       #继续循环

wb.save('example2.xlsx')        #保存为2
