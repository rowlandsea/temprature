__author__ = 'DYB'
import random
import xlrd
import xlwt
import time
from _datetime import datetime
from xlutils.copy import copy

#def wb():
    #book=open('酸轧电气全员体温日报.xls','w+')
book=xlwt.Workbook(encoding='utf-8',style_compression=0)
sheet=book.add_sheet('体温',cell_overwrite_ok=True)
style = xlwt.XFStyle()#格式信息
font = xlwt.Font()#字体基本设置
font.name = u'微软雅黑'
font.color = 'black'
font.height= 220 #字体大小，220就是11号字体，大概就是11*20得来的吧
style.font = font
alignment = xlwt.Alignment() # 设置字体在单元格的位置
alignment.horz = xlwt.Alignment.HORZ_CENTER #水平方向
alignment.vert = xlwt.Alignment.VERT_CENTER #竖直方向
style.alignment = alignment
border = xlwt.Borders()  #给单元格加框线
border.left = xlwt.Borders.THIN  #左
border.top=xlwt.Borders.THIN     #上
border.right=xlwt.Borders.THIN   #右
border.bottom=xlwt.Borders.THIN  #下
border.left_colour = 0x40  #设置框线颜色，0x40是黑色，颜色真的巨多，都晕了
border.right_colour = 0x40
border.top_colour = 0x40
border.bottom_colour = 0x40
style.borders = border
dateformat=xlwt.XFStyle()
dateformat.num_format_str='yyyy/mm/dd'
sheet.write_merge(0,0,0,4,'酸轧电气全员体温日报',style)
sheet.write(1,0,'单位',style)
sheet.write(1,1,'电气车间',style)
date=datetime.now()
sheet.write_merge(1,1,2,4,datetime.now(),dateformat)
sheet.write(2,0,'序号',style)
sheet.write(2,1,'姓名',style)
sheet.write(2,2,'上午体温',style)
sheet.write(2,3,'下午体温',style)
sheet.write(2,4,'备注',style)
sheet.write(3,1,'董英彬',style)
sheet.write(4,1,'李向阳',style)
sheet.write(5,1,'李鑫',style)
sheet.write(6,1,'姬健翔',style)
sheet.write(7,1,'吉凯东',style)
sheet.write(8,1,'高志宽',style)
sheet.write(9,1,'秦宁',style)
sheet.write(10,1,'张天啸',style)
for row in range(3,11):
    sheet.write(row,4," ",style)
sheet.col(2).width=3333
r=[]
s=range(1,10)
for i in range(8):
    da=[]
    p=s[i]
    da.append(p)
    r.append(da)
    index=3
for n in r:
        for t,item in enumerate(n):
            sheet.write(index,t,item,style)
        index=index+1
now=time.strftime('%Y-%m-%d',time.localtime(time.time()))
book.save("E:\\work\\practice\\staff temperature\\"+now+"酸轧电气全员体温日报.xls")

def sheet_wr(row,col):
    now=time.strftime('%Y-%m-%d',time.localtime(time.time()))
    wd=xlrd.open_workbook("E:\\work\\practice\\staff temperature\\"+now+r"酸轧电气全员体温日报.xls",formatting_info=True)
    wb=copy(wd)
    ws=wb.get_sheet(0)
    l=[]
    for i in range(8):
       t=[]
       x=random.uniform(36.0,36.7)
       t.append('%.1f'%x)
       l.append(t)
    for n in l:
        for m,item in enumerate(n):
            ws.write(row,col,item,style)
        row=row+1
    now=time.strftime('%Y-%m-%d',time.localtime(time.time()))
    wb.save("E:\\work\\practice\\staff temperature\\"+now+"酸轧电气全员体温日报.xls")

tm=datetime.now()
if tm.hour<12:
    ws1=sheet_wr(3,2)
if tm.hour>=12:
    ws1=sheet_wr(3,2)
    ws2=sheet_wr(3,3)








