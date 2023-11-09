import openpyxl as xl
import datetime

from tkinter.messagebox import *

wb=xl.Workbook()
ws=wb.active

ws.append(['序号','日期','金额','品名','国别'])

for i in range(1,1030000):
    country=''
    if str(i)[len(str(i))-1:]=='1' or str(i)[len(str(i))-1:]=='3' or str(i)[len(str(i))-1:]=='5':
        country='中国'
    elif str(i)[len(str(i))-1:]=='2' or str(i)[len(str(i))-1:]=='4':
        country='泰国'
    else:
        country='日本'
    line=[i,datetime.date.today(),str(i)+'泰铢','出口商品'+str(i)+'号',country]
    ws.append(line)
    print(f'已生成EXCEL第{i}行!')

wb.save('C:\\Users\\lu\\Desktop\\1234.xlsx')
showinfo(title='提示',message='EXCEL表已创建完成！')



