from language_dict import *
import openpyxl as xl

wb=xl.Workbook()
ws=wb.active
dic=EngDict_Msg
for each in dic.keys():
    print(each,dic[each])
    ws.append([each,dic[each]])
wb.save('dictionary.xlsx')
print('Successfully!')
