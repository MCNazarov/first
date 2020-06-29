#https://tokmakov.msk.ru/blog/item/71

import openpyxl
# создаем Exel файл
wb = openpyxl.Workbook()
# название листа
wb.create_sheet(title= 'Первый лист', index= 0)
sheet = wb ['Первый лист']
sheet['A1'] = 'Здравствуй мир'
sheet['A2'] = 'Здравствуй и тебе мир'
wb.save('example.xlsx')