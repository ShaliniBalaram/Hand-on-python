
import json
import xlsxwriter

# Opening JSON file
f = open('stock_data.json', )

# returns JSON object as
# a dictionary
data = json.load(f)
print(data)
# Iterating through the json
# list
x=0
y=0
for i in data:
    if i['Stock'] == 'NIFTY':
        x = x+1
    if i['Stock'] == 'BANKNIFTY':
        y = y+1
print(x)
print(y)

res = list(filter(lambda ii: ii['Stock'] == 'NIFTY', data))
res1 = list(filter(lambda ii: ii['Stock'] == 'BANKNIFTY', data))
print(type(res))
print(res1)



with xlsxwriter.Workbook('nifty and banknifty.xlsx') as workbook:
    worksheet = workbook.add_worksheet('NIFTY')
    for row_num, data_xl in enumerate(res):
        if row_num == 0:
            worksheet.write_row('A1',data_xl)
        else:
            worksheet.write_row(row_num,0, data_xl.values())
    worksheet = workbook.add_worksheet('BANKNIFTY')
    for row_num, data_xl1 in enumerate(res1):
        if row_num == 0:
            worksheet.write_row('A1', data_xl1)
        else:
            worksheet.write_row(row_num,0, data_xl1.values())









