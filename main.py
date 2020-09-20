import xlrd
import Sales

# book = xlrd.open_workbook('./data/1_企业信息.xlsx')
# sheet = book.sheets()[0]
# list = []
#
# for i in range(sheet.nrows):
#     list.append(sheet.row_values(i)[0])

# for j in range(len(list)):
#     print(list[j])
#     Sales.writeExcel(list[j], '销项发票信息', 'buy', 4)
Sales.write_excel('销项发票信息表', '销项发票信息', 'buy', 4)
# Sales.write_excel('进项发票信息表', '进项发票信息', 'sales', 6)
