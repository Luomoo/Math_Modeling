import xlrd
import xlwt


def write_excel(table_name, work_name, file_name, a):
    wb = xlwt.Workbook()
    book = xlrd.open_workbook('data/1_企业信息.xlsx')
    book1 = xlrd.open_workbook('./data/' + work_name + '.xlsx')

    sheet = book.sheets()[0]
    sheet1 = book1.sheets()[0]
    data = sheet1.row_values(0)
    list = []

    for i in range(sheet.nrows):
        if i != 0:
            list.append(sheet.row_values(i)[0])
    print(list)

    # write_data(wb, 'E2', data, sheet1, a)
    for j in range(len(list)):
        write_data(wb, list[j], data, sheet1, a)
        print(list[j])
    wb.save(r"./" + file_name + "/" + table_name + '.xls')


def write_data(wb, table_name, data, sheet, a):
    money_sum = float(0)
    curr = 0

    ws = wb.add_sheet(table_name)  # 增加sheet
    for i in range(len(data)):
        ws.write(0, i, data[i])

    for i in range(sheet.nrows):
        if sheet.row_values(i)[0] == table_name:
            if sheet.row_values(i)[7] == "有效发票":
                money_sum = money_sum + float(sheet.row_values(i)[a])
                curr += 1
                for j in range(len(sheet.row_values(i))):
                    ws.write(curr, j, sheet.row_values(i)[j])
                # print(sheet.row_values(i))
    print(money_sum)
    ws.write(0, 9, '总计')
    ws.write(1, 9, money_sum)
    print(table_name)
    print("---------")
