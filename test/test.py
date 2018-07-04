import configparser
import xlrd
import xlwt

config = configparser.ConfigParser()
config.read('channels-config')

xml_name = u'1_标准报表.xls'

workbook = xlrd.open_workbook('../in/' + xml_name)
sheet_names = workbook.sheet_names()



# 隐藏名字
def modifyName():
    for sheet_name in sheet_names:
        sheet = workbook.sheet_by_name(sheet_name)
        cell0 = sheet.cell(2, 9)
        cell0.value = get_modified_name(cell0)
        print(cell0.value)
        cell1 = sheet.cell(2, 24)
        cell1.value = get_modified_name(cell1)
        print(cell1.value)
        cell2 = sheet.cell(2, 39)
        cell2.value = get_modified_name(cell2)
        print(cell2.value)


# 遍历表 打印所有非空值及其所在的位置
def test():
    wb = xlwt.Workbook()
    # 遍历表
    for sheet_name in sheet_names:
        ws = wb.add_sheet(sheet_name)
        sheet = workbook.sheet_by_name(sheet_name)
        for row in range(sheet.nrows):
            values = sheet.row_values(row)
            for col in range(len(values)):
                cell = sheet.cell(row, col)
                # if cell.value:
                if(row is 2 and (col is 9 or col is 24 or col is 39)):
                    cell.value = get_modified_name(cell)
                ws.write(row, col, cell.value)
                print(row, col, cell)
    wb.save("../in/new_" + xml_name)


def get_modified_name(cell):
    return cell.value[0] + '*' * (len(cell.value) - 1)


test()
