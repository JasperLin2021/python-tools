import os

import openpyxl


def test():
    # 获取当前工作目录
    # current_directory = os.getcwd()

    # new_file = os.path.join(current_directory, "test1.xlsx")
    # 打开 Excel 文件
    workbook = openpyxl.load_workbook( "test1.xlsx")

    # 选择 Sheet1
    sheet = workbook['Sheet1']

    # 定义一个列表用于存储单元格的值
    cell_values = []

    # 遍历 A 列的每个单元格
    for cell in sheet['A']:

        if cell.number_format.find("US$") != -1:
            print("字符串中包含'US$'")
            a = round(float(str(cell.value)) * 1.1, 4)
            cell_values.append(a)
        if str(cell.value).find("CA$") != -1:
            print("字符串中包含'CA$'")
            a =  round(float(str(cell.value).split("CA$")[1]) * 1.1, 4)
            cell_values.append(a)


    # 输出列表中的值
    print(cell_values)

if __name__ == '__main__':
    test()