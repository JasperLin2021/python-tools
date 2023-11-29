import os
import re

import pandas as pd

import openpyxl
from openpyxl.reader.excel import load_workbook

from utils import deleteRow, get_ad_sku_dict


def createDirectory():
    # 获取当前工作目录
    current_directory = os.getcwd()

    # 新建文件夹
    folder_name = "输出2"
    folder_path = os.path.join(current_directory, folder_name)
    os.makedirs(folder_path)


def arrangeCostHeadProcess():
    # 获取当前工作目录
    current_directory = os.getcwd()
    # 新建文件夹
    folder_name = "输出2"

    # 创建一个新的Excel文件来保存数据透视表
    output_filename1 = "Arrange_成本头程1.xlsx"
    output_filename2 = "Arrange_成本头程2.xlsx"
    new_file1 = os.path.join(current_directory, folder_name, output_filename1)
    new_file2 = os.path.join(current_directory, folder_name, output_filename2)

    target_workbook1 = openpyxl.Workbook()
    target_workbook2 = openpyxl.Workbook()
    target_sheet1 = target_workbook1.active
    target_sheet2 = target_workbook2.active

    files = [file for file in os.listdir() if file.startswith('成本头程') ]
    for file in files:
        # 读取原始的Excel文件
        wb = load_workbook(file, read_only=True, data_only=True)
        sheet = wb.active

        for row in sheet.iter_rows(min_row=1, values_only=True):
            data_a = row[0]
            data_k = row[10]
            data_o = row[14]
            data_p = row[15]

            # 复制到目标文件，并在D列和F列之间插入空白列
            target_sheet1.append([data_a, data_k])
            target_sheet2.append([data_o, data_p])

        # 将B1单元格的值设置为"SKU"
        target_sheet2['B1'] = 'SKU'


    # 保存目标文件

    target_workbook1.save(new_file1)
    target_workbook2.save(new_file2)


def vlookupCostHeadProcess():
    # 获取当前工作目录
    current_directory = os.getcwd()
    # 新建文件夹
    folder_name = "输出2"

    # 创建一个新的Excel文件来保存数据透视表
    output_filename1 = "Arrange_成本头程1.xlsx"
    output_filename2 = "Arrange_成本头程2.xlsx"
    output_filename3 = "Arrange_end_成本头程.xlsx"
    file1 = os.path.join(current_directory, folder_name, output_filename1)
    file2 = os.path.join(current_directory, folder_name, output_filename2)
    new_file = os.path.join(current_directory, folder_name, output_filename3)

    # 读取abc.xlsx文件中的数据
    df_file1 = pd.read_excel(file1)

    # 读取efg.xlsx文件中的数据
    df_file2 = pd.read_excel(file2)

    # 进行VLOOKUP操作
    merged_df = pd.merge(df_file1, df_file2, how='left', on='SKU')

    merged_df.to_excel(new_file,  sheet_name='成本头程', index=False)





def arrangeFBA():
    # 获取当前工作目录
    current_directory = os.getcwd()
    # 新建文件夹
    folder_name1 = "输出"
    folder_name2 = "输出2"

    # 创建一个新的Excel文件来保存数据透视表
    output_filename = "Arrange_fba库存.xlsx"
    new_file = os.path.join(current_directory, folder_name2, output_filename)

    #
    CostHeadProcess_filename = "Arrange_end_成本头程.xlsx"
    CostHeadProcess_file = os.path.join(current_directory, folder_name2, CostHeadProcess_filename)
    CostHeadProcess_dict = get_ad_sku_dict(CostHeadProcess_file, "成本头程", 2, 1)

    Summary_filename = "汇总.xlsx"
    Summary_file = os.path.join(current_directory, folder_name1, Summary_filename)
    Seven_day_sales_dict = get_ad_sku_dict(Summary_file, "Sheet1", 43, 44)
    fifteen_day_sales_dict = get_ad_sku_dict(Summary_file, "Sheet1", 48, 49)
    thirty_day_sales_dict = get_ad_sku_dict(Summary_file, "Sheet1", 53, 54)



    target_workbook = openpyxl.Workbook()
    target_sheet = target_workbook.active

    files = [file for file in os.listdir() if file.startswith('fba') ]
    for file in files:
        # 读取原始的Excel文件
        wb = load_workbook(file)
        sheet = wb.active

        pattern = r'Y(\d+)'
        for row in sheet.iter_rows(min_row=1, values_only=True):
            match = re.search(pattern, row[1])
            if match:
                data_b = match.group(1) + "店"
            else:
                data_b = row[1]

            data_d = row[3]  # 获取D列数据
            data_m = row[12]  # 获取M列数据
            data_f = row[5]  # 获取F列数据
            data_z = row[25]  # 获取Z列数据
            data_aa = row[26]  # 获取AA列数据
            data_ac = row[28]  # 获取AC列数据
            data_ad = row[29]  # 获取AD列数据
            data_ak = row[36]  # 获取AK列数据

            available_stock = data_z+data_aa+data_ac+data_ad+data_ak
            average_cost = round(CostHeadProcess_dict.get(data_d, [0])[0], 2)
            end_value = available_stock * average_cost
            Seven_day_sales = Seven_day_sales_dict.get(data_d, [0])[0]
            fifteen_day_sales = fifteen_day_sales_dict.get(data_d, [0])[0]
            thirty_day_sales = thirty_day_sales_dict.get(data_d, [0])[0]


            # 复制到目标文件，并在D列和F列之间插入空白列
            target_sheet.append(
                [data_b, data_d, data_m, '', data_f, data_z, data_aa, data_ac, data_ad, data_ak,
                 available_stock,
                 average_cost,
                 end_value,
                 Seven_day_sales,
                 fifteen_day_sales,
                 thirty_day_sales
                 ])


    # 保存目标文件

    target_workbook.save(new_file)
    deleteRow(new_file, 2)
    deleteRow(new_file, 3, "小满1店")



def copyArrangeFBA():
    # 获取当前工作目录
    current_directory = os.getcwd()
    # 新建文件夹
    folder_name = "输出2"

    # 创建一个新的Excel文件来保存数据透视表
    output_filename = "New_fba库存.xlsx"
    new_file = os.path.join(current_directory, folder_name, output_filename)

    # 打开源文件
    source_file = os.path.join(current_directory, folder_name, "Arrange_fba库存.xlsx")
    source_workbook = openpyxl.load_workbook(source_file)
    source_sheet = source_workbook.active

    # 打开目标文件
    target_file = "亚马逊库存分析模板.xlsx"
    target_workbook = openpyxl.load_workbook(target_file)
    target_sheet = target_workbook.active

    # 复制数据
    start_row = 4
    start_column = 2  # 列号B对应索引2
    for row in source_sheet.iter_rows():
        for cell in row:
            target_sheet.cell(row=start_row, column=start_column).value = cell.value
            start_column += 1
        start_row += 1
        start_column = 2  # 重置列号B对应索引2

    # 保存目标文件
    target_workbook.save(new_file)


if __name__ == '__main__':
    print("这个脚本正在直接运行。")

    # createDirectory()
    arrangeCostHeadProcess()
    vlookupCostHeadProcess()
    arrangeFBA()






    copyArrangeFBA()
