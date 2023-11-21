import glob
import os

import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook
import pandas as pd

def my_fuc():
    # 获取当前目录下的所有Excel文件
    files = [file for file in os.listdir() if file.endswith('.xlsx')]
    for file in files:
        # 加载Excel文件
        wb = load_workbook(file)
        for sheet in wb.sheetnames:
            # 获取当前工作表
            ws = wb[sheet]
            # 在第一列的前面插入一列
            ws.insert_cols(1)
            # 设置第一行的值为"店铺"
            ws.cell(row=1, column=1).value = "店铺"
            # 获取表名的前两个字符
            shop_name = file[:2]

            # 设置余下行的值为表名的前两个字符
            i = ws.max_row
            real_max_row = 0
            while i > 0:
                row_dict = {i.value for i in ws[i]}
                if row_dict == {None}:
                    i = i - 1
                else:
                    real_max_row = i
                    break

            for row in range(2, real_max_row + 1):
                ws.cell(row=row, column=1).value = shop_name
            # 构建新的文件名
        new_file = file.replace('.xlsx', '-new.xlsx')
            # 保存修改后的Excel文件为新文件
        wb.save(new_file)

def hebing():
    # 获取当前目录下所有以"-new.xlsx"结尾的Excel文件
    files = [file for file in os.listdir() if file.endswith('-new.xlsx')]

    # 字典用于存储工作簿名称和对应的工作表数据
    workbook_data = {}

    # 遍历Excel文件
    for file in files:
        # 加载Excel文件
        wb = load_workbook(file)

        # 遍历工作表
        for sheetname in wb.sheetnames:
            # 获取当前工作表
            ws = wb[sheetname]

            # 如果工作簿名称已存在于字典中，则将当前工作表数据追加到已存在的列表中
            if sheetname in workbook_data:
                workbook_data[sheetname].extend(ws.iter_rows(values_only=True))
            else:
                # 如果工作簿名称不存在于字典中，则将当前工作表数据添加到字典中
                workbook_data[sheetname] = list(ws.iter_rows(values_only=True))

    # 创建新的Workbook对象
    merged_wb = Workbook()

    # 遍历工作簿数据字典
    for sheetname, data in workbook_data.items():
        # 创建新的工作表
        ws = merged_wb.create_sheet(title=sheetname)

        # 写入数据到工作表
        for row_index, row_data in enumerate(data):
            for col_index, cell_value in enumerate(row_data):
                column_letter = get_column_letter(col_index + 1)
                ws[column_letter + str(row_index + 1)].value = cell_value

    # 删除默认创建的Sheet
    del merged_wb['Sheet']

    # 保存合并后的Excel文件
    merged_wb.save('merged-new-all.xlsx')

def  delxiangtonghang():
    # 获取当前目录下以"merged-new-all.xlsx"命名的Excel文件
    file_paths = [f for f in os.listdir() if f == "merged-new-all.xlsx"]
    for file in file_paths:
        # 打开Excel文件
        wb = load_workbook(file)
        # 遍历所有工作表
        for ws in wb.worksheets:
            rows = list(ws.rows)
            # 删除第二行及之后的包含"店铺"的行
            for row in rows[1:]:
                if "店铺" in row[0].value:
                    ws.delete_rows(row[0].row)
        # 保存修改
        wb.save(file)

def  xiaoshou_toushibiao():
    # 读取Excel文件
    wb = openpyxl.load_workbook('merged-new-all.xlsx')
    ws = wb.active

    # 读取数据到DataFrame
    # data = ws.values
    # columns = next(data)[1:]
    # data = list(data)
    # idx = [r[0] for r in data]
    # data = (r[1:] for r in data)
    # df = pd.DataFrame(data, index=idx, columns=columns)

    # 读取数据到DataFrame，指定第一列为索引列
    df = pd.read_excel('merged-new-all.xlsx', sheet_name=ws.title, index_col=0)

    # 确保第一列作为列名而非索引
    df.reset_index(inplace=True)

    print(df.columns)

    # 创建数据透视表
    pivot_table = pd.pivot_table(df, values='已订购商品数量', index=['店铺', 'SKU'], aggfunc='sum')

    # 创建Excel写入器
    writer = pd.ExcelWriter('pivot_table.xlsx', engine='openpyxl')
    writer.book = wb

    # 写入数据透视表到Excel文件
    pivot_table.to_excel(writer, sheet_name='数据透视表')

    # 获取数据透视表工作表对象
    pivot_ws = writer.sheets['数据透视表']

    # 设置数据透视表布局和打印设置
    pivot_ws.sheet_view.showGridLines = False
    pivot_ws.sheet_properties.outlinePr.summaryBelow = True
    pivot_ws.sheet_properties.outlinePr.summaryRight = True

    # # 设置分类汇总为无
    # pivot_ws.pivotTables.pivotTable[0].showRowLabels = False
    #
    # # 选择重复项目标签
    # pivot_ws.pivotTables.pivotTable[0].rowFields[0].addItems('店铺')

    # 关闭Excel写入器
    writer.save()
    writer.close()

def  xiaoshou_toushibiao2():
    # 读取Excel文件
    filename = "merged-new-all.xlsx"
    excel_file = pd.ExcelFile(filename)

    # 获取所有工作簿名称
    sheet_names = excel_file.sheet_names

    # 创建一个新的Excel文件来保存数据透视表
    output_filename = "pivot_tables_xiaoshou.xlsx"
    writer = pd.ExcelWriter(output_filename, engine='openpyxl')

    # 遍历每个工作簿
    for sheet_name in sheet_names:
        if sheet_name.startswith("销售"):
            # 读取工作簿数据
            df = excel_file.parse(sheet_name)

            # 创建数据透视表
            pivot_table = pd.pivot_table(df, values='已订购商品数量', index=['店铺', 'SKU'], aggfunc='sum')

            # 将数据透视表写入新的工作簿
            pivot_table.to_excel(writer, sheet_name=sheet_name, index=True, startrow=1)

            # 获取工作簿的worksheet对象
            worksheet = writer.sheets[sheet_name]

            # 设置数据透视表布局和打印设置
            worksheet.sheet_view.showGridLines = False
            worksheet.sheet_properties.outlinePr.summaryBelow = True
            worksheet.sheet_properties.outlinePr.summaryRight = True


            # # 设置分类汇总为无
            # worksheet.pivot_tables[sheet_name].showTotals = False
            #
            # # 设置布局和打印设置
            # worksheet.pivot_tables[sheet_name].outline = True
            # worksheet.pivot_tables[sheet_name].outlineData = True
            #
            # # 选择重复项目标签
            # worksheet.pivot_tables[sheet_name].rowColGrandTotals = 'both'
            # worksheet.pivot_tables[sheet_name].rowColRepeatLabels = True

    # 保存并关闭Excel文件
    writer.save()
    writer.close()

    quxiao_hebingdanyuange(output_filename)

def  shangpintuiguang_toushibiao():
    # 读取Excel文件
    filename = "merged-new-all.xlsx"
    excel_file = pd.ExcelFile(filename)

    # 获取所有工作簿名称
    sheet_names = excel_file.sheet_names

    # 创建一个新的Excel文件来保存数据透视表
    output_filename = "pivot_tables_shangpintuiguang.xlsx"
    writer = pd.ExcelWriter(output_filename, engine='openpyxl')

    # 遍历每个工作簿
    for sheet_name in sheet_names:
        if sheet_name.startswith("商品推广"):
            # 读取工作簿数据
            df = excel_file.parse(sheet_name)

            # 创建数据透视表
            pivot_table = pd.pivot_table(df, values=['花费', '7天总销售额'], index=['店铺','广告SKU'], aggfunc='sum')

            # 将数据透视表写入新的工作簿
            pivot_table.to_excel(writer, sheet_name=sheet_name, index=True, startrow=1)

            # 获取工作簿的worksheet对象
            worksheet = writer.sheets[sheet_name]

            # 设置数据透视表布局和打印设置
            worksheet.sheet_view.showGridLines = False
            worksheet.sheet_properties.outlinePr.summaryBelow = True
            worksheet.sheet_properties.outlinePr.summaryRight = True

            # # 设置分类汇总为无
            # worksheet.pivot_tables[sheet_name].showTotals = False
            #
            # # 设置布局和打印设置
            # worksheet.pivot_tables[sheet_name].outline = True
            # worksheet.pivot_tables[sheet_name].outlineData = True
            #
            # # 选择重复项目标签
            # worksheet.pivot_tables[sheet_name].rowColGrandTotals = 'both'
            # worksheet.pivot_tables[sheet_name].rowColRepeatLabels = True

    # 保存并关闭Excel文件
    writer.save()
    writer.close()

    quxiao_hebingdanyuange(output_filename)

def  zhanshituiguang_toushibiao():
    # 读取Excel文件
    filename = "merged-new-all.xlsx"
    excel_file = pd.ExcelFile(filename)

    # 获取所有工作簿名称
    sheet_names = excel_file.sheet_names

    # 创建一个新的Excel文件来保存数据透视表
    output_filename = "pivot_tables_zhanshituiguang.xlsx"
    writer = pd.ExcelWriter(output_filename, engine='openpyxl')

    # 遍历每个工作簿
    for sheet_name in sheet_names:
        if sheet_name.startswith("展示推广"):
            # 读取工作簿数据
            df = excel_file.parse(sheet_name)

            # 创建数据透视表
            pivot_table = pd.pivot_table(df, values=['花费', '14天总销售额'], index=['店铺','广告SKU'], aggfunc='sum')

            # 将数据透视表写入新的工作簿
            pivot_table.to_excel(writer, sheet_name=sheet_name, index=True, startrow=1)

            # 获取工作簿的worksheet对象
            worksheet = writer.sheets[sheet_name]

            # 设置数据透视表布局和打印设置
            worksheet.sheet_view.showGridLines = False
            worksheet.sheet_properties.outlinePr.summaryBelow = True
            worksheet.sheet_properties.outlinePr.summaryRight = True

            # # 设置分类汇总为无
            # worksheet.pivot_tables[sheet_name].showTotals = False
            #
            # # 设置布局和打印设置
            # worksheet.pivot_tables[sheet_name].outline = True
            # worksheet.pivot_tables[sheet_name].outlineData = True
            #
            # # 选择重复项目标签
            # worksheet.pivot_tables[sheet_name].rowColGrandTotals = 'both'
            # worksheet.pivot_tables[sheet_name].rowColRepeatLabels = True

    # 保存并关闭Excel文件
    writer.save()
    writer.close()

    quxiao_hebingdanyuange(output_filename)

def quxiao_hebingdanyuange(file_path):
    # 获取目标文件
    # file_path = 'pivot_tables_xiaoshou.xlsx'

    # 加载Excel文件
    workbook = load_workbook(file_path)

    # 遍历所有工作簿
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        # 获取所有合并单元格的范围和值
        merged_cells = list(sheet.merged_cells)  # 创建副本
        merged_values = {}

        # 遍历合并单元格
        for merged_range in merged_cells:
            # 获取合并单元格的范围边界
            min_row, min_col, max_row, max_col = merged_range.min_row, merged_range.min_col, merged_range.max_row, merged_range.max_col

            # 获取合并单元格的值
            merged_value = sheet.cell(row=min_row, column=min_col).value

            # 保存合并单元格的值和范围
            merged_values[merged_range.coord] = {
                'value': merged_value,
                'range': (min_row, min_col, max_row, max_col)
            }

            # 取消合并单元格
            sheet.unmerge_cells(str(merged_range))

        # 恢复取消合并的单元格的值
        for coord, data in merged_values.items():
            merged_value = data['value']
            min_row, min_col, max_row, max_col = data['range']
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    sheet.cell(row=row, column=col).value = merged_value

    # 保存修改后的Excel文件
    workbook.save(file_path)


def huizong():
    # 设置文件夹路径和新文件名
    folder_path = './'  # 设置包含要合并的Excel文件的文件夹路径
    output_file = '汇总.xlsx'  # 设置输出文件名

    # 获取文件夹中以"pivot_"开头的所有Excel文件
    files = [file for file in os.listdir(folder_path) if file.startswith('pivot_') and file.endswith('.xlsx')]

    # 创建一个空的DataFrame用于存储合并后的数据
    merged_data = pd.DataFrame()

    # 遍历每个Excel文件并将所有工作簿横向合并
    for file in files:
        file_path = os.path.join(folder_path, file)

        # 读取Excel文件中的所有工作簿
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names

        # 遍历每个工作簿并将其横向合并到merged_data中
        for sheet_name in sheet_names:
            # 读取当前工作簿的数据
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            # 在最右侧新增两个空白列
            df_with_columns = df.assign(NewColumn1="", NewColumn2="")

            # 将数据横向合并到merged_data中
            merged_data = pd.concat([merged_data, df_with_columns], axis=1)

    # 将合并后的数据写入新的Excel文件
    merged_data.to_excel(output_file, index=False)

if __name__ == '__main__':
    print("这个脚本正在直接运行。")

    # my_fuc()
    # hebing()
    # delxiangtonghang()
    # xiaoshou_toushibiao()
    xiaoshou_toushibiao2()
    shangpintuiguang_toushibiao()
    zhanshituiguang_toushibiao()
    huizong()