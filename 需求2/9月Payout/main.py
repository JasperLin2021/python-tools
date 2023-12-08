import os
import csv
from base64 import encode

import pandas as pd


from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import Workbook

from utils import *

def monthlyPaymentDetails(current_directory):
    image_extensions = [".jpg", ".jpeg", ".png", ".gif", ".bmp"]


    folder_list = [folder for folder in os.listdir('.') if os.path.isdir(folder)]

    big_dict = {}
    for folder in folder_list:
        dict = {}
        directory=os.path.join(current_directory, folder)
        for filename in os.listdir(directory):
            file_path = os.path.join(directory, filename)
            # 检查文件是否是图片文件
            if os.path.isfile(file_path) and any(filename.lower().endswith(ext) for ext in image_extensions):
                dict.update(readMyImage(file_path))
        big_dict[folder] = dict

    return big_dict

def determineDailyBills(current_directory):
    # monthlyPaymentDetails_dict = monthlyPaymentDetails(current_directory)
    monthlyPaymentDetails_dict = {'AA': {'30': '61.51', '29': '42.07', '28': '53.10', '27': '133.10', '26': '140.68', '25': '94.79', '24': '104.54', '23': '102.96', '22': '123.88', '21': '87.23', '20': '106.06', '19': '43.07', '18': '79.70', '17': '102.27', '16': '106.14', '15': '118.38', '14': '91.13', '13': '101.16', '12': '6.32', '11': '33.14', '10': '72.37', '9': '160.93', '8': '123.08', '7': '72.03', '6': '211.49', '5': '133.70', '3': '39.03', '2': '88.25', '1': '54.28'}, 'AE': {'22': '191.55', '21': '151.03', '20': '191.28', '19': '203.04', '18': '128.37', '17': '88.11', '16': '172.58', '15': '262.08', '14': '136.61', '13': '113.68', '12': '160.67', '11': '268.95', '10': '244.69', '9': '157.54', '8': '119.80', '7': '128.42', '6': '307.93', '5': '292.02', '3': '152.70', '2': '22.67', '1': '128.40', '30': '356.58', '29': '176.29', '28': '197.77', '27': '198.55', '26': '138.41', '25': '134.81', '24': '106.87', '23': '158.21'}, '__pycache__': {}}


    folder_list = [folder for folder in os.listdir('.') if os.path.isdir(folder) and folder != '__pycache__']
    for folder in folder_list:
        dict = {}
        directory=os.path.join(current_directory, folder)
        for filename in os.listdir(directory):
            file_path = os.path.join(directory, filename)
            # 检查文件是否是图片文件
            if file_path.endswith(".csv"):
                new_file_path = os.path.join(directory, filename.split(".csv")[0] + ".xlsx")
                # 读取CSV文件
                # # 读取CSV文件
                f = open(file_path, 'r', encoding='utf-8')
                # 创建一个workbook 设置编码
                workbook = Workbook()
                # 创建一个worksheet
                worksheet = workbook.active
                workbook.title = 'sheet'

                for line in f:
                    if '","' in line:
                        line.replace('","', '*')
                        line = line.split('@')
                        worksheet.append(line)
                    else:
                        a = line.replace('"', '')
                        # line.split(',')
                        worksheet.append([a])
                    # row = line.split(',')
                    # worksheet.append(line)
                    # if row[0].endswith('00'):    # 每一百行打印一次
                    #     print(line, end="")

                workbook.save(new_file_path)
                # print(file_path)
                # csv = pd.read_csv(file_path, encoding='utf-8', header=None, sep=',')
                #
                # csv.to_excel(new_file_path, sheet_name='data')

if __name__ == '__main__':
    # 获取当前工作目录
    current_directory = os.getcwd()

    # monthlyPaymentDetails(current_directory)

    determineDailyBills(current_directory)
    # data = pd.read_csv('./9.1.csv', )
    # data.to_excel('./9.1.xlsx', sheet_name='data')
