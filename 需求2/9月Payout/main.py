import json
import os
from datetime import datetime
from utils import *
import pyautogui


def monthlyPaymentDetails(current_directory, month_content, current_year):
    image_extensions = [".jpg", ".jpeg", ".png", ".gif", ".bmp"]

    folder_list = [folder for folder in os.listdir('.') if os.path.isdir(folder)]

    big_dict = {}
    error_img_list = []
    manual_list = []
    for folder in folder_list:
        dict1 = {}
        directory = os.path.join(current_directory, folder)
        for filename in os.listdir(directory):
            file_path = os.path.join(directory, filename)
            # 检查文件是否是图片文件
            if os.path.isfile(file_path) and any(filename.lower().endswith(ext) for ext in image_extensions):
                print(file_path)
                value1, value2, value3 = readMyImage(folder, file_path, month_content, current_year)
                dict1.update(value1)
                if value2:
                    error_img_list.append(value2)
                if value3:
                    manual_list.append(value3)

        if not dict1.get("error"):
            big_dict[folder] = dict1

    return big_dict, error_img_list, manual_list


if __name__ == '__main__':
    # 获取当前工作目录
    current_directory = os.getcwd()
    file = open('月份.txt', 'r')
    month_content = int(file.read())
    file.close()

    current_year = str(datetime.now().year)

    monthlyPaymentDetails_dict, error_monthlyPaymentDetails_list, manual_monthlyPaymentDetails_list = monthlyPaymentDetails(
        current_directory, month_content, current_year)
    # 将字典序列化为JSON字符串并写入文件
    with open('各店铺当月每天回款情况.txt', 'w', encoding='utf-8') as file:
        json.dump(monthlyPaymentDetails_dict, file, ensure_ascii=False)

    with open('需校正回款情况.txt', 'w', encoding='utf-8') as file:
        json.dump(error_monthlyPaymentDetails_list, file, ensure_ascii=False)

    with open('需人工处理.txt', 'w', encoding='utf-8') as file:
        json.dump(manual_monthlyPaymentDetails_list, file, ensure_ascii=False)

    message = "完成！"
    title = "截图识别完成"
    pyautogui.alert(message, title)