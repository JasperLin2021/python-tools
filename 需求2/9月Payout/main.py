import json
import os
from datetime import datetime
# from utils import *
import pyautogui

import re
from paddleocr import PaddleOCR

def readMyImage(folder, img_path, month_content, current_year):

    # need to run only once to download and load model into memory
    if month_content == 12:
        check_month_list = [month_dict[str(month_content - 1)], month_dict[str(month_content)],
                            month_dict[str(1)]]
    elif month_content == 1:
        check_month_list = [month_dict[str(12)], month_dict[str(month_content)],
                            month_dict[str(month_content + 1)]]
    else:
        check_month_list = [month_dict[str(month_content - 1)], month_dict[str(month_content)],
                        month_dict[str(month_content + 1)]]
    result = ocr.ocr(img_path, cls=True)

    res = []
    raw_dict = {}

    # print(result[0])
    error_folder = None
    manual_folder = None
    for sub_item in result[0]:
        if re.match(r'US\s*\$|US\s*S|USS|US\$', sub_item[1][0]):
            res_1 = sub_item[1][0].replace(' ', '')
            res.append(re.split(r'USS|US\$', res_1)[1].replace(",", ""))
        elif re.search(r"\$\d+.*", sub_item[1][0]):
            res.append(re.search(r"\$\d+.*", sub_item[1][0])[0].split('$')[1].replace(",", ""))
        elif "'" in sub_item[1][0] and "$" in sub_item[1][0]:
            res.append("error")
            error_folder = folder
        elif all(keyword not in sub_item[1][0] for keyword in
                 ['Start', 'End', 'uary', 'March', 'April', 'May', 'June', 'July', 'August', 'ber']):
            res_1 = sub_item[1][0].replace(' ', '').replace(',', '').replace('.', '')
            for cml in check_month_list:
                if cml in res_1:
                    res.append(re.split(r'' + str(current_year) + '', res_1)[0])

    try:
        for i in range(0, len(res), 2):
            key = res[i]
            value = res[i + 1]
            dictionary = {key: value}
            raw_dict.update(dictionary)


        filtered_dict = {key.split(month_dict[str(month_content)])[1]: value for key, value in raw_dict.items() if
                         month_dict[str(month_content)] in key}
    except Exception as e:
        filtered_dict = {'error':'error'}
        manual_folder = folder

    return filtered_dict, error_folder, manual_folder



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

    ocr = PaddleOCR(use_angle_cls=True, lang="en")

    month_dict = {
        '1': 'Jan', '2': 'Feb', '3': 'Mar', '4': 'Apr',
        '5': 'May', '6': 'Jun', '7': 'Jul', '8': 'Aug',
        '9': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
    }

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
    title = "截图识别功能"
    pyautogui.alert(message, title)