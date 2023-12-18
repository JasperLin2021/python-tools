import re
from paddleocr import PaddleOCR

ocr = PaddleOCR(use_angle_cls=True, lang="en")

month_dict = {
    '1': 'Jan', '2': 'Feb', '3': 'Mar', '4': 'Apr',
    '5': 'May', '6': 'Jun', '7': 'Jul', '8': 'Aug',
    '9': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
}


def readMyImage(folder, img_path, month_content, current_year):
    # need to run only once to download and load model into memory
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
