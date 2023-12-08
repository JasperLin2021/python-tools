from paddleocr import PaddleOCR


ocr = PaddleOCR(use_angle_cls=True, lang="en")
def readMyImage(img_path):
      # need to run only once to download and load model into memory

    result = ocr.ocr(img_path, cls=True)

    res = []
    dict = {}
    month = ['Sep']
    money = ['$']

    for sub_item in result[0]:
        for item in month:
            if item in sub_item[1][0]:
                res.append(sub_item[1][0].split(',')[0].split()[1])
        for item in money:
            if item in sub_item[1][0]:
                res.append(sub_item[1][0].split('$')[1])

    for i in range(0, len(res), 2):
        key = res[i]
        value = res[i + 1]
        dictionary = {key: value}
        dict.update(dictionary)

    return dict