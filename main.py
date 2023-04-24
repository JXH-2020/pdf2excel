import os

from paddleocr import PaddleOCR

import utils

import os

os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'
if __name__ == '__main__':
    ocr = PaddleOCR(use_angle_cls=True, lang="ch")

    file_path = './pdf2img'
    last_str = file_path.split('.')[-1]  # 检测待检测文件的尾缀
    if not os.path.exists('pdf2txt/'):  # 检测pdf2txt是否存在，没有则重新创建
       os.mkdir('pdf2txt/')
    if last_str == 'pdf':
        utils.pdf2png(file_path, v_res_dir='pdf2img')  # 将pdf转化为图片,待后续检测
        for root, dirs, files in os.walk('./pdf2img'):
            for i, file in enumerate(files):
                path = os.path.join(root, file)
                line_Y = utils.horizontal_line_detection(path)
                res = ocr.ocr(path, cls=True)
                txtSavePath = 'pdf2txt/' + os.path.basename(path).split('.')[0] + '.txt'

                utils.analyze_ocr(res, line_Y, txtSavePath)
                utils.pdf2excel()

    elif last_str.lower == 'jpg' or last_str.lower() == 'png' or last_str.lower() == 'jpeg':

        line_Y = utils.horizontal_line_detection(file_path)
        res = ocr.ocr(file_path, cls=True)

        txtSavePath = 'pdf2txt/' + os.path.basename(file_path).split('.')[0] + '.txt'

        utils.analyze_ocr(res, line_Y, txtSavePath)
        utils.pdf2excel()

    elif os.path.isdir(file_path):
        for root, dirs, files in os.walk(file_path):
            for i, file in enumerate(files):
                path = os.path.join(root, file)
                line_Y = utils.horizontal_line_detection(path)
                res = ocr.ocr(path, cls=True)
                txtSavePath = 'pdf2txt/' + os.path.basename(path).split('.')[0] + '.txt'

                utils.analyze_ocr(res, line_Y, txtSavePath)

                utils.pdf2excel()
