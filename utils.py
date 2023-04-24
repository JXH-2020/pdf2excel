# -*- coding: utf-8 -*-
#############################################################################
# Copyright (c) 2023  - Shanghai Davis Tech, Inc.  All rights reserved
"""
文件名: utils.py
说明: pdf2excel功能实现的工具集合
2022-04-17: 江绪好, Davy @Davis Tech
"""
import os
import fitz
import cv2
import numpy as np
import pandas as pd
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter


def pdf2png(v_filename, v_res_dir='pdf2img'):
    """
    pdf2png(v_filename, v_res_dir='')
    .   @brief 将pdf文件按页转为png图片
    .   @param v_filename: 指定原始pdf文件，如："pdf/1.pdf"
    .   @param v_res_dir: 指定页png保存路径，默认为‘’
    """
    print("pdf2png开始处理 {}".format(v_filename))
    if not os.path.exists(v_res_dir):  # 检测pdf2img是否存在，没有则重新创建
        os.mkdir(v_res_dir)
    f_pdf = fitz.open(v_filename)
    i_pgen = f_pdf.page_count
    for i in range(i_pgen):
        n_png = os.path.join(v_res_dir, '{}.png'.format(i))
        f_pdf[i].get_pixmap(matrix=fitz.Matrix(4, 4)).save(n_png)
    f_pdf.close()


def horizontal_line_detection(v_filename):
    """
    horizontal_line_detection(v_filename, v_res_dir='')
    .   @brief 识别图像中的横线，并返回横线纵坐标
    .   @param v_filename: 待检测的图像地址
    .   @return line_Y: 返回由每行横线组成的纵坐标集合
    """

    # 加载图像
    img = cv2.imread(v_filename, 1)

    # 灰度图片
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    binary = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 35, -5)

    rows, cols = binary.shape
    scale = 40
    # 识别横线:
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (cols // scale, 1))
    eroded = cv2.erode(binary, kernel, iterations=1)
    dilated_col = cv2.dilate(eroded, kernel, iterations=1)

    lines = cv2.HoughLinesP(dilated_col, 1, np.pi / 180, 100, minLineLength=100, maxLineGap=50)

    # 筛选出长度符合要求的横线
    half_length = img.shape[1] // 2
    filtered_lines = []
    for line in lines:
        x1, y1, x2, y2 = line[0]
        length = np.sqrt((y2 - y1) ** 2 + (x2 - x1) ** 2)
        if abs(y1 - y2) < 5 and half_length < length < half_length * 2:
            filtered_lines.append(line)

    # 按照纵坐标排序
    lines = sorted(filtered_lines, key=lambda x: x[0][1])

    # 删除纵向距离很小的直线
    filtered_lines = []
    prev_line = None
    for line in lines:
        if prev_line is None:
            prev_line = line
            continue
        x1, y1, x2, y2 = line[0]
        px1, py1, px2, py2 = prev_line[0]
        distance = abs(y1 - py2)
        if distance > 10:
            filtered_lines.append(prev_line)
        prev_line = line
    filtered_lines.append(prev_line)

    line_Y = []
    for line in filtered_lines:
        x1, y1, x2, y2 = line[0]
        cv2.line(img, (x1, y1), (x2, y2), (0, 0, 255), 1)
        line_Y.append((y1 + y2) / 2)

    return line_Y


def find_interval(v_num, v_intervals):
    """
    .   @brief 寻找 num 所在的区间
    .   @param v_num: 要查找的数
    .   @param v_intervals: 区间列表，每个元素为 (start, end)
    .   @return num 所在的区间索引，如果没有找到则返回 -1
    """
    for i, interval in enumerate(v_intervals):
        if interval[0] <= v_num <= interval[1]:
            return i
    return -1


# 定义排序规则：按照元组的第二个元素排序
def sort_by_second_item(item):
    return item[1][0]


def analyze_ocr(v_res, v_line_Y, v_savefile_path):
    """
    analyze_ocr(v_res, v_line_Y, v_savefile_path)
    .   @brief 根据识别到的横线纵坐标以及ocr识别出结果中的坐标，对result进行解析，并以.txt文件方式保存到./pdf2txt/文件夹下
    .   @param v_res: ocr识别的结果
    .   @param v_line_Y: 横线纵坐标集合
    .   @param v_savefile_path: 待保存解析结果txt的位置
    """

    vertical_interval = []
    for i, value in enumerate(v_line_Y):
        if i == 0:
            pass
        elif i < len(v_line_Y) - 1:
            vertical_interval.append((value, v_line_Y[i + 1]))
    result = []  # 中心点(x,y),行数-1,ocr结果
    # ---------------------------------------根据返回的line_Y 将ocr结果分割 ---------------------------------------------#
    for i, item in enumerate(v_res):
        item_y = (v_res[i][0][0][1] + v_res[i][0][2][1]) / 2
        item_x = (v_res[i][0][0][0] + v_res[i][0][2][0]) / 2
        index = find_interval(item_y, vertical_interval)
        if index != -1:
            result.append((index, (item_x, item_y), v_res[i][1][0]))

    # ---------------------------------------对每一大行里 进行细小分行处理---------------------------------------------#
    large_lines = []
    small_lines = []
    prev_line = None
    for i, line in enumerate(result):
        if prev_line is None:
            prev_line = line
            continue
        x, y = line[1]
        px, py = prev_line[1]
        distance = abs(y - py)
        if distance > 10:
            small_lines.append(prev_line)
            large_lines.append(small_lines)
            small_lines = []
        else:
            small_lines.append(prev_line)
        prev_line = line
        if i == len(result) - 1:
            small_lines = [line]
            large_lines.append(small_lines)
    # -------------------------------对细小分行中没有识别到的单元格进行再识别操作--------------------------------------#
    # -------------------------------主要还是每一大行中的第一行第二行重要信息识别--------------------------------------#

    # ---------------------------------------对细小分行进行x轴排序处理---------------------------------------------#
    lines = list(map(lambda x: sorted(x, key=sort_by_second_item), large_lines))
    title = ''
    line_item = []
    text = ''  # item之间以\t分割
    for i, item in enumerate(lines):
        if i == 0 or i == 1:
            for im in item:
                if '/' in im[2]:
                    title = title + im[2].split('/')[0] + '\t' + im[2].split('/')[1] + '\t'

                else:
                    title = title + im[2] + '\t'

        elif item[0][2].strip() == 'Ubertrag':  # 若小行里含有“Ubertrag"字符 直接跳过这一行
            pass
        else:
            if i < len(lines) - 1:
                if item[0][0] == lines[i + 1][0][0]:
                    for j, im in enumerate(item):
                        if j == 1:
                            if "48 Std." in item[1][2]:
                                text = text + item[1][2].split("48 Std.")[0].split(' ')[0] + '\t' + \
                                       item[1][2].split("48 Std.")[0].split(' ')[1] + '\t' + "48 Std.\t "
                            elif "Ter.12:00" in item[1][2]:
                                text = text + item[1][2].split("Ter.12:00")[0].split(' ')[0] + '\t' + \
                                       item[1][2].split("Ter.12:00")[0].split(' ')[1] + '\t' + "Ter.12:00\t "
                            else:
                                text = text + item[j][2] + '\t'
                        elif j == 2:
                            if "48 Std" in item[2][2]:
                                text = text + item[2][2].split("48 Std")[0] + '\t' + "48 Std.\t"
                            else:
                                text = text + item[j][2] + '\t'

                        else:
                            text = text + item[j][2] + '\t'
                else:
                    for j, im in enumerate(item):
                        if j == 2:
                            if "48 Std." in item[2][2]:
                                text = text + item[2][2].split("48 Std.")[0] + '\t' + "48 Std.\t"
                        else:
                            text = text + item[j][2] + '\t'
                    line_item.append(text)
                    text = ''
            else:
                for j, im in enumerate(item):
                    if j == 2:
                        if "48 Std." in item[2][2]:
                            text = text + item[2][2].split("48 Std.")[0] + '\t' + "48 Std.\t"
                    else:
                        text = text + item[j][2] + '\t'
                line_item.append(text)
                text = ''
    title = title + 'bemerkungen'  # 最后一列 备注（德语）
    with open(v_savefile_path, 'w+', encoding='utf-8') as file:
        file.write(title)
        file.write('\n')
        for item in line_item:
            file.write(item)
            file.write('\n')


def pdf2excel():
    """
    pdf2excel()
    .   @brief 将./pdf2txt目录下的txt中的数据转为excel
    """
    for root, dirs, files in os.walk('./pdf2txt'):
        for i, file in enumerate(files):
            all_data = []
            line_data = []
            last = ''
            path = os.path.join(root, file)
            with open(path, 'r', encoding='utf-8') as FILE:
                lines = FILE.readlines()
                for line in lines:
                    items = line.split('\t')
                    for i, item in enumerate(items):
                        if len(items) - 1 > i > 11:
                            last = last + item + ' | '
                            if i == len(items) - 2:
                                line_data.append(last)
                        else:
                            if item == '\n':
                                continue
                            else:
                                line_data.append(item)
                    all_data.append(line_data)
                    line_data = []
                    last = ''

            # 将数据列表和列名称传入DataFrame中
            df = pd.DataFrame(all_data[1:], columns=all_data[0])
            v_res_dir = './pdf2excel/'
            if not os.path.exists(v_res_dir):  # 检测pdf2excel是否存在，没有则重新创建
                os.mkdir(v_res_dir)
            # 将数据写入Excel文件
            with pd.ExcelWriter(v_res_dir + file.split('.')[0] + '.xlsx', engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Sheet1', index=False)
                worksheet = writer.sheets['Sheet1']

                # 遍历每列，找到最长的值并设置列宽度
                for col in worksheet.columns:
                    column = [cell for cell in col if cell.value]
                    if column:
                        column_width = max(len(str(cell.value)) for cell in column)
                        col_letter = get_column_letter(col[0].column)
                        worksheet.column_dimensions[col_letter].width = column_width + 2
                        for cell in column:
                            cell.alignment = Alignment(horizontal='center')
