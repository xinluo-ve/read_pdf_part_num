import os
import re
import time

import pandas as pd
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
from PIL import Image, ImageOps, ImageEnhance
import pytesseract

def pdf_to_image(pdf_path, dpi=300):
    """PDF转图片"""
    images = convert_from_path(pdf_path, dpi=dpi)
    return images

def process_one_config(args):
    img, enhance_config, picture_number = args
    gray = img.convert('L')
    enhancer = ImageEnhance.Contrast(gray)
    gray_enhanced = enhancer.enhance(enhance_config)
    threshold = 150
    binary = gray_enhanced.point(lambda p: 255 if p > threshold else 0)
    # binary.save(f'part_number_img_{img_idx}_{picture_number}.png')
    data = pytesseract.image_to_data(binary, lang="eng", config="--psm 1", output_type=pytesseract.Output.DICT)
    n_boxes = len(data['text'])
    # 获取part number位置
    part_x1 = part_x2 = None
    width = None
    for i in range(n_boxes - 1):
        text = data['text'][i].strip().upper()
        next_text = data['text'][i + 1].strip().upper()
        if text[-3:] == "PART"[-3:] and next_text == 'NUMBER':
            part_x1 = data['left'][i]
            width = data['width'][i]
            part_x2 = data['left'][i + 1] + data['width'][i + 1]
            break
    if part_x1 is None or part_x2 is None:
        return []

    result = []
    right_x = part_x2 + width
    left_x = part_x1 - width
    # 匹配
    for i in range(n_boxes):
        text = data['text'][i].strip().upper()
        i_x1 = data['left'][i]
        i_x2 = data['left'][i] + data['width'][i]
        if i_x1 >= left_x and i_x2 <= right_x:
            new_text = part_number_match(picture_number, text)
            if new_text is not None and new_text not in result:
                print('part_numer:', new_text)
                result.append(new_text)
    return result

from multiprocessing import Pool, Manager
from PIL import ImageEnhance

def get_mult_part_number(img_idx, img, picture_number):
    enhance_configs = [1, 2, 5, 10]
    args_list = [(img, config, picture_number) for config in enhance_configs]

    with Pool(processes=4) as pool:  # 4个进程，可以根据CPU数调整
        results = pool.map(process_one_config, args_list)
    # 合并去重结果
    flat_result = list(set(r for lst in results for r in lst))
    return flat_result


def get_part_number(img_idx, img, picture_number):
    result = []
    gray = img.convert('L')  # 'L'模式是灰度
    for enhance_config in [1, 2, 5, 10]:
        # 2. 提升对比度（可选）
        enhancer = ImageEnhance.Contrast(gray)
        gray_enhanced = enhancer.enhance(enhance_config)  # 参数 >1 增强，1.5~2.0比较合适
        # 3. 简单二值化（阈值处理）
        threshold = 150
        binary = gray_enhanced.point(lambda p: 255 if p > threshold else 0)
        # binary.save(f'part_number_img_{img_idx}_{picture_number}.png')
        data = pytesseract.image_to_data(binary, lang="eng", config="--psm 1", output_type=pytesseract.Output.DICT)
        n_boxes = len(data['text'])
        # 获取part number位置
        part_x1 = part_x2 = None
        width = None
        for i in range(n_boxes - 1):
            text = data['text'][i].strip().upper()
            next_text = data['text'][i + 1].strip().upper()
            if text[-3:] == "PART"[-3:] and next_text == 'NUMBER':
                part_x1 = data['left'][i]
                width = data['width'][i]
                part_x2 = data['left'][i + 1] + data['width'][i + 1]
                break
        if part_x1 is None or part_x2 is None:
            continue
        right_x = part_x2 + width
        left_x = part_x1 - width
        # 匹配
        for i in range(n_boxes):
            text = data['text'][i].strip().upper()
            i_x1 = data['left'][i]
            i_x2 = data['left'][i] + data['width'][i]
            if i_x1 >= left_x and i_x2 <= right_x:
                new_text = part_number_match(picture_number, text)
                if new_text is not None and new_text not in result:
                    print('part_numer:', new_text)
                    result.append(new_text)
    return result


def find_revision_position(img):
    data = pytesseract.image_to_data(img, lang="eng", config="--psm 6", output_type=pytesseract.Output.DICT)
    n_boxes = len(data['text'])
    part_x1 = part_x2 = part_y1 = part_y2 = None
    for i in range(n_boxes - 1, -1, -1):
        text = data['text'][i].strip().upper()
        if text == "REVISION" or text == 'REV':
            part_x1 = data['left'][i]
            width = data['width'][i]
            part_x2 = part_x1 + width
            part_y2 = data['top'][i]
            break
    right_x = part_x2 + 50
    left_x = part_x1 - 50
    return left_x, right_x, part_y2


def crop_header_area(image, x1, x2):
    """裁剪表头区域"""
    cropped = image.crop((x1, 0, x2, image.height))
    return cropped


def clean_text(line):
    """清理掉行开头的符号，比如 | [ ] { } 空格"""
    return re.sub(r'^[\|\[\]\{\}\s]+', '', line)

def part_number_match(picture_number, input_data):
    if len(input_data.split('-')) != 3:
        return None
    picture_number1, picture_number2, picture_number3 = picture_number.split('-')
    cleaned_line1, cleaned_line2, cleaned_line3 = input_data.split('-')
    if picture_number1[-2:] == cleaned_line1[-2:] and picture_number2 == cleaned_line2:
        cleaned_line3 = cleaned_line3[:len(picture_number3)]
        new_input_data = picture_number1 + '-' + picture_number2 + '-' + cleaned_line3
        return new_input_data
    else:
        return None


def get_revision(img, file_prod):
    x1, x2, y2 = find_revision_position(img)
    cropped = img.crop((x1, y2, x2, img.height))
    cropped.save(f'revision_img_{0}_{file_prod}revision_img.png')
    text = pytesseract.image_to_string(cropped, lang='eng', config='--psm 1')
    lines = text.split('\n')
    result = []
    for line in lines:
        line = line.strip()
        if not line or 'REV' in line or "REVISION" in line:
            continue  # 跳过空行
        cleaned_line = clean_text(line)
        print(f"vision: {cleaned_line}")
        result.append(cleaned_line)
    return result


if __name__ == "__main__":
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

    rows = []
    error_file = []
    cols = ['图纸号', '页码', 'REVISION', 'PART NUMBER']
    error_cols = ['异常图纸号', '页码', '异常原因']
    s = time.time()
    for root, dirs, files in os.walk(r".\图纸"):
        for file_idx, file in enumerate(files):
            if file.split('-')[0][0] != '7':
                continue
            file_path = os.path.join(root, file)
            file_prod = file.split('.pdf')[0]
            print(f'当前运行pdf：{file_prod}', time.time() - s)
            images = pdf_to_image(file_path)
            # revision
            try:
                revision_result = get_revision(images[0], file_prod.split('_')[0])
                revision = (revision_result[0] if len(revision_result) == 1 else '')
            except:
                error_file.append([file_prod, 'vision错误'])
                print(f'{file_prod}, vision错误')
                revision = ''

            # PART NUMBER
            for img_idx, img in enumerate(images):
                try:
                    cleaned_line_list = []
                    cleaned_line_list.extend(get_mult_part_number(img_idx, img, file_prod.split('_')[0]))
                except:
                    error_file.append([file_prod, img_idx, 'part_number错误'])
                    cleaned_line_list = []
                    print(file_prod, img_idx, 'part_number错误')
                if len(cleaned_line_list) == 0:
                    error_file.append([file_prod, img_idx, 'part_number为空'])
                    print(file_prod, img_idx, 'part_number为空')
                for c in cleaned_line_list:
                    rows.append([file_prod, img_idx, revision, c])

    df = pd.DataFrame(rows, columns=cols)
    error_df = pd.DataFrame(error_file, columns=error_cols)
    with pd.ExcelWriter('PART_NUMBER_20250428_2.xlsx') as writer:
        df.to_excel(writer, sheet_name='PART_NUMBER')
        error_df.to_excel(writer, sheet_name='error')



