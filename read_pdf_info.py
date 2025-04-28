import os
import re

import pandas as pd
from pdf2image import convert_from_path
import pytesseract
from PIL import Image


def pdf_to_image(pdf_path, dpi=300):
    """PDF转图片"""
    images = convert_from_path(pdf_path, dpi=dpi)
    return images


def find_part_number_position(img):
    data = pytesseract.image_to_data(img, lang="eng", config="--psm 6", output_type=pytesseract.Output.DICT)
    n_boxes = len(data['text'])
    part_x1 = part_x2 = None
    width = None
    for i in range(n_boxes):
        text = data['text'][i].strip().upper()
        next_text = data['text'][i + 1].strip().upper()
        if text == "PART" and next_text == 'NUMBER':
            part_x1 = data['left'][i]
            width = data['width'][i]
            part_x2 = data['left'][i + 1] + data['width'][i + 1]
            break

    if not part_x1 and not part_x2:
        return None, None

    right_x = part_x2 + width
    # for right in range(i, n_boxes):
    #     if data['text'][right] == '|':
    #         if data['left'][right] <= right_x:
    #             right_x = data['left'][right]
    #         break

    left_x = part_x1 - width
    # for left in range(i, 0, -1):
    #     if data['text'][left] == '|':
    #         print(left)
    #         if data['left'][left] >= left_x:
    #             left_x = data['left'][left]
    #         break
    return left_x, right_x


def find_revision_position(img):
    data = pytesseract.image_to_data(img, lang="eng", config="--psm 6", output_type=pytesseract.Output.DICT)
    n_boxes = len(data['text'])
    part_x1 = part_x2 = part_y1 = part_y2 = None
    i = 0
    for i in range(n_boxes - 1, -1, -1):
        text = data['text'][i].strip().upper()
        if text == "REVISION" or text == 'REV':
            part_x1 = data['left'][i]
            width = data['width'][i]
            part_x2 = part_x1 + width
            part_y2 = data['top'][i]
            break

    right_x = part_x2 + 50
    # for right in range(i, n_boxes):
    #     if data['text'][right] == '|':
    #         print(right)
    #         right_x = data['left'][right]
    #         break

    left_x = part_x1 - 50
    # for left in range(i, 0, -1):
    #     if data['text'][left] == '|':
    #         print(left)
    #         if data['left'][left] <= part_x1:
    #             left_x = data['left'][left]
    #         break
    return left_x, right_x, part_y2


def crop_header_area(image, x1, x2):
    """裁剪表头区域"""
    cropped = image.crop((x1, 0, x2, image.height))
    return cropped


def clean_text(line):
    """清理掉行开头的符号，比如 | [ ] { } 空格"""
    return re.sub(r'^[\|\[\]\{\}\s]+', '', line)


def get_part_number(img, picture_number):
    x1, x2 = find_part_number_position(img)
    result = []
    if not x1 and not x2:
        return result

    cropped = img.crop((x1, 0, x2, img.height))
    cropped.save('part_number_img.png')
    text = pytesseract.image_to_string(cropped, lang='eng', config='--psm 1')
    lines = text.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue  # 跳过空行
        cleaned_line = clean_text(line)
        print(f"行内容: {cleaned_line}")
        if len(cleaned_line.split('-')) != 3:
            continue
        # 提取料号
        picture_number1, picture_number2, picture_number3 = picture_number.split('-')
        cleaned_line1, cleaned_line2, cleaned_line3 = cleaned_line.split('-')
        if picture_number1 in cleaned_line.split('-')[0] and picture_number2 ==  cleaned_line2:
            cleaned_line3 = cleaned_line3[:len(picture_number3)]
            new_cleaned_line = picture_number1 + '-' + picture_number2 + '-' + cleaned_line3
            result.append(new_cleaned_line)
            print(new_cleaned_line)
        # match = re.search(r'\b(\d{3}-\d{3})-\d{4}\b', cleaned_line)
        #
        # if match:
        #     found_picture_number = match.group(1)  # 提取948-000
        #     # 判断是否和传入的picture_number一致
        #     if found_picture_number == picture_number:
        #         print({'cleaned_line': cleaned_line})
        #         result.append(cleaned_line)
    return result


def get_revision(img):
    x1, x2, y2 = find_revision_position(img)
    cropped = img.crop((x1, y2, x2, img.height))
    cropped.save('revision_img.png')
    text = pytesseract.image_to_string(cropped, lang='eng', config='--psm 6')
    lines = text.split('\n')
    result = []
    for line in lines:
        line = line.strip()
        if not line:
            continue  # 跳过空行
        cleaned_line = clean_text(line)
        print(f"行内容: {cleaned_line}")
        result.append(cleaned_line)
    return result


if __name__ == "__main__":
    picture_number = '121-400-7000'
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    file_path = r'C:\Users\hf\Documents\WeChat Files\wxid_l44618or5fh522\FileStorage\File\2025-04\121-400-9000_CG.pdf'
    images = pdf_to_image(file_path)
    # revision
    revision_result = get_revision(images[0])
    part_number_result = get_part_number(images[0], picture_number)
    print(revision_result, part_number_result)
    exit()



    rows = []
    error_file = []
    cols = ['料号', 'REVISION', 'PART NUMBER']
    error_cols = ['异常pdf']
    for root, dirs, files in os.walk(r"D:\data\desktop\pdf图纸编辑修改\图纸"):
        for file in files:
            file_path = os.path.join(root, file)
            file_prod = file.split('.pdf')[0]
            print(file_prod)
            try:
                picture_number = file_prod.split('-')[0] + '-' + file_prod.split('-')[1]
                images = pdf_to_image(file_path)
                # revision
                revision_result = get_revision(images[0])
                vision = (revision_result[0] if len(revision_result) != 0 else '')

                # PART NUMBER
                cleaned_line_list = []
                for img in images:
                    cleaned_line_list.extend(get_part_number(img, picture_number))

                if len(cleaned_line_list) == 0:
                    error_file.append(file_prod)
                    continue
                for cleaned_line in cleaned_line_list:
                    rows.append([file_prod, vision, cleaned_line])
            except:
                error_file.append(file_prod)

    df = pd.DataFrame(rows, columns=cols)
    error_df = pd.DataFrame(error_file, columns=error_cols)
    with pd.ExcelWriter('PART_NUMBER_20250428.xlsx') as writer:
        df.to_excel(writer, sheet_name='PART_NUMBER')
        error_df.to_excel(writer, sheet_name='error')




