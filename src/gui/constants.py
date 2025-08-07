# -*- coding: utf-8 -*-
"""
常量定义文件
"""

import os

# 获取当前文件的绝对路径
current_file_path = os.path.abspath(__file__)
# 获取项目根目录
project_root = os.path.abspath(os.path.join(current_file_path, '..', '..', '..'))

# 文件路径常量
TEMP_SINGLE_STORAGE_EXCEL_PATH = os.path.join("src", "data", "input", "manual", "temp_manual_input_data.xls")
TEMP_SINGLE_STORAGE_EXCEL_PATH2 = os.path.join("src", "data", "input", "manual", "temp_manual_input_data2.xls")
PHOTO_TEMP_SINGLE_STORAGE_EXCEL_PATH = os.path.join("src", "data", "input", "manual", "temp_img_input.xlsx")
PHOTO_TEMP_SINGLE_STORAGE_EXCEL_PATH2 = os.path.join("src", "data", "input", "manual", "temp_img_input.xls")

TEMP_IMAGE_DIR = os.path.join(".", "src", "data", "input", "img")

MAIN_WORK_EXCEL_FOLDER = "src\\data\\storage\\work\\主表\\"
SUB_WORK_EXCEL_FOLDER = "src\\data\\storage\\work\\子表\\"
WELFARE_EXCEL_FOLDER = "src\\data\\storage\\work\\福利表\\"
ITEM_EXCEL_FOLDER = "src\\data\\storage\\work\\条目表\\"

# 使用绝对路径
MAIN_WORK_EXCEL_FOLDER = os.path.join(project_root, MAIN_WORK_EXCEL_FOLDER)
SUB_WORK_EXCEL_FOLDER = os.path.join(project_root, SUB_WORK_EXCEL_FOLDER)
WELFARE_EXCEL_FOLDER = os.path.join(project_root, WELFARE_EXCEL_FOLDER)
ITEM_EXCEL_FOLDER = os.path.join(project_root, ITEM_EXCEL_FOLDER)

# OCR模型路径
OCR_MODEL_PATH = "./src/.paddleocr"

# 其他常量
TOTAL_FIELD_NUMBER = 10  # 录入信息总条目数
TEMP_STORAGED_NUMBER_LISTS = 1  # 初始编辑条目索引号
TEMP_LIST_ROLLBACK_SIGNAL = True  # 信号量，标记是否需要回滚

# 模式相关
MODE = 0  # 0表示入库，1表示出库

# 功能开关常量
ADD_DAY_SUMMARY = False
ADD_MONTH_SUMMARY = False
ADD_PAGE_SUMMARY = False
ADD_TOTAL_SUMMARY = False
ONLY_WELFARE_TABLE = False
SAVE_OK_SIGNAL = True
PAGE_COUNTER_SIGNAL = True
SERIALS_NUMBER = 1
DEBUG_SIGN = True