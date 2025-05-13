import os
import xlrd
from pyperclip import copy

# 获取当前目录下所有 .xls 文件
xls_files = [f for f in os.listdir('.') if f.endswith('.xls')]

for file in xls_files:
    try:
        workbook = xlrd.open_workbook(file)
        sheet_names = workbook.sheet_names()
        copy(str(sheet_names))
        print(f"{file} 的 sheets: {sheet_names}")
    except Exception as e:
        print(f"{file} 打开失败: {e}")

a = input()