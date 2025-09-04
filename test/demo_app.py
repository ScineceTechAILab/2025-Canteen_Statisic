import os
import sys

def get_resource_path(relative_path):
    try:
        # 如果程序是打包后的 .exe 文件
        base_path = sys._MEIPASS
    except Exception:
        # 如果程序是在开发环境中运行
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# 打印工作目录，帮助调试
print("当前工作目录：", os.getcwd())

# 定义相对路径
MAIN_WORK_EXCEL_FOLDER = get_resource_path("src/data/storage/work/主表/")
SUB_WORK_EXCEL_FOLDER = get_resource_path("src/data/storage/work/子表/")

print("主工作表路径:", MAIN_WORK_EXCEL_FOLDER)
print("子工作表路径:", SUB_WORK_EXCEL_FOLDER)
