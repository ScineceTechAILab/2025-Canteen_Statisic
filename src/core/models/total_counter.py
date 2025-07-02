##### 作者：ESJIAN
# 日期: 2025.7.1
# 版本：v1.1
# 功能：
#   1. 为主表、福利表、子表主食表、子表副食表添加页计行





import __main__
import datetime
import sys
import os
import threading
import shutil
import multiprocessing
import subprocess
import time
import xlwings as xw
import PySide6
from PySide6.QtWidgets import QMessageBox

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt, QEvent)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform, Qt)
from PySide6.QtWidgets import (QAbstractScrollArea,QApplication, QButtonGroup, QFormLayout, QGridLayout,
    QGroupBox, QHBoxLayout, QLabel, QLayout,
    QLineEdit, QPlainTextEdit, QPushButton, QScrollArea,
    QSizePolicy, QSpinBox, QTabWidget, QVBoxLayout,
    QWidget, QFileDialog, QDialog, QVBoxLayout, QCheckBox)

from core.models.page_counter import get_first_blank_row_index

def counting_total_value(excel_type:str,work_book:xw.Book,work_sheet:xw.Sheet):
    """
    入库在将条目添加到表并进行日记、月计之后,为该页面加上页计行(v1.1 逻辑流版本)
    
    Parameters:
      excel_type: 正在处理的表的类型,值为"主表"、"福利表"、"子表主食表"、"子表副食表"
      work_book: 要修改的 Excel 传入workbook对象
      work_sheet: 要修改的 Excel 传入worksheet对象
    """
        
    "判断哪个表需要页计"
    if excel_type == "主表":

        if work_sheet.name in [s.name for s in work_book.sheets]:

            if work_sheet.name == "食堂物品收发存库存表":

                print(f"Notice: 开始为 {excel_type} {work_sheet.name} 页执行页计功能")

                "设定一页的行数"
                sheet_ratio = 25                                    # 主表中 食堂物品收发存库存表 表为26行一页

                "跳过已经写好了的表页，定位存有空行的一行，计算出其所在的页码"
                blank_row_index = get_first_blank_row_index(work_sheet) 
                page_index = int(blank_row_index / sheet_ratio) + 1 # int(13 / 26) = 0 但是是第一页，所以要加1  
                print(f"Notice: 发现 {excel_type} {work_sheet.name} 中第 {page_index} 页存在空行")

                "所有页的页计行是否在所有页的倒数二行"
                for i in range(1,page_index+1):
                    
                    # 若每一页倒数第二行为页计行
                    if work_sheet.range((i * sheet_ratio - 1, 1)).value.replace(" ","") == "页计":
                        continue
                    
                    # 若每一页倒数第二行存在非页计行
                    else:
                        print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页倒数第二行不是页计行,终止本次提交")
                        __main__.SAVE_OK_SIGNAL = False
                        return
                
                "累加本页以及前页所有页计行的值"
                page_item_sum = {"C":0,"D":0,"E":0,"F":0,"G":0,"H":0}
                for item in page_item_sum:
                    for i in range(1,page_index+1):
                    # 分别累加本页及之前页的页计行的C~H列的值
                        page_item_sum[item] += work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("C") + 1)).value

                "将这些值设为总计行的C~H列的值,放置于从前往后的第一行空行"
                work_sheet.range((blank_row_index, 1)).value = "总计"
                for item in page_item_sum:
                    work_sheet.range((blank_row_index, ord(item) - ord("C") + 1)).value = page_item_sum[item]

                print(f"Notice: 结束为 {excel_type} 中 {work_sheet.name} 进行总计")
                


                    
                    


    
                    