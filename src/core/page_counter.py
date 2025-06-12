


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


def counting_page_value(page_counter_signal:bool,excel_type:str,work_book,work_sheet):
    """
    在将条目添加到表之后,为该页面加上页计行
    
    Parameters:
      page_counter_signal: 页计功能是否开启的信号，值为 True 或 False
      excel_type: 正在处理的表的类型,值为"主表"、"福利表"
      work_book: 要修改的 Excel 传入工作簿对象
      work_sheet: 要修改的 Excel 传入工作表对象
    """
    if page_counter_signal == True:
        
        if excel_type == "主表":

            if work_sheet.name in [s.name for s in work_book.sheets]:

                print("Notice: 开始为主表 {sheet_name} 页执行页计功能".format(sheet_name=work_sheet.name))
            
                if work_sheet.name == "食堂物品收发存库存表":

                    # 定位存有空行的一行，计算出其所在的页码
                    blank_row_index = get_blank_row_index(work_sheet) 
                    page_index = int(blank_row_index / 26) + 1 # 主表中该表为26行一页，int(13 / 26) = 0 但是是第一页，所以要加1  
                
                    #TODO

                else:

                    # 定位存有空行的一行，计算出其所在的页码
                    blank_row_index = get_blank_row_index(work_sheet)
                    page_index = int(blank_row_index / 33) + 1 # 主表中的其他主食表皆为 33 行一页
                
                    #TODO



            else:
                print("Error: 主表中 {sheet_name} 页不存在,跳过执行页计功能".format(sheet_name=work_sheet.name))
                return

        elif excel_type == "福利表":

            if  work_sheet.name in [s.name for s in work_book.sheets]:
                
                print("Notice: 开始为福利表 {sheet_name} 页执行页计功能".format(sheet_name=work_sheet.name))

                # 定位存有空行的一行，计算出其所在的页码
                blank_row_index = get_blank_row_index(work_sheet)
                page_index = int(blank_row_index / 32) + 1 # 福利表中的所有表皆为 32 行一页
                
                #TODO



            else:
                print("Error: 福利表中 {sheet_name} 页不存在,跳过执行页计功能".format(sheet_name=work_sheet.name))
                return
            
        else:
            print("Error: 未知的Excel_Type值: {Excel_Type} 跳过执行页计功能".format(Excel_Type=excel_type))
            return
        
    
    elif page_counter_signal == False:
        print("Notice: 页计功能被关闭，跳过为{Excel_Type} 执行页计".format(Excel_Type=excel_type))
        return
    
    else:
        print("Error: PAGE_COUNTER_SIGNAL值错误,请检查代码逻辑")
        return
    
def get_blank_row_index(work_sheet):
    """
    定位在表中出现的第一个有效空行(在某页内出现的第一个空行不包括页与页间的),返回该空行的行索引(1索引格式)
    
    Parameters:
        work_sheet: 要进行检测的 xlwings sheet 对象
    
    Returns:
        effective_row_index:有效空行索引(1索引格式)
    """
    # 查找第一行空行，记录下空行行标（从表格的第二行开始）


    for row_index in range(0, work_sheet.used_range.rows.count):
        if work_sheet.range((row_index + 1, 1)).value is None and row_index != 0:
            # 检查前一行是否包含“领导”二字
            if row_index > 0:
                previous_row_values = [
                str(work_sheet.range((row_index, col)).value).strip()
                for col in range(1, work_sheet.used_range.columns.count + 1)
                if work_sheet.range((row_index, col)).value is not None
                ]
                if any("领导" in value for value in previous_row_values):
                    print(f"Notice: 第 {row_index} 行包含“领导”二字，继续查找下一行")
                    continue

            # 检查当前列的前几行是否包含“序号”二字
            column_values = [
                str(work_sheet.range((row, 1)).value).strip()
                for row in range(1, row_index + 1)
                if work_sheet.range((row, 1)).value is not None
            ]
            if not any("序号" in value for value in column_values):
                print(f"Notice: 前 {row_index} 行未找到“序号”二字，继续查找下一行")
                continue
            break
        
        effective_row_index = row_index + 1  # 转换为1索引格式

        return effective_row_index


def get_page_item_indexes():
    """
    统计当前页有真正物品登记的条目索引,返回该页的条目索引的一维列表
    
    Returns:
        page_item_indexes: 当前页的条目索引列表
    """

if __name__ == "__main__":
