


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
                    blank_row_index = get_first_blank_row_index(work_sheet) 
                    sheet_ratio = 26                                    # 主表中该表为26行一页
                    page_index = int(blank_row_index / sheet_ratio) + 1 # int(13 / 26) = 0 但是是第一页，所以要加1  
                    # 统计该页范围内除了 "日记"、"页计"等的行
                    page_item_indexes = get_page_item_indexes(work_sheet,page_index,sheet_ratio)
                    # 检测该页的倒数第二行是否为空
                    if work_sheet.range((page_index * sheet_ratio - 1, 1)).value is None:
                        print("Notice: {sheet_name} 页的倒数第二行为空，跳过页计功能".format(sheet_name=work_sheet.name))
                        return
                    else:
                        print("Notice: {sheet_name} 页的倒数第二行不为空，继续执行页计功能".format(sheet_name=work_sheet.name))
                        
                        # 检测为空时是否是页做，自
                        if work_book.sheets[work_sheet.name].range((page_index * sheet_ratio, 1)).value is None:
                            print("Notice: {sheet_name} 页的页计行为空，继续执行页计功能".format(sheet_name=work_sheet.name))
                            
                    

                else:

                    # 定位存有空行的一行，计算出其所在的页码
                    blank_row_index = get_first_blank_row_index(work_sheet)
                    sheet_ratio = 33                                    # 主表中其他主食表皆为 33 行一页
                    page_index = int(blank_row_index / sheet_ratio) + 1 

                    #TODO



            else:
                print("Error: 主表中 {sheet_name} 页不存在,跳过执行页计功能".format(sheet_name=work_sheet.name))
                return

        elif excel_type == "福利表":

            if  work_sheet.name in [s.name for s in work_book.sheets]:
                
                print("Notice: 开始为福利表 {sheet_name} 页执行页计功能".format(sheet_name=work_sheet.name))

                # 定位存有空行的一行，计算出其所在的页码
                blank_row_index = get_first_blank_row_index(work_sheet)
                sheet_ratio = 32                                    # 福利表中所有表皆为 32 行一页
                page_index = int(blank_row_index / sheet_ratio) + 1 
                
                # 统计该页范围内除了 "日记"、"页计"等的行
                page_item_indexes = get_page_item_indexes(work_sheet,page_index,sheet_ratio)
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
    
def get_first_blank_row_index(page_counter_signal:bool,excel_type:str,work_book,work_sheet):
    """
    定位在表中出现的第一个有效空行(在某页内出现的第一个空行不包括页与页间的),返回该空行的行索引(1索引格式)
    
    Parameters:
        work_sheet: 要进行检测的 xlwings sheet 对象
    
    Returns:
        effective_row_index:有效空行索引(1索引格式)
    """
    # 查找第一行空行，记录下空行行标（从表格的第二行开始）

    effective_row_index = 0  # 初始化有效行索引

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


def get_page_item_indexes(work_sheet,page_index:int,sheet_ratio:int):
    """
    统计当前页有物品登记行的条目索引,返回存储该页的商品行索引
    
    Parameters:
        work_sheet: 要进行统计的 xlwings sheet 对象
        page_index: 当前页的页码(1索引格式)
        sheet_ratio: 相应 Sheet 中每页的行数(1索引格式)，例如主表中自购主食入库每页33行，福利表中每页32行

    Returns:
        page_item_indexes: 当前页的条目索引列表(1索引格式)
    """
    page_item_indexes = []  # 初始化当前页的条目索引列表

    # 计算当前页的起始行和结束行(1索引格式)
    start_row = (page_index - 1) * sheet_ratio + 1  # 每页起始行，带入 page_index =1,sheet_ratio=33: (1-1)*33 + 1 = 1 
    end_row = page_index * sheet_ratio              # 每页结束行, sheet_ratio 行，带入 page_index =1,sheet_ratio=33: 1*33 = 33

    for row_index in range(start_row, end_row + 1): # 由于range()函数的结束值是不包含的，所以要加1
        
        # 观察到非"日记"、"页计"等行的第一列都是字符，利用强制类型转换会报错的特性筛选出有效行
        try:
            int(work_sheet.range((row_index, 1)).value)                  
            page_item_indexes.append(row_index)                      
            continue
        
        # 如果强制转换失败，说明该行不是有效的条目行
        except:    
            continue

    return page_item_indexes 





if __name__ == "__main__":
