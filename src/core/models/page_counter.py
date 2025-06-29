##### 作者：ESJIAN
# 日期: 2025.6.21
# 版本：v1.1
# 功能：
#   1. 为主表、福利表、子表主食表、子表副食表添加页计行


#### 记录：
# 日期: 2025.6.21
# 作者：ESJIAN
# 内容
#   1. 完成了福利表的逻辑
#   2. 完成了子表主食表的逻辑
#   3. 完成了子表副食表的逻辑
# 要做：
#   1. 代码融合到 main_window.py 的测试



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


def counting_page_value(excel_type:str,work_book:xw.Book,work_sheet:xw.Sheet):
    """
    在将条目添加到表之后,为该页面加上页计行(v1.1 逻辑流版本)
    
    Parameters:
      excel_type: 正在处理的表的类型,值为"主表"、"福利表"、"子表主食表"、"子表副食表"
      work_book: 要修改的 Excel 传入workbook对象
      work_sheet: 要修改的 Excel 传入worksheet对象
    """
        
    "判断哪个表需要页计"
    if excel_type == "主表":

        if work_sheet.name in [s.name for s in work_book.sheets]:
        
            if work_sheet.name == "食堂物品收发存库存表":

                print("Notice: 开始为 {excel_type} {sheet_name} 页执行页计功能".format(sheet_name=work_sheet.name))

                "设定一页的行数"
                sheet_ratio = 26                                    # 主表中 食堂物品收发存库存表 表为26行一页

                "跳过已经写好了的表页，定位存有空行的一行，计算出其所在的页码"
                blank_row_index = get_first_blank_row_index(work_sheet) 
                page_index = int(blank_row_index / sheet_ratio) + 1 # int(13 / 26) = 0 但是是第一页，所以要加1  
                print("Notice: 当前页码为: {page_index}".format(page_index=page_index))

                "统计该页范围内除了 日记、页计等的行"
                page_item_indexes = get_page_item_indexes(work_sheet,page_index,sheet_ratio)
                print("Notice: 该页范围内除了 日记、页计等的行号为: {page_item_indexes}".format(page_item_indexes=page_item_indexes))

                "累加每一页除了日计、月计、总计每一项 J 列的值"
                page_item_sum = 0
                for item_index in page_item_indexes:
                    item_value = work_sheet.range((item_index, 10)).value
                    if isinstance(item_value, (int, float)):
                        page_item_sum += item_value
                print("Notice: 该页范围内除了 日记、页计等的行号 J 列的值总和为: {page_item_sum}".format(page_item_sum=page_item_sum))

                "将该值设置为页计行 J 列的新值"
                print("Notice: 将页计行 J 列的当前值为: {page_item_sum}".format(page_item_sum=work_sheet.range((page_index * sheet_ratio - 2, 10)).value))
                work_sheet.range((page_index * sheet_ratio - 2, 10)).value = page_item_sum
                print("Notice: 已将页计行 J 列的新值设置为: {page_item_sum}".format(page_item_sum=page_item_sum))

            elif work_sheet.name in ["自购主食入库","食堂副食入库","厂调面食入库","扶贫主食入库","扶贫副食入库"]:
                
                print("Notice: 开始为 {excel_type} {sheet_name} 页执行页计功能".format(sheet_name=work_sheet.name))

                "设定一页的行数"
                sheet_ratio = 33                                    # 主表中 其他主食表 皆为 33 行一页
                "跳过已经写好了的表页，定位存有空行的一行，计算出其所在的页码"
                blank_row_index = get_first_blank_row_index(work_sheet)
                page_index = int(blank_row_index / sheet_ratio) + 1 
                print("Notice: 当前页码为: {page_index}".format(page_index=page_index))

                "统计该页范围内除了 日记、页计等的行"
                page_item_indexes = get_page_item_indexes(work_sheet,page_index,sheet_ratio)
                print("Notice: 该页范围内除了 日记、页计等的行号为: {page_item_indexes}".format(page_item_indexes=page_item_indexes))

                "累加每一页除了日计、月计、总计每一项 J 列的值"
                page_item_sum = 0
                for item_index in page_item_indexes:
                    item_value = work_sheet.range((item_index, 10)).value
                    if isinstance(item_value, (int, float)):
                        page_item_sum += item_value
                print("Notice: 该页范围内除了 日记、页计等的行号 J 列的值总和为: {page_item_sum}".format(page_item_sum=page_item_sum))

                "将该值设置为页计行 J 列的新值"
                print("Notice: 将页计行 J 列的当前值为: {page_item_sum}".format(page_item_sum=work_sheet.range((page_index * sheet_ratio - 2, 10)).value))
                work_sheet.range((page_index * sheet_ratio - 2, 10)).value = page_item_sum
                print("Notice: 已将页计行 J 列的新值设置为: {page_item_sum}".format(page_item_sum=page_item_sum))

        else:
            print("Error: {excel_type} 中 {sheet_name} 页不存在,跳过执行页计功能".format(sheet_name=work_sheet.name))
            return

    elif excel_type == "福利表":

        if  work_sheet.name in [s.name for s in work_book.sheets]:
            
            print("Notice: 开始为 {excel_type} {sheet_name} 页执行页计功能".format(sheet_name=work_sheet.name))

            "设定一页的行数"
            sheet_ratio = 32                                    # 福利表中该表为32行一页

            "跳过已经写好了的表页，定位存有空行的一行，计算出其所在的页码"
            blank_row_index = get_first_blank_row_index(work_sheet) 
            page_index = int(blank_row_index / sheet_ratio) + 1 # int(13 / 32) = 0 但是是第一页，所以要加1  
            
            "统计该页范围内除了 日记、页计等的行"
            page_item_indexes = get_page_item_indexes(work_sheet,page_index,sheet_ratio)
                            
            "累加每一页除了日计、月计、总计每一项 J 列的值"
            page_item_sum = 0
            for item_index in page_item_indexes:
                item_value = work_sheet.range((item_index, 10)).value
                if isinstance(item_value, (int, float)):
                    page_item_sum += item_value
            
            "将该值设置为页计行 J 列的新值"
            work_sheet.range((page_index * sheet_ratio - 2, 10)).value = page_item_sum

        else:
            print("Error: {excel_type} 中 {sheet_name} 页不存在,跳过执行页计功能".format(sheet_name=work_sheet.name))
            return

    elif excel_type == "子表主食表":
    
        if  work_sheet.name in [s.name for s in work_book.sheets]:

            print("Notice: 开始为 {excel_type} {sheet_name} 页执行页计功能".format(sheet_name=work_sheet.name))
            
            "设定一页的行数"
            sheet_ratio = 33                                    # 子表主食表中该表为33行一页

            "跳过已经写好了的表页，定位存有空行的一行，计算出其所在的页码"
            blank_row_index = get_first_blank_row_index(work_sheet) 
            page_index = int(blank_row_index / sheet_ratio) + 1 # int(13 / 33) = 0 但是是第一页，所以要加1  
            
            "统计该页范围内除了 日记、页计等的行"
            page_item_indexes = get_page_item_indexes(work_sheet,page_index,sheet_ratio)
                            
            "取所有行中行号最大的一行的 库存 列组的 数量(J列)、金额(K列) 列作为页计行 库存 列组的 数量(J列)、金额(K列) 列的值"
            max_row_index = max(page_item_indexes)
            page_item_sum = work_sheet.range((max_row_index, 10)).value                # 获取数量(J列)的值
            amount_value = work_sheet.range((max_row_index, 11)).value                 # 获取金额(K列)的值
            work_sheet.range((page_index * sheet_ratio - 2, 10)).value = page_item_sum # 设置页计行数量(J列)的新值
            work_sheet.range((page_index * sheet_ratio - 2, 11)).value = amount_value  # 设置页计行金额(K列)的新值

            "根据 page_item_indexes 存储行摘要列(D列)的字符值(入库/出库)将统计到的行以 {类型(入库/出库):[行号一维列表]}的方式重新打包为一个字典"
            page_item_types = {"入库": [], "出库": []}  # 初始化一个字典来存储类型和对应的行号列表
            
            for item_index in page_item_indexes:       
                
                item_type = work_sheet.range((item_index, 4)).value # 获取该行的摘要列(D列)的值
                
                if item_type == "入库":                             
                    page_item_types[item_type].append(item_index)   # 将行号添加到字典的"入库"键值对的值中
                
                elif item_type == "出库":
                    page_item_types[item_type].append(item_index)   # 将行号添加到字典的"出库"键值对的值中
                
                else:
                    print(f"Error: 在行 {item_index} 中发现非 入库/出库 的未知类型: {item_type}，已终止为{excel_type} 执行页计")
                    return

            "分别累加类型为 “入库” 的每一行的 “数量” 列(F列)、“金额“(G列) 列的值,分别对应存入页计行 “入库”列组 下的 “数量”列(F列)、“金额”列(G列) 列中"
            page_item_in_sum = 0
            page_item_out_sum = 0
            
            for item_index in page_item_types["入库"]:
                item_value = work_sheet.range((item_index, 6)).value
                if isinstance(item_value, (int, float)):
                    page_item_in_sum += item_value

            for item_index in page_item_types["出库"]:
                item_value = work_sheet.range((item_index, 8)).value
                if isinstance(item_value, (int, float)):
                    page_item_out_sum += item_value
            
            work_sheet.range((page_index * sheet_ratio - 2, 10)).value = page_item_in_sum  # 设置页计行 入库 列组的 数量(F列)的新值
            work_sheet.range((page_index * sheet_ratio - 2, 11)).value = page_item_out_sum # 设置页计行 入库 列组的 金额(G列)的新值
            
            "分别累加类型为 “出库” 的每一行的 “数量” 列(H列)、“金额“(I列) 列的值,分别对应存入页计行 “出库” 列组 下的 “数量”列(H列)、“金额”列(I列) 列中"
            page_item_in_sum = 0
            page_item_out_sum = 0 

            for item_index in page_item_types["出库"]:
                item_value = work_sheet.range((item_index, 6)).value
                if isinstance(item_value, (int, float)):
                    page_item_out_sum += item_value
            for item_index in page_item_types["出库"]:
                item_value = work_sheet.range((item_index, 8)).value
                if isinstance(item_value, (int, float)):
                    page_item_out_sum += item_value
            
            work_sheet.range((page_index * sheet_ratio - 2, 10)).value = page_item_in_sum  # 设置页计行 入库 列组的 数量(F列)的新值
            work_sheet.range((page_index * sheet_ratio - 2, 11)).value = page_item_out_sum # 设置页计行 出库 列组的 金额(G列)的新值

    elif excel_type == "子表副食表":

        if  work_sheet.name in [s.name for s in work_book.sheets]:

            print("Notice: 开始为 {excel_type} {sheet_name} 页执行页计功能".format(sheet_name=work_sheet.name))
            
            "设定一页的行数"
            sheet_ratio = 32                                    # 子表副食表中该表为32行一页

            "跳过已经写好了的表页，定位存有空行的一行，计算出其所在的页码"
            blank_row_index = get_first_blank_row_index(work_sheet) 
            page_index = int(blank_row_index / sheet_ratio) + 1 # int(13 / 32) = 0 但是是第一页，所以要加1  
            
            "统计该页范围内除了 日记、页计等的行"
            page_item_indexes = get_page_item_indexes(work_sheet,page_index,sheet_ratio)
                            
            "取所有行中行号最大的一行的 库存 列组的 数量(J列)、金额(K列) 列作为页计行 库存 列组的 数量(J列)、金额(K列) 列的值"
            max_row_index = max(page_item_indexes)
            page_item_sum = work_sheet.range((max_row_index, 10)).value                # 获取数量(J列)的值
            amount_value = work_sheet.range((max_row_index, 11)).value                 # 获取金额(K列)的值
            work_sheet.range((page_index * sheet_ratio - 2, 10)).value = page_item_sum # 设置页计行数量(J列)的新值
            work_sheet.range((page_index * sheet_ratio - 2, 11)).value = amount_value  # 设置页计行金额(K列)的新值

            "根据 page_item_indexes 存储行摘要列(D列)的字符值(入库/出库)将统计到的行以 {类型(入库/出库):[行号一维列表]}的方式重新打包为一个字典"
            page_item_types = {"入库": [], "出库": []}  # 初始化一个字典来存储类型和对应的行号列表
            
            for item_index in page_item_indexes:       
                
                item_type = work_sheet.range((item_index, 4)).value # 获取该行的摘要列(D列)的值
                
                if item_type == "入库":                             
                    page_item_types[item_type].append(item_index)   # 将行号添加到字典的"入库"键值对的值中
                
                elif item_type == "出库":
                    page_item_types[item_type].append(item_index)   # 将行号添加到字典的"出库"键值对的值中
                
                else:
                    print(f"Error: 在行 {item_index} 中发现非 入库/出库 的未知类型: {item_type}，已终止为{excel_type} 执行页计")
                    return

            "分别累加类型为 “入库” 的每一行的 “数量” 列(F列)、“金额“(G列) 列的值,分别对应存入页计行 “入库”列组 下的 “数量”列(F列)、“金额”列(G列) 列中"
            page_item_in_sum = 0
            page_item_out_sum = 0
            
            for item_index in page_item_types["入库"]:
                item_value = work_sheet.range((item_index, 6)).value
                if isinstance(item_value, (int, float)):
                    page_item_in_sum += item_value

            for item_index in page_item_types["出库"]:
                item_value = work_sheet.range((item_index, 8)).value
                if isinstance(item_value, (int, float)):
                    page_item_out_sum += item_value
            
            work_sheet.range((page_index * sheet_ratio - 2, 10)).value = page_item_in_sum  # 设置页计行 入库 列组的 数量(F列)的新值
            work_sheet.range((page_index * sheet_ratio - 2, 11)).value = page_item_out_sum # 设置页计行 入库 列组的 金额(G列)的新值
            
            "分别累加类型为 “出库” 的每一行的 “数量” 列(H列)、“金额“(I列) 列的值,分别对应存入页计行 “出库” 列组 下的 “数量”列(H列)、“金额”列(I列) 列中"
            page_item_in_sum = 0
            page_item_out_sum = 0 

            for item_index in page_item_types["出库"]:
                item_value = work_sheet.range((item_index, 6)).value
                if isinstance(item_value, (int, float)):
                    page_item_out_sum += item_value
            for item_index in page_item_types["出库"]:
                item_value = work_sheet.range((item_index, 8)).value
                if isinstance(item_value, (int, float)):
                    page_item_out_sum += item_value
            
            work_sheet.range((page_index * sheet_ratio - 2, 10)).value = page_item_in_sum  # 设置页计行 入库 列组的 数量(F列)的新值
            work_sheet.range((page_index * sheet_ratio - 2, 11)).value = page_item_out_sum # 设置页计行 出库 列组的 金额(G列)的新值

    else:
        print("Error: 未知的表类型值: {Excel_Type} 跳过执行页计功能".format(Excel_Type=excel_type))
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
