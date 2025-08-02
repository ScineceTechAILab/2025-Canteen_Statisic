



import __main__
import xlwings as xw
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


def counting_total_value(excel_type:str,work_book:xw.Book,work_sheet:xw.Sheet,sheet_name:str):
    """
    入库在将条目添加到表并进行日记、月计之后,为该页面加上页计行(v1.1 逻辑流版本)
    
    Parameters:
      excel_type: 正在处理的表的类型,值为"主表"、"福利表"、"子表主食表"、"子表副食表"
      work_book: 要修改的 Excel 传入workbook对象
      work_sheet: 要修改的 Excel 传入worksheet对象
      sheet_name: 要修改的 Excel 存储的表名
    """
        
    "判断哪个表需要合计"
    if excel_type == "主表":

        if work_sheet.name in [s.name for s in work_book.sheets]:

            if work_sheet.name == "食堂物品收发存库存表":

                print(f"\nNotice: 开始为 `{excel_type}` `{sheet_name}` 页执行页计功能")

                "设定一页的行数"
                sheet_ratio = 25                                    # 主表中 食堂物品收发存库存表 表为26行一页

                "跳过已经写好了的表页，定位存有空行的一行，计算出其所在的页码"
                blank_row_index = get_first_blank_row_index(work_sheet) 
                page_index = int(blank_row_index / sheet_ratio) + 1 # int(13 / 25) = 0 但是是第一页，所以要加1  
                print(f"Notice: 发现 {excel_type} `{work_sheet.name}` 中第 `{page_index}` 页存在空行")

                "所有页的页计行是否在所有页的倒数二行"
                for i in range(1,page_index+1):
                    
                    page_line_name = str(work_sheet.range((i * sheet_ratio - 1, 1)).value).replace(" ","")
                    # 若每一页倒数第二行为页计行
                    if page_line_name == "页计":
                        continue

                    # 若每一页倒数第二行为空行
                    elif page_line_name == "" or page_line_name == "None":
                        print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页倒数第二行是空行,终止本次提交")
                        __main__.SAVE_OK_SIGNAL = False
                        raise Exception
                    
                    # 若每一页倒数第二行存在非页计行
                    else:
                        print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页倒数第二行不是页计行,终止本次提交")
                        __main__.SAVE_OK_SIGNAL = False
                        raise Exception
                
                "累加本页以及前页所有页计行的值"
                page_item_sum = {"F":0,"G":0,"H":0,"I":0,"J":0,"K":0,"L":0,"M":0,"N":0}
                for item in page_item_sum:
                    for i in range(1,page_index+1):
                        # 分别累加本页及之前页的页计行的J列的值
                        if work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value is None:
                            print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列的值为None,终止本次提交")
                            __main__.SAVE_OK_SIGNAL = False
                            raise Exception

                        elif work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value == "":
                            print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列的值为空,终止本次提交")
                            __main__.SAVE_OK_SIGNAL = False
                            raise Exception
                        
                        print(f"Notice: {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列的当前值为{work_sheet.range((i * sheet_ratio - 1, ord(item) - ord('A') + 1)).value}")
                        page_item_sum[item] += float(work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value)


                "将这些值设为总计行的F~N列的值,放置于从前往后的第一行空行"
                work_sheet.range((blank_row_index, 1)).value = "总计"
                for item in page_item_sum:
                        
                    print(f"Notice: {excel_type} {work_sheet.name} 中更新前的总计值为{work_sheet.range((blank_row_index, ord(item) - ord('A') + 1)).value}")
                    work_sheet.range((blank_row_index, ord(item) - ord("A") + 1)).value = page_item_sum[item]
                    print(f"Notice: {excel_type} {work_sheet.name} 中更新后的总计值为{work_sheet.range((blank_row_index, ord(item) - ord('A') + 1)).value}")
                
                print(f"Notice: 结束为 {excel_type} 中 {work_sheet.name} 进行总计")
                
            elif work_sheet.name in ["自购主食入库等","食堂副食入库","食堂副食入库 ","厂调面食入库","扶贫主食入库","扶贫副食入库"]:
                
                print(f"\nNotice: 开始为 `{excel_type}` `{sheet_name}` 页执行页计功能")

                "设定一页的行数"
                sheet_ratio = 33                                    # 主表中 食堂物品收发存库存表 表为33行一页

                "跳过已经写好了的表页，定位存有空行的一行，计算出其所在的页码"
                blank_row_index = get_first_blank_row_index(work_sheet) 
                page_index = int(blank_row_index / sheet_ratio) + 1 # int(13 / 33) = 0 但是是第一页，所以要加1  
                print(f"Notice: 发现 {excel_type} `{work_sheet.name}` 中第 `{page_index}` 页存在空行")

                "所有页的页计行是否在所有页的倒数二行"
                for i in range(1,page_index+1):
                    
                    # 若每一页倒数第二行为页计行
                    page_line_name = str(work_sheet.range((i * sheet_ratio - 1, 1)).value).replace(" ","")
                    if page_line_name == "页计":
                        continue
                    
                    # 若每一页倒数第二行为空行
                    elif page_line_name == "" or page_line_name == "None":
                        print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页倒数第二行是空行,终止本次提交")
                        __main__.SAVE_OK_SIGNAL = False
                        raise Exception

                    # 若每一页倒数第二行存在非页计行
                    else:
                        print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页倒数第二行不是页计行,终止本次提交")
                        __main__.SAVE_OK_SIGNAL = False
                        raise Exception
        
                "累加本页以及前页所有页计行的值"
                page_item_sum = {"J":0}

                for item in page_item_sum:
                    for i in range(1,page_index+1):
                        
                        # 分别累加本页及之前页的页计行的J列的值
                        if work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value is None:
                            print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列的值为None,终止本次提交")
                            __main__.SAVE_OK_SIGNAL = False
                            raise Exception

                        elif work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value == "":
                            print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列的值为空,终止本次提交")
                            __main__.SAVE_OK_SIGNAL = False
                            raise Exception

                        print(f"Notice: {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列的当前值为{work_sheet.range((i * sheet_ratio - 1, ord(item) - ord('A') + 1)).value}")
                        page_item_sum[item] += float(work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value)
                        
                
                "将这些值设为总计行的J列的值,放置于从前往后的第一行空行"
                work_sheet.range((blank_row_index, 1)).value = "总计"
                for item in page_item_sum:

                    print(f"Notice: {excel_type} {work_sheet.name} 中更新前的总计值为{work_sheet.range((blank_row_index, ord(item) - ord('A') + 1)).value}")
                    work_sheet.range((blank_row_index, ord(item) - ord("A") + 1)).value = page_item_sum[item]
                    print(f"Notice: {excel_type} {work_sheet.name} 中更新后的总计值为{work_sheet.range((blank_row_index, ord(item) - ord('A') + 1)).value}")

                print(f"Notice: 结束为 {excel_type} 中 {work_sheet.name} 进行总计")

        else:
            print(f"Error: {excel_type} 中 {work_sheet.name} 页不存在,跳过执行页计功能")
            __main__.SAVE_OK_SIGNAL = False

            return
            
    elif excel_type == "福利表":
                    
        if work_sheet.name == "过年福利入":

            print(f"\nNotice: 开始为 `{excel_type}` `{sheet_name}` 页执行页计功能")

            "设定一页的行数"
            sheet_ratio = 32                                   

            "跳过已经写好了的表页，定位存有空行的一行，计算出其所在的页码"
            blank_row_index = get_first_blank_row_index(work_sheet) 
            page_index = int(blank_row_index / sheet_ratio) + 1 # int(13 / 32) = 0 但是是第一页，所以要加1  
            print(f"Notice: 发现 {excel_type} `{work_sheet.name}` 中第 `{page_index}` 页存在空行")

            "所有页的页计行是否在所有页的倒数二行"
            for i in range(1,page_index+1):
                
                # 若每一页倒数第二行为页计行
                page_line_name = str(work_sheet.range((i * sheet_ratio - 1, 1)).value).replace(" ","")
                if page_line_name == "页计"  :
                    continue

                # 若每一页倒数第二行为空行
                elif page_line_name == "" or page_line_name == "None":
                    print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页倒数第二行是空行,终止本次提交")
                    __main__.SAVE_OK_SIGNAL = False
                    raise Exception                

                # 若每一页倒数第二行存在非页计行
                else:
                    print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页倒数第二行不是页计行,终止本次提交")
                    __main__.SAVE_OK_SIGNAL = False
                    raise Exception
            
            "累加本页以及前页所有页计行的值"
            page_item_sum = {"J":0}
            for item in page_item_sum:
                for i in range(1,page_index+1):
                    # 分别累加本页及之前页的页计行的J列的值
                    if work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value is None:
                        print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列的值为None,终止本次提交")
                        __main__.SAVE_OK_SIGNAL = False
                        raise Exception

                    elif work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value == "":
                        print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列的值为空,终止本次提交")
                        __main__.SAVE_OK_SIGNAL = False
                        raise Exception
                    
                    print(f"Notice: {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列更新前的值为{work_sheet.range((i * sheet_ratio - 1, ord(item) - ord('A') + 1)).value}")
                    page_item_sum[item] += float(work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value)
                    
            
            "将这些值设为总计行的J列的值,放置于从前往后的第一行空行"
            work_sheet.range((blank_row_index, 1)).value = "总计"
            for item in page_item_sum:
                    
                print(f"Notice: {excel_type} {work_sheet.name} 中更新前的总计值为{work_sheet.range((blank_row_index, ord(item) - ord('A') + 1)).value}")
                work_sheet.range((blank_row_index, ord(item) - ord("A") + 1)).value = page_item_sum[item]
                print(f"Notice: {excel_type} {work_sheet.name} 中更新后的总计值为{work_sheet.range((blank_row_index, ord(item) - ord('A') + 1)).value}")
            
            print(f"Notice: 结束为 {excel_type} 中 {work_sheet.name} 进行总计")

        else:
            print(f"Error: {excel_type} 中 {work_sheet.name} 页不存在,跳过执行页计功能")
            __main__.SAVE_OK_SIGNAL = False
            raise Exception

    elif excel_type == "子表主食表":
    
        if work_sheet.name in [s.name for s in work_book.sheets]:

            print(f"\nNotice: 开始为 `{excel_type}` `{sheet_name}` 页执行页计功能")

            "设定一页的行数"
            sheet_ratio = 33                                    

            "跳过已经写好了的表页，定位存有空行的一行，计算出其所在的页码"
            blank_row_index = get_first_blank_row_index(work_sheet) 
            page_index = int(blank_row_index / sheet_ratio) + 1 # int(13 / 32) = 0 但是是第一页，所以要加1  
            print(f"Notice: 发现 {excel_type} `{work_sheet.name}` 中第 `{page_index}` 页存在空行")

            "所有页的页计行是否在所有页的倒数二行"
            for i in range(1,page_index+1):
                
                # 若每一页倒数第二行为页计行
                page_line_name = str(work_sheet.range((i * sheet_ratio - 1, 1)).value).replace(" ","")
                if page_line_name == "页计":
                    continue
                
                # 若每一页倒数第二行为空行
                elif page_line_name == "" or page_line_name == "None":
                    print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页倒数第二行是空行,终止本次提交")
                    __main__.SAVE_OK_SIGNAL = False
                    raise Exception

                # 若每一页倒数第二行存在非页计行
                else:
                    print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页倒数第二行不是页计行,终止本次提交")
                    __main__.SAVE_OK_SIGNAL = False
                    raise Exception
            
            "累加本页以及前页所有页计行的值"
            page_item_sum = {"F":0,"G":0,"H":0,"I":0,"J":0,"K":0}
            for item in page_item_sum:
                for i in range(1,page_index+1):
                    # 分别累加本页及之前页的页计行的J列的值
                    if work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value is None:
                        print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列的值为None,终止本次提交")
                        __main__.SAVE_OK_SIGNAL = False
                        raise Exception

                    elif work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value == "":
                        print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列的值为空,终止本次提交")
                        __main__.SAVE_OK_SIGNAL = False
                        raise Exception
                    
                    print(f"Notice: {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列更新前的值为{work_sheet.range((i * sheet_ratio - 1, ord(item) - ord('A') + 1)).value}")
                    page_item_sum[item] += float(work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value)

            "将这些值设为总计行的F~K列的值,放置于从前往后的第一行空行"
            work_sheet.range((blank_row_index, 1)).value = "总计"
            for item in page_item_sum:
                    
                print(f"Notice: {excel_type} {work_sheet.name} 中更新前的总计值为{work_sheet.range((blank_row_index, ord(item) - ord('A') + 1)).value}")
                work_sheet.range((blank_row_index, ord(item) - ord("A") + 1)).value = page_item_sum[item]
                print(f"Notice: {excel_type} {work_sheet.name} 中更新后的总计值为{work_sheet.range((blank_row_index, ord(item) - ord('A') + 1)).value}")
            
            print(f"Notice: 结束为 {excel_type} 中 {work_sheet.name} 进行总计")

        else:
            print(f"Error: {excel_type} 中 {work_sheet.name} 页不存在,跳过执行页计功能")
            __main__.SAVE_OK_SIGNAL = False
            raise Exception

    elif excel_type == "子表副食表":

        if work_sheet.name in [s.name for s in work_book.sheets]:

            print(f"\nNotice: 开始为 `{excel_type}` `{sheet_name}` 页执行页计功能")

            "设定一页的行数"
            sheet_ratio = 32 # 子表副食表皆为 32 行一页，类型在第 4 列                                

            "跳过已经写好了的表页，定位存有空行的一行，计算出其所在的页码"
            blank_row_index = get_first_blank_row_index(work_sheet) 
            page_index = int(blank_row_index / sheet_ratio) + 1 # int(13 / 32) = 0 但是是第一页，所以要加1  
            print(f"Notice: 发现 {excel_type} `{work_sheet.name}` 中第 `{page_index}` 页存在空行")

            "所有页的页计行是否在所有页的倒数二行"
            for i in range(1,page_index+1):
                
                # 若每一页倒数第二行为页计行
                page_line_name = str(work_sheet.range((i * sheet_ratio - 1, 4)).value).replace(" ","")
                if page_line_name == "页计":
                    continue

                # [ ] TODO:按照平板上的甲方 EXCEL 要求重构甲方的 EXCEL ，解决每一页行数不统一；存在无意义行（过次页...）等问题 
                # 若每一页倒数第二行为空行
                elif page_line_name == "" or page_line_name == "None":
                    print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页倒数第二行是空行,终止本次提交")
                    __main__.SAVE_OK_SIGNAL = False
                    raise Exception

                # 若每一页倒数第二行存在非页计行
                else:
                    print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页倒数第二行不是页计行,终止本次提交")
                    __main__.SAVE_OK_SIGNAL = False
                    raise Exception
            
            "累加本页以及前页所有页计行的值"
            page_item_sum = {"F":0,"G":0,"H":0,"I":0,"J":0,"K":0}
            for item in page_item_sum:
                for i in range(1,page_index+1):
                    # 分别累加本页及之前页的页计行的J列的值
                    if work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value is None:
                        print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列的值为None,终止本次提交")
                        __main__.SAVE_OK_SIGNAL = False
                        raise Exception

                    elif work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value == "":
                        print(f"Error: 发现 {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列的值为空,终止本次提交")
                        __main__.SAVE_OK_SIGNAL = False
                        raise Exception
                    
                    print(f"Notice: {excel_type} {work_sheet.name} 中第 {i} 页的页计行的{item}列更新前的值为{work_sheet.range((i * sheet_ratio - 1, ord(item) - ord('A') + 1)).value}")
                    page_item_sum[item] += float(work_sheet.range((i * sheet_ratio - 1, ord(item) - ord("A") + 1)).value)

            "将这些值设为总计行的F~K列的值,放置于从前往后的第一行空行"
            work_sheet.range((blank_row_index, 1)).value = "总计"
            for item in page_item_sum:
                    
                print(f"Notice: {excel_type} {work_sheet.name} 中更新前的总计值为{work_sheet.range((blank_row_index, ord(item) - ord('A') + 1)).value}")
                work_sheet.range((blank_row_index, ord(item) - ord("A") + 1)).value = page_item_sum[item]
                print(f"Notice: {excel_type} {work_sheet.name} 中更新后的总计值为{work_sheet.range((blank_row_index, ord(item) - ord('A') + 1)).value}")
            
            print(f"Notice: 结束为 {excel_type} 中 {work_sheet.name} 进行总计")

        else:
            print(f"Error: {excel_type} 中 {work_sheet.name} 页不存在,跳过执行页计功能")
            __main__.SAVE_OK_SIGNAL = False
            raise Exception


def get_first_blank_row_index(work_sheet):
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
        sheet_ratio: 相应 Sheet 中每页的行数(1索引格式)，例如主表中自购主食入库等每页33行，福利表中每页32行

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

    
                    