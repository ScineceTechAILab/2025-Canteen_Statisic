"""
item_data_operate.py: Utility functions for item inventory operations.

This module provides a collection of functions for updating and retrieving
item information in the inventory index during item check-in and check-out
operations.

Author: ESJIAN
Email: esjian@outlook.com

Copyright (c) 2025 STA LAB. All rights reserved.
Licensed under the MIT License.

Version: 1.0
Date: 2025-7-6

Dependencies:
    - [List any dependencies, e.g., pandas, sqlite3]

Example:
    ido.add_or_update_item(item_id, item_info)
    stock = ido.get_item_stock(item_id)
    print(f"Current stock: {stock}")

Description:
    - Adds or updates item information in the inventory index during check-in.
    - Updates stock data in the index during check-out.
    - Provides interfaces for querying current stock and item attributes.

This module is designed for the data operation layer of a canteen item
management system, ensuring the accuracy and consistency of inventory data.
"""

import openpyxl
import os
import sys
import __main__


def item_data_operate(model, year, month, day, product_name, unit_name, price, quantity, amount, remark, company_name, sigle_name):
    """
    This function is the main entry point for the module.
    It provides a collection of functions for updating and retrieving
    item information in the inventory index during item check-in and check-out
    operations.

    Parameters:
        model:入库\出库
        year: 入库出库年份
        month: 该条目入库出库的月份
        day: 该条目入库出库的日期
        product_name: 该条目入库出库的食品名
        unit_name: 该条目入库出库的单位
        price: 该条目入库出库的单价
        quantity: 该条目入库出库的数量
        amount: 该条目入库出库的金额
        remark: 该条目入库出库的备注
        company_name: 该条目入库出库的公司名称
        sigle_name: 该条目入库出库的简称

    Returns:
        dict:{品名:[单位-单价-数量-数额-日期]}

    """
    print(f"\nNotice:条目表 {product_name} 数据信息开始更新！")
    wb = openpyxl.Workbook()
    ws = wb.active

    export_data = {product_name: []}

    "物品索引库文件夹是否存在"
    if not os.path.exists(__main__.ITEM_EXCEL_FOLDER):

        "新建索引数据表"

        # 创建条目表文件夹
        os.mkdir(__main__.ITEM_EXCEL_FOLDER)

        # 写入表头
        ws["A1"] = "单位"
        ws["B1"] = "单价"
        ws["C1"] = "存量"
        ws["D1"] = "存值"
        ws["E1"] = "日期"

        # 存为条目表.xlsx 文件
        wb.save(__main__.ITEM_EXCEL_FOLDER + "条目表.xlsx")

    "打开此索引库"
    item_excel_list = [item for item in os.listdir(__main__.ITEM_EXCEL_FOLDER) if item == "条目表.xlsx"]
    for item in item_excel_list:
        # 若 __main__.ITEM_EXCEL_FOLDER 下 条目表.xlsx 文件存在
        if item == "条目表.xlsx":
            # 打开条目表.xlsx 文件(Mistake: load_workbook 方法采用绝对路径)
            wb = openpyxl.load_workbook(__main__.ITEM_EXCEL_FOLDER + item)

    "检查索引库中是否有此类"
    for sheet in wb.worksheets:
        
        if sheet.title != product_name:
            continue
        
        elif sheet.title == product_name:
            
            ws = wb[product_name]

            ws["A1"] = "单位"
            ws["B1"] = "单价"
            ws["C1"] = "存量"
            ws["D1"] = "存值"
            ws["E1"] = "日期"

            # 从D2开始，D列值为C列值与B列值同行相乘的积
            for row in range(2, ws.max_row + 1):
                ws[f"D{row}"] = float(ws[f"C{row}"].value)*float(ws[f"B{row}"].value)
            break

        else:

            wb.create_sheet(product_name)
            ws = wb[product_name]

            ws["A1"] = "单位"
            ws["B1"] = "单价"
            ws["C1"] = "存量"
            ws["D1"] = "存值"
            ws["E1"] = "日期"

            # 从D2开始，D列值为C列值与B列值同行相乘的积
            for row in range(2, ws.max_row + 1):
                ws[f"D{row}"] = float(ws[f"C{row}"].value)*float(ws[f"B{row}"].value)

    "打开此Sheet"
    ws = wb[product_name]


    "检查操作类型"
    if model == "入库":
        
        "检测除表头外是否有数据"
        if ws.max_row == 1:
            # 追加写入数据
            ws.append([unit_name, price, quantity, amount, f"{year}-{month}-{day}"])

        else:
            "检查价格列中有无匹配条目价格的行"
            # 筛查 B 列中是否有与 price 等值的行
            for row in range(2, ws.max_row + 1):
                if ws[f"B{row}"].value == price:
                    # 更新存量列(C列)、存值列(D列)、日期列(E列)
                    print(f"Notice: 条目表 {product_name} 页更新前第 {row} 行 单位:{ws[f'A{row}'].value} 单价:{ws[f'B{row}'].value} 存量:{ws[f'C{row}'].value} 存值:{ws[f'D{row}'].value} 日期:{ws[f'E{row}'].value}")
                    ws[f"C{row}"] = float(ws[f"C{row}"].value) + float(quantity)
                    ws[f"D{row}"] = float(ws[f"C{row}"].value)*float(ws[f"B{row}"].value)
                    ws[f"E{row}"] = f"{year}-{month}-{day}"        
                    print(f"Notice: 条目表 {product_name} 页更新后第 {row} 行 单位:{ws[f'A{row}'].value} 单价:{ws[f'B{row}'].value} 存量:{ws[f'C{row}'].value} 存值:{ws[f'D{row}'].value} 日期:{ws[f'E{row}'].value}")

    elif model == "出库":

        "暂存该表中的条目"
        data_rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True))

        "检测除表头外是否有数据"
        if ws.max_row == 1:
            __main__.SAVE_OK_SIGNAL = False
            print(f"Error: 条目表中 {product_name} 没有存储数据，请先入库！ ")
            return

        else:
            "遍历行列"
        
            for row in range(2, ws.max_row + 1):  
                "比较该行存量与出库数量关系"
                if float(ws[f"C{row}"].value) >= float(quantity):
                    
                    #将出库清单添加到 export_data
                    export_data[product_name].append([ws[f"A{row}"].value,ws[f"B{row}"].value , quantity , float(quantity) * float(ws[f'B{row}'].value), ws[f"E{row}"].value])

                    # 更新存量列(C列)、存值列(D列)、日期列(E列)|Mistake: 单元格在运算的时候需要先进行强制类型转换
                    print(f"Notice: 条目表 {product_name} 页更新前第 {row} 行 单位:{ws[f'A{row}'].value} 单价:{ws[f'B{row}'].value} 存量:{ws[f'C{row}'].value} 存值:{ws[f'D{row}'].value} 日期:{ws[f'E{row}'].value}")
                    ws[f"C{row}"] = float(ws[f"C{row}"].value) - float(quantity)
                    ws[f"D{row}"] = float(ws[f"C{row}"].value)*float(ws[f"B{row}"].value)
                    ws[f"E{row}"] = f"{year}-{month}-{day}"
                    print(f"Notice: 条目表 {product_name} 页更新后第 {row} 行 单位:{ws[f'A{row}'].value} 单价:{ws[f'B{row}'].value} 存量:{ws[f'C{row}'].value} 存值:{ws[f'D{row}'].value} 日期:{ws[f'E{row}'].value}")
                    
                    quantity = 0

                elif float(ws[f"C{row}"].value) < float(quantity):   
                    
                    # 将出库清单添加到 export_data
                    export_data[product_name].append([ws[f"A{row}"].value,ws[f"B{row}"].value , ws[f"C{row}"].value,float(ws[f"B{row}"].value) * float(ws[f'C{row}'].value), ws[f"E{row}"].value])

                    # 更新存量列(C列)、存值列(D列)、日期列(E列)
                    ws[f"C{row}"] = 0
                    ws[f"D{row}"] = float(ws[f"C{row}"].value)*float(ws[f"B{row}"].value)
                    ws[f"E{row}"] = f"{year}-{month}-{day}"

                    quantity = float(quantity) - float(ws[f"C{row}"].value)

                    continue    
                                
                if quantity == 0:
                    break 
            
            "检查库存数是否够出库数"
            if float(quantity) != 0:
                
                "复原条目"
                for row in range(2, ws.max_row + 1):
                    ws[f"C{row}"] = data_rows[row - 2][2]
                    ws[f"D{row}"] = data_rows[row - 2][3]
                    ws[f"E{row}"] = data_rows[row - 2][4]
                
                __main__.SAVE_OK_SIGNAL = False
                print(f"Error: 条目表中 {product_name} 存量不足，请重新输入！ ")
                return
            

    "将从第二行开始的所有行按照单价值从小到大重新排序"
    # 提取数据（不包括表头）
    data_rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True))
    # 按单价（B列，索引1）排序
    data_rows.sort(key=lambda x: x[1])
    # 清除原有数据（不包括表头）
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None
    # 重新写入排序后的数据
    for idx, row_data in enumerate(data_rows, start=2):
        for col_idx, value in enumerate(row_data, start=1):
            ws.cell(row=idx, column=col_idx, value=value)

    "保存此Sheet"
    wb.save(__main__.ITEM_EXCEL_FOLDER + "条目表.xlsx")
    print(f"Notice: 条目表 {product_name} 页更新完成！")
    return export_data

    
def index_item_data():
    """
    This function is the main entry point for the module.
    It provides a collection of functions for updating and retrieving
    item information in the inventory index during item check-in and check-out
    operations.
    """