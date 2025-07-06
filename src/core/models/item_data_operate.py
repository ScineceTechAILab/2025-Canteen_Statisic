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


def item_data_operate(model, year, month, day, category_name, unit_name, price, quantity, amount, remark, company_name, sigle_name):
    """
    This function is the main entry point for the module.
    It provides a collection of functions for updating and retrieving
    item information in the inventory index during item check-in and check-out
    operations.

    Parameters:
        model:入库\出库
        year:
        month:
        day:
        category_name:
        unit_name:
        price:
        quantity:
        amount:
        remark:
        company_name:
        sigle_name:

    Returns:
        list:在模式为出库时候返回出库清单的二维列表

    """
    wb = openpyxl.Workbook()
    ws = wb.active

    export_data = {category_name: []}

    "物品索引库是否存在"
    if not os.path.exists(__main__.ITEM_EXCEL_FOLDER):
        # 若 __main__.ITEM_EXCEL_FOLDER 不存在
        os.mkdir(__main__.ITEM_EXCEL_FOLDER)

    item_excel_list = [item for item in os.listdir(__main__.ITEM_EXCEL_FOLDER) if item == "条目表.xlsx"]
    if item_excel_list == []:
        # 创建条目表.xlsx 文件
        wb.save(__main__.ITEM_EXCEL_FOLDER + "条目表.xlsx")
        item_excel_list = [item for item in os.listdir(__main__.ITEM_EXCEL_FOLDER) if item == "条目表.xlsx"]
    
    "打开此索引库"
    for item in item_excel_list:
        # 若 __main__.ITEM_EXCEL_FOLDER 下 条目表.xlsx 文件存在
        if item == "条目表.xlsx":
            # 打开条目表.xlsx 文件
            wb = openpyxl.load_workbook(item_excel_list[0])

    "检查索引库中是否有此类"
    for sheet in wb.worksheets:
        
        if sheet.title == category_name:
            continue
        else:

            wb.create_sheet(category_name)
            ws = wb[category_name]

            ws["A1"] = "单位"
            ws["B1"] = "单价"
            ws["C1"] = "存量"
            ws["D1"] = "存值"
            ws["E1"] = "日期"

            # 从D2开始，D列值为C列值与B列值同行相乘的积
            for row in range(2, ws.max_row + 1):
                ws[f"D{row}"] = f"=C{row}*B{row}"

    "打开此Sheet"
    ws = wb[category_name]

    "检测单位是否匹配"
    if ws[f"A{row}"].value == unit_name:
    
        "检查操作类型"
        if model == "入库":
            
            "检查价格列中有无匹配条目价格的行"
            # 筛查 B 列中是否有与 price 等值的行
            for row in range(2, ws.max_row + 1):
                if ws[f"B{row}"].value == price:
                    # 更新存量列(C列)、存值列(D列)、日期列(E列)
                    ws[f"C{row}"] = ws[f"C{row}"].value + quantity
                    ws[f"D{row}"] = f"=C{row}*B{row}"
                    ws[f"E{row}"] = f"{year}-{month}-{day}"        
            
        elif model == "出库":
     
            "便历行列"
            for row in range(2, ws.max_row + 1):  
                "比较该行存量与出库数量关系"
                if ws[f"C{row}"].value >= quantity:

                    # 更新存量列(C列)、存值列(D列)、日期列(E列)
                    ws[f"C{row}"] = ws[f"C{row}"].value - quantity
                    ws[f"D{row}"] = f"=C{row}*B{row}"
                    ws[f"E{row}"] = f"{year}-{month}-{day}"

                    #将出库清单添加到 export_data
                    export_data[category_name].append([ws[f"A{row}"],ws[f"B{row}"] ,ws[f"C{row}"], ws[f"D{row}"], ws[f"E{row}"]])


                elif ws[f"C{row}"].value < quantity:    

                    # 更新存量列(C列)、存值列(D列)、日期列(E列)
                    ws[f"C{row}"] = 0
                    ws[f"D{row}"] = f"=C{row}*B{row}"
                    ws[f"E{row}"] = f"{year}-{month}-{day}"

                    quantity = quantity - ws[f"C{row}"].value

                    # 将出库清单添加到 export_data
                    export_data[category_name].append([ws[f"A{row}"],ws[f"B{row}"] ,ws[f"C{row}"], ws[f"D{row}"], ws[f"E{row}"]])
                    continue    
                              
                if quantity == 0:
                    break 
            
    "将从第二行开始的所有行按照单价值从小到大重新排序"
    ws.sort(key=lambda x: x[1].value, reverse=False)

    "保存此Sheet"
    wb.save(__main__.ITEM_EXCEL_FOLDER + "条目表.xlsx")
    
    return export_data

    
