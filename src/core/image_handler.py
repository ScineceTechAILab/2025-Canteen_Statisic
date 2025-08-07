import os
import threading
import pandas as pd
from paddleocr import PPStructure
import paddleocr

def image_to_excel(
    image_path: str,
    save_folder: str = "./src/data/input/manual",
    ocr_model_path: str = None  
):
    """
    使用PaddleOCR和PPStructure识别图片中的表格并导出为Excel文件（支持追加写入）。

    :param image_path: 输入图片路径
    :param save_folder: Excel文件保存目录
    :param ocr_model_path: OCR模型路径
    """
    
    os.makedirs(save_folder, exist_ok=True)
    file_stem = os.path.splitext(os.path.basename(image_path))[0]
    excel_path = os.path.join(save_folder, f"temp_img_input.xlsx")

    # 初始化表格结构识别引擎
    try:
        # 保存原始BASE_DIR值
        original_base_dir = None
        if ocr_model_path:
            original_base_dir = paddleocr.BASE_DIR
            paddleocr.BASE_DIR = os.path.abspath(ocr_model_path)
            
        table_engine = PPStructure()
        
        # 恢复原始BASE_DIR值
        if original_base_dir:
            paddleocr.BASE_DIR = original_base_dir
            
    except Exception as e:
        print(f"Error: 线程 {threading.current_thread().name}（ID={threading.get_ident()}）无法初始化表格结构识别引擎。请检查模型路径是否正确。")
        return

    # 进行表格结构识别
    result = table_engine(image_path)

    # 提取表格内容
    new_tables = []
    for i, item in enumerate(result):
        if item['type'] == 'table':
            table_html = item['res']['html']
            dfs = pd.read_html(table_html)
            for df in dfs:
                new_tables.append(df)

    # 追加写入逻辑
    if os.path.exists(excel_path):
        # 读取已有数据
        try:
            existing_df = pd.read_excel(excel_path)
        except Exception:
            existing_df = pd.DataFrame()
        # 合并所有新表格
        combined_new = pd.concat(new_tables, ignore_index=True) if new_tables else pd.DataFrame()
        # 合并已有和新数据
        final_df = pd.concat([existing_df, combined_new], ignore_index=True)
    else:
        # 仅保存新表格
        final_df = pd.concat(new_tables, ignore_index=True) if new_tables else pd.DataFrame()

    # 保存到Excel
    final_df.to_excel(excel_path, index=False)
    print(f"Notice: 线程 {threading.current_thread().name}（ID={threading.get_ident()}）已完成 {image_path} 的表格识别并追加导出到Excel文件。")