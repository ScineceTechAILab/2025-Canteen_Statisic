# mini_ocr.py
from paddleocr import PaddleOCR

def main():
    ocr = PaddleOCR(use_angle_cls=True, lang="ch")  # 简体中文
    result = ocr.ocr("test.png", cls=True)
    for line in result[0]:
        print(line[1][0])  # 输出识别文本

if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()  # 🔑 打包 Windows 必备
    main()
