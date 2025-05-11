import sys
import os
import threading
import shutil
import multiprocessing
import subprocess

from PySide6.QtWidgets import QMessageBox

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt, QEvent)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform, Qt)
from PySide6.QtWidgets import (QApplication, QButtonGroup, QFormLayout, QGridLayout,
    QGroupBox, QHBoxLayout, QLabel, QLayout,
    QLineEdit, QPlainTextEdit, QPushButton, QScrollArea,
    QSizePolicy, QSpinBox, QTabWidget, QVBoxLayout,
    QWidget, QFileDialog, QDialog, QVBoxLayout, QCheckBox)

# 获取当前文件的绝对路径
current_file_path = os.path.abspath(__file__) # Fixed1:将项目包以绝对形式导入,解决了相对导入不支持父包的报错
# 获取项目根目录
project_root = os.path.abspath(os.path.join(current_file_path, '..', '..', '..')) # Fixed1:将项目包以绝对形式导入,解决了相对导入不支持父包的报错
# 将项目根目录添加到 sys.path
sys.path.insert(0, project_root) # Fixed1:将项目包以绝对形式导入,解决了相对导入不支持父包的报错

from config.config import FIRST_START

def first_start_detect(Form):
    """检测是否首次启动，如果是首次启动则弹出提示框，并修改配置文件"""
    print("Notice:检测是否首次启动")
    # 检查是否首次启动
    if FIRST_START :
        QMessageBox.information(Form, "提示", "首次启动应用，请点击重导表格", QMessageBox.Ok)
        # 打开目标文件将值写成 False
        with open("./config/config.py", "w") as f: # 创建一个文件，将值写成 False
            # 将FIRST_START  = True 的 True 替换成 False
            f.write(f"FIRST_START = False")            
            print("Notice:首次启动成功,将首次启动标识覆写为 False ")
    else:
        print("Notice:非首次启动，正常进入应用")