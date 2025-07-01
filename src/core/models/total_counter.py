##### 作者：ESJIAN
# 日期: 2025.7.1
# 版本：v1.1
# 功能：
#   1. 为主表、福利表、子表主食表、子表副食表添加页计行





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
        