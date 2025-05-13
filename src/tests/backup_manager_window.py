# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'backup_window_v1tkqhOe.ui'
##
## Created by: Qt User Interface Compiler version 6.9.0
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

import sys
from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QAbstractScrollArea, QApplication, QHBoxLayout, QLabel,
    QLayout, QPushButton, QScrollArea, QSizePolicy,
    QVBoxLayout, QWidget)

class Ui_Form(object):
    def setupUi(self, Form):
        if not Form.objectName():
            Form.setObjectName(u"Form")
        # 设置窗口大小
        Form.resize(712, 370)
            # 创建垂直布局
        self.verticalLayout = QVBoxLayout(Form)
        self.verticalLayout.setObjectName(u"verticalLayout")

        # 滚动区域
        self.scrollArea = QScrollArea(Form)
        self.scrollArea.setObjectName(u"scrollArea")
            # 设置滚动区域的大小策略
        sizePolicy = QSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.scrollArea.sizePolicy().hasHeightForWidth())
            # 应用大小策略
        self.scrollArea.setSizePolicy(sizePolicy)
        self.scrollArea.setSizeAdjustPolicy(QAbstractScrollArea.SizeAdjustPolicy.AdjustIgnored)
        self.scrollArea.setWidgetResizable(True)
            # 添加滚动区域
        self.verticalLayout.addWidget(self.scrollArea)
            # 设置滚动区域的内容
                # 创建滚动区域内容
        self.scrollAreaWidgetContents = QWidget()
        self.scrollAreaWidgetContents.setObjectName(u"scrollAreaWidgetContents")
        self.scrollAreaWidgetContents.setEnabled(True)
        self.scrollAreaWidgetContents.setGeometry(QRect(0, 0, 692, 350))
                # 设置滚动区域内容的大小策略
        sizePolicy1 = QSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        sizePolicy1.setHorizontalStretch(0)
        sizePolicy1.setVerticalStretch(0)
        sizePolicy1.setHeightForWidth(self.scrollAreaWidgetContents.sizePolicy().hasHeightForWidth())
        self.scrollAreaWidgetContents.setSizePolicy(sizePolicy1)
        self.scrollAreaWidgetContents.setMinimumSize(QSize(200, 200))
                # 为该滚动区域创建垂直布局
        self.verticalLayout_2 = QVBoxLayout(self.scrollAreaWidgetContents)
        self.verticalLayout_2.setSpacing(0)
        self.verticalLayout_2.setObjectName(u"verticalLayout_2")
        self.verticalLayout_2.setSizeConstraint(QLayout.SizeConstraint.SetDefaultConstraint)
        self.verticalLayout_2.setContentsMargins(-1, 0, -1, 0)
                # 设置滚动区域布局
        self.scrollArea.setWidget(self.scrollAreaWidgetContents) # Notice:这句很关键，少了这一句全部变空白
        
        
        # 创建滚动区域中的列表 Widget
            # 创建 widget
        self.widget = QWidget(self.scrollAreaWidgetContents)
        self.widget.setObjectName(u"widget")
            # 设置widget的大小策略
        sizePolicy2 = QSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Minimum)
        sizePolicy2.setHorizontalStretch(0)
        sizePolicy2.setVerticalStretch(0)
        sizePolicy2.setHeightForWidth(self.widget.sizePolicy().hasHeightForWidth())
        self.widget.setSizePolicy(sizePolicy2)
            # 为该 widget 创建水平布局
        self.horizontalLayout = QHBoxLayout(self.widget)
        self.horizontalLayout.setObjectName(u"horizontalLayout")
            # 将该 widget 添加到垂直布局中
        self.verticalLayout_2.addWidget(self.widget)
        # 创建文件显示标签
            # 创建Widget
        self.label = QLabel(self.widget)
        self.label.setObjectName(u"label")
            # 设置label的大小策略
        sizePolicy3 = QSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)
        sizePolicy3.setHorizontalStretch(0)
        sizePolicy3.setVerticalStretch(0)
        sizePolicy3.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy3)
            # 将label添加到水平布局中
        self.horizontalLayout.addWidget(self.label)

        # 创建按钮
            # 创建Widget
        self.pushButton = QPushButton(self.widget)
        self.pushButton.setObjectName(u"pushButton")
            # 添加按钮到水平布局中
        self.horizontalLayout.addWidget(self.pushButton)
        # 创建按钮
            # 创建Widget
        self.pushButton_2 = QPushButton(self.widget)
        self.pushButton_2.setObjectName(u"pushButton_2")
            # 添加按钮到水平布局中
        self.horizontalLayout.addWidget(self.pushButton_2)
        # 创建按钮
            # 创建Widget
        self.pushButton_3 = QPushButton(self.widget)
        self.pushButton_3.setObjectName(u"pushButton_3")
            # 添加按钮到水平布局中
        self.horizontalLayout.addWidget(self.pushButton_3)


        
        


        self.retranslateUi(Form)

        QMetaObject.connectSlotsByName(Form)
    # setupUi

    def retranslateUi(self, Form):
        Form.setWindowTitle(QCoreApplication.translate("Form", u"Form", None))
        self.label.setText(QCoreApplication.translate("Form", u"TextLabel", None))
        self.pushButton.setText(QCoreApplication.translate("Form", u"PushButton", None))
        self.pushButton_2.setText(QCoreApplication.translate("Form", u"PushButton", None))
        self.pushButton_3.setText(QCoreApplication.translate("Form", u"PushButton", None))

    # retranslateUi

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ui = Ui_Form()
    Form = QWidget()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())
