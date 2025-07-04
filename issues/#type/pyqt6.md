# 错误的清空QlineWidget的方式

执行以下代码：

**self**.date_2 **=** **""**

**self**.foodType_2 **=** **""**

**self**.name_2 **=** **""**

**self**.info_2 **=** **""**

**self**.amount_2 **=** **""**

**self**.LineEdit **=** **""**

**self**.LineEdit_2 **=** **""**

**self**.LineEdit_3 **=** **""**

### 发生的情况

1. **逻辑层面** ：

* 这些语句会在 [self](vscode-file://vscode-app/e:/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html) 对象上创建新的属性（如 [self.date_2](vscode-file://vscode-app/e:/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html)、[self.foodType_2](vscode-file://vscode-app/e:/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html) 等），并将它们的值设置为空字符串 `""`。
* 如果这些属性之前已经存在，它们的值会被覆盖为 `""`。

1. **图形化界面（GUI）层面** ：

* 如果这些属性（如 [self.date_2](vscode-file://vscode-app/e:/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html) 等）原本是绑定到 `QLineEdit`（或其他输入框控件）的对象，这些语句会将它们替换为普通的字符串 `""`。
* 替换后，这些属性将不再是 `QLineEdit` 对象，而是普通的字符串，因此无法再通过这些属性操作对应的输入框控件。
* **结果：图形化界面上的输入框内容不会被清空，仍然显示上次输入的数据** 。

---

### 为什么会这样？

* 在 PyQt/PySide 中，[self.date_2](vscode-file://vscode-app/e:/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html) 等属性通常是绑定到 `QLineEdit` 控件的对象。
* 当你执行 [self.date_2 = &#34;&#34;](vscode-file://vscode-app/e:/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html) 时，[self.date_2](vscode-file://vscode-app/e:/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html) 不再指向原来的 `QLineEdit` 对象，而是被重新赋值为一个字符串 `""`。
* 因此，原来的 `QLineEdit` 控件仍然存在于界面上，但你已经失去了对它的引用，无法通过 [self.date_2](vscode-file://vscode-app/e:/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html) 操作它。

---

### 正确的做法

如果你想清空图形化界面上的输入框内容，应该调用 `QLineEdit` 的 [setText(&#34;&#34;)](vscode-file://vscode-app/e:/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html) 方法，而不是直接赋值。

修改代码如下：

**self**.date_2.setText**(**""**)**  **# 清空 QLineEdit 的内容**

**self**.foodType_2.setText**(**""**)**

**self**.name_2.setText**(**""**)**

**self**.info_2.setText**(**""**)**

**self**.amount_2.setText**(**""**)**

**self**.LineEdit.setText**(**""**)**

**self**.LineEdit_2.setText**(**""**)**

**self**.LineEdit_3.setText**(**""**)**

---

### 执行后的效果

1. **逻辑层面** ：

* [self.date_2](vscode-file://vscode-app/e:/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html) 等属性仍然是 `QLineEdit` 对象。
* 调用 [setText(&#34;&#34;)](vscode-file://vscode-app/e:/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html) 方法会清空这些输入框的内容。

1. **图形化界面（GUI）层面** ：

* 界面上的输入框内容会被清空，显示为空。

---

### 总结

* **错误的方式** ：直接赋值 [self.date_2 = &#34;&#34;](vscode-file://vscode-app/e:/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html) 会导致属性被替换为字符串，失去对原控件的引用。
* **正确的方式** ：使用 [setText(&#34;&#34;)](vscode-file://vscode-app/e:/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html) 方法清空输入框内容，同时保持属性仍然指向原控件。



# 问题二

```python-repl

    def back_up_manager(self):
        """
        总备份管理
        :param: self
        :return: None
        """
        # 创建界面
            # 创建窗口
        self.BackUpWindow = QWidget()
        self.BackUpWindow.setWindowTitle("备份管理")
        self.BackUpWindow.resize(800, 600)
        self.BackUpWindow.setObjectName("BackUpWindow")
        self.BackUpWindow.show()
            # 添加主布局
        self.window_layout_dom1 = QVBoxLayout(self.BackUpWindow)                                                            # 创建垂直布局
        self.BackUpWindow.setLayout(self.window_layout_dom1)                                                                # 应用其到窗口
                
        # 添加滚动区域  
            # 创建滚动区域                                         
        self.window_scroll_area_dom2 = QScrollArea(self.BackUpWindow)                                                       # 创建滚动区域
        self.window_scroll_area_dom2.setObjectName("window_scroll_area_dom2")                                               # 设置对象名称                                        
            # 设置滚动区域大小策略
        window_scroll_area_dom2_sizepolicy = QSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)        # 设置滚动区域的大小策略
        window_scroll_area_dom2_sizepolicy.setHorizontalStretch(0)                                                          # 设置水平拉伸为0
        window_scroll_area_dom2_sizepolicy.setVerticalStretch(0)                                                            # 设置垂直拉伸为0
        window_scroll_area_dom2_sizepolicy.setHeightForWidth(self.window_scroll_area_dom2.sizePolicy().hasHeightForWidth()) # 设置滚动区域大小策略
            # 应用滚动区域大小策略
        self.window_scroll_area_dom2.setSizePolicy(window_scroll_area_dom2_sizepolicy) 
        self.window_scroll_area_dom2.setSizeAdjustPolicy(QAbstractScrollArea.SizeAdjustPolicy.AdjustIgnored)  
        self.window_scroll_area_dom2.setWidgetResizable(True) 
            # 将滚动区域添加到主布局
        self.window_layout_dom1.addWidget(self.window_scroll_area_dom2)
            # 设置滚动区域的内容容器
                # 创建内容容器
        self.window_scroll_area_contents_dom3 = QWidget() # 创建内容容器不需要传入父widget
        self.window_scroll_area_contents_dom3.setObjectName("window_scroll_area_contents_dom3")
                # 设置内容容器大小策略
        window_scroll_area_contents_dom3_sizepolicy = QSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        window_scroll_area_contents_dom3_sizepolicy.setHorizontalStretch(0)
        window_scroll_area_contents_dom3_sizepolicy.setVerticalStretch(0)
        window_scroll_area_contents_dom3_sizepolicy.setHeightForWidth(self.window_scroll_area_contents_dom3.sizePolicy().hasHeightForWidth())
        self.window_scroll_area_contents_dom3.setSizePolicy(window_scroll_area_contents_dom3_sizepolicy)
                # 为内容容器创建垂直布局
        self.window_scroll_area_contents_layout_dom3 = QVBoxLayout(self.window_scroll_area_contents_dom3)
        self.window_scroll_area_contents_layout_dom3.setObjectName("window_scroll_area_contents_layout_dom3")
        self.window_scroll_area_contents_layout_dom3.setSpacing(0)
        self.window_scroll_area_contents_layout_dom3.setSizeConstraint(QLayout.SizeConstraint.SetDefaultConstraint)
        self.window_scroll_area_contents_layout_dom3.setContentsMargins(-1, 0,-1, 0)
                # 将垂直布局应用其到内容容器
        self.window_scroll_area_contents_dom3.setLayout(self.window_scroll_area_contents_layout_dom3)
                # 将内容容器应用其到滚动区域
        self.window_scroll_area_dom2.setWidget(self.window_scroll_area_contents_dom3) # Notice:这句如果消失界面会变空白

        """
        读取 backup 目录下的文件夹名，存储成一维列表
        """
        backup_folder_name = [ folder_name for folder_name in os.listdir(".\\src\\data\\storage\\backup") if os.path.isdir(os.path.join(".\\src\\data\\storage\\backup", folder_name))]
        # 判断 backup 目录下是否有备份文件夹
        if backup_folder_name  != []:
            print(f"Notice: 读取备份 backup 目录下的文件夹名: {backup_folder_name}")
        else:
            print("Notice: backup 目录下没有文件夹")
            QMessageBox.information(None, "提示", "备份目录下没有文件夹", QMessageBox.Ok)
            return

        # 为读取到的每个子文件夹名创建一个widget，包含 文件夹名和查看备份-还原备份-删除备份 3 个按钮
        for name_dom4 in backup_folder_name:
        
            # 动态创建备份按钮属性名
            back_up_item_check_button_dom5 = f"{name_dom4}_back_up_button_dom5"
            # 动态创建还原按钮属性名
            back_up_item_restore_button_dom5 = f"{name_dom4}_restore_button_dom5"
            # 动态创建删除按钮属性名
            back_up_item_delete_button_dom5 = f"{name_dom4}_delete_button_dom5"
  
            # 从内容容器中创建存放每一个列表显示条目的 widget 
                # 创建 widget
            self.name_dom4 = QWidget(self.window_scroll_area_contents_dom3)
            self.name_dom4.setObjectName(name_dom4) 
                # 设置 widget 大小策略
            name_dom4_sizepolicy = QSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
            name_dom4_sizepolicy.setHorizontalStretch(0)
            name_dom4_sizepolicy.setVerticalStretch(0)
            name_dom4_sizepolicy.setHeightForWidth(self.name_dom4.sizePolicy().hasHeightForWidth())
            self.name_dom4.setSizePolicy(name_dom4_sizepolicy)
                # 为该 widget 创建一个布局
            self.back_up_item_layout_dom4 = QHBoxLayout(self.name_dom4)
            self.name_dom4.setLayout(self.back_up_item_layout_dom4)
                # 将 widget 应用其到父内容容器的垂直布局中
            self.window_scroll_area_contents_layout_dom3.addWidget(self.name_dom4)     #
            

            # 创建文件夹名标签
            self.back_up_item_label_dom5 = QLabel(name_dom4, self.name_dom4)
            self.back_up_item_layout_dom4.addWidget(self.back_up_item_label_dom5)               # 加入到布局

            # 创建查看备份按钮 
            self.back_up_item_check_button_dom5 = QPushButton("查看备份", self.name_dom4)
            self.back_up_item_check_button_dom5.setObjectName(f"{name_dom4}")
            self.back_up_item_check_button_dom5.clicked.connect(lambda:view_backup(self,self.back_up_item_check_button_dom5.objectName())) # 使用lambda函数，避免按钮点击时，参数被提前执行，同时也能够进行传参操作
            self.back_up_item_layout_dom4.addWidget(self.back_up_item_check_button_dom5)             # 加入到布局

            # 创建还原备份按钮
            self.back_up_item_restore_button_dom5 = QPushButton("还原备份", self.name_dom4)
            self.back_up_item_restore_button_dom5.setObjectName(f"{name_dom4}")
            self.back_up_item_restore_button_dom5.clicked.connect(lambda:restore_backup(self,self.back_up_item_restore_button_dom5.objectName()))
            self.back_up_item_layout_dom4.addWidget(self.back_up_item_restore_button_dom5)             # 加入到布局
  
            # 创建删除备份按钮
            self.back_up_item_delete_button_dom5 = QPushButton("删除备份", self.name_dom4)
            self.back_up_item_delete_button_dom5.setObjectName(f"{name_dom4}")
            self.back_up_item_delete_button_dom5.clicked.connect(lambda:delete_backup(self,self.back_up_item_delete_button_dom5.objectName()))
            self.back_up_item_layout_dom4.addWidget(self.back_up_item_delete_button_dom5)             # 加入到布局 
```

Q:这段代码有个问题，因为无法动态创建属性造成每次删除条目的时候都是删除back_up_item_delete_button_dom5这个条目的

在您的代码中，问题的核心在于 **Lambda 函数未能正确捕获每次循环中 `name_dom4` 的当前值**，导致所有按钮的点击事件最终都传递的是最后一次循环的 `name_dom4` 值。这是 Python 中常见的闭包陷阱问题。

---

### ✅ 问题分析

在循环中，您为每个备份文件夹创建了三个按钮（查看、还原、删除），并为每个按钮的 `clicked` 信号连接了一个 Lambda 函数。例如：

```python
self.back_up_item_delete_button_dom5.clicked.connect(
    lambda: delete_backup(self, self.back_up_item_delete_button_dom5.objectName())
)
```

这段代码的问题在于：**Lambda 函数在定义时不会立即执行，而是在按钮被点击时才执行**。此时，`self.back_up_item_delete_button_dom5` 已经被循环覆盖为最后一个按钮，导致所有按钮点击后都传递的是最后一个文件夹的名称。

---

### ✅ 解决方案

为了避免这个问题，我们需要 **在每次循环中将当前的 `name_dom4` 值绑定到 Lambda 函数中**。可以使用 **默认参数** 来实现这一点，因为默认参数在函数定义时就会被求值。

#### ✅ 修改后的 Lambda 写法（推荐）

```python
self.back_up_item_delete_button_dom5.clicked.connect(
    lambda _, name=name_dom4: delete_backup(self, name)
)
```

这样，每次循环中，`name` 参数会被绑定为当前的 `name_dom4` 值，而不是在点击时再去查找变量。

> ⚠️ 注意：`_` 是一个占位符，用于接收 Qt 传递的 `checked` 参数（对于 `QPushButton.clicked` 信号来说是布尔值），但我们不需要它，所以可以忽略。

---

### ✅ 完整修复示例（以删除按钮为例）

```python
self.back_up_item_delete_button_dom5 = QPushButton("删除备份", self.name_dom4)
self.back_up_item_delete_button_dom5.setObjectName(name_dom4)

# 使用默认参数绑定当前 name_dom4 的值
self.back_up_item_delete_button_dom5.clicked.connect(
    lambda _, name=name_dom4: delete_backup(self, name)
)

self.back_up_item_layout_dom4.addWidget(self.back_up_item_delete_button_dom5)
```

同样的方式适用于“查看”和“还原”按钮：

```python
self.back_up_item_check_button_dom5.clicked.connect(
    lambda _, name=name_dom4: view_backup(self, name)
)

self.back_up_item_restore_button_dom5.clicked.connect(
    lambda _, name=name_dom4: restore_backup(self, name)
)
```

---

### ✅ 其他优化建议

- **避免使用 `self.name_dom4`**：在循环中创建的 `QWidget` 不需要绑定为 `self` 的属性，可以改为局部变量，避免污染命名空间和潜在的属性冲突。
- **不要使用 `self.back_up_item_check_button_dom5` 等属性**：这些属性在每次循环中都会被覆盖，最终只保留最后一个按钮的引用，没有实际意义，建议改为局部变量。

---

### ✅ 总结

| 问题                           | 原因                          | 解决方案                        |
| ------------------------------ | ----------------------------- | ------------------------------- |
| 删除按钮始终删除最后一个备份   | Lambda 函数未正确捕获循环变量 | 使用默认参数绑定当前变量值      |
| 所有按钮点击后都执行相同操作   | Lambda 函数引用的是变量而非值 | 使用默认参数捕获当前值          |
| 控件被错误绑定为 `self` 属性 | 属性被循环覆盖，导致引用混乱  | 改用局部变量，避免污染 `self` |

---

### ✅ 最终推荐写法（简化版）

```python
for folder_name in backup_folder_name:
    item_widget = QWidget(self.window_scroll_area_contents_dom3)
    layout = QHBoxLayout(item_widget)

    label = QLabel(folder_name)
    check_button = QPushButton("查看备份")
    restore_button = QPushButton("还原备份")
    delete_button = QPushButton("删除备份")

    layout.addWidget(label)
    layout.addWidget(check_button)
    layout.addWidget(restore_button)
    layout.addWidget(delete_button)

    check_button.clicked.connect(lambda _, name=folder_name: view_backup(self, name))
    restore_button.clicked.connect(lambda _, name=folder_name: restore_backup(self, name))
    delete_button.clicked.connect(lambda _, name=folder_name: delete_backup(self, name))

    self.window_scroll_area_contents_layout_dom3.addWidget(item_widget)
```

这样写法更清晰、更安全，避免了属性污染和 Lambda 捕获错误的问题。
