import sys
import os
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from ut_excel2word import TestDocumentGenerator
from jt_htm2excel import HtmlTableParser

# 定义各个页面类
class ColorPage(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        label = QLabel("单元测试 表格转文档", self)
        label.setFont(QFont("Roman times", 12, QFont.Bold))
        layout.addWidget(label)

        btn = QPushButton("选择文件", self)
        btn.clicked.connect(self.open_file)
        layout.addWidget(btn)
        # 显示以及选择的文件
        self.textEdit = QTextEdit(self)
        self.textEdit.setStyleSheet("color: black")
        self.textEdit.setReadOnly(True)
        layout.addWidget(self.textEdit)

        btn2 = QPushButton("生成文档", self)
        btn2.clicked.connect(self.generate_docx)
        layout.addWidget(btn2)
        self.choise_file = None

    def generate_docx(self):
        if not self.choise_file:
            self.textEdit.append("<font color='red'>need choise a file</font>")
            return
        try:
            generator = TestDocumentGenerator("template/template.docx")
            save_file_name = os.path.splitext(self.choise_file)[0] + "_oa.docx"

            def progress_callback(message):
                self.textEdit.append(f"<font color='blue'>{message}</font>")
                QApplication.processEvents()  # 保证实时刷新 UI

            count = generator.generate(
                self.choise_file, save_file_name, progress_callback
            )
            self.textEdit.append(
                f"<font color='green'>{save_file_name} is created with {count} test cases</font>"
            )
        except Exception as e:
            self.textEdit.append(f"<font color='red'>Error: {e}</font>")

    def open_file(self):
        self.choise_file, _ = QFileDialog.getOpenFileName(
            self, " ", "", "All Files (*)"
        )
        if self.choise_file:
            self.textEdit.append(
                f"<font color='green'>choise file: <b>{self.choise_file}</b></font>"
            )


class TexturePage(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        label = QLabel("静态测试 testbed报告转表格", self)
        label.setFont(QFont("Roman times", 12, QFont.Bold))
        layout.addWidget(label)
        
        btn = QPushButton("选择文件", self)
        btn.clicked.connect(self.open_file)
        layout.addWidget(btn)
         # 显示以及选择的文件
        self.textEdit = QTextEdit(self)
        self.textEdit.setStyleSheet("color: black")
        self.textEdit.setReadOnly(True)
        layout.addWidget(self.textEdit)
        
        btn2 = QPushButton("生成表格", self)
        btn2.clicked.connect(self.generate_excel)
        layout.addWidget(btn2)
        self.choise_file = None
        
    def open_file(self):
        self.choise_file, _ = QFileDialog.getOpenFileName(
            self, " ", "", "All Files (*)"
        )
        if self.choise_file:
            self.textEdit.append(
                f"<font color='green'>choise file: <b>{self.choise_file}</b></font>"
            )

    def generate_excel(self):
        if not self.choise_file:    
            self.textEdit.append("<font color='red'>need choise a file</font>")
            return
        try:
            parser = HtmlTableParser(self.choise_file)
            save_file_name = os.path.splitext(self.choise_file)[0] + "_oa.xlsx"
            parser.parse_tables()
            parser.save_to_excel(save_file_name)
            self.textEdit.append(
                f"<font color='green'>{save_file_name} is created</font>"
            )
            
        except Exception as e:
            self.textEdit.append(f"<font color='red'>Error: {e}</font>")
        

class VGGPage(QWidget):
    def __init__(self):
        super().__init__()
        layout = QHBoxLayout(self)
        label = QLabel("VGG", self)
        label.setAlignment(Qt.AlignCenter)
        label.setFont(QFont("Roman times", 50, QFont.Bold))
        layout.addWidget(label)


class ResNetPage(QWidget):
    def __init__(self):
        super().__init__()
        layout = QHBoxLayout(self)
        label = QLabel("ResNet", self)
        label.setAlignment(Qt.AlignCenter)
        label.setFont(QFont("Roman times", 50, QFont.Bold))
        layout.addWidget(label)


# 主窗口类
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OA tools @ 2024")
        self.resize(1440, 773)

        # 创建菜单栏
        menubar = self.menuBar()
        menu_color = menubar.addMenu("文档转换")

        menu_deep = menubar.addMenu("基于深度学习查找")

        # 创建动作
        action_rgb = QAction("单元测试 表格转文档", self)
        action_gabor = QAction("静态测试 报告转表格", self)

        action_vgg = QAction("VGG", self)
        action_resnet = QAction("ResNet", self)

        # 将动作添加到菜单
        # menu_color.addAction(action_rgb)
        menu_color.addActions([action_rgb, action_gabor])

        menu_deep.addActions([action_vgg, action_resnet])

        # 创建页面
        self.stackedWidget = QStackedWidget()
        self.setCentralWidget(self.stackedWidget)

        self.pages = {
            "color": ColorPage(),
            "texture": TexturePage(),
            "vgg": VGGPage(),
            "resnet": ResNetPage(),
        }

        for page in self.pages.values():
            self.stackedWidget.addWidget(page)

        # 动作绑定槽函数
        action_rgb.triggered.connect(lambda: self.switchPage("color"))
        action_gabor.triggered.connect(lambda: self.switchPage("texture"))
        action_vgg.triggered.connect(lambda: self.switchPage("vgg"))
        action_resnet.triggered.connect(lambda: self.switchPage("resnet"))

    def switchPage(self, page_name):
        page = self.pages.get(page_name)
        if page:
            self.stackedWidget.setCurrentWidget(page)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.show()
    sys.exit(app.exec_())
