import sys
import os
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from ut_excel2word import UnitTestTestCaseDocumentGenerator
from jt_htm2excel import HtmlTableParser
from build.version_data import version_data


class Page(QWidget):
    def __init__(self, title):
        super().__init__()
        self.setAcceptDrops(True)  # 开启拖拽支持
        layout = QVBoxLayout(self)
        label = QLabel(title, self)
        label.setFont(QFont("Roman times", 12, QFont.Bold))
        layout.addWidget(label)

        self.textEdit = QTextEdit(self)
        self.textEdit.setStyleSheet("color: black")
        self.textEdit.setReadOnly(True)
        layout.addWidget(self.textEdit)

        self.choise_file = None

    def open_file(self):
        self.choise_file, _ = QFileDialog.getOpenFileName(
            self, " ", "", "All Files (*)"
        )
        if self.choise_file:
            self.textEdit.append(
                f"<font color='green'>choise file: <b>{self.choise_file}</b></font>"
            )

    def dragEnterEvent(self, event):
        """当拖拽进入窗口时触发"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        """当拖拽放下时触发"""
        if event.mimeData().hasUrls():
            file_urls = event.mimeData().urls()
            file_path = file_urls[0].toLocalFile()
            if os.path.isfile(file_path):
                self.choise_file = file_path
                self.textEdit.append(
                    f"<font color='green'>choise file: <b>{self.choise_file}</b></font>"
                )
            else:
                self.textEdit.append("<font color='red'>Invalid file</font>")
        else:
            self.textEdit.append("<font color='red'>No valid file detected</font>")


class UnitTest_Excel2ReportDocx_Page(Page):
    def __init__(self):
        super().__init__("单元测试 表格转文档(报告)")

        btn = QPushButton("选择文件", self)
        btn.clicked.connect(self.open_file)
        self.layout().addWidget(btn)

        btn2 = QPushButton("生成文档", self)
        btn2.clicked.connect(self.generate_docx)
        self.layout().addWidget(btn2)

    def generate_docx(self):
        if not self.choise_file:
            self.textEdit.append("<font color='red'>need choise a file</font>")
            return
        try:
            generator = UnitTestTestCaseDocumentGenerator()
            save_file_name = os.path.splitext(self.choise_file)[0] + "_报告oa.docx"

            def progress_callback(message):
                self.textEdit.append(f"<font color='blue'>{message}</font>")
                QApplication.processEvents()  # 保证实时刷新 UI

            count = generator.generate_document(
                self.choise_file,
                save_file_name,
                mode="report",
                callback=progress_callback,
            )
            self.textEdit.append(
                f"<font color='green'>{save_file_name} <br> is created with {count} test cases</font>"
            )
        except Exception as e:
            self.textEdit.append(f"<font color='red'>Error: {e}</font>")


class UnitTest_Excel2InstructionDocx_Page(Page):
    def __init__(self):
        super().__init__("单元测试 表格转文档(说明)")

        btn = QPushButton("选择文件", self)
        btn.clicked.connect(self.open_file)
        self.layout().addWidget(btn)

        btn2 = QPushButton("生成文档", self)
        btn2.clicked.connect(self.generate_docx)
        self.layout().addWidget(btn2)

    def generate_docx(self):
        if not self.choise_file:
            self.textEdit.append("<font color='red'>need choise a file</font>")
            return
        try:
            generator = UnitTestTestCaseDocumentGenerator()
            save_file_name = os.path.splitext(self.choise_file)[0] + "_说明oa.docx"

            def progress_callback(message):
                self.textEdit.append(f"<font color='blue'>{message}</font>")
                QApplication.processEvents()  # 保证实时刷新 UI

            count = generator.generate_document(
                self.choise_file,
                save_file_name,
                mode="instructions",
                callback=progress_callback,
            )
            self.textEdit.append(
                f"<font color='green'>{save_file_name} <br> is created with {count} test cases</font>"
            )
        except Exception as e:
            self.textEdit.append(f"<font color='red'>Error: {e}</font>")


class StaticTestPage(Page):
    def __init__(self):
        super().__init__("静态测试 testbed报告转表格")

        btn = QPushButton("选择文件", self)
        btn.clicked.connect(self.open_file)
        self.layout().addWidget(btn)

        btn2 = QPushButton("生成表格", self)
        btn2.clicked.connect(self.generate_excel)
        self.layout().addWidget(btn2)

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


# 程序信息页面
class ProgramInfo(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)

        label = QLabel("about OA tools", self)
        label.setFont(QFont("Roman times", 20, QFont.Bold))
        layout.addWidget(label)

        version_label = QLabel(f"软件版本: {version_data['version']}", self)
        version_label.setFont(QFont("Arial", 14))
        layout.addWidget(version_label)

        branch_label = QLabel(f"构建分支: {version_data['git_branch']}", self)
        branch_label.setFont(QFont("Arial", 14))
        layout.addWidget(branch_label)

        hash_label = QLabel(f"基于代码版本: {version_data['git_hash']}", self)
        hash_label.setFont(QFont("Arial", 14))
        layout.addWidget(hash_label)

        build_time_label = QLabel(f"构建时间: {version_data['build_time']}", self)
        build_time_label.setFont(QFont("Arial", 14))
        layout.addWidget(build_time_label)

        change_log_label = QLabel(f"修改记录: \n{version_data['git_log']}", self)
        change_log_label.setFont(QFont("Arial", 12))
        layout.addWidget(change_log_label)


def resource_path(relative_path):
    """获取资源文件的路径，兼容pyinstaller打包"""
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# 主窗口类
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # 设置窗口属性
        self.setWindowTitle(
            f"OA Tools {version_data['version']} {version_data['build_time']}"
        )
        self.setWindowIcon(QIcon(resource_path("assets/app_icon.svg")))
        self.resize(500, 400)

        # 创建菜单栏和页面
        self.createMenus()
        self.createPages()

    def createMenus(self):
        # 创建菜单栏
        menubar = self.menuBar()

        # 定义菜单项和动作
        menu_items = {
            "文档转换": [
                ("单元测试 表格转文档(报告)", "ut_excel2Rword"),
                ("单元测试 表格转文档(说明)", "ut_excel2Iword"),
                ("静态测试 报告转表格", "static_test"),
            ],
            "基于深度学习查找": [
                ("VGG", "vgg"),
                ("ResNet", "resnet"),
                ("about OA tools", "program_info"),
            ],
        }

        # 动态生成菜单和绑定动作
        for menu_name, actions in menu_items.items():
            menu = menubar.addMenu(menu_name)
            for action_name, page_key in actions:
                action = QAction(action_name, self)
                action.triggered.connect(lambda _, key=page_key: self.switchPage(key))
                menu.addAction(action)

    def createPages(self):
        # 创建页面
        self.stackedWidget = QStackedWidget()
        self.setCentralWidget(self.stackedWidget)

        # 定义页面字典
        self.pages = {
            "ut_excel2Rword": UnitTest_Excel2ReportDocx_Page(),
            "ut_excel2Iword": UnitTest_Excel2InstructionDocx_Page(),
            "static_test": StaticTestPage(),
            "vgg": VGGPage(),
            "resnet": ResNetPage(),
            "program_info": ProgramInfo(),
        }

        # 将页面添加到堆栈组件中
        for page in self.pages.values():
            self.stackedWidget.addWidget(page)

    def switchPage(self, page_key):
        # 切换页面
        page = self.pages.get(page_key)
        if page:
            self.stackedWidget.setCurrentWidget(page)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.show()
    sys.exit(app.exec_())
