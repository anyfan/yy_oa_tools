import sys
import os
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from ut_excel2word import UnitTestTestCaseDocumentGenerator
from jt_htm2excel import HtmlTableParser
from build.version_data import version_data


# 定义各个页面类
class UnitTest_Excel2ReportDocx_Page(QWidget):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)  # 开启拖拽支持
        layout = QVBoxLayout(self)
        label = QLabel("单元测试 表格转文档(报告)", self)
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
            generator = UnitTestTestCaseDocumentGenerator()
            save_file_name = os.path.splitext(self.choise_file)[0] + "_oa.docx"

            def progress_callback(message):
                self.textEdit.append(f"<font color='blue'>{message}</font>")
                QApplication.processEvents()  # 保证实时刷新 UI

            count = generator.generate_report_document(
                self.choise_file, save_file_name, progress_callback
            )
            self.textEdit.append(
                f"<font color='green'>{save_file_name} <br> is created with {count} test cases</font>"
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


class UnitTest_Excel2InstructionDocx_Page(QWidget):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)  # 开启拖拽支持
        layout = QVBoxLayout(self)
        label = QLabel("单元测试 表格转文档(说明)", self)
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
            generator = UnitTestTestCaseDocumentGenerator()
            save_file_name = os.path.splitext(self.choise_file)[0] + "_oa.docx"

            def progress_callback(message):
                self.textEdit.append(f"<font color='blue'>{message}</font>")
                QApplication.processEvents()  # 保证实时刷新 UI

            count = generator.generate_instruction_document(
                self.choise_file, save_file_name, progress_callback
            )
            self.textEdit.append(
                f"<font color='green'>{save_file_name} <br> is created with {count} test cases</font>"
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


class StaticTestPage(QWidget):
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)  # 开启拖拽支持

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
            self, "选择文件", "", "All Files (*)"
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
            # 假设 HtmlTableParser 是一个定义好的类
            parser = HtmlTableParser(self.choise_file)
            save_file_name = os.path.splitext(self.choise_file)[0] + "_oa.xlsx"
            parser.parse_tables()
            parser.save_to_excel(save_file_name)
            self.textEdit.append(
                f"<font color='green'>{save_file_name} is created</font>"
            )
        except Exception as e:
            self.textEdit.append(f"<font color='red'>Error: {e}</font>")

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
        label.setFont(QFont("Roman times", 50, QFont.Bold))
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


# 主窗口类
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(
            "OA tools " + version_data["version"] + "  " + version_data["build_time"]
        )
        self.resize(500, 400)

        # 创建菜单栏
        menubar = self.menuBar()
        menu_conversion = menubar.addMenu("文档转换")

        menu_deep_learning = menubar.addMenu("基于深度学习查找")

        # 创建动作
        action_ut_excel2Rword = QAction("单元测试 表格转文档(报告)", self)
        action_ut_excel2Iword = QAction("单元测试 表格转文档(说明)", self)
        action_static_test = QAction("静态测试 报告转表格", self)

        action_vgg = QAction("VGG", self)
        action_resnet = QAction("ResNet", self)
        action_program_info = QAction("about OA tools", self)

        # 将动作添加到菜单
        menu_conversion.addActions(
            [action_ut_excel2Rword, action_ut_excel2Iword, action_static_test]
        )

        menu_deep_learning.addActions([action_vgg, action_resnet, action_program_info])

        # 创建页面
        self.stackedWidget = QStackedWidget()
        self.setCentralWidget(self.stackedWidget)

        self.pages = {
            "ut_excel2Rword": UnitTest_Excel2ReportDocx_Page(),
            "ut_excel2Iword": UnitTest_Excel2InstructionDocx_Page(),
            "static_test": StaticTestPage(),
            "vgg": VGGPage(),
            "resnet": ResNetPage(),
            "program_info": ProgramInfo(),
        }

        for page in self.pages.values():
            self.stackedWidget.addWidget(page)

        # 动作绑定槽函数
        action_ut_excel2Rword.triggered.connect(
            lambda: self.switchPage("ut_excel2Rword")
        )
        action_ut_excel2Iword.triggered.connect(
            lambda: self.switchPage("ut_excel2Iword")
        )
        action_static_test.triggered.connect(lambda: self.switchPage("static_test"))
        action_vgg.triggered.connect(lambda: self.switchPage("vgg"))
        action_resnet.triggered.connect(lambda: self.switchPage("resnet"))
        action_program_info.triggered.connect(lambda: self.switchPage("program_info"))

    def switchPage(self, page_name):
        page = self.pages.get(page_name)
        if page:
            self.stackedWidget.setCurrentWidget(page)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.show()
    sys.exit(app.exec_())
