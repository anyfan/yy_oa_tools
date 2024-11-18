# coding=utf-8
import sys
import os
import pandas as pd
from copy import deepcopy
from docx import Document
from string import Template


class TestDocumentGenerator:
    def __init__(self, template_path):
        self.template_path = template_path
        self.template_docx = Document(self.resource_path(template_path))
        self.new_docx = Document(self.resource_path("template/new.docx"))
        self.template_count = 0

    @staticmethod
    def resource_path(relative_path):
        """获取资源文件的路径，兼容pyinstaller打包"""
        if getattr(sys, "frozen", False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    @staticmethod
    def fill_merged_cells(df):
        """填充DataFrame中的合并单元格"""
        for col in df.columns:
            df[col] = df[col].ffill().infer_objects(copy=False)
        return df

    @staticmethod
    def process_data(data_obj):
        """处理数据对象，适配模板"""
        rm_br = lambda x: (
            x.replace("\n", "") if isinstance(x, str) else x
        )  # 去除所有换行符
        add_br = lambda x: "\n" + x if "\n" in x else x  # 存在换行符,句首添加换行符
        str_time = lambda x: (
            x.date() if isinstance(x, pd.Timestamp) else x
        )  # timestamp转str
        get_tester = lambda x: x.split("/")[0] if "/" in x else x  # 获取测试者
        get_checker = lambda x: x.split("/")[1] if "/" in x else x  # 获取校对者

        return {
            "dybs": data_obj["单元标识"],
            "sjzs": data_obj["设计追踪"],
            "gnms": data_obj["功能描述"],
            "qdmk": data_obj["驱动模块"],
            "dzmk": rm_br(data_obj["打桩模块"]),
            "fgljl": rm_br(data_obj["覆盖率记录"]),
            "bz": data_obj["备注"],
            "csylmc": data_obj["测试用例名称"],
            "csylbs": data_obj["测试用例标识"],
            "cslx": data_obj["测试类型"],
            "hdz": data_obj["活动桩"],
            "cryj": add_br(data_obj["插入语句"]),
            "qjbl": add_br(data_obj["创建全局变量"]),
            "cssm": data_obj["测试说明"],
            "srblm": data_obj["输入变量名称"],
            "srblqz": data_obj["取值"],
            "scblm": data_obj["输出变量名称"],
            "yqsc": data_obj["预期输出"],
            "sjsc": data_obj["实际输出"],
            "sjz": get_tester(data_obj["测试者/校对者"]),
            "xdz": get_checker(data_obj["测试者/校对者"]),
            "csry": get_tester(data_obj["测试者/校对者"]),
            "cszxsj": str_time(data_obj["测试时间"]),
        }

    def add_table(self, data_obj):
        """根据数据生成表格并添加到文档"""
        data = self.process_data(data_obj)

        new_table = deepcopy(self.template_docx.tables[0])

        for row in new_table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run_template = Template(run.text)
                        run.text = run_template.substitute(data)

        paragraph = self.new_docx.add_paragraph(f"{data_obj['测试用例名称']}")
        paragraph._p.addnext(new_table._element)
        self.new_docx.add_page_break()

        self.template_count += 1

    def generate(self, excel_file, save_path, callback=None):
        """根据Excel生成文档"""
        excel_data = pd.read_excel(excel_file, sheet_name=None)
        excel_data = {
            sheet_name: self.fill_merged_cells(df)
            for sheet_name, df in excel_data.items()
        }
        excel_data = {
            sheet_name: df.astype(str) for sheet_name, df in excel_data.items()
        }

        self.template_count = 0
        for sheet_name, df in excel_data.items():
            if sheet_name in ["test", "索引"]:
                continue
            for _, row in df.iterrows():
                self.add_table(row)
                if callback:
                    callback(f"{sheet_name} {row["测试用例名称"]}")
        self.new_docx.save(save_path)
        return self.template_count


# 调用示例
if __name__ == "__main__":
    DEBUG_MODE = False
    generator = TestDocumentGenerator("template/template.docx")

    if len(sys.argv) != 2 and not DEBUG_MODE:
        print("\033[0;31mneed excel file\033[0m")
        os.system("pause")
        sys.exit(1)

    excel_file = "dist/单元测试协同表格.xlsx" if DEBUG_MODE else sys.argv[1]
    save_file_name = os.path.splitext(excel_file)[0] + "_oa.docx"

    try:
        count = generator.generate(excel_file, save_file_name)
        print(f"\033[0;32m{save_file_name} is created with {count} test cases\033[0m")
    except Exception as e:
        print(f"\033[0;31mError: {e}\033[0m")

    os.system("pause")
