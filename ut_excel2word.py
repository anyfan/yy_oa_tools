from docx import Document
from string import Template
from docx.oxml.parser import parse_xml
import pandas as pd
import re
import sys
import os


def resource_path(relative_path):
    """获取资源文件的路径，兼容pyinstaller打包"""
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class UnitTestTestCaseDocumentGenerator:

    def __init__(
        self,
        template_doc_path=resource_path("template/new.docx"),
        table_template_path=resource_path("template/ut_table.xml"),
        row_template_path=resource_path("template/ut_table_tr.xml"),
    ):
        # 初始化模板文件路径
        self.template_doc_path = template_doc_path
        self.table_template_path = table_template_path
        self.row_template_path = row_template_path

        with open(self.table_template_path, "r", encoding="utf-8") as file:
            self.table_template_xml = file.read()

        with open(self.row_template_path, "r", encoding="utf-8") as file:
            self.row_template_xml = file.read()

    @staticmethod
    def escape_ampersand(text):
        """将 & 替换为 &amp;"""
        return text.replace("&", "&amp;")

    def parse_input_data(self, input_names, input_values):
        """解析输入数据"""
        name_list = re.split(r"[;,，\n]", input_names)
        parsed_names = []
        for name in name_list:
            name = name.strip()
            if not name:
                continue
            matches = re.findall(r"\*?\s*([a-zA-Z_]\w*)", name)
            if matches:
                # 取最后一个匹配
                name = self.escape_ampersand(matches[-1])
                parsed_names.append(name)

        value_list = re.split(r"[;,，\n]", input_values)
        parsed_values = []
        for value in value_list:
            value = value.strip()
            if not value:
                continue
            parsed_values.append(self.escape_ampersand(value))
        return parsed_names, parsed_values

    @staticmethod
    def parse_output_data(output_string):
        """解析输出数据"""
        output_string = str(output_string)
        output_list = re.split(r"[;,，\n]", output_string)
        parsed_outputs = [output.strip() for output in output_list if output.strip()]
        return parsed_outputs

    @staticmethod
    def parse_insert_statements(statement_string):
        """解析插入语句"""
        statement_list = re.split(r"[;,，\n]", statement_string)
        parsed_names, parsed_values = [], []
        for statement in statement_list:
            statement = statement.strip()
            if not statement:
                continue
            matches = re.findall(r"(\w+\.\w+)\s*=\s*(\w+)", statement)
            if matches:
                parsed_names.append(matches[0][0])
                parsed_values.append(matches[0][1])
        return parsed_names, parsed_values

    @staticmethod
    def fill_empty_cells(data_frame):
        """填充 DataFrame 中的合并单元格"""
        for column in data_frame.columns:
            data_frame[column] = data_frame[column].ffill().infer_objects(copy=False)
        return data_frame

    def process_test_case_data(self, test_case_data):
        """处理单个测试用例数据"""
        input_index = 0
        output_index = 0
        io_data_frame = pd.DataFrame(
            columns=[
                "input_names",
                "input_values",
                "output_names",
                "expected_values",
                "actual_values",
            ]
        )

        insert_names, insert_values = self.parse_insert_statements(
            test_case_data["插入语句"]
        )
        input_names, input_values = self.parse_input_data(
            test_case_data["输入变量名称"], test_case_data["取值"]
        )
        input_data = pd.DataFrame(
            {
                "input_names": insert_names + input_names,
                "input_values": insert_values + input_values,
            }
        )

        # 填充输入数据
        for _, row in input_data.iterrows():
            io_data_frame.loc[input_index, ["input_names", "input_values"]] = row[
                ["input_names", "input_values"]
            ]
            input_index += 1

        output_data = pd.DataFrame(
            {
                "output_names": self.parse_output_data(test_case_data["输出变量名称"]),
                "expected_values": self.parse_output_data(test_case_data["预期输出"]),
                "actual_values": self.parse_output_data(test_case_data["实际输出"]),
            }
        )

        # 填充输出数据
        for _, row in output_data.iterrows():
            io_data_frame.loc[
                output_index, ["output_names", "expected_values", "actual_values"]
            ] = row[["output_names", "expected_values", "actual_values"]]
            output_index += 1

        return io_data_frame

    def add_test_case_table(self, test_case_data, mode="report"):
        """将测试用例数据添加为表格"""

        nan2blank = lambda x: "" if pd.isna(x) else x

        def format_date(value):
            return value.date() if isinstance(value, pd.Timestamp) else value

        def extract_tester(value):
            return value.split("/")[0] if "/" in value else value

        def extract_checker(value):
            return value.split("/")[1] if "/" in value else value

        test_case_io_data = self.process_test_case_data(test_case_data)
        temp_row_template = self.row_template_xml
        test_case_mapping = {}
        match mode:
            case "report":
                row_xml = ""
                for _, row in test_case_io_data.iterrows():
                    row_data = {
                        "srblm": nan2blank(row["input_names"]),
                        "srblqz": nan2blank(row["input_values"]),
                        "scblm": nan2blank(row["output_names"]),
                        "yqsc": nan2blank(row["expected_values"]),
                        "sjsc": nan2blank(row["actual_values"]),
                    }
                    row_xml += Template(temp_row_template).substitute(row_data)
                test_case_mapping = {
                    "dybs": test_case_data["单元标识"],
                    "sjzs": test_case_data["设计追踪"],
                    "gnms": test_case_data["功能描述"],
                    "qdmk": test_case_data["驱动模块"],
                    "dzmk": test_case_data["打桩模块"].replace("\n", ""),
                    "fgljl": test_case_data["覆盖率记录"]
                    .replace("\n", "")
                    .replace("。", ""),
                    "bz": test_case_data["备注"],
                    "csylmc": test_case_data["测试用例名称"],
                    "csylbs": test_case_data["测试用例标识"],
                    "cslx": test_case_data["测试类型"],
                    "hdz": test_case_data["活动桩"],
                    "cssm": test_case_data["测试说明"],
                    "table": row_xml,
                    "tgzz": "实际测试结果与预期结果一致。",  # 通过准则
                    "ycxxms": "无",  # 异常现象描述
                    "rjwtbh": "无",  # 软件问题编号
                    "zxqk_zx": "☑",  # 执行情况—执行
                    "zxqk_wzx": "□",  # 执行情况—未执行
                    "zxjg_tg": "☑",  # 执行结果—通过
                    "zxjg_wtg": "□",  # 执行结果—未通过
                    "sjz": extract_tester(test_case_data["测试者/校对者"]),
                    "xdz": extract_checker(test_case_data["测试者/校对者"]),
                    "csry": extract_tester(test_case_data["测试者/校对者"]),
                    "cszxsj": format_date(test_case_data["测试时间"]),
                }

            case "instructions":
                row_xml = ""
                for _, row in test_case_io_data.iterrows():
                    temp_row_template = self.row_template_xml
                    row_data = {
                        "srblm": nan2blank(row["input_names"]),
                        "srblqz": nan2blank(row["input_values"]),
                        "scblm": nan2blank(row["output_names"]),
                        "yqsc": nan2blank(row["expected_values"]),
                        "sjsc": "",
                    }
                    row_xml += Template(temp_row_template).substitute(row_data)
                test_case_mapping = {
                    "dybs": test_case_data["单元标识"],
                    "sjzs": test_case_data["设计追踪"],
                    "gnms": test_case_data["功能描述"],
                    "qdmk": "系统自动生成（需要驱动模块时，写具体模块名）",
                    "dzmk": test_case_data["打桩模块"].replace("\n", ""),
                    "fgljl": "",
                    "bz": test_case_data["备注"],
                    "csylmc": test_case_data["测试用例名称"],
                    "csylbs": test_case_data["测试用例标识"],
                    "cslx": test_case_data["测试类型"],
                    "hdz": test_case_data["活动桩"],
                    "cssm": test_case_data["测试说明"],
                    "table": row_xml,
                    "tgzz": "实际测试结果与预期结果一致。",
                    "ycxxms": "",
                    "rjwtbh": "",
                    "zxqk_zx": "□",
                    "zxqk_wzx": "□",
                    "zxjg_tg": "□",
                    "zxjg_wtg": "□",
                    "sjz": extract_tester(test_case_data["测试者/校对者"]),
                    "xdz": extract_checker(test_case_data["测试者/校对者"]),
                    "csry": "",
                    "cszxsj": "",
                }

        table_xml = Template(self.table_template_xml).substitute(test_case_mapping)
        return table_xml

    def generate_document(self, excel_path, output_path, mode="report", callback=None):
        """从 Excel 文件生成 Word 文档"""
        output_doc = Document(self.template_doc_path)
        excel_sheets = pd.read_excel(excel_path, sheet_name=None)
        excel_sheets = {
            sheet_name: self.fill_empty_cells(sheet_data)
            for sheet_name, sheet_data in excel_sheets.items()
        }
        case_count = 0
        for sheet_name, sheet_data in excel_sheets.items():
            if sheet_name in ["test", "索引"]:
                continue
            for _, row_data in sheet_data.iterrows():
                if callback:
                    callback(f"{sheet_name} {row_data["测试用例名称"]}")
                table_xml = self.add_test_case_table(row_data, mode=mode)
                table_element = parse_xml(table_xml)
                paragraph = output_doc.add_paragraph(f"{row_data['测试用例名称']}")
                paragraph._p.addnext(table_element)
                output_doc.add_page_break()
                case_count += 1

        output_doc.save(output_path)
        return case_count


if __name__ == "__main__":
    # 使用示例
    generator = UnitTestTestCaseDocumentGenerator()
    # generator.generate_document(
    #     "dist/信息综合软件单元测试.xlsx",
    #     "dist/test.docx",
    #     mode="report",
    #     callback=lambda x: print(x),
    # )
    generator.generate_document(
        "dist/信息综合软件单元测试.xlsx",
        "dist/test.docx",
        mode="instructions",
        callback=lambda x: print(x),
    )
