from docx import Document
from string import Template
from docx.oxml.parser import parse_xml
import pandas as pd
import re

# 读取 XML 模板文件
with open("temp/table.xml", "r", encoding="utf-8") as file:
    table_template_xml = file.read()

with open("temp/tr.xml", "r", encoding="utf-8") as file:
    row_template_xml = file.read()

# 创建新文档
output_doc = Document("template/new.docx")


# 工具函数：将 "&" 替换为 "&amp;"
def escape_ampersand(text):
    return text.replace("&", "&amp;")


def parse_input_data(input_names, input_values):
    """解析输入数据"""
    name_list = re.split(r"[;,，\n]", input_names)
    parsed_names = []
    for name in name_list:
        name = name.strip()
        if not name:
            continue
        matches = re.findall(r"\*?\s*([a-zA-Z_]\w*)", name)
        name = escape_ampersand(matches[1])
        parsed_names.append(name)

    value_list = re.split(r"[;,，\n]", input_values)
    parsed_values = []
    for value in value_list:
        value = value.strip()
        if not value:
            continue
        parsed_values.append(escape_ampersand(value))
    return parsed_names, parsed_values


def parse_output_data(output_string):
    """解析输出数据"""
    output_string = str(output_string)
    output_list = re.split(r"[;,，\n]", output_string)
    parsed_outputs = [output.strip() for output in output_list if output.strip()]
    return parsed_outputs


def parse_insert_statements(statement_string):
    """
    解析插入语句
    例如: g_ps_pDetailFaultCode.I_VMSFaultCode=0x1000;
    返回: 名称和值的列表
    """
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


def process_test_case_data(test_case_data):
    """处理测试用例数据，生成输入和输出表格"""
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

    insert_names, insert_values = parse_insert_statements(test_case_data["插入语句"])
    input_names, input_values = parse_input_data(
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
            "output_names": parse_output_data(test_case_data["输出变量名称"]),
            "expected_values": parse_output_data(test_case_data["预期输出"]),
            "actual_values": parse_output_data(test_case_data["实际输出"]),
        }
    )

    # 填充输出数据
    for _, row in output_data.iterrows():
        io_data_frame.loc[
            output_index, ["output_names", "expected_values", "actual_values"]
        ] = row[["output_names", "expected_values", "actual_values"]]
        output_index += 1

    return io_data_frame


def fill_empty_cells(data_frame):
    """填充 DataFrame 中的合并单元格"""
    for column in data_frame.columns:
        data_frame[column] = data_frame[column].ffill().infer_objects(copy=False)
    return data_frame


def add_test_case_table(test_case_data):
    """根据测试用例数据生成表格并添加到文档"""
    row_xml = ""
    test_case_io_data = process_test_case_data(test_case_data)

    for _, row in test_case_io_data.iterrows():
        temp_row_template = row_template_xml
        row_data = {
            "srblm": "" if pd.isna(row["input_names"]) else row["input_names"],
            "srblqz": "" if pd.isna(row["input_values"]) else row["input_values"],
            "scblm": "" if pd.isna(row["output_names"]) else row["output_names"],
            "yqsc": "" if pd.isna(row["expected_values"]) else row["expected_values"],
            "sjsc": "" if pd.isna(row["actual_values"]) else row["actual_values"],
        }
        row_xml += Template(temp_row_template).substitute(row_data)

    # 处理数据映射
    def remove_newlines(text):
        return text.replace("\n", "") if isinstance(text, str) else text

    def prepend_newline_if_needed(text):
        return "\n" + text if isinstance(text, str) and "\n" in text else text

    def format_date(value):
        return value.date() if isinstance(value, pd.Timestamp) else value

    def extract_tester(value):
        return value.split("/")[0] if "/" in value else value

    def extract_checker(value):
        return value.split("/")[1] if "/" in value else value

    test_case_mapping = {
        "dybs": test_case_data["单元标识"],
        "sjzs": test_case_data["设计追踪"],
        "gnms": test_case_data["功能描述"],
        "qdmk": test_case_data["驱动模块"],
        "dzmk": remove_newlines(test_case_data["打桩模块"]),
        "fgljl": remove_newlines(test_case_data["覆盖率记录"]),
        "bz": test_case_data["备注"],
        "csylmc": test_case_data["测试用例名称"],
        "csylbs": test_case_data["测试用例标识"],
        "cslx": test_case_data["测试类型"],
        "hdz": test_case_data["活动桩"],
        "cssm": test_case_data["测试说明"],
        "table": row_xml,
        "sjz": extract_tester(test_case_data["测试者/校对者"]),
        "xdz": extract_checker(test_case_data["测试者/校对者"]),
        "csry": extract_tester(test_case_data["测试者/校对者"]),
        "cszxsj": format_date(test_case_data["测试时间"]),
    }
    table_xml = Template(table_template_xml).substitute(test_case_mapping)
    table_element = parse_xml(table_xml)
    paragraph = output_doc.add_paragraph(f"{test_case_data['测试用例名称']}")
    paragraph._p.addnext(table_element)
    output_doc.add_page_break()


# 读取 Excel 数据
excel_sheets = pd.read_excel("dist/信息综合软件单元测试.xlsx", sheet_name=None)
excel_sheets = {
    sheet_name: fill_empty_cells(sheet_data)
    for sheet_name, sheet_data in excel_sheets.items()
}

for sheet_name, sheet_data in excel_sheets.items():
    if sheet_name in ["test", "索引"]:
        continue
    for _, row_data in sheet_data.iterrows():
        add_test_case_table(row_data)

# 保存输出文档
output_doc.save("test.docx")
