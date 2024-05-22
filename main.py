import sys
import pandas as pd
from docx import Document

def add_table(docx_obj,data_obj): 
    
    docx_obj.add_paragraph(f"{data_obj['测试用例名称']}")
    
    t = docx_obj.add_table(rows=19, cols=5, style="Table Grid")
    # 第一行单元格
    t.cell(0, 0).text = f"单元标识：{data_obj['单元标识']}"
    t.cell(1, 0).text = f"设计追溯：{data_obj['设计追踪']}"
    t.cell(2, 0).text = f"功能描述：{data_obj['功能描述']}"
    t.cell(3, 0).text = f"驱动模块：{data_obj['驱动模块']}"
    t.cell(4, 0).text = f"打桩模块：{data_obj['打桩模块']}"
    t.cell(5, 0).text = f"覆盖率记录：{data_obj['覆盖率记录']}"
    t.cell(6, 0).text = f"备注：{data_obj['覆盖率记录']}"
    t.cell(0, 0).merge(t.cell(6, 4))

    # 第二行单元格
    t.cell(7, 0).text = f"测试用例名称"
    t.cell(7, 1).text = f"{data_obj['测试用例名称']}"
    t.cell(7, 2).text = f"测试用例标识"
    t.cell(7, 3).text = f"{data_obj['测试用例标识']}"
    t.cell(7, 3).merge(t.cell(7, 4))

    # 第三四五行单元格
    t.cell(8, 0).text = f"测试类型：{data_obj['测试类型']}"
    t.cell(9, 0).text = f"活动桩：{data_obj['活动桩']}"
    t.cell(10, 0).text = f"测试说明：{data_obj['测试说明']}"
    t.cell(8, 0).merge(t.cell(8, 4))
    t.cell(9, 0).merge(t.cell(9, 4))
    t.cell(10, 0).merge(t.cell(10, 4))

    t.cell(11, 0).text = f"输入"
    t.cell(11, 2).text = f"输出"
    t.cell(11, 0).merge(t.cell(11, 1)) 
    t.cell(11, 2).merge(t.cell(11, 4))

    t.cell(12, 0).text = f"变量名"
    t.cell(12, 1).text = f"取值"
    t.cell(12, 2).text = f"变量名"
    t.cell(12, 3).text = f"预期输出"
    t.cell(12, 4).text = f"实际输出"

    t.cell(13, 0).text = f"{data_obj['输入变量名称']}"
    t.cell(13, 1).text = f"{data_obj['取值']}"
    t.cell(13, 2).text = f"{data_obj['输出变量名称']}"
    t.cell(13, 3).text = f"{data_obj['预期输出']}"
    t.cell(13, 4).text = f"{data_obj['实际输出']}"

    t.cell(14, 0).text = f"通 过 准 则"
    t.cell(14, 1).text = f"实际测试结果与预期结果一致。"
    t.cell(14, 1).merge(t.cell(14, 4))

    t.cell(15, 0).text = f"异常现象描述"
    t.cell(15, 1).text = f"无"
    t.cell(15, 2).text = f"软件问题单编号"
    t.cell(15, 3).text = f"无"
    t.cell(15, 3).merge(t.cell(15, 4))

    t.cell(16, 0).text = f"设计者"
    t.cell(16, 1).text = f"{data_obj['测试者/校对者']}"
    t.cell(16, 2).text = f"校对者"
    t.cell(16, 3).text = f"{data_obj['测试者/校对者']}"
    t.cell(16, 3).merge(t.cell(16, 4))

    t.cell(17, 0).text = f"执 行 情 况"
    t.cell(17, 1).text = f"□√ 执行  □ 未执行"
    t.cell(17, 2).text = f"执 行 结 果"
    t.cell(17, 3).text = f"□√ 通过    □ 未通过"
    t.cell(17, 3).merge(t.cell(17, 4))

    t.cell(18, 0).text = f"测 试 人 员"
    t.cell(18, 1).text = f"{data_obj['测试者/校对者']}"
    t.cell(18, 2).text = f"测试执行时间"
    t.cell(18, 3).text = f"2023-05-15"
    t.cell(18, 3).merge(t.cell(18, 4))
    
    docx_obj.add_page_break()

    # t.add_column(t.cell(18, 3).width)

    # # 获取第2行
    # rows = t.rows[1]
    # for cell in rows.cells:
    #     cell.text = f"第2行第..列"
    # # 获取第3列
    # cols = t.columns[2]
    # for cell in cols.cells:
    #     cell.text = f"第3列第..行"
    # # 增加行、增加列
    # for cell in t.add_row().cells:
    #     cell.text = "新增行"
    # t.add_column(cell.width)

if __name__ == '__main__':
    # if len(sys.argv) != 2:
    #     print('Usage: %s <excel file>' % sys.argv[0])
    #     sys.exit(1)
    # excel_file = sys.argv[1]
    excel_file = 'test.xlsx'
    excel_datas = pd.read_excel(excel_file,sheet_name=None)

    d = Document()
    for sheet_name, df in excel_datas.items():
        if sheet_name in ['test', '索引']:
            continue
        for index, row in df.iterrows():
            print(sheet_name, row['测试用例名称'])
            add_table(d,row)

    d.save("test.docx")

# □
# √