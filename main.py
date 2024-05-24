# coding=utf-8

import sys
import os
import pandas as pd
from copy import deepcopy
from docx import Document
from string import Template

# 资源文件打包到exe时，生成资源文件的访问路径
def resource_path(relative_path):
    if getattr(sys, 'frozen', False): #是否Bundle Resource
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def add_table(docx_obj,mb_obj,data_obj): 
    
    data = {
        'dybs'  : data_obj['单元标识'],
        'sjzs'  : data_obj['设计追踪'],
        'gnms'  : data_obj['功能描述'],
        'qdmk'  : data_obj['驱动模块'],
        'dzmk'  : data_obj['打桩模块'],
        'fgljl' : data_obj['覆盖率记录'],
        'bz'    : data_obj['备注'],
        'csylmc': data_obj['测试用例名称'],
        'csylbs': data_obj['测试用例标识'],
        'cslx'  : data_obj['测试类型'],
        'hdz'   : data_obj['活动桩'],
        'cssm'  : data_obj['测试说明'],
        'srblm' : data_obj['输入变量名称'],
        'srblqz': data_obj['取值'],
        'scblm' : data_obj['输出变量名称'],
        'yqsc'  : data_obj['预期输出'],
        'sjsc'  : data_obj['实际输出'],
        'sjz'   : data_obj['测试者/校对者'],
        'xdz'   : data_obj['测试者/校对者'],
        'csry'  : data_obj['测试者/校对者'],
        'cszxsj': data_obj['测试时间'],
    }
    
    new_table = deepcopy(mb_obj.tables[0])
    
    for row in new_table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run_template = Template(run.text)
                    run.text = run_template.substitute(data)
                        
    paragraph = docx_obj.add_paragraph(f"{data_obj['测试用例名称']}")
    paragraph._p.addnext(new_table._element)
    docx_obj.add_page_break()

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('\033[0;31mneed excel file\033[0m')
        os.system("pause")
        sys.exit(1)
    excel_file = sys.argv[1]
    # excel_file = 'test.xlsx'
    excel_datas = pd.read_excel(excel_file,sheet_name=None)

    template_docx = Document(resource_path("mb.docx"))
    new_docx = Document(resource_path("new.docx"))
    
    for sheet_name, df in excel_datas.items():
        if sheet_name in ['test', '索引']:
            continue
        for index, row in df.iterrows():
            print(sheet_name, row['测试用例名称'])
            add_table(new_docx,template_docx,row)

    new_docx.save("test.docx")
    
    print("\033[0;32mtest.docx is created\033[0m")

    os.system("pause")

# ☑☐□