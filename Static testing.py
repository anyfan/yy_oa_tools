import email
# from bs4 import BeautifulSoup

import pandas as pd


filename = 'dist/LDRA Testbed Code Review Report.mht'
# with open(filename, 'rb') as f:
#     mht_content = f.read()
mht_content = open(filename,'r')
mht_message = email.message_from_file(mht_content)
mht_content.close()
html_content = ''

for part in mht_message.walk():
    content_type = part.get_content_type()
    if content_type == 'text/html':
        html_content = part.get_payload(decode=True)
        break

tables = pd.read_html(html_content)


for i in range(len(tables)):
    columns_name = tables[i].columns.tolist() #索引名称
    if (len(columns_name) == 1 and ' - FAIL' in columns_name[0]): # 名称包含 ‘- FAIL’
        index = i + 1
        while set(tables[index].columns.tolist()).issuperset(set(['Code', 'Violation', 'Standard'])):
            index = index + 1
        # 合并列表
        
        
        print(columns_name)
        print(index - i-1)
            



# for table in tables:
#     print(table.columns)
    # 保留  table.columns[0] 里包含- FAIL的表
    # columns_name = table.columns.tolist()
    # if (len(columns_name) == 1 and 'FAIL' in columns_name[0]) or (columns_name  == ['Code', 'Line', 'Violation', 'Standard']):
    #     print(table)

    # print(table.columns.to_list()[0])
    
    # if table.columns.tolist() == ['Code', 'Line', 'Violation', 'Standard']:
    #     # print(table)
    #     print(table.columns)
    #     pass
    

# print(list(tables[18].columns))
# 保存到本地
# tables[18].to_excel('dist/sample.xlsx')