import pandas as pd
from lxml import html

class HtmlTableParser:
    def __init__(self, file_path):
        self.file_path = file_path
        self.tree = None
        self.g_data = pd.DataFrame(columns=['Code', 'File', 'Function', 'Line', 'Violation', 'Standard', 'Parameter'])
        self.function_href = ""
        self.function_name = ""
        self.function_file = ""
        self.g_state = 0  # 初始状态
        self._load_html()

    def _load_html(self):
        with open(self.file_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        self.tree = html.fromstring(html_content)

    def _case1_table2df(self, table):
        df = pd.read_html(html.tostring(table))[0]
        df['File'] = self.function_file
        df['Function'] = self.function_name

        if 'Standard' in df.columns:
            df['Standard'] = df['Standard'].str.replace('GJB_8114 ', '', regex=False)

        self.g_data = pd.concat([self.g_data, df], ignore_index=True)

    def _case2_table2df(self, table):
        df = pd.read_html(html.tostring(table))[0]
        df[['File', 'Line']] = df['File: Src Line'].str.split(': ', expand=True)

        df['Line'] = pd.to_numeric(df['Line'], errors='coerce')

        if 'Standard' in df.columns:
            df['Standard'] = df['Standard'].str.replace('GJB_8114 ', '', regex=False)
        df.drop(columns=['File: Src Line'], inplace=True)

        self.g_data = pd.concat([self.g_data, df], ignore_index=True)

    def _table_case_0(self, table):
        th_elements = table.xpath('.//th[@bgcolor = "#FF0000"]')
        if th_elements and (a_element := th_elements[0].xpath('.//a')):
            self.function_href = a_element[0].get('href')
            self.function_name = a_element[0].text.strip()
            self.function_file = self.function_href.split('File=')[1].split('&')[0].rsplit('\\', 1)[-1]
            self.g_state = 1
        elif th_elements and ("Globals / code outside procedures - FAIL" in th_elements[0].text_content()):
            self.function_href = ""
            self.function_name = ""
            self.function_file = ""
            self.g_state = 2

    def _table_case_1(self, table):
        match table.get('bgcolor'):
            case "#FFFFF2":
                if "Key to Terms   |  Procedure Table" in table.text_content():
                    self.g_state = 0
            case "#ECE2E2":
                self._case1_table2df(table)

    def _table_case_2(self, table):
        match table.get('bgcolor'):
            case "#FFFFF2":
                if "Key to Terms   |  Procedure Table" in table.text_content():
                    self.g_state = 0
            case "#ECE2E2":
                self._case2_table2df(table)

    def parse_tables(self):
        tables = self.tree.xpath('//table')
        for table in tables:
            match self.g_state:
                case 0:
                    self._table_case_0(table)
                case 1:
                    self._table_case_1(table)
                case 2:
                    self._table_case_2(table)

    def save_to_excel(self, output_path):
        self.g_data.to_excel(output_path, index=False)

# 使用方法
if __name__ == "__main__":
    parser = HtmlTableParser('1.htm')
    parser.parse_tables()
    parser.save_to_excel('output.xlsx')
