import pandas as pd
from haystack import Pipeline
from haystack.nodes import BaseComponent

class ExcelReader(BaseComponent):
    def run(self, file_path: str):
        data = {}
        xl = pd.ExcelFile(file_path)
        for sheet_name in xl.sheet_names:
            sheet = xl.parse(sheet_name)
            data[sheet_name] = self.detect_headers(sheet)
        return {"data": data}, "output_1"

    def detect_headers(self, sheet):
        headers = {}
        current_header = None
        for i, row in sheet.iterrows():
            if self.is_header(row):
                current_header = list(row)
                headers[i] = current_header
            elif current_header:
                # Collect data under the current header
                headers.setdefault(current_header, []).append(list(row))
        return headers

    def is_header(self, row):
        return row.notnull().all() and row.apply(lambda x: isinstance(x, str)).all()

excel_reader = ExcelReader()

pipeline = Pipeline()
pipeline.add_node(component=excel_reader, name="ExcelReader", inputs=["File"])

result = pipeline.run(file_path="path_to_your_excel_file.xlsx")
for sheet, headers in result['data'].items():
    print(f"Sheet: {sheet}")
    for header_row, header in headers.items():
        print(f"Header found at row {header_row}: {header}")
        print(f"Data under this header: {headers[header]}")
