import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

input_file_path = 'input.xlsx'
output_file_path = 'output.xlsx'
generated_report_path = './generated_report_preserving_format.xlsx'

input_df = pd.read_excel(input_file_path)

workbook = openpyxl.load_workbook(output_file_path)
sheet = workbook.active

def set_merged_cell_value(sheet, cell, value):
    for merged_range in sheet.merged_cells.ranges:
        if cell in merged_range:
            top_left_cell = merged_range.start_cell.coordinate
            sheet[top_left_cell] = value
            return
    sheet[cell] = value

# Process only the first row
if not input_df.empty:
    input_first_row = input_df.iloc[4]
    print(first_row["Customer Name"])
    set_merged_cell_value(sheet, 'B4', input_first_row["Customer Name"])
    set_merged_cell_value(sheet, 'K5', input_first_row["Loan Amount"])


workbook.save(generated_report_path)