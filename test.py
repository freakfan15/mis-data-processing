from openpyxl import load_workbook
from openpyxl.drawing.image import Image

# Load the original Excel file
input_file_path = 'output.xlsx'
output_file_path = './copied_output_with_formatting.xlsx'

workbook = load_workbook(input_file_path)
workbook.save(output_file_path)