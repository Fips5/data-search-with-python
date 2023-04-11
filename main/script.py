from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl import Workbook, load_workbook

row_input = input("Enter row number: ")

col_input = input("Enter column letter: ")

col_index = column_index_from_string(col_input)

cell_input = f"{get_column_letter(col_index)}{row_input}"

wb = load_workbook("test.xlsx")
ws = wb.active

cell_value = ws[cell_input].value

print(f"Cell {cell_input}: {cell_value}")
