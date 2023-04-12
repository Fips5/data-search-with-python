import os
from openpyxl import load_workbook

dir_path = "./files"

xlsx_files = [f for f in os.listdir(dir_path) if f.endswith(".xlsx")]

row_num = int(input("Enter the row number: "))
col_letter = input("Enter the column letter: ")

for filename in xlsx_files:
    filepath = os.path.join(dir_path, filename)
    workbook = load_workbook(filepath)
    sheet = workbook.active
    cell = sheet["{}{}".format(col_letter, row_num)]
    print("{}: {}".format(filename, cell.value))
