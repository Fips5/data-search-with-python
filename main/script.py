import os
from openpyxl import load_workbook

while True:
    dir_path = "./files"

    xlsx_files = [f for f in os.listdir(dir_path) if f.endswith(".xlsx")]

    row_num = int(input("Enter the row number: "))
    col_letter = input("Enter the column letter: ")

    cell_values = {}

    for filename in xlsx_files:
        filepath = os.path.join(dir_path, filename)
        workbook = load_workbook(filepath)
        sheet = workbook.active
        cell = sheet["{}{}".format(col_letter, row_num)]
        cell_values[filename] = cell.value

    all_values_same = all(value == list(cell_values.values())[0] for value in cell_values.values())

    if all_values_same:
        print("All files have the same value: {}".format(list(cell_values.values())[0]))
    else:

        for filename, value in cell_values.items():
            if value != list(cell_values.values())[0]:
                print("{}: {}".format(filename, value))
        
        normal_value = list(cell_values.values())[0]
        print("Normal value from other files: {}".format(normal_value))

    user_choice = input("Do you want to exit? y/n: ")
    if user_choice.lower() == "y":
        break
    elif user_choice.lower() == "n":
        continue
    else:
        print("Invalid choice. Please enter y or n.")

print("Program has ended.")
