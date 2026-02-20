
# # Working with the File with using OS and another one

# # import os
# # import shutil

# # source_file = r"C:\Users\Toli\Desktop\action"

# # with os.scandir(source_file) as  entries:
# #   for entry in entries:
# #     print(entry.name)
# # list_file =  os.listdir()
# # for pdf in list_file:
# #   if pdf
# # print(f"Here the list of your file {list_file}")

# import os
# import shutil
# source_file = r"C:\Users\Toli\Desktop\action"
# with os.scandir(source_file) as  entries:
#   for entry in entries:
#     print(entry.name)


# list_file =  os.listdir()
# for pdf in list_file:
#   if pdf
# print(f"Here the list of your file {list_file}")


# # print(os.getcwd)
# # # os.removedirs("os_create/file.py")
# # for root, dir, file in os.walk("D:\ML Project"):
# #     print("current folder ", root)
# #     print("subfolders", dir)
# #     print("current file", file)

# # folder = r"C:\Users\Toli\Desktop\ana"
# # for file in os.listdir(folder):
# #     full_path = os.path.join(folder, file)
# #     word = file.endswith(".docx")
# #     if not word:
# #         print("not have docs file move to next")
# #     else:
# #         shutil.move(full_path, r"C:\Users\Toli\Desktop\WORD")


# # Renaming the file using the simpl python code easily

# # source = r"C:\Users\Toli\Desktop\WORD"

# # for i, file in enumerate(os.listdir(source), start=1):
# #     old_file_path = os.path.join(source, file)
# #     name, ext = os.path.splitext(file)
# #     new_name = f"{str(i).zfill(3)}_Document{ext}"
# #     new_path = os.path.join(source, new_name)
# #     os.rename(old_file_path, new_path)

# # print("The file was renamed as expected")

# # working with the Excel using the openpyxl

# from openpyxl import Workbook, load_workbook
# from openpyxl.utils import get_column_letter
# from openpyxl.styles import Font
# # wb = load_workbook(r"D:\Ecxel\student_grades.xlsx")
# # # ws = wb.active
# # # ws['A2'].value = "Tolera Imiru"
# # # wb.save(r"D:\Ecxel\student_grades.xlsx")
# # # print(wb.sheetnames)
# # wb.create_sheet("Test")

# # creating the new sheet
# # wb = Workbook()
# # ws = wb.active

# # ws.append(["Name" , "Sex" , "Age", "Dept"])
# # ws.append(["Toli" , "M" ,'23', "Soft"])
# # ws.append(["Anaan" , "F" ,21, "IT"])

# # wb = load_workbook(r"D:\Ecxel\Form.xlsx")
# # ws = wb.active

# # for row in range( 1, 11):
# #     for col in range(1,5):
# #         char = get_column_letter(col)
# # print(ws[char + str(row)].value)

# # Merging the cell
# # ws.merge_cells("A1:D1")
# # ws['A1'] = "Students Information"

# # Unmerging the cell
# # ws.unmerge_cells('A1:D1')

# # ws['B1'] = 'Sex'
# # wb.save(r"D:\Ecxel\Form.xlsx")

# # you can also insert and delete the rows and cols of the data using
# # insert_cols and insert_rows and also you can delete it using the same delete on the above one


# # data = {
# #     "Toli": {
# #         "math": 79,
# #         "engl": 56,
# #         "gym": 70,
# #         "phy": 70,
# #     },
# #     "Ayyu": {
# #         "math": 70,
# #         "engl": 56,
# #         "gym": 70,
# #         "phy": 70,
# #     },
# #     "Miklo": {
# #         "math": 79,
# #         "engl": 56,
# #         "gym": 70,
# #         "phy": 70,
# #     },
# #     "Simo": {
# #         "math": 79,
# #         "engl": 56,
# #         "gym": 70,
# #         "phy": 70,
# #     },
# # }

# # wb = Workbook()
# # ws = wb.active

# # ws.title = "Grades"

# # headings = ['Name'] + list(data['Toli'].keys())
# # ws.append(headings)

# # for person in data:
# #     grades = list(data[person].values())
# #     ws.append([person] + grades)

# # # finding the average of the data
# # for col in range(2, len(data['Toli']) + 2):
# #     char = get_column_letter(col)
# #     # ws[char + "7"] = f"=SUM({char}2:{char}6/{len(data)})"
# #     ws[char + "7"] = f"=SUM({char}2:{char}6)/{len(data)}"


# # for cols in range(1, 6):
# #     ws[get_column_letter(col) + '1'].font = Font(bold=True, color='FF0099CC')
# # wb.save(r"D:\Ecxel\NewGrades.xlsx")

# # print(wb.sheetnames)


#         #working with the sheet
# wb = load_workbook(r"D:\Ecxel\Form.xlsx")
# grade= wb['Sheet']
# print(grade.title)

# # Apply filter on the data
# filtered_ws = wb.create_sheet(title = "Filtered_Data")
# header = [cell.value for cell in grade[1]]
# filtered_ws.append(header)
# print(header.index("Dept"))

# # apply the filter condition on the data 
# filter_condition = 1000

# # filtered = [row for row in data if row["math"] >= 75]
# for row_num, row in enumerate(grade.iter_rows(min_row=2, values_only=True)):

#         name = row[0]
#         sex = row[1] 
#         age = row[2]
#         dept = row[3]
#         age = int(age)

#         if sex == "M" and age > 20:
#                 print(f"{name} is male and older than 20")

#         else:
#                 print("no")

# ws2 = wb.create_sheet(" new Grade")
# ws2['A1'] = "Name"
# ws2['B1'] = "Year"
# ws2['C1'] = "Background"

# # wb.save(r"D:\Ecxel\Form.xlsx")

# for num_run , row in enumerate(grade.iter_rows(min_row = 2 , values_only=True)):
#         headers = [cell.value for cell in grade[1]] 
#         student = dict(zip(headers, row))
#         # print(student)

# print(grade.max_column)



# from openpyxl.styles import Font, PatternFill, Alignment
#         # working with the Dataframe 
# # which was used to write to the excel the processed data after it was done using the pandas and another one 

# from openpyxl import Workbook, load_workbook
# from openpyxl.utils.dataframe import dataframe_to_rows
# import pandas as pd

# df = pd.read_excel(r"D:\Ecxel\Form.xlsx")
# print(df.columns)

#         # creating the Reports based on the pivot Table 
# pivot_table = pd.pivot_table(
# df, values='Age', index="Students Information", columns="Dept", aggfunc='sum', fill_value=0)
# print(pivot_table)

# wb = load_workbook((r"D:\Ecxel\Form.xlsx"))
# wb.create_sheet('pivot_table')
# stinfo = wb['pivot_table']
# for row in dataframe_to_rows(pivot_table , index = True , header=True  ):
#         stinfo.append(row)
# # wb.save(r"D:\Ecxel\Form_pt.xlsx")


# for sheet in wb.sheetnames:
#         if sheet.startswith("p"):
#                 for cell in wb[sheet][1]:
#                         cell.font = Font(color="FF0000", bold=True, size=12)
#                         cell.fill = PatternFill(
#                         start_color='FFFF00',
#                         end_color='FFFF00',
#                         fill_type='solid'
#                         )
#                         cell.alignment = Alignment(horizontal="center")

# wb.save(r"D:\Ecxel\update.xlsx")


                # working with xlwings for file 
# import xlwings as xw 
# wb = xw.Book()
# # print(wb.sheets)

# # Adding the worksheets 
# wb.sheets.add(name = "Test" , before = "Sheet1")
# print(wb.sheets.count)
# ws = wb.sheets['Test1']

# path = r"D:\Ecxel\Form.xlsx"
# obj = xw.books.open(path)
# # print (obj.name)
# import os 
# print(os.getcwd())


    # Working with the CSV and handle the case related to this one 


import csv 

filename =  r"D:\Ecxel\Test.csv"

with open(filename , "w" , newline = "") as f:
  wrt = csv.writer (f,delimiter= ' ')
  wrt.writerow(['first-column', 'second-column'])
