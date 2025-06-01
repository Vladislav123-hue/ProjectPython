import gspread
from gspread_formatting import *

gc = gspread.service_account(filename='credentials.json')
wks = gc.open("asd").sheet1

red_fill = CellFormat(backgroundColor=Color(1, 0.5, 0.5))
yellow_fill = CellFormat(backgroundColor=Color(1, 1, 0))        
green_fill = CellFormat(backgroundColor=Color(0.5, 1, 0.5))

class SubjectGrade:
    def __init__(self, name, grade):
        self.name = name
        self.grade = grade


all_values = wks.get('E1:H10')  

header = all_values[0]  
data_rows = all_values[1:]  

for i, row_values in enumerate(data_rows, start=2):  
    classes = []
    grades = []

    for col_index, grade_str in enumerate(row_values):
        try:
            grade = int(grade_str)
        except ValueError:
            grade = 0  

        subject_name = header[col_index]
        classes.append(SubjectGrade(subject_name, grade))
        grades.append(grade)

    if len(grades) == 0:
        continue

    mean = round(sum(grades) / len(grades))

    cell_address = f'I{i}'
    if mean in (1, 2):
        format_cell_range(wks, cell_address, red_fill)
    elif mean == 3:
        format_cell_range(wks, cell_address, yellow_fill)
    else:
        format_cell_range(wks, cell_address, green_fill)

    wks.update_cell(i, 9, mean)

    max_grade = max(obj.grade for obj in classes)
    max_classes = [obj for obj in classes if obj.grade == max_grade]
    top_classes_str = ', '.join(obj.name for obj in max_classes)
    wks.update_cell(i, 10, top_classes_str)

    min_grade = min(obj.grade for obj in classes)
    min_classes = [obj for obj in classes if obj.grade == min_grade]
    worst_classes_str = ', '.join(obj.name for obj in min_classes)
    wks.update_cell(i, 11, worst_classes_str)





