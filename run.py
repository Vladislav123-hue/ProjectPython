import gspread
from gspread_formatting import *

gc = gspread.service_account(filename='credentials.json')

wks = gc.open("asd").sheet1

row = 2
coloumn = 5
grades = []
class Class:
    def __init__(self, name, grade):
        self.name = name  # имя предмета
        self.grade = grade   # оценка

classes = []

red_fill = CellFormat(backgroundColor=Color(1, 0.5, 0.5))
yellow_fill = CellFormat(backgroundColor=Color(1, 1, 0))        
green_fill = CellFormat(backgroundColor=Color(0.5, 1, 0.5))  

for row in range(2, 11):
    grades = []
    for coloumn in range(5, 9):
        classes.append(Class(wks.cell(1, coloumn).value, int(wks.cell(row, coloumn).value)))
        print(F"Class:{wks.cell(1, coloumn).value}, grade:{wks.cell(row, coloumn).value}")
        grades.append(int(wks.cell(row, coloumn).value))
        mean = round(sum(grades)/len(grades))
    if mean == 1 or mean == 2:
        cell_address = f'I{row}'
        format_cell_range(wks, cell_address, red_fill)
    if mean == 3:
        cell_address = f'I{row}'
        format_cell_range(wks, cell_address, yellow_fill)
    if mean == 4 or mean == 5: 
        cell_address = f'I{row}'
        format_cell_range(wks, cell_address, green_fill)

    wks.update_cell(row, 9, mean)
    max_Class = max(classes, key=lambda obj: obj.grade)
    wks.update_cell(row, 10, max_Class.name)
    min_Class = min(classes, key=lambda obj: obj.grade)
    wks.update_cell(row, 11, min_Class.name)
    grades.clear()
    classes.clear()







