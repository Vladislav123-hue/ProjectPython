import gspread

gc = gspread.service_account(filename='credentials.json')

wks = gc.open("asd").sheet1

row = 2
coloumn = 5

for row in range(2, 11):
    grades = []
    for coloumn in range(5, 9):
        grades.append(int(wks.cell(row, coloumn).value))
    mean = sum(grades)/len(grades)
    wks.update_cell(row, 9, mean)
    grades.clear()







