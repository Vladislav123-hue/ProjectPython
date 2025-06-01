import gspread

gc = gspread.service_account(filename='credentials.json')

wks = gc.open("asd").sheet1

grades = []

row = 2
coloumn = 5

for coloumn in range(5, 9):
    grades.append(int(wks.cell(row, coloumn).value))
 
mean = sum(grades)/len(grades)
wks.update_cell(2, 9, mean)







