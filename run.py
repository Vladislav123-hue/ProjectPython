import gspread
from gspread_formatting import format_cell_range, Color, CellFormat

# Authorization and connection to the spreadsheet
try:
    gc = gspread.service_account(filename='credentials.json')
    wks = gc.open("asd").sheet1
except Exception as e:
    print(f"[ERROR] Failed to connect to Google Sheets: {e}")
    exit(1)

# Color styles
red_fill = CellFormat(backgroundColor=Color(1, 0.5, 0.5))
yellow_fill = CellFormat(backgroundColor=Color(1, 1, 0))        
green_fill = CellFormat(backgroundColor=Color(0.5, 1, 0.5))

# Class to store subject and grade
class SubjectGrade:
    def __init__(self, name, grade):
        self.name = name
        self.grade = grade

# Retrieve data from the range
try:
    all_values = wks.get('E1:H10')
except Exception as e:
    print(f"[ERROR] Failed to retrieve data from sheet: {e}")
    exit(1)

if not all_values or len(all_values) < 2:
    print("[WARNING] Not enough data (less than two rows).")
    exit(1)

header = all_values[0]
data_rows = all_values[1:]

# Process each row
for i, row_values in enumerate(data_rows, start=2):
    classes = []
    grades = []

    for col_index, grade_str in enumerate(row_values):
        try:
            grade = int(grade_str)
        except (ValueError, TypeError):
            grade = 0  # If empty or not a number
        subject_name = header[col_index] if col_index < len(header) else f"Column {col_index+1}"
        classes.append(SubjectGrade(subject_name, grade))
        grades.append(grade)

    if not grades:
        print(f"[WARNING] Row {i} is empty or contains non-numeric values. Skipping.")
        continue

    mean = round(sum(grades) / len(grades))
    cell_address = f'I{i}'

    # Update average grade
    try:
        wks.update_cell(i, 9, mean)
        if mean in (1, 2):
            format_cell_range(wks, cell_address, red_fill)
        elif mean == 3:
            format_cell_range(wks, cell_address, yellow_fill)
        else:
            format_cell_range(wks, cell_address, green_fill)
    except Exception as e:
        print(f"[ERROR] Failed to update or format cell I{i}: {e}")

    # Best subjects
    try:
        max_grade = max(obj.grade for obj in classes)
        max_classes = [obj for obj in classes if obj.grade == max_grade]
        top_classes_str = ', '.join(obj.name for obj in max_classes)
        wks.update_cell(i, 10, top_classes_str)
    except Exception as e:
        print(f"[ERROR] Failed to write top subjects in row {i}: {e}")

    # Worst subjects
    try:
        min_grade = min(obj.grade for obj in classes)
        min_classes = [obj for obj in classes if obj.grade == min_grade]
        worst_classes_str = ', '.join(obj.name for obj in min_classes)
        wks.update_cell(i, 11, worst_classes_str)
    except Exception as e:
        print(f"[ERROR] Failed to write worst subjects in row {i}: {e}")




