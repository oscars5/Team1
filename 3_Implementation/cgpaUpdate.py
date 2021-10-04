# ~~~~~~~~~~~~~~~~~~~~~~~PACKAGE IMPORTING~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
import openpyxl as xl
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment, Color, PatternFill
from copy import copy

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~LOADING SHEETS~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

baseData = "Base_Data - Copy.xlsx"  # accessing base data sheet
wb1 = xl.load_workbook(baseData)
workSheet1 = wb1.active

keyData = "key.xlsx"                # accessing key sheet
wb2 = xl.load_workbook(keyData)
workSheet2 = wb2.worksheets[0]

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~NEW COLUMN STRUCTURE CREATION~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

add_border = Border(left=Side(style='thin'),    # setting border
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin'))

workSheet1.insert_cols(6)        # insert a column

mr = workSheet1.max_row         # max row of base data
mc = workSheet1.max_column      # max column of base data

fill_color = PatternFill(patternType='solid', fgColor='FFFF00')     # choosing color
new_cell = workSheet1.cell(row=1, column=6, value="CGPA")
old_cell = workSheet1.cell(row=1, column=5)
new_cell.border = add_border
new_cell.alignment = Alignment(horizontal='center', vertical='center')
new_cell.font = copy(old_cell.font)
new_cell.fill = copy(old_cell.fill)
print(mr)
print(mc)

# ~~~~~~~~~~~~~~~~~~~~~~~CALCULATING CGPA UPDATING TO THE SAME SHEET~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

for i in range(2, mr + 1):
    workSheet1.cell(row=i, column=6).border = add_border
    workSheet1.cell(row=i, column=6).alignment = Alignment(horizontal='center')
    workSheet1.cell(row=i, column=6).fill = fill_color
    x = 0
    total_cg_score = 0
    total_credit = 0
    emp_cgpa = 0
    for data_col in range(7, mc + 1):
        credit = workSheet2.cell(row=2, column=data_col-4+x)
        total_credit = total_credit + credit.value
        x = x + 1

        score = workSheet1.cell(row=i, column=data_col)
        total_cg_score = total_cg_score + ((score.value/10) * credit.value)

        cgpa = total_cg_score/total_credit
        workSheet1.cell(row=i, column=6).value = round(cgpa, 2)
        # z1 = z1 + 2

print(credit.value)
print(score.value)
print(total_cg_score)

# ~~~~~~~~~~~~~~~~~~~~~~~~~~~SAVING THE SHEET~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
wb1.save(str(baseData))
