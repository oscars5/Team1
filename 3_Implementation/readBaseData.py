import openpyxl as xl
# import pandas as pd
# import numpy as np

# opening the source excel file
baseData = "Base_Data.xlsx"
wb1 = xl.load_workbook(baseData)
workSheet1 = wb1.worksheets[0]

keyData = "key.xlsx"
wb2 = xl.load_workbook(keyData)
workSheet2 = wb2.worksheets[0]


# opening the destination excel file
finalSheet = "final.xlsx"
wb3 = xl.load_workbook(finalSheet)
workSheet3 = wb3.active

# calculate total number of rows and
# columns in source excel file
mr = workSheet1.max_row
mc = workSheet1.max_column
print(mr, mc)

# ___________________________creating structure of base data and raw marks entry__________________________________
# copying the cell values from source
# excel file to destination excel file
x = 0
for l in range(6, mc+1):
	b = workSheet1.cell(row=1, column=l)
	workSheet3.cell(row=1, column=l + 2 * x + 1).value = b.value
	x = x+1

y = 0
for k in range(6, mc+1):
	workSheet3.cell(row=2, column=k + y, value='Score')
	workSheet3.cell(row=2, column=k + y + 1, value='Grades')
	workSheet3.cell(row=2, column=k + y + 2, value='Credits')
	y = y + 2
workSheet3.cell(row=2, column=k + y + 1, value='Average Score')
workSheet3.cell(row=2, column=k + y + 2, value='CGPA')

# _______________________________RAW SCORE ENTRY_______________________________________
for m in range(3, mr+2):
	z = 0
	total = 0
	for n in range(6, mc+1):
		c = workSheet1.cell(row=m - 1, column=n)
		total = total + c.value
		workSheet3.cell(row=m, column=n + z).value = c.value
		z = z+2
	workSheet3.cell(row=m, column=k + y + 1, value=total / (mc - 5))

print(c.value)

for i in range(1, mr + 1):
	for j in range(1, 6):
		# reading cell value from source excel file
		c = workSheet1.cell(row=i, column=j)

		# writing the read value to destination excel file
		workSheet3.cell(row=i + 1, column=j).value = c.value

# ________________________________Credits Copying____________________________________________

for m1 in range(3, mr+2):
	z1 = 0
	x1 = 0

	for n1 in range(6, mc+1):
		c = workSheet2.cell(row=2, column=n1 + x1 - 3)
		x1 = x1 + 1

		workSheet3.cell(row=m1, column=n1 + z1 + 2).value = c.value
		z1 = z1+2

# total = c.value + c.value
# ws3.cell(row=25, column=3).value = total

print(c.value)

# _____________________________________Saving_________________________________________

# saving the destination excel file
wb3.save(str(finalSheet))