import openpyxl
import Yield2 as yd

wb = openpyxl.load_workbook("Adjusted.xlsx")

for j in range(5):
    for i in range(5):
        n = 0.2
        n1 = n*(i+1)

        wb.save("Adjusted"+str(j)+str(i)+".xlsx")