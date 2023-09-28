import openpyxl
from openpyxl import Workbook
import os


# Check for file
if os.path.exists('Data.xlsx') == True:
    Book = Workbook('Data.xlsx')
    print('Workbook Found!')
else:
    Book = Workbook()
    Top = [["ID", "NAME", "COUNT"]]
    for row in Top:
        Book.active.append(row)
    print("Created Workbook / Not found")


Sheet = Book.active




Book.save('Data.xlsx')


