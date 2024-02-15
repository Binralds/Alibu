import os.path
from Function import sakEkrn

import openpyxl

from openpyxl import Workbook
from openpyxl import load_workbook

if os.path.exists("Dati.xlsx"):
    Book = load_workbook("Dati.xlsx")
    print("Atver datni")
else:
    Book = Workbook()
    Top = [["ID", "NOSK.", "SKAITS"]]
    for row in Top:
        Book.active.append(row)
    Book.active.cell(row=1, column=999).value = 1
    print("Izveidota jauna datne")
    Book.save("Dati.xlsx")




sakEkrn()




