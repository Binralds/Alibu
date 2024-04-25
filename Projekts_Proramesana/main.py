from openpyxl import Workbook
from openpyxl import load_workbook
from Function import sakEkrn
import os




if os.path.exists("Dati.xlsx"):
    ex = load_workbook("Dati.xlsx")
    print("Atveru datni")

#Nestrada jauna faila izveide
else:
    wb = Workbook()
    Top = [["ID", "NOSK.", "SKAITS"]]
    for row in Top:
        wb.active.append(row)
    print("Izveidoju datni")
    wb.save("Dati.xlsx")


sakEkrn()