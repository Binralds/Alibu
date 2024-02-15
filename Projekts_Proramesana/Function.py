from openpyxl import Workbook
from openpyxl import load_workbook
def sakEkrn():
    print("""Sveiki! Izvēlieties opciju : \n
    Produkti ll Skaits ll Rediģēt ll Palīdzība ll Iziet
    """)
    Izvele = input("Izvēle : " + "").lower()

    if Izvele == "produkti":
        tirit()
        #produkti()
    elif Izvele == "skaits":
        tirit()
        #skaits()
    elif Izvele == "rediģēt":
        tirit()
        #rediget()
    elif Izvele == "palīdzība":
        tirit()
        #palidziba()
    elif Izvele == "iziet":
        iz = input("Vai tiešām vēlaties iziet no programmas? y vai n :" + " ").lower()
        if iz == "y":
            print("Visu labu!")
            sakEkrn()
        elif iz == "n":
            sakEkrn()
    else:
        print("Nederīga opcija, lūdzu mēģiniet vēlreiz...")
        sakEkrn()



def produkti():
    ex = load_workbook("Dati.xlsx")
    ac = ex.active

    print()





def tirit():
    i = 0
    while i < 10:
        i = i+1
        print("""
        
        
        
        
        
        
        
        """)