from openpyxl import Workbook
from openpyxl import load_workbook
import time

ex = load_workbook("Dati.xlsx")
ac = ex.active
lapa = ac["Sheet"]
def sakEkrn():
    print("""Izvēlieties opciju : \n
    Produkti ll Skaits ll Rediģēt ll Palīdzība ll Iziet
    """)
    Izvele = input("Izvēle : " + "").lower()

    if Izvele == "produkti":
        tirit()
        #produkti()
    elif Izvele == "skaits":
        tirit()
        #skaits()
    elif Izvele == "rediget":
        tirit()
        #rediget()
    elif Izvele == "palīdzība":
        tirit()
        #palidziba()
    elif Izvele == "iziet":
        iz = input("Vai tiešām vēlaties iziet no programmas? y vai n :" + " ").lower()
        if iz == "y":
            print("Visu labu!")
            tirit()
            sakEkrn()
        elif iz == "n":
            tirit()
            sakEkrn()
    else:
        print("Nederīga opcija, lūdzu mēģiniet vēlreiz...")
        sakEkrn()



def rediget():
    id = input("Ievadiet produkta ID :" + " ")

    if id.isdigit() == False:
        print("Nederīga vērtība, lūdzu ievadiet ID")
        rediget()

    else:

        pass
    i = 0
    for row in lapa:
        i = i+1
        if ac.cell(row=i, column=1) == int(id):
            nosk = ac.cell(row=i, column=2).value
            skaits = ac.cell(row=i, column=3).value
        else:
            print("Nederīgs ID, lūdzu mēģiniet vēlreiz")
            time.sleep(0.5)
            id = input("Ievadiet ID produktam, kura skaitu vēlaties rediģēt :" + " ")

        print("Tiks rediģets produkts" + nosk + ", kura skaits ir" + skaits)
        time.sleep(1.0)
        atb = input("Turpināt? : Y vai N").lower()


        if atb == "y":
            opc = input("Pievienot vai noņemt? :" + " ").lower()
            if opc == "pievienot":
                piev = input("Cik daudz vēlaties pievienot/noņemt?:" + " ")
                while piev.isdigit() == False:
                    print("Nederīga vērtība, lūdzu ievadiet skaitu")
                    time.sleep(0.5)
                    piev = input("Cik daudz vēlaties pievienot/noņemt?:" + " ")
                else:
                    skaits = skaits+piev
                    print("Produkta" + nosk + "jaunais daudzums ir" + skaits)
                    ex.save("Dati.xlsx")
            if opc == "nonemt":
                skaits = skaits - piev
                print("Produkta" + nosk + "jaunais daudzums ir" + skaits)
                ex.save("Dati.xlsx")

            else:
                print("Nederīga vērtība")
                sakEkrn()

        if atb == "n":
            print("Novirzu atpakaļ uz sākuma ekrānu")
            sakEkrn()


def skaits():
    id = input("Ievadiet ID produktam, kura skaitu vēlaties apskatīt :" + " ")

    if id.isdigit() == False:
        print("Nederīga vērtība, lūdzu ievadiet derīgu ID")
        time.sleep(0.5)
        id = input("Ievadiet ID produktam, kura skaitu vēlaties apskatīt :" + " ")

    else:

        pass

    i = 0
    for row in lapa:
        i = i + 1
        if ac.cell(row=i, column=1) == int(id):
            nosk = ac.cell(row=i, column=2).value
            skaits = ac.cell(row=i, column=3).value
            print("Produkta" + nosk + "daudzums ir" + skaits)
        else:
            print("Nederīgs ID, lūdzu mēģiniet vēlreiz")
            time.sleep(0.5)
            id = input("Ievadiet ID produktam, kura skaitu vēlaties apskatīt :" + " ")

    velv = input("Vai vēlaties apskatīt vēl kāda cita produkta skaitu? Y vai N" + " ")
    if velv == "y":
        skaits()
    elif velv == "n":
        print("Novadu uz sākuma ekrānu")
        time.sleep(0.5)
        sakEkrn()
    else:
        print("Nesapratu atbildi, novirzu uz sākuma ekrānu")
        time.sleep(0.5)
        sakEkrn()







def tirit():
    i = 0
    while i < 10:
        i = i+1
        print("""
        
        
        
        
        
        
        
        """)