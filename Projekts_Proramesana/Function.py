from openpyxl import Workbook
from openpyxl import load_workbook
import time

ex = load_workbook("Data.xlsx")
ac = ex["Sheet"]
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
        skaits()
    elif Izvele == "rediget":
        tirit()
        rediget()
    elif Izvele == "palīdzība":
        tirit()
        #palidziba()
    elif Izvele == "iziet":
        iz = input("Vai tiešām vēlaties iziet no programmas? Y vai N :" + " ").lower()
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
    tirit()
    id = input("Ievadiet produkta ID :" + " ")

    if id.isdigit() == False:
        tirit()
        print("Nederīga vērtība, lūdzu ievadiet ID")
        rediget()

    else:

        pass
    i = 0
    for row in lapa:
        i = i+1
        if ex.active.cell(row=i,column=1).value == int(id):
            nosk = str(ex.active.cell(row=i, column=2).value)
            skaits = str(ex.active.cell(row=i, column=3).value)
            print("Tiks rediģets produkts" + " " + nosk + " " + ", kura skaits ir" + " " + skaits)
            atb = input("Turpināt? : Y vai N").lower()
        else:
            tirit()
            print("Nav atrasts produkts ar ID :" + id)
            time.sleep(0.5)
            rediget()

        if atb == "y":
            opc = input("Pievienot vai noņemt? :" + " ").lower()
            if opc == "pievienot":
                piev = input("Cik daudz vēlaties pievienot?:" + " ")
                while piev.isdigit() == False:
                    tirit()
                    print("Nederīga vērtība, lūdzu ievadiet skaitu")
                    time.sleep(0.5)
                    piev = input("Cik daudz vēlaties pievienot/noņemt?:" + " ")
                else:
                    skaits = skaits+piev
                    print("Produkta" + " " + nosk + " " + "jaunais daudzums ir" + " " + skaits)
                    ex.save("Dati.xlsx")
            if opc == "nonemt":
                piev = input("Cik daudz vēlaties noņemt?:" + " ")
                while piev.isdigit() == False:
                    tirit()
                    print("Nederīga vērtība, lūdzu ievadiet skaitu")
                    time.sleep(0.5)
                    piev = input("Cik daudz vēlaties pievienot/noņemt?:" + " ")
                else:
                    skaits = skaits - piev
                    time.sleep(0.5)
                    print("Produkta" + " " + nosk + " " + "jaunais daudzums ir" + " " + skaits)
                    ex.save("Dati.xlsx")

            else:
                print("Nederīga vērtība")
                sakEkrn()

        if atb == "n":
            print("Novirzu atpakaļ uz sākuma ekrānu...")
            time.sleep(0.5)
            sakEkrn()

def skaits():

    tirit()
    id = input("Ievadiet ID produktam, kura skaitu vēlaties apskatīt :" + " ")

    while id.isdigit() == False:
        tirit()
        print("Nederīga vērtība, lūdzu ievadiet derīgu ID")
        time.sleep(0.5)
        skaits()

    else:

        pass
        i = 0
        for row in lapa:
            i = i + 1
            if ex.active.cell(row=i, column=1).value == int(id):
                nosk = str(ex.active.cell(row=i, column=2).value)
                skaits = str(ex.active.cell(row=i, column=3).value)
                print("Produkta" + " " + nosk + " " + "daudzums ir" + " " + skaits)
            else:
                tirit()
                print("Ievadītais ID :" + " " + str(id) + " " + "nav atpzīts, lūdzu mēģiniet vēlreiz")
                time.sleep(0.5)
                skaits()




    velv = input("Vai vēlaties apskatīt vēl kāda cita produkta skaitu? Y vai N" + " ")
    if velv == "y":
        tirit()
        skaits()
    elif velv == "n":
        print("Novadu uz sākuma ekrānu...")
        time.sleep(0.5)
        sakEkrn()
    else:
        print("Nesapratu atbildi, novirzu uz sākuma ekrānu...")
        time.sleep(0.5)
        sakEkrn()




# varetu but ka 132 - 143 neizpildas








def tirit():
    i = 0
    while i < 10:
        i = i+1
        print("""
        
        
        
        
        
        
        
        """)

