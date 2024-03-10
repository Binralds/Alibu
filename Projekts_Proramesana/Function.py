from openpyxl import Workbook
from openpyxl import load_workbook
import time
import timeit

ex = load_workbook("Dati.xlsx")
sht = ex["Sheet"]
active = ex.active


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
    tirit()
    id = input("Ievadiet produkta ID :" + " ")

    if id.isdigit() == False:
        tirit()
        print("Nederīga vērtība, lūdzu ievadiet ID")
        rediget()

    else:

        pass
    i = 0
    for row in sht:
        i = i+1
        if active.cell(row=i, column=1) == int(id):
            nosk = active.cell(row=i, column=2).value
            skaits = active.cell(row=i, column=3).value
            print("Tiks rediģets produkts" + nosk + ", kura skaits ir" + skaits)
            atb = input("Turpināt? : Y vai N").lower()
        else:
            tirit()
            print("Nav atrasts produkts ar ID :" + id)
            time.sleep(0.5)
            rediget()

        if atb == "y":
            opc = input("Pievienot vai noņemt? :" + " ").lower()
            if opc == "pievienot":
                piev = int(input("Cik daudz vēlaties pievienot?:" + " "))
                while piev.isdigit() == False:
                    tirit()
                    print("Nederīga vērtība, lūdzu ievadiet skaitu")
                    time.sleep(0.5)
                    piev = int(input("Cik daudz vēlaties pievienot/noņemt?:" + " "))

                else:
                    skaits = int(skaits)+piev
                    print("Produkta" + nosk + "jaunais daudzums ir" + skaits)
                    ex.save("Dati.xlsx")
                    sakEkrn()
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
                    print("Produkta" + nosk + "jaunais daudzums ir" + skaits)
                    ex.save("Dati.xlsx")
                    sakEkrn()

            else:
                print("Nederīga vērtība")
                sakEkrn()

        if atb == "n":
            print("Novirzu atpakaļ uz sākuma ekrānu...")
            time.sleep(0.5)
            sakEkrn()
    atb = input("Rediģēt citus produktus? Y vai N :" + " ").lower()
    if atb == "y":
        rediget()
    elif atb == "n":
        sakEkrn()
    else:
        print("Nesapratu, novirzu uz sākuma ekrānu...")
        time.sleep(0.5)
        sakEkrn()
def skaits():

    tirit()
    id = input("Ievadiet ID produktam, kura skaitu vēlaties apskatīt :" + " ")

    while id.isdigit() == False:
        tirit()
        print("Nederīga vērtība, lūdzu ievadiet derīgu ID")
        time.sleep(0.5)
        sakEkrn()


    else:

        pass
        i = 0
        for row in sht:
            i = i + 1
            if active.cell(row=i, column=1) == int(id):
                nosk = active.cell(row=i, column=2).value
                daudz = active.cell(row=i, column=3).value
                print("Produkta" + nosk + "daudzums ir" + daudz)
                break
        else:
            tirit()
            print("Ievadītais ID :" + id + "nav atpzīts, lūdzu mēģiniet vēlreiz")
            time.sleep(0.5)
            sakEkrn()

    atb = input("Skatīt citus produktus? Y vai N:" + " ").lower()
    if atb == "y":
        skaits()
    elif atb == "n":
        sakEkrn()
    else:
        print("Nesapratu, novirzu uz sākuma ekrānu...")
        time.sleep(0.5)
        sakEkrn()

def produkti():
    print("""Izvēlieties darbību 
        Pievienot ll Dzēst ll Atrast
    """)
    izv = input("Jūsu izvēle:" + " ").lower
    if izv == "pievienot":
        tirit()
        prod_piev()
    elif izv == "dzēst" or "dzest":
        tirit()
        prod_dzest()
    elif izv == "atrast":
        tirit()
        prod_atr()
    else:
        print("Nesapratu, lūdzu mēģiniet vēlreiz...")
        produkti()

def prod_piev():
    newID = input("Lūdzu ievadiet jaunā produkta ID:" + " ")
    nosk = input("Lūdzu ievadiet jaunā produkta NOSAUKUMU:" + " ")
    daudz = input("Lūdzu ievadiet jaunā produkta SKAITU" + " ")
    print("Produktu ar ID : "+str(newID)+" ,Nosaukums: "+str(nosk)+" un daudzums: "+str(daudz)+" pievienošu izklājlapai")
    atb = input("Turpināt? Y vai N :" + " ").lower()
    if atb == "n":
        print("Dzēšu datus...")
        time.sleep(0.3)
        print("Novirzu atpakaļ uz sākuma ekrānu")
        time.sleep(0.3)
        sakEkrn()
    elif atb == "y":
        i = 0
        for row in sht:
            i = i+1
            if cell.value is None:
                active.cell(row=i, column=1).value = int(newID)
                active.cell(row=i, column=2).value = nosk
                active.cell(row=i, column=3).value = daudz
                ex.save("Dati.xslx")
                break
            print("Izmaiņas saglabātas rindā " + i)
            time.sleep(0.5)
            sakEkrn()
        else:
            active.append([newID, nosk, daudz])
            ex.save("Dati,xslx")
    else:
        print("Neizprotu atbildi, nosūtu uz sākuma ekrānu...")
        time.sleep(0.5)
        sakEkrn()


def skatit():
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














def tirit():
    i = 0
    while i < 10:
        i = i+1
        print("""
        
        
        
        
        
        
        
        """)

#rahhh
