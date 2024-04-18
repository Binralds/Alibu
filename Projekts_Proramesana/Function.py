
from openpyxl import load_workbook
import time



ex = load_workbook("Dati.xlsx")
sht = ex["Sheet"]



def sakEkrn():
    print("""Izvēlieties opciju : \n
    Produkti || Skaits || Rediģēt || Palīdzība
    """)
    Izvele = input("Izvēle : " + "").lower()

    if Izvele == "produkti":
        tirit()
        produkti()
    elif Izvele == "skaits":
        tirit()
        skaits()
    elif Izvele == "rediget" or "rediģēt":
        tirit()
        rediget()
    elif Izvele == "palīdzība" or "palidziba":
        tirit()
        # palidziba()
    else:
        print("Nederīga opcija, lūdzu mēģiniet vēlreiz...")
        sakEkrn()


def rediget():
    ex = load_workbook("Dati.xlsx")
    sht = ex["Sheet"]
    id = input("Ievadiet produkta ID :" + " ")

    if not id.isdigit():
        print("Nederīga vērtība, lūdzu ievadiet ID")
        time.sleep(1.0)
        sakEkrn()

    else:

        pass
    i = 0
    for row in sht:
        i = i + 1
        if str(ex.active.cell(row=i, column=1).value) == str(id):
            nosk = ex.active.cell(row=i, column=2).value
            skaits = ex.active.cell(row=i, column=3).value
            print("Tiks rediģets produkts " + str(nosk) + ", kura skaits ir " + str(skaits))
            atb = input("Turpināt? Y vai N : ").lower()



            if atb == "y":
                opc = input("Pievienot vai noņemt? :" + " ").lower()
                if opc == "pievienot":
                    piev = input("Cik daudz vēlaties pievienot?:" + " ")
                    if not piev.isdigit():
                        print("Nederīga vērtība, lūdzu ievadiet skaitu")
                        time.sleep(0.5)
                        rediget()

                    else:
                        ex.active.cell(row=i, column=3).value = int(ex.active.cell(row=i, column=3).value) + int(piev)
                        print("Produkta " + str(nosk) + " jaunais daudzums ir " + str(ex.active.cell(row=i, column=3).value))
                        ex.save("Dati.xlsx")
                        sakEkrn()
                if opc == "nonemt" or "noņemt":
                    noNemt = input("Cik daudz vēlaties noņemt?:" + " ")
                    if not noNemt.isdigit():
                        print("Nederīga vērtība, lūdzu ievadiet skaitu")
                        time.sleep(0.5)
                        rediget()
                    else:
                        ex.active.cell(row=i, column=3).value = int(ex.active.cell(row=i, column=3).value) - int(noNemt)
                        time.sleep(0.5)
                        print("Produkta " + str(nosk) + " jaunais daudzums ir " + str(ex.active.cell(row=i, column=3).value))
                        ex.save("Dati.xlsx")
                        sakEkrn()

                else:
                    print("Nederīga vērtība")
                    sakEkrn()

            elif atb == "n":
                print("Novirzu atpakaļ uz sākuma ekrānu...")
                time.sleep(0.5)
                sakEkrn()

            else:
                print("Nesapratu, novirzu atpakaļ uz sākuma ekrānu...")
                time.sleep(1.0)
                sakEkrn()
    else:
        print("Nav atrasts produkts ar ID : " + id)
        time.sleep(1.0)
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
    ex = load_workbook("Dati.xlsx")
    sht = ex["Sheet"]

    tirit()
    id = input("Ievadiet ID produktam, kura skaitu vēlaties apskatīt :" + " ")

    if not id.isdigit():
        print("Nederīga vērtība, lūdzu ievadiet derīgu ID")
        time.sleep(1.0)
        sakEkrn()


    else:

        pass
        i = 0
        for row in sht:
            i = i + 1
            if str(ex.active.cell(row=i, column=1).value) == str(id):
                nosk = sht.cell(row=i, column=2).value
                daudz = sht.cell(row=i, column=3).value
                print("Produkta " + str(nosk) + " daudzums ir " + str(daudz))
                break
        else:
            tirit()
            print("Ievadītais ID :" + str(id) + " nav atpzīts, lūdzu mēģiniet vēlreiz")
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
    ex = load_workbook("Dati.xlsx")
    sht = ex["Sheet"]

    print("""Izvēlieties darbību 
        Pievienot || Dzēst || Atrast
    """)
    izv = input("Jūsu izvēle:" + " ").lower()

    if izv == "pievienot":
        prod_piev()
    elif izv == "dzēst" or izv == "dzest":
        prod_dzest()
    elif izv == "atrast":
        prod_atr()
    else:
        print("Nesapratu, lūdzu mēģiniet vēlreiz...")
        produkti()


def prod_piev():
    tirit()
    ex = load_workbook("Dati.xlsx")
    sht = ex["Sheet"]
    max = sht.max_row + 1

    newID = input("Lūdzu ievadiet jaunā produkta ID:" + " ")
    if not newID.isdigit():
        print("Nederīga vērtība, lūdzu ievadiet vērtību, kas atbilst ID")
        time.sleep(0.5)
        sakEkrn()
    nosk = input("Lūdzu ievadiet jaunā produkta NOSAUKUMU:" + " ")
    daudz = input("Lūdzu ievadiet jaunā produkta SKAITU:" + " ")
    if not daudz.isdigit():
        print("Nederīga vērtība, lūdzu ievadiet SKAITU")
        time.sleep(0.5)
        sakEkrn()
    print("Produktu ar ID : " + str(newID) + " ,nosaukumu: " + str(nosk) + " un daudzumu: " + str(
        daudz) + " pievienošu izklājlapai")
    time.sleep(0.5)
    atb = input("Turpināt? Y vai N :" + " ").lower()
    if atb == "n":
        print("Dzēšu datus...")
        time.sleep(0.3)
        print("Novirzu atpakaļ uz sākuma ekrānu")
        time.sleep(0.3)
        sakEkrn()
    elif atb == "y":
        sht.cell(row=max, column=1).value = newID
        sht.cell(row=max, column=2).value = nosk
        sht.cell(row=max, column=3).value = int(daudz)
        ex.save("Dati.xlsx")
        print("Izmaiņas saglabātas rindā " + str(max))
        time.sleep(0.5)
        izv = input("Vai vēlaties pievienot vēlvienu produktu? Y vai N " + "").lower()
        if izv == "y":
            prod_piev()
        elif izv == "n":
            print("Novirzu uz sākuma ekrānu...")
            sakEkrn()
        else:
            print("Nesapratu atbildi, novirzu uz sākuma ekrānu...")
            time.sleep(1.0)
            sakEkrn()
    else:
        print("Neizprotu atbildi, nosūtu uz sākuma ekrānu...")
        time.sleep(0.5)
        sakEkrn()


def prod_atr():
    tirit()
    ex = load_workbook("Dati.xlsx")
    sht = ex["Sheet"]

    aizm = input("Ievadiet nosaukumu produktam, kura ID/Skaitu vēlaties noskaidrot:" + " ")


    i = 0
    for row in sht:
        i = i+1
        if ex.active.cell(row=i, column=2).value == aizm:
            id = str(ex.active.cell(row=i, column=1).value)
            skaits = str(ex.active.cell(row=i, column=3).value)
            print("Produkta" + " " + aizm + " " + "ID ir" + " " + id + " " + "un skaits ir" + " " + skaits)
            time.sleep(0.5)
            izv = input("Vai vēlaties atrast vēlvienu produktu? Y vai N: " + "").lower()
            if izv == "y":
                prod_piev()
            elif izv == "n":
                print("Novirzu uz sākuma ekrānu...")
                sakEkrn()
            else:
                print("Nesapratu atbildi, novirzu uz sākuma ekrānu...")
                time.sleep(1.0)
                sakEkrn()

    else:
        print("Nevarēju atrast produktu ar norādīto nosaukumu, lūdzu mēģiniet vēlreiz...")
        time.sleep(0.5)
        produkti()




def skatit():
    ex = load_workbook("Dati.xlsx")
    sht = ex["Sheet"]

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

def prod_dzest():
    tirit()
    ex = load_workbook("Dati.xlsx")
    sht = ex["Sheet"]

    id = input("Ievadiet tā produkta ID, kuru vēlaties dzēst:" + " ")
    i = 0
    for row in sht:
        i = i+1
        if str(ex.active.cell(row=i, column=1).value) == str(id):
            nosk = ex.active.cell(row=i, column=2).value
            daudzums = ex.active.cell(row=i, column=3).value
            print("Tiks dzēsts produkts " + str(nosk) + " ar daudzumu " + str(daudzums))
            atbilde = input("Vai tiešām vēlaties dzēst šo produktu? Datus atgūt nebūs iespējams. Y vai N: ").lower()
            if atbilde == "y":
                sht.delete_rows(i)
                print("Veiksmīgi izdzēsu rindu " + str(i))
                ex.save("Dati.xlsx")
                atbildeNext = input("Vai vēlaties dzēst vēl kādus datus? Y vai N: ").lower()
                if atbildeNext == "y":
                    prod_dzest()
                elif atbildeNext == "n":
                    print("Novirzu atpakaļ uz sākuma ekrānu...")
                    time.sleep(1.5)
                    sakEkrn()
                else:
                    print("Nesaprotu atbildi, novirzu uz sākuma ekrānu...")
                    time.sleep(1.5)
                    sakEkrn()
            elif atbilde == "n":
                print("Novadu uz sākuma ekrānu...")
                time.sleep(1.5)
                sakEkrn()
            else:
                print("Nesaprotu atbildi, novirzu uz sākuma ekrānu...")
                time.sleep(1.5)
                sakEkrn()

    else:
        print("Neatradu produktu ar doto ID: " + str(id) + " ,lūdzu mēģiniet vēlreiz")
        time.sleep(0.5)
        aiz = input("Vai aizmirsāt kāda produkta id?: Y  vai N " + " ").lower()
        if aiz == "y":
            print("Atļaujiet man jums palīdzēt, novadīšu Jūs uz produktu atrašanas funkciju.")
            time.sleep(0.5)
            prod_atr()
        elif aiz == "n":
            print("Novirzu uz sākuma ekrānu...")
            time.sleep(0.5)
            sakEkrn()
        else:
            print("Atbildi nesapratu, novirzu uz sākuma ekrānu...")
            time.sleep(0.5)
            sakEkrn()

def tirit():
    i = 0
    while i < 10:
        i = i + 1
        print("""







        """)

# rahhhh

