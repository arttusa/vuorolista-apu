import xlrd
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook


def laskeTyoaika(alkamisaika, loppumisaika):
    #muistetaan muuttaa minuutit sadasosiksi
    alkamistunnit = float(alkamisaika[:len(alkamisaika)-3])
    alkamisminuutit = float(alkamisaika[len(alkamisaika)-2:]) / 60.0
    alkamisaika = alkamistunnit + alkamisminuutit
    loppumistunnit = float(loppumisaika[:len(loppumisaika)-3])
    loppumisminuutit = float(loppumisaika[len(loppumisaika)-2:]) / 60.0
    loppumisaika = loppumistunnit + loppumisminuutit
    tyoaika = loppumisaika - alkamisaika
    return round(tyoaika,2)

  
       


location1 = "~/OneDrive/Python_testit/Kesakuu-converted.xlsx"
tyovuorolista = xlrd.open_workbook(location1)
tyoaikakortti = load_workbook('Tyoaikakortti.xlsx')
tyoaikakortin_taulukko = tyoaikakortti["tunnit"]
#on olemassa alkukuu ja loppukuu erikseen, koska ne ovat eri sheeteilla
alkukuu = tyovuorolista.sheet_by_index(0)
loppukuu = tyovuorolista.sheet_by_index(1)


#aivan ensimmaiseksi etsitaan kaikki paivat ja lisataan ne tyoaikakorttiin
paivat = []
viikonpaiva1_vuorolistassa = alkukuu.cell_value(0,1)
#etsitaan y-koordinaatti ensimmaisesta tyopaivasta koska x aina 1
xviikonpaiva = int
for i in range(9, 16):
    if(tyoaikakortin_taulukko.cell(row=i, column=2).value == viikonpaiva1_vuorolistassa):
        xviikonpaiva = i
#etsitaan kaikki tyopaivat kuulta
for i in range (1, alkukuu.ncols):
    paiva1 = alkukuu.cell_value(1, i)
    paivat.append(paiva1)
for i in range (1, loppukuu.ncols):
    paiva2 = loppukuu.cell_value(1, i)
    paivat.append(paiva2)
#lisataan etsityt paivat tyoaikakorttiin
indeksi = 0
for i in range(xviikonpaiva, xviikonpaiva + len(paivat)):
    taytettava_viikonpaiva = tyoaikakortin_taulukko.cell(row=i, column=3)
    taytettava_viikonpaiva.value = paivat[indeksi]
    indeksi += 1



#otetaan kasittelyyn alkukuun tyoajat
print("alkukuu:")
#etsitaan y-koordinaatti
arttuY1 = int
for i in range (0, alkukuu.nrows):
    koordinaatti1 = alkukuu.cell_value(i, 0)
    if(koordinaatti1 == "ARTTU"):
        arttuY1 = i

#etsitaan tyotunnit samalla myos lisataan suoraan tyoaikakorttiin
for i in range (1, alkukuu.ncols):
    viikonpaiva1 = alkukuu.cell_value(0, i)
    day1 = alkukuu.cell_value(1, i)
    cell1 = alkukuu.cell_value(arttuY1, i)
    alkamisaika = cell1[:len(cell1)-7]
    loppumisaika = cell1[len(cell1)-5:]
    if (cell1 != ""):
        print(viikonpaiva1 + " " + day1 + " " + cell1)
        #lisataan tyoaikakorttiin jos ajankohta tasmaa
        for j in range(xviikonpaiva, xviikonpaiva + len(paivat)):
            #tarkistetaan onko oikea paiva
            if(tyoaikakortin_taulukko.cell(row=j, column=3).value == day1):
                tyoaikakortin_taulukko.cell(row=j, column=4).value = alkamisaika 
                tyoaikakortin_taulukko.cell(row=j, column=6).value = loppumisaika
                tyoaikakortin_taulukko.cell(row=j, column=8).value = laskeTyoaika(alkamisaika, loppumisaika)

#otetaan kasittelyyn loppukuun tyoajat
print("loppukuu:")
#etsitaan y-koordinaatti
arttuY2 = None
for i in range (0, loppukuu.nrows):
    koordinaatti2 = loppukuu.cell_value(i, 0)
    if(koordinaatti2 == "ARTTU"):
        arttuY2 = i
#etsitaan tyotunnit
for i in range (1, loppukuu.ncols):
    viikonpaiva2 = loppukuu.cell_value(0, i)
    day2 = loppukuu.cell_value(1, i)
    cell2 = loppukuu.cell_value(arttuY2, i)
    #pitaa erotella nain koska python ei tunnista viivaa, joka on alku- ja loppuajan valissa
    alkamisaika = cell2[:len(cell2)-7] 
    loppumisaika = cell2[len(cell2)-5:]
    if (cell2 != ""):
        print(viikonpaiva2 + " " + day2 + " " + cell2)
        #lisataan tyoaikakorttiin jos ajankohta tasmaa
        for j in range(xviikonpaiva, xviikonpaiva + len(paivat)):
            #tarkistetaan onko oikea paiva
            if(tyoaikakortin_taulukko.cell(row=j, column=3).value == day2):
                tyoaikakortin_taulukko.cell(row=j, column=4).value = alkamisaika
                tyoaikakortin_taulukko.cell(row=j, column=6).value = loppumisaika
                tyoaikakortin_taulukko.cell(row=j, column=8).value = laskeTyoaika(alkamisaika, loppumisaika)
tyoaikakortti.save('Tyoaikakortti_valmis.xlsx')

