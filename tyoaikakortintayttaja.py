import sys
def laskeIltalisa(loppumisaika):
    if(float(loppumisaika[:2]) >= 18):
        ylimeneva_tunteina = float(loppumisaika[:2]) - 18
        ylimeneva_minuutteina = float(loppumisaika[len(loppumisaika)-2:])
        aika = ylimeneva_tunteina + ylimeneva_minuutteina / 60
        return aika
    else:
        return 0

annettu_parametri = sys.argv[1]
lopullinen_parametri = annettu_parametri[:1].upper() + annettu_parametri[1:].lower()


print(lopullinen_parametri)