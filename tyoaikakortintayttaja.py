import sys
loppumisaika = sys.argv[1]
def laskeIltalisa(loppumisaika):
    if(float(loppumisaika[:2]) >= 18):
        ylimeneva_tunteina = float(loppumisaika[:2]) - 18
        ylimeneva_minuutteina = float(loppumisaika[len(loppumisaika)-2:])
        aika = ylimeneva_tunteina + ylimeneva_minuutteina / 60
        return aika
    else:
        return 0
print (loppumisaika + "jatkuuko")
#print(laskeIltalisa(loppumisaika))