import time
import numpy as np
from operator import itemgetter

def lev5(source, target):
    if len(source) < len(target):
        return lev5(target, source)

    # So now we have len(source) >= len(target).
    if len(target) == 0:
        return len(source)

    # We call tuple() to force strings to be used as sequences
    # ('c', 'a', 't', 's') - numpy uses them as values by default.
    source = np.array(tuple(source))
    target = np.array(tuple(target))

    # We use a dynamic programming algorithm, but with the
    # added optimization that we only need the last two rows
    # of the matrix.
    previous_row = np.arange(target.size + 1)
    for s in source:
        # Insertion (target grows longer than source):
        current_row = previous_row + 1

        # Substitution or matching:
        # Target and source items are aligned, and either
        # are different (cost of 1), or are the same (cost of 0).
        current_row[1:] = np.minimum(
                current_row[1:],
                np.add(previous_row[:-1], target != s))

        # Deletion (target grows shorter than source):
        current_row[1:] = np.minimum(
                current_row[1:],
                current_row[0:-1] + 1)

        previous_row = current_row

    return previous_row[-1]

startzeit = time.time()
listeDateien = []
for i in range(1,733): # Alle Dateien in eine Liste einlesen:
    dateiname = str(i) + ".dat"
    dateiinhalt = ""
    # print(dateiname)
    fobj = open("textdateien/" + dateiname, "r") 
    for line in fobj: 
        dateiinhalt = dateiinhalt + line# + "\n"
    fobj.close()
    # print(dateiinhalt + "---")
    listeDateien.append(dateiinhalt)

print(listeDateien[0])

# Ausgabe in Datei speichern:
fobj = open("ausgabe" + str(time.time()) + ".txt", "w")

ausgabe = []
for i in range(0,len(listeDateien)):
    #print(listeDateien[i])
    for j in range(i,len(listeDateien)):
        prozentualeUebereinstimmung = 1-2*lev5(listeDateien[i],listeDateien[j])/(len(listeDateien[i])+len(listeDateien[j]))
        ausgabe.append(str(i+1) + " " + str(j+1) + " " + str(prozentualeUebereinstimmung) + " " + str(time.time()-startzeit))
        fobj.write(ausgabe[len(ausgabe)-1] + "\n")
        if j%50 == 0:
            print(ausgabe[len(ausgabe)-1])
            fobj.flush()

endzeit = time.time()
rechenzeit = endzeit -startzeit

fobj.write("Rechenzeit (s): " + str(rechenzeit) + "\n")
fobj.write("Länge Ausgabe: " + str(len(ausgabe)) + "\n" + "\n")

#for i in range(0,len(ausgabe)):
#    fobj.write(ausgabe[i] + "\n")
fobj.close()

print("Fertig!")
