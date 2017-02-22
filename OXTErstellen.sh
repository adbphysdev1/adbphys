#!/bin/bash
# Dieses Skript erstellt ein Archiv mit allen Dateien im
# Ordner "Makros"

#echo "Bitte Name für neue Makrosammlung eingebeben:"
#read NAME
cd Makros
NAME=$(cat description.xml | grep '<version value="' | cut -d '"' -f 2)
NAMEMakros="MakrosAufgDB."$NAME".oxt"
PfadZuMakrosammlung="../../"$NAMEMakros
zip -qr $PfadZuMakrosammlung *
echo ""
echo "------- OXT-Erstellen: -------"
echo ""
echo "Die neu erstellte Makrosammlung hat den Namen: -->"$NAMEMakros"<--"
echo "                                               ------------------------------"
echo "... und befindet sich im Verzeichnis:"
cd ../..
pwd
echo ""
echo "------- OXT-Erstellen beendet -------"
echo ""
