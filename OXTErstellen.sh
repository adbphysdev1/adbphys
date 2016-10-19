#!/bin/bash
# Dieses Skript erstellt ein Archiv mit allen Dateien im
# Ordner "Makros"
echo "Bitte Name für neue Makrosammlung eingebeben:"
read NAME
PfadZuMakrosammlung="../../"$NAME".oxt"
#echo $PfadZuMakrosammlung
cd Makros
zip -r $PfadZuMakrosammlung *
