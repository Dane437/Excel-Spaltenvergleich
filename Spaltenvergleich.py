import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import string
from pathlib import Path
import pandas as pd

#Datenvergleich zwischen Netzleitzahlen.xlsm und GIS-Export_Station.xlsx

#file paths
path_folder = Path(r"\\estw-01\Bereich-N\NG\NGE\Statistiken ESTW\20-kV-Stationen")
path_Netzleitzahlen = path_folder / "20260210_Netzleitzahlen.xlsm"
path_GIS_Export = path_folder / "GIS-Export_Station_20260209.xlsx"

#Excel einlesen
excel_Netzleitzahlen = pd.read_excel(path_Netzleitzahlen, sheet_name="Stationsliste")
excel_Netzleitzahlen['NLZ'] = pd.to_numeric(excel_Netzleitzahlen['NLZ'], errors='coerce')
excel_GIS_Export = pd.read_excel(path_GIS_Export, sheet_name="Station", header=4) #header=4: erste vier Zeilen werden ignoriert, Zeile 5 in Spaltenname
excel_GIS_Export['Stations ID'] = pd.to_numeric(excel_GIS_Export['Stations ID'], errors='coerce')

#Datensätze zusammenführen
merge_nlz: pd.DataFrame = pd.merge(excel_Netzleitzahlen, excel_GIS_Export, left_on='NLZ', right_on='Stations ID', how='outer')
merge_nlz["Netz"] = merge_nlz["Netz"] + "netz"

#Unterschiede finden
difference_nkz = merge_nlz[merge_nlz['NKZ'].fillna('') != merge_nlz['Kurzname'].fillna('')]
only_nkz = difference_nkz[['NKZ', 'Kurzname']]
difference_nlz = merge_nlz[merge_nlz['NLZ'].fillna('') != merge_nlz['Stations ID'].fillna('')]
only_nlz = difference_nlz[['NLZ', 'Stations ID']]
difference_address = merge_nlz[merge_nlz['Stationsname'].fillna('') != merge_nlz['Name'].fillna('')]    #eine leere Zelle zählt als Unterschied
only_address = difference_address[['Stationsname', 'Name']]
difference_netz =  merge_nlz[merge_nlz['Netz'].notna() & merge_nlz['Teilnetz'].notna() & (merge_nlz['Netz'] != merge_nlz['Teilnetz'])]  #eine leere Zelle zählt nicht als Unterschied, da sie mit fillna('') behandelt wird
difference_netz['Netz'] = difference_netz['Netz'].str.replace('netz', '', regex=False)  # "netz" aus der Spalte "Netz" entfernen
only_netz = difference_netz[['Netz', 'Teilnetz', 'NKZ', 'Kurzname']]

#Excel-Datei kreieren
#merge_NKZ.to_excel('Gesamt.xlsx', index=False)
with pd.ExcelWriter('Unterschiede_GIS_Netzleitzahlen.xlsx', engine='openpyxl') as writer:
    only_nlz.to_excel(writer, sheet_name='NLZ', index=False)
    only_nkz.to_excel(writer, sheet_name='NKZ', index=False)
    only_address.to_excel(writer, sheet_name='Adresse', index=False)
    only_netz.to_excel(writer, sheet_name='Netz', index=False)
print("Datei wurde erfolgreich gespeichert!")