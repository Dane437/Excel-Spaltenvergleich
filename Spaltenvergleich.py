import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import string
from pathlib import Path
import pandas as pd

#file paths
path_folder = Path(r"\\estw-01\Bereich-N\NG\NGE\Statistiken ESTW\20-kV-Stationen")
path_Netzleitzahlen = path_folder / "20260202_Netzleitzahlen.xlsm"
path_GIS_Export = path_folder / "GIS-Export_Station_20260202.xlsx"
#path_K3_Export = path_folder / "K3-Export_Vergleich k3v zu Netz.xlsx"

#
excel_Netzleitzahlen = pd.read_excel(path_Netzleitzahlen, sheet_name="Stationsliste")
excel_Netzleitzahlen['NLZ'] = pd.to_numeric(excel_Netzleitzahlen['NLZ'], errors='coerce')
excel_GIS_Export = pd.read_excel(path_GIS_Export, sheet_name="Station", header=4) #header=4: erste vier Zeilen werden ignoriert, Zeile 5 in Spaltenname
excel_GIS_Export['Stations ID'] = pd.to_numeric(excel_GIS_Export['Stations ID'], errors='coerce')

#excel_K3_Export = pd.read_excel(path_K3_Export, sheet_name="Tabelle1")

#
#for i, buchstabe in enumerate(string.ascii_uppercase):
#    globals()[buchstabe] = i

#print(excel_Netzleitzahlen.iloc[20-2, B])
#print(excel_GIS_Export.shape)
#print(excel_Netzleitzahlen.shape)

#
merge_nlz: pd.DataFrame = pd.merge(excel_Netzleitzahlen, excel_GIS_Export, left_on='NLZ', right_on='Stations ID', how='outer')
#merge_nlz["Netz"] = merge_nlz["Netz"] + "netz"
#merge_nlz['NLZ'] = pd.to_numeric(merge_nlz['NLZ'], errors='coerce')
#merge_nlz['Stations ID'] = pd.to_numeric(merge_nlz['Stations ID'], errors='coerce')

# Filter für fehlende Daten (NaN)
#not_in_excel_Netzleitzahlen = merge_nkz[merge_nkz['NKZ'].isna()]
#not_in_excel_GIS_Export = merge_nkz[merge_nkz['Kurzname'].isna()]

difference_nkz = merge_nlz[merge_nlz['NKZ'].fillna('') != merge_nlz['Kurzname'].fillna('')]
only_nkz = difference_nkz[['NKZ', 'Kurzname']]
difference_nlz = merge_nlz[merge_nlz['NLZ'].fillna('') != merge_nlz['Stations ID'].fillna('')]
only_nlz = difference_nlz[['NLZ', 'Stations ID']]
difference_address = merge_nlz[merge_nlz['Stationsname'].fillna('') != merge_nlz['Name'].fillna('')]
only_address = difference_address[['Stationsname', 'Name']]
difference_netz = merge_nlz

#if not not_in_excel_Netzleitzahlen.empty:
#    print("Folgende Nummern fehlen in Netzleitzahlen:")
#    print(not_in_excel_Netzleitzahlen['NKZ'])

#if not not_in_excel_GIS_Export.empty:
#    print("\nFolgende Nummern fehlen im Gis Export:")
#    print(not_in_excel_GIS_Export['Kurzname'])

#if not difference.empty:
#    print("\nInhaltliche Unterschiede bei gleicher Nummer:")
    # Zeigt Nummer, Info aus beiden Exceln
#    print(difference[['NKZ', 'Stationsname', 'Name']])

#merge_NKZ.to_excel('Gesamt.xlsx', index=False)
#difference.to_excel('Unterschiede.xlsx', index=False)
#only_nkz.to_excel('Unterschiede.xlsx', index=False)
#only_address.to_excel('Unterschiede2.xlsx', index=False)

with pd.ExcelWriter('Unterschiede5.xlsx', engine='openpyxl') as writer:
    only_nlz.to_excel(writer, sheet_name='NLZ', index=False)
    only_nkz.to_excel(writer, sheet_name='NKZ', index=False)
    only_address.to_excel(writer, sheet_name='Adresse', index=False)
print("Datei wurde erfolgreich gespeichert!")