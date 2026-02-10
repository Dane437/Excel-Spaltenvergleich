import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import string
from pathlib import Path
import pandas as pd

#Datenvergleich zwischen Netzleitzahlen.xlsm und K3V-Export.xlsx

#file paths
path_folder = Path(r"\\estw-01\Bereich-N\NG\NGE\Statistiken ESTW\20-kV-Stationen")
path_Netzleitzahlen = path_folder / "20260210_Netzleitzahlen.xlsm"
path_K3_Export = path_folder / "K3V-Export_20260205.xlsx"

#Excel einlesen
excel_Netzleitzahlen = pd.read_excel(path_Netzleitzahlen, sheet_name="Stationsliste")
excel_Netzleitzahlen['NLZ'] = pd.to_numeric(excel_Netzleitzahlen['NLZ'], errors='coerce')
excel_K3_Export = pd.read_excel(path_K3_Export, sheet_name="Tabelle1")
excel_K3_Export[['Nummer NLZ', 'Nummer NKZ']] = (excel_K3_Export['Nummer'].str.extract(r'TS\s*(\d{3})\s*/\s*(.+)'))
excel_K3_Export['Nummer NLZ'] = pd.to_numeric(excel_K3_Export['Nummer NLZ'], errors='coerce')

#Datensätze zusammenführen
merge_nlz: pd.DataFrame = pd.merge(excel_Netzleitzahlen, excel_K3_Export, left_on='NLZ', right_on='Nummer NLZ', how='outer')
merge_nlz['NLZ'] = pd.to_numeric(merge_nlz['NLZ'], errors='coerce')
merge_nlz["Netz"] = merge_nlz["Netz"] + "netz"

# Unterschiede finden
difference_nlz = merge_nlz[merge_nlz['NLZ'].fillna('') != merge_nlz['Nummer NLZ'].fillna('')]
only_nlz = difference_nlz[['NLZ', 'Nummer NLZ']]
difference_nkz = merge_nlz[merge_nlz['NKZ'].fillna('') != merge_nlz['Nummer NKZ'].fillna('').str.upper()]
only_nkz = difference_nkz[['NKZ', 'Nummer NKZ']]
difference_address = merge_nlz[merge_nlz['Stationsname'].fillna('') != merge_nlz['K3v ist'].fillna('')]    #eine leere Zelle zählt als Unterschied
only_address = difference_address[['Stationsname', 'K3v ist']]
difference_netz = merge_nlz[merge_nlz['Netz'].fillna('') != merge_nlz['Teilnetz'].fillna('')]    #eine leere Zelle zählt als Unterschied
difference_netz['Netz'] = difference_netz['Netz'].str.replace('netz', '', regex=False)  # "netz" aus der Spalte "Netz" entfernen
only_netz = difference_netz[['Netz', 'Teilnetz', 'NKZ', 'Nummer NKZ']]

# Excel-Datei kreiren
#merge_nlz.to_excel('Gesamt.xlsx', index=False)
with pd.ExcelWriter('Unterschiede_K3V_Netzleitzahlen.xlsx', engine='openpyxl') as writer:
    only_nlz.to_excel(writer, sheet_name='NLZ', index=False)
    only_nkz.to_excel(writer, sheet_name='NKZ', index=False)
    only_address.to_excel(writer, sheet_name='Adresse', index=False)
    only_netz.to_excel(writer, sheet_name='Netz', index=False)
print("Datei wurde erfolgreich gespeichert!")