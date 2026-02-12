import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import string
from pathlib import Path
import pandas as pd

#Datenvergleich zwischen Netzleitzahlen.xlsm und K3V-Export.xlsx

#file paths
path_folder = Path(r"\\estw-01\Bereich-N\NG\NGE\Statistiken ESTW\20-kV-Stationen")
path_Netzleitzahlen = path_folder / "20260211_Netzleitzahlen.xlsm"
path_K3_Export = path_folder / "K3V-Export_20260211.xlsx"

#Excel einlesen
excel_Netzleitzahlen = pd.read_excel(path_Netzleitzahlen, sheet_name="Stationsliste")
excel_K3_Export = pd.read_excel(path_K3_Export, sheet_name="Tabelle1")

#Datensätze zusammenführen
merge_adress: pd.DataFrame = pd.merge(excel_Netzleitzahlen, excel_K3_Export, left_on='Stationsname', right_on='K3v', how='outer')



# Unterschiede finden
only_adress_IBS = merge_adress[['Stationsname', 'K3v', 'Inbetriebnahme']]


# Excel-Datei kreiren
#merge_nlz.to_excel('Gesamt.xlsx', index=False)
with pd.ExcelWriter('Stationsname_IBS.xlsx', engine='openpyxl') as writer:
    only_adress_IBS.to_excel(writer, sheet_name='IBS', index=False)
print("Datei wurde erfolgreich gespeichert!")