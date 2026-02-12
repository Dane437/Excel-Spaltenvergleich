import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import string
from pathlib import Path
import matplotlib.pyplot as plt

#Datenvergleich zwischen Netzleitzahlen.xlsm und K3V-Export.xlsx

#file paths

path_folder_Netzleitzahlen = Path(r"B:\# N-Gemeinsam\Trafostationsliste")
path_Netzleitzahlen = path_folder_Netzleitzahlen /  "Netzleitzahlen_Makro.xlsm"
path_folder_K3_Export = Path(r"\\estw-01\Bereich-N\NG\NGE\Statistiken ESTW\20-kV-Stationen")
path_K3_Export = path_folder_K3_Export / "K3V-Export_20260211.xlsx"

#Excel einlesen
excel_Netzleitzahlen = pd.read_excel(path_Netzleitzahlen, sheet_name="Stationsliste")
excel_K3_Export = pd.read_excel(path_K3_Export, sheet_name="Tabelle1")

#Datensätze zusammenführen
merge_adress: pd.DataFrame = pd.merge(excel_Netzleitzahlen, excel_K3_Export, left_on='Stationsname', right_on='K3v', how='outer')

merge_adress["Inbetriebnahme"] = pd.to_datetime(merge_adress["Inbetriebnahme"], dayfirst=True, errors='coerce') 
merge_adress["Inbetriebnahme"] = merge_adress["Inbetriebnahme"].dt.year.astype('Int64')

netzstationen = merge_adress[merge_adress["Stationstyp"] == "Netzstation"]
anzahl_pro_jahr = netzstationen.groupby("Inbetriebnahme").size()


# Unterschiede finden
#only_adress_IBS = merge_adress[['Stationsname', 'K3v', 'Inbetriebnahme']]


# Excel-Datei kreiren
#only_adress_IBS.to_excel('Temp.xlsx', index=False)
#with pd.ExcelWriter('Auswertung.xlsx', engine='openpyxl') as writer:
#    merge_adress.to_excel(writer, sheet_name='IBS', index=False)
#print("Datei wurde erfolgreich gespeichert!")

# Diagramm erzeugen
plt.figure(figsize=(14, 8))
anzahl_pro_jahr.plot(kind="bar")
plt.title("Netzstationen pro Jahr")
plt.xlabel("Jahr")
plt.ylabel("Anzahl")
plt.xticks(rotation=90)
plt.tight_layout()
plt.show()