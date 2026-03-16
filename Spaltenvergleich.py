import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import FormulaRule
import io
from pathlib import Path
import matplotlib.pyplot as plt
from datetime import datetime
from utils import highlight_differences, de_sort_key

# Spalte einer Excel-Datei filtern nach bestimmten Zeichen

# !!! Überpüfen: !!!
path_folder = Path(r'C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp\Quellen')    #Dateipfad überprüfen
file_1_name = 'Auszug Marktstammdatenregister 23.02.2026'
file_1_sheet_name = 'Stromerzeuger (77)'
file_1_header = 0
file_1_needed_columns = ['MaStR-Nr. der Einheit', 'Anzeige-Name der Einheit', 'Betriebsstatus', 'Systemstatus', 'NBP-Status', 'Energieträger', 'Bruttoleistung der Einheit', 'Inbetriebnahmedatum der Einheit', 'Straße']

# Excel einlesen
excel_file_1 = pd.read_excel(path_folder / (file_1_name + '.xlsx'), file_1_sheet_name, header=file_1_header)

#excel_file_1 = excel_file_1[excel_file_1[file_1_filter_column].astype(str).str.contains(file_1_filter_character, na=False)]
excel_file_1["Straße"] = excel_file_1["Straße"] + " " + excel_file_1["Hausnummer"].fillna("").astype(str)
excel_file_1 = excel_file_1[file_1_needed_columns]  # auskommentieren, wenn alle Spalten angezeigt werden sollen
excel_file_1 = excel_file_1[excel_file_1["Energieträger"] == "Solare Strahlungsenergie"]
df_multiple = excel_file_1[excel_file_1.duplicated(subset=["Straße", "Bruttoleistung der Einheit"], keep=False)]

df_multiple["_sort_key"] = (df_multiple["Straße"].fillna(df_multiple['Anzeige-Name der Einheit']).astype(str).str.strip().apply(de_sort_key))
df_multiple = (df_multiple.sort_values(by="_sort_key").drop(columns="_sort_key").reset_index(drop=True))

# Excel-Datei kreiren
file_name = 'Auswertung_MaStR'
date_today = datetime.now().strftime('%Y%m%d') + '_'
path_folder = Path(r'C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp\Ergebnisse')
file_path = path_folder / (date_today + file_name + '.xlsx')

df_multiple.to_excel(file_path, index=False)
print("Datei wurde erfolgreich gespeichert!")

