import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import FormulaRule
import io
from pathlib import Path
import matplotlib.pyplot as plt
from datetime import datetime
from utils import highlight_differences, de_sort_key, automatic_column_width

# Dieses Script erstellt aus dem MaStR die Werte wie viel PV-Leistung für die einzelnen Jahre zugebaut wurde

# !!! Überpüfen: !!!
path_folder = Path(r'C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp\MaStR')    #Dateipfad überprüfen
file_1_name = 'Marktstammdatenregister 09.04.2026'
file_1_sheet_name = 'Stromerzeuger (81)'
file_1_header = 0
file_1_needed_columns = ['MaStR-Nr. der Einheit', 'Anzeige-Name der Einheit', 'Betriebsstatus', 'Systemstatus', 'NBP-Status', 'Energieträger', 'Bruttoleistung der Einheit', 
                         'Nettonennleistung der Einheit', 'Inbetriebnahmedatum der Einheit', 'Registrierungsdatum der Einheit',
                         'Straße']

# Excel einlesen
excel_file_1 = pd.read_excel(path_folder / (file_1_name + '.xlsx'), file_1_sheet_name, header=file_1_header)

#excel_file_1 = excel_file_1[excel_file_1[file_1_filter_column].astype(str).str.contains(file_1_filter_character, na=False)]
excel_file_1['Straße'] = excel_file_1['Straße'] + ' ' + excel_file_1['Hausnummer'].fillna('').astype(str)
excel_file_1 = excel_file_1[file_1_needed_columns]  # auskommentieren, wenn alle Spalten angezeigt werden sollen
excel_file_1 = excel_file_1[excel_file_1['Energieträger'] == 'Solare Strahlungsenergie']
excel_file_1 = excel_file_1[excel_file_1['Betriebsstatus'] == 'In Betrieb']
excel_file_1 = excel_file_1[excel_file_1['Systemstatus'] == 'Aktiviert']

# Filtern nach doppelten und diese sortieren
df_multiple = excel_file_1[excel_file_1.duplicated(subset=['Straße', 'Bruttoleistung der Einheit'], keep=False)]
df_multiple['_sort_key'] = (df_multiple['Straße'].fillna(df_multiple['Anzeige-Name der Einheit']).astype(str).str.strip().apply(de_sort_key))
df_multiple = (df_multiple.sort_values(by='_sort_key').drop(columns='_sort_key').reset_index(drop=True))

# DataFrame mit der Leistung nach Jahren erstellen
excel_file_1['Inbetriebnahmedatum der Einheit'] = excel_file_1['Inbetriebnahmedatum der Einheit'].dt.year
df_power_per_year = (excel_file_1.groupby('Inbetriebnahmedatum der Einheit')['Bruttoleistung der Einheit'].sum().reset_index())

# Excel-Datei kreiren
file_name = 'Auswertung_MaStR_2024'
date_today = datetime.now().strftime('%Y%m%d') + '_'
file_path = path_folder / (date_today + file_name + '.xlsx')

with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    excel_file_1.to_excel(writer, sheet_name='Gesamte PVs', index=False)
    df_power_per_year.to_excel(writer, sheet_name='Leistung_pro_Jahr', index=False)
    df_multiple.to_excel(writer, sheet_name='Doppelte Anlagen', index=False)
automatic_column_width(file_path, sheet_name='Gesamte PVs')
automatic_column_width(file_path, sheet_name='Doppelte Anlagen')
print('Datei wurde erfolgreich gespeichert!')

