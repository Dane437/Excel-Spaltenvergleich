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
run_in_vs_code = True
path_folder = Path(r'C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp\MaStR_Ortsteil')
file_1_name = 'MaStR'
file_1_sheet_name = 'Stromerzeuger'
file_1_header = 0
file_1_needed_columns = ['MaStR-Nr. der Einheit', 'Anzeige-Name der Einheit', 'Betriebsstatus', 'Systemstatus', 'NBP-Status', 'Energieträger', 'Bruttoleistung der Einheit', 
                         'Nettonennleistung der Einheit', 'Inbetriebnahmedatum der Einheit', 'Registrierungsdatum der Einheit',
                         'Postleitzahl', 'Straße', 'Nutzbare Speicherkapazität in kWh']

file_2_name = 'Straßen Westnetz'
file_2_sheet_name = 'List1'
file_2_header = 0
print('Lädt...')

# Excel einlesen
if run_in_vs_code:
    file_path_1 = path_folder / (file_1_name + '.xlsx')
    file_path_2 = path_folder / (file_2_name + '.xlsx')
else:
    file_path_1 = file_1_name + '.xlsx' 
    file_path_2 = file_2_name + '.xlsx' 
excel_file_1 = pd.read_excel(file_path_1, file_1_sheet_name, header=file_1_header)
excel_file_2 = pd.read_excel(file_path_2, file_2_sheet_name, header=file_2_header)

# Daten bearbeiten
excel_file_1 = excel_file_1[file_1_needed_columns]  # auskommentieren, wenn alle Spalten angezeigt werden sollen
excel_file_1 = excel_file_1[excel_file_1['Betriebsstatus'] == 'In Betrieb']
excel_file_1 = excel_file_1[excel_file_1['Systemstatus'] == 'Aktiviert']
excel_file_1['Inbetriebnahmedatum der Einheit'] = excel_file_1['Inbetriebnahmedatum der Einheit'].dt.year

df_all_pv = excel_file_1[excel_file_1['Energieträger'] == 'Solare Strahlungsenergie']

# DataFrame Westnetz installierte PV nach Jahren
df_all_pv_west = df_all_pv[df_all_pv['Postleitzahl'] == 91056]

merged_df: pd.DataFrame = pd.merge(df_all_pv_west, excel_file_2, on='Straße', how='left')

# 
list_districts = merged_df.loc[merged_df['Ortsteil'] != '-', 'Ortsteil'].dropna().unique().tolist() # Liste aller Ortsteile

# Excel-Datei kreiren
file_name = 'Auswertung_MaStR_Westnetz'
date_today = datetime.now().strftime('%Y%m%d') + '_'
if run_in_vs_code:
    file_path = path_folder / (date_today + file_name + '.xlsx')
    file_path_2 = path_folder / (date_today + 'alOrTe.xlsx')
else:
    file_path = date_today + file_name + '.xlsx'

merged_df.to_excel(file_path_2, index=False)
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    #df_all_pv.to_excel(writer, sheet_name='Gesamte PVs', index=False)
    for district in list_districts:
        df_district = merged_df[merged_df['Ortsteil'] == district]
        df_district_power_per_year = (df_district.groupby('Inbetriebnahmedatum der Einheit')['Bruttoleistung der Einheit'].sum().reset_index())
        df_district_power_per_year = df_district_power_per_year.rename(columns={'Bruttoleistung der Einheit': 'Zuwachs pro Jahr', 'Inbetriebnahmedatum der Einheit': 'Jahr'})
        df_district_power_per_year['Summe'] = (df_district_power_per_year['Zuwachs pro Jahr'].cumsum())
        df_district_power_per_year.to_excel(writer, sheet_name=district, index=False)
    #df_all_storage.to_excel(writer, sheet_name='Gesamte Speicher', index=False)

sheet_names = list_districts
for sheet_name in sheet_names: automatic_column_width(file_path, sheet_name=sheet_name)

print('Datei wurde erfolgreich erstellt!')

