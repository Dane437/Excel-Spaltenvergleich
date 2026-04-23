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
run_in_vs_code = False
path_folder = Path(r'C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp\MaStR')
file_1_name = 'MaStR'
file_1_sheet_name = 'Stromerzeuger'
file_1_header = 0
file_1_needed_columns = ['MaStR-Nr. der Einheit', 'Anzeige-Name der Einheit', 'Betriebsstatus', 'Systemstatus', 'NBP-Status', 'Energieträger', 'Bruttoleistung der Einheit', 
                         'Nettonennleistung der Einheit', 'Inbetriebnahmedatum der Einheit', 'Registrierungsdatum der Einheit',
                         'Postleitzahl', 'Straße']

print('Lädt...')

# Excel einlesen
if run_in_vs_code:
    path_folder = path_folder / (file_1_name + '.xlsx')
else:
    path_folder = file_1_name + '.xlsx' 
excel_file_1 = pd.read_excel(path_folder, file_1_sheet_name, header=file_1_header)

# Daten bearbeiten
excel_file_1['Straße'] = excel_file_1['Straße'] + ' ' + excel_file_1['Hausnummer'].fillna('').astype(str)
excel_file_1 = excel_file_1[file_1_needed_columns]  # auskommentieren, wenn alle Spalten angezeigt werden sollen
excel_file_1 = excel_file_1[excel_file_1['Betriebsstatus'] == 'In Betrieb']
excel_file_1 = excel_file_1[excel_file_1['Systemstatus'] == 'Aktiviert']
excel_file_1['Inbetriebnahmedatum der Einheit'] = excel_file_1['Inbetriebnahmedatum der Einheit'].dt.year

df_all_pv = excel_file_1[excel_file_1['Energieträger'] == 'Solare Strahlungsenergie']
df_all_storage = excel_file_1[excel_file_1['Energieträger'] == 'Speicher']

# Dataframe gesamt installierte PV nach Jahren
df_all_pv_power_per_year = (df_all_pv.groupby('Inbetriebnahmedatum der Einheit')['Bruttoleistung der Einheit'].sum().reset_index())
df_all_pv_power_per_year = df_all_pv_power_per_year.rename(columns={'Bruttoleistung der Einheit': 'Zuwachs pro Jahr', 'Inbetriebnahmedatum der Einheit': 'Jahr'})
df_all_pv_power_per_year['Summe'] = (df_all_pv_power_per_year['Zuwachs pro Jahr'].cumsum())

# DataFrame gesamt PV nach Größe
sum_smaller_25 = df_all_pv.loc[df_all_pv['Bruttoleistung der Einheit'] < 25, 'Bruttoleistung der Einheit'].sum()
sum_25_until_100 = df_all_pv.loc[(df_all_pv['Bruttoleistung der Einheit'] >= 25) & (df_all_pv['Bruttoleistung der Einheit'] <= 100), 'Bruttoleistung der Einheit'].sum()
sum_bigger_100 = df_all_pv.loc[df_all_pv['Bruttoleistung der Einheit'] > 100, 'Bruttoleistung der Einheit'].sum()
df_all_pv_power_per_cap = pd.DataFrame({'< 25 kW': [sum_smaller_25], '25 - 100 kW': [sum_25_until_100], '> 100 kW': [sum_bigger_100]})

# DataFrame Westnetz installierte PV nach Jahren
df_all_pv_west = df_all_pv[df_all_pv['Postleitzahl'] == 91056]
df_wn_pv_power_per_year = (df_all_pv_west.groupby('Inbetriebnahmedatum der Einheit')['Bruttoleistung der Einheit'].sum().reset_index())
df_wn_pv_power_per_year = df_wn_pv_power_per_year.rename(columns={'Bruttoleistung der Einheit': 'Zuwachs pro Jahr', 'Inbetriebnahmedatum der Einheit': 'Jahr'})
df_wn_pv_power_per_year['Summe'] = (df_wn_pv_power_per_year['Zuwachs pro Jahr'].cumsum())

# DataFrame Westnetz PV nach Größe
wn_sum_smaller_25 = df_all_pv_west.loc[df_all_pv_west['Bruttoleistung der Einheit'] < 25, 'Bruttoleistung der Einheit'].sum()
wn_sum_25_until_100 = df_all_pv_west.loc[(df_all_pv_west['Bruttoleistung der Einheit'] >= 25) & (df_all_pv_west['Bruttoleistung der Einheit'] <= 100), 'Bruttoleistung der Einheit'].sum()
wn_bigger_100 = df_all_pv_west.loc[df_all_pv_west['Bruttoleistung der Einheit'] > 100, 'Bruttoleistung der Einheit'].sum()
df_wn_pv_power_per_cap = pd.DataFrame({'< 25 kW': [wn_sum_smaller_25], '25 - 100 kW': [wn_sum_25_until_100], '> 100 kW': [wn_bigger_100]})

# DataFrame Speicher
df_all_storage_per_year = (df_all_storage.groupby('Inbetriebnahmedatum der Einheit')['Bruttoleistung der Einheit'].sum().reset_index())
df_all_storage_per_year = df_all_storage_per_year.rename(columns={'Bruttoleistung der Einheit': 'Zuwachs pro Jahr', 'Inbetriebnahmedatum der Einheit': 'Jahr'})
df_all_storage_per_year['Summe'] = (df_all_storage_per_year['Zuwachs pro Jahr'].cumsum())

# DataFrame Speicher nach Größe
df_all_storage['Bruttoleistung der Einheit gerundet'] = df_all_storage['Bruttoleistung der Einheit'].round(0).astype(int)
max_storage_capacity = df_all_storage['Bruttoleistung der Einheit gerundet'].max()
data = []
for i in range(max_storage_capacity + 1):
    sum = df_all_storage.loc[df_all_storage['Bruttoleistung der Einheit gerundet'] == i, 'Bruttoleistung der Einheit'].sum()
    data.append(sum)
df_all_storage_per_cap = pd.DataFrame({'Kapazität in kW': list(range(max_storage_capacity + 1)), 'Summe': data})

# Excel-Datei kreiren
file_name = 'Auswertung_MaStR'
date_today = datetime.now().strftime('%Y%m%d') + '_'
if run_in_vs_code:
    file_path = path_folder / (date_today + file_name + '.xlsx')
else:
    file_path = date_today + file_name + '.xlsx'

with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    #df_all_pv.to_excel(writer, sheet_name='Gesamte PVs', index=False)
    df_all_pv_power_per_year.to_excel(writer, sheet_name='PV-Leistung pro Jahr', index=False)
    df_all_pv_power_per_cap.to_excel(writer, sheet_name='PV Größenverteilung', index=False)
    df_wn_pv_power_per_year.to_excel(writer, sheet_name='Westnetz PV-Leistung pro Jahr', index=False)
    df_wn_pv_power_per_cap.to_excel(writer, sheet_name='Westnetz PV Größenverteilung', index=False)
    df_all_storage_per_year.to_excel(writer, sheet_name='Speicherleistung pro Jahr', index=False)
    df_all_storage_per_cap.to_excel(writer, sheet_name='Speicher Größenverteilung', index=False)
    #df_all_storage.to_excel(writer, sheet_name='Gesamte Speicher', index=False)

sheet_names = ['PV-Leistung pro Jahr', 'PV Größenverteilung', 'Westnetz PV-Leistung pro Jahr', 'Westnetz PV Größenverteilung', 'Speicherleistung pro Jahr', 
               'Speicher Größenverteilung']
for sheet_name in sheet_names: automatic_column_width(file_path, sheet_name=sheet_name)

print('Datei wurde erfolgreich erstellt!')

