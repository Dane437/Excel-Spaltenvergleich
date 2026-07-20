import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import FormulaRule
import io
from pathlib import Path
import matplotlib.pyplot as plt
from datetime import datetime
from utils import highlight_differences, de_sort_key, automatic_column_width, bereinige_hausnummer

# Dieses Script erstellt aus dem MaStR die Werte wie viel PV-Leistung für die einzelnen Jahre zugebaut wurde

# !!! Überpüfen: !!!
run_in_vs_code = True
path_folder = Path(r'C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp\MaStR\MaStR_Adress_Bereinigung')
file_1_name = 'MaStR_Anlagenbetreiber'
file_1_sheet_name = 'List1'
file_1_merging_column = 'Straße'
file_1_header = 0
file_1_needed_columns = ['Betriebsstatus', 'Systemstatus', 'Energieträger', 'Bruttoleistung der Einheit', 'Inbetriebnahmedatum der Einheit', 'Straße', 'Name des Anlagenbetreibers', 'MaStR-Nr. der Einheit']

file_2_name = 'Strom_Verbrauchsstatistik_2018-2024'
file_2_sheet_name = 'SLP_Zusammengeführt'
file_2_merging_column = 'Ad'
file_2_header = 0
file_2_needed_columns = ['Ad']

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
excel_file_1['Hausnummer'] = excel_file_1['Hausnummer'].apply(bereinige_hausnummer)
#excel_file_1["Hausnummer"] = (excel_file_1["Hausnummer"].astype(str).str.split(r"[-/]").str[0])
excel_file_1["Straße"] = (excel_file_1["Straße"].str[0].str.upper() +excel_file_1["Straße"].str[1:])
excel_file_1["Straße"] = (excel_file_1["Straße"].str.replace(r"strasse$", "straße", case=False, regex=True))
excel_file_1['Straße'] = excel_file_1['Straße'] + ' ' + excel_file_1['Hausnummer'].fillna('').astype(str)
#excel_file_1 = excel_file_1[file_1_needed_columns]  # auskommentieren, wenn alle Spalten angezeigt werden sollen
excel_file_2 = excel_file_2[file_2_needed_columns]
#excel_file_1 = excel_file_1[excel_file_1['Betriebsstatus'] == 'In Betrieb']
#excel_file_1 = excel_file_1[excel_file_1['Systemstatus'] == 'Aktiviert']
#excel_file_1['Inbetriebnahmedatum der Einheit'] = excel_file_1['Inbetriebnahmedatum der Einheit'].dt.year

#excel_file_1 = excel_file_1[excel_file_1['Energieträger'] == 'Solare Strahlungsenergie']
# df_all_storage = excel_file_1[excel_file_1['Energieträger'] == 'Speicher']

#agg = {col: "first" for col in excel_file_1.columns}
#agg["Bruttoleistung der Einheit"] = "sum"

#excel_file_1 = (
#    excel_file_1
#    .groupby("Straße", as_index=False, sort=False)
#    .agg(agg)
#)

#merged_files: pd.DataFrame = pd.merge(excel_file_1, excel_file_2, left_on= file_1_merging_column, right_on= file_2_merging_column, how='outer', indicator=False)

# Excel-Datei kreiren
file_name = 'Anlagenbetreiber_Bereinigte_Adr'
date_today = datetime.now().strftime('%Y%m%d') + '_'
if run_in_vs_code:
    file_path_1 = path_folder / (date_today + file_name + '.xlsx')
else:
    file_path_1 = date_today + file_name + '.xlsx'

with pd.ExcelWriter(file_path_1, engine='openpyxl') as writer:
    excel_file_1.to_excel(writer, sheet_name='Bereinigte Adr', index=False)
    #merged_files.to_excel(writer, sheet_name='Zuordnung', index=False)

sheet_names = ['Bereinigte Adr']#, 'Zuordnung']
for sheet_name in sheet_names: automatic_column_width(file_path_1, sheet_name=sheet_name)

print('Datei wurde erfolgreich erstellt!')

