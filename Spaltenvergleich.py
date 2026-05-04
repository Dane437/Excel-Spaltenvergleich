import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import FormulaRule
import io
from pathlib import Path
from datetime import datetime
import matplotlib.pyplot as plt
from utils import highlight_differences, de_sort_key, create_district_sheets

# Hausanschlüsse Analyse

# !!! Überpüfen: !!!
run_in_vs_code = True
print('Lädt...')
path_folder = Path(r'C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp\Hausanschlüsse')    #Dateipfad überprüfen
file_1_name = 'Alle Anschlüsse'
file_1_sheet_name = 'Hausanschluss'
file_1_merging_column = 'Absicherung'
file_1_header = 4   # Spaltenname in Zeile 5
file_1_needed_columns = ['Absicherung', 'Teilnetz', 'Ortsteil']

file_2_name = 'Zuordnungstabelle'
file_2_sheet_name = 'List1'
file_2_merging_column = 'Sicherungseinsätze'
file_2_header = 0
file_2_needed_columns = ['Sicherungseinsätze', 'Leistung']

# Excel einlesen
if run_in_vs_code:
    file_1_path = path_folder / (file_1_name + '.xlsx')
    file_2_path = path_folder / (file_2_name + '.xlsx')
else:
    file_1_path = file_1_name + '.xlsx'
    file_2_path = file_2_name + '.xlsx' 
df_Hausanschluss = pd.read_excel(file_1_path, sheet_name=file_1_sheet_name, header=file_1_header)
df_Zuordnungstabelle = pd.read_excel(file_2_path, sheet_name=file_2_sheet_name, header=file_2_header)

# Daten individuel bearbeiten
df_Hausanschluss = df_Hausanschluss[file_1_needed_columns]
df_Zuordnungstabelle = df_Zuordnungstabelle[file_2_needed_columns]
df_Zuordnungstabelle['Leistung'] = df_Zuordnungstabelle['Leistung'].str.replace(',', '.', regex=False)  # Komma → Punkt
df_Zuordnungstabelle['Leistung'] = pd.to_numeric(df_Zuordnungstabelle['Leistung'], errors='coerce')
df_Hausanschluss['Absicherung'] = df_Hausanschluss['Absicherung'].str.extract(r'\s(.+)')
df_Hausanschluss['Absicherung'] = df_Hausanschluss['Absicherung'].replace('Unterlagen vorh.', None)

# Datensätze zusammenführen
df_merged_files: pd.DataFrame = pd.merge(df_Hausanschluss, df_Zuordnungstabelle, left_on= file_1_merging_column, right_on= file_2_merging_column, how='left', indicator=False)
df_missing_values = df_merged_files[df_merged_files['Sicherungseinsätze'].isna()]
df_only = df_missing_values.drop_duplicates(subset=['Absicherung'])

#anzahl = df_missing_values['Absicherung'].isna().sum().astype('int64')
#print(anzahl)

list_districts = df_merged_files.loc[df_merged_files['Ortsteil'] != '-', 'Ortsteil'].dropna().unique().tolist() # Liste aller Ortsteile

# Excel-Datei kreiren
list_values_all_districts =  []
date_today = datetime.now().strftime('%Y%m%d') + '_'
file_name = 'Hausanschluss_Auswertung'
if run_in_vs_code:
    file_1_path = path_folder / (date_today + file_name + '.xlsx')
else:
    file_1_path = date_today + file_name + '.xlsx'
sheet_name = 'Gesamt'
with pd.ExcelWriter(file_1_path, engine='openpyxl') as writer:
    #df_merged_files.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
    list_values_district = create_district_sheets(writer, df_merged_files, district_name=sheet_name)
    list_values_all_districts.append(list_values_district)
    df_missing_values.to_excel(writer, sheet_name='Fehlender Sicherungswert', index=False, startrow=0)
    df_only.to_excel(writer, sheet_name='Doppelte entfernt', index=False, startrow=0)
    for district in list_districts:
        list_values_district = create_district_sheets(writer, df_merged_files, district_name=district)
        list_values_all_districts.append(list_values_district)
    df_values_all_districts = pd.DataFrame(list_values_all_districts, columns=['Ortsteil', 'Anzahl kein Sicherungswert', 'Prozent kein Sicherungswert', 'Anzahl Gesamt'])
    df_values_all_districts.to_excel(writer, sheet_name='Vergleich Ortsteile', index=False, startrow=0)

print('Datei wurde erfolgreich gespeichert!')

