import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import FormulaRule
import io
from pathlib import Path
import matplotlib.pyplot as plt
from utils import highlight_differences, de_sort_key, create_differences_sheet

# Vergleich PV-Werte MaStR und EEGDB Datenbank

# !!! Überpüfen: !!!
path_folder = Path(r'C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp\MaStR_EEGDB')    #Dateipfad überprüfen
file_1_name = '20260316_Auswertung_MaStR'
file_1_sheet_name = 'Gesamte PVs'
file_1_merging_column = 'Straße'
file_1_header = 0
file_1_needed_columns = ['Anzeige-Name der Einheit', 'Bruttoleistung der Einheit', 'Nettonennleistung der Einheit', 
                         'Inbetriebnahmedatum der Einheit', 'Straße']

file_2_name = '20260316_ESTW-EEG-DB'
file_2_sheet_name = 'Tabelle1'
file_2_merging_column = 'Strasse der Anlage'
file_2_header = 0
file_2_needed_columns = ['Strasse der Anlage', 'PLZ der Anlage', 'Objekt Info', 'LeistungErzeugung']

compare_columns = [('Bruttoleistung der Einheit', 'LeistungErzeugung')]

# Excel einlesen
excel_file_1 = pd.read_excel(path_folder / (file_1_name +'.xlsx'), sheet_name=file_1_sheet_name, header=file_1_header)
excel_file_1 = excel_file_1[file_1_needed_columns]
excel_file_2 = pd.read_excel(path_folder / (file_2_name +'.xlsx'), sheet_name=file_2_sheet_name, header=file_2_header)
excel_file_2 = excel_file_2[excel_file_2['Art der Erzeugung'].isin(['PV', 'PV-Balkonanlage', 'PV-Freiflächeanlage'])]
excel_file_2 = excel_file_2[file_2_needed_columns]

# Daten individuel bearbeiten
excel_file_1.rename(columns={'AlterName': 'NeuerName'}, inplace=True)

# Anzahl Spalten bestimmen
nb_columns_excel_file_1 = excel_file_1.shape[1]
nb_columns_excel_file_2 = excel_file_2.shape[1]
nb_columns = nb_columns_excel_file_1 + nb_columns_excel_file_2

# Datensätze zusammenführen
merged_files: pd.DataFrame = pd.merge(excel_file_1, excel_file_2, left_on= file_1_merging_column, right_on= file_2_merging_column, how='left', indicator=False)
#merged_files.to_excel('Gesamt.xlsx', index=False)

# Excel-Datei kreiren
file_name = 'Unterschiede_' + file_1_name.split('.')[0] + '.xlsx'
file_path = path_folder / file_name
sheet_name='Vergleich'
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    merged_files.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)

    # Zugriff auf das Workbook und Worksheet
    workbook = writer.book
    worksheet = writer.sheets['Vergleich']
    # Überschrift erstellen und formatieren
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nb_columns_excel_file_1)
    worksheet.merge_cells(start_row=1, start_column=nb_columns_excel_file_1+1, end_row=1, end_column=nb_columns)
    cell_left = worksheet.cell(row=1, column=1)
    cell_right = worksheet.cell(row=1, column=nb_columns_excel_file_1+1)
    cell_left.value = file_1_name.split('.')[0]
    cell_right.value = file_2_name.split('.')[0]
    center_align = Alignment(horizontal='center', vertical='center')
    cell_left.fill = PatternFill(start_color='C6EFCE', fill_type='solid')   # grüne Formatierung
    cell_left.font = Font(bold=True, size=16)
    cell_left.alignment = center_align
    cell_right.fill = PatternFill(start_color='BDD7EE', fill_type='solid') # blaue Foramtierung
    cell_right.font = Font(bold=True, size=16)
    cell_right.alignment = center_align
    for col in range(1, nb_columns + 1):
        cell = worksheet.cell(row=2, column=col)
        cell.fill = PatternFill(start_color='E7E6E6', fill_type='solid')    # graue Formatierung
        cell.font = Font(bold=True)
    # Trennungslinie erzeugen
    max_row = worksheet.max_row
    for row in range(1, max_row + 1):
        cell = worksheet.cell(row=row, column=nb_columns_excel_file_1)
        cell.border = Border(right=Side(style='thick'))
    # Automatisch Spaltenbreite
    for column in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2  # etwas Puffer
        if adjusted_width > 30 :    # Maximalbreite setzen
            adjusted_width = 30
        worksheet.column_dimensions[column_letter].width = adjusted_width

header_row = 2  # in dieser Reihe stehen meine Spaltenüberschriften

# Temporäres Sheet erstellen, wird für die Funktion create_differences_sheet benötigt
workbook = load_workbook(file_path)
worksheet = workbook[sheet_name]
worksheet_copy = workbook.copy_worksheet(worksheet)
worksheet_copy.title ='Temp'
workbook.save(file_path)

# Unterschiede zwischen den Spalten grafisch hervorheben
all_rows_different = []
for idx, compare_column in enumerate(compare_columns):
    rows_different = highlight_differences(file_path, sheet_name, compare_column[0], compare_column[1], header_row, valid_combinations=None)
    all_rows_different.append(rows_different)

# # Schnittmenge aller Listen finden
# common_numbers = set(all_rows_different[0])
# for row in all_rows_different[1:]:
#     common_numbers &= set(row)
# all_rows_different = [[x for x in row if x not in common_numbers] for row in all_rows_different]    # Zahlen aus jeder Liste entfernen
# all_rows_different.append(list(common_numbers)) # Gemeinsame Zahlen als neue Liste anhängen

# Sheets mit den Unterschieden erstellen
for idx, compare_column in enumerate(compare_columns):
    create_differences_sheet(file_path, compare_column, 'Leistung', all_rows_different[idx], header_row)
# create_differences_sheet(file_path, ('leer', 'leer'), 'leer', all_rows_different[-1], header_row)

# Temporäres Sheet löschen
workbook = load_workbook(file_path)
del workbook['Temp']
workbook.save(file_path)

print('Datei wurde erfolgreich gespeichert!')

