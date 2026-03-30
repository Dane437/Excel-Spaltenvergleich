import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import FormulaRule
import io
from pathlib import Path
import matplotlib.pyplot as plt
from utils import highlight_differences, de_sort_key

# Datenvergleich zwischen zwei Excel-Dateien, hier Trafoliste K3V und Trafoliste aus dem GIS

# !!! Überpüfen: !!!
path_folder = Path(r"C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp")    #Dateipfad überprüfen
file_1_name = 'Trafoliste_K3V.xlsx'
file_1_sheet_name = 'Tabelle1'
file_1_sorting_column = 'Nummer NKZ'
file_1_needed_columns = ['Nummer NKZ', 'Anlage', 'Typ (Anlage)', 'SN [kVA]', 'Betriebsstatus (Anlage)']

file_2_name = "Transformator_GIS.xlsx"
file_2_sheet_name = 'Transformator (MSP)'
file_2_sorting_column = 'Kurzname'
file_2_needed_columns = ['Kurzname', 'Name Station', 'Leistung (kVA)', 'Status']

compare_columns = [('SN [kVA]', 'Leistung (kVA)')]
                   #, ('Betriebsstatus (Anlage)', 'Status')]


# Excel einlesen
excel_file_1 = pd.read_excel(path_folder / file_1_name, sheet_name=file_1_sheet_name)
excel_file_1[['Nummer NLZ', 'Nummer NKZ']] = (excel_file_1['Nummer'].str.extract(r'.*?(\d{3})\s*/\s*(.+)')) # NLZ und NKZ in eigenen Spalten
excel_file_1 = excel_file_1[file_1_needed_columns]  # auskommentieren, wenn alle Spalten angezeigt werden sollen
excel_file_2 = pd.read_excel(path_folder / file_2_name, sheet_name=file_2_sheet_name)
excel_file_2 = excel_file_2[file_2_needed_columns] # auskommentieren, wenn alle Spalten angezeigt werden sollen


# Anzahl Spalten bestimmen
nb_columns_excel_file_1 = excel_file_1.shape[1]
nb_columns_excel_file_2 = excel_file_2.shape[1]
nb_columns = nb_columns_excel_file_1 + nb_columns_excel_file_2

# Hilfsindex, damit wenn in beiden Tabellen 2 Zeilen mit dem gleichen Sortierschlüssel existieren, keine vier Zeilen entstehen
excel_file_1['idx'] = excel_file_1.groupby('Nummer NKZ').cumcount()
excel_file_2['idx'] = excel_file_2.groupby('Kurzname').cumcount() 

# Datensätze zusammenführen
merged_files: pd.DataFrame = pd.merge(excel_file_1, excel_file_2, left_on= [file_1_sorting_column, 'idx'], right_on= [file_2_sorting_column, 'idx'], how='outer', indicator=False)
merged_files = merged_files.drop(columns='idx') 

# Excel-Datei kreiren
file_name = 'Unterschiede_' + file_1_name.split('.')[0] + '-' + file_2_name.split('.')[0] + '.xlsx'
with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    merged_files.to_excel(writer, sheet_name='Vergleich', index=False, startrow=1)

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
    center_align = Alignment(horizontal="center", vertical="center")

    cell_left.fill = PatternFill(start_color="C6EFCE", fill_type="solid")   # grüne Formatierung
    cell_left.font = Font(bold=True, size=16)
    cell_left.alignment = center_align

    cell_right.fill = PatternFill(start_color="BDD7EE", fill_type="solid") # blaue Foramtierung
    cell_right.font = Font(bold=True, size=16)
    cell_right.alignment = center_align

    for col in range(1, nb_columns + 1):
        cell = worksheet.cell(row=2, column=col)
        cell.fill = PatternFill(start_color="E7E6E6", fill_type="solid")    # graue Formatierung
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

# Unterschiede zwischen den Spalten grafisch hervorheben
for column_1, column_2 in compare_columns:
    highlight_differences(file_name, 'Vergleich', column_1, column_2, header_row=2, valid_combinations=None, empty_is_difference=False)

print("Datei wurde erfolgreich gespeichert!")

