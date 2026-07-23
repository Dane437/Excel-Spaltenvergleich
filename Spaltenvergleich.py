import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image
from openpyxl.formatting.rule import FormulaRule
import io
from pathlib import Path
import matplotlib.pyplot as plt
from utils import highlight_differences, de_sort_key, create_differences_sheet, bereinige_hausnummer

# Datenvergleich zwischen zwei Excel-Dateien, hier Trafoliste K3V und Trafoliste aus dem GIS

# !!! Überpüfen: !!!
path_folder = Path(r"C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp\Ladesäulen_Vergleich")    #Dateipfad überprüfen
file_1_name = 'Ladesaeulen_SharePoint'
file_1_header = 0 # Spaltenname in Zeile 1
file_1_sheet_name = 'List1'
file_1_sorting_column = 'Straßenname'
file_1_needed_columns = [file_1_sorting_column, 'Zugang zur Ladeeinrichtung', 'Gesamtleistung Ladepunkte', 'Inbetriebssetzungs-datum', 'Bearbeitungsstatus']

file_2_name = "Ladesaeulenregister_BNetzA"
file_2_header = 0 # Spaltenname in Zeile 1
file_2_sheet_name = 'Tabelle1'
file_2_sorting_column = 'Straße'
file_2_needed_columns = [file_2_sorting_column, 'Nennleistung Ladeeinrichtung [kW]', 'Inbetriebnahmedatum', 'Status']

compare_columns = [('Inbetriebssetzungs-datum', 'Inbetriebnahmedatum'),
                   ('Gesamtleistung Ladepunkte', 'Nennleistung Ladeeinrichtung [kW]'),
                   ('Bearbeitungsstatus', 'Status')]

# Excel einlesen
excel_file_1 = pd.read_excel(path_folder / (file_1_name + '.xlsx'), sheet_name=file_1_sheet_name, header=file_1_header)
excel_file_1['Hs.- bzw. Flur-Nr.'] = excel_file_1['Hs.- bzw. Flur-Nr.'].apply(bereinige_hausnummer)
excel_file_1['Straßenname'] = excel_file_1['Straßenname'] + ' ' + excel_file_1['Hs.- bzw. Flur-Nr.'].fillna('').astype(str)
excel_file_1["Straßenname"] = excel_file_1["Straßenname"].str.lower()
excel_file_1 = excel_file_1[file_1_needed_columns]  # auskommentieren, wenn alle Spalten angezeigt werden sollen
excel_file_1 = excel_file_1[excel_file_1['Zugang zur Ladeeinrichtung'].isin(['öffentlich', 'privat - öffentlich'])]  # nur öffentlich
excel_file_1 = excel_file_1[~excel_file_1['Bearbeitungsstatus'].isin(['außer Betrieb'])]

excel_file_2 = pd.read_excel(path_folder / (file_2_name + '.xlsx'), sheet_name=file_2_sheet_name, header=file_2_header)
excel_file_2["Straße"] = excel_file_2["Straße"].str.replace(r"str\.(?=\s|$)", "straße", regex=True, case=False)
excel_file_2["Straße"] = excel_file_2["Straße"].str.lower()
excel_file_2['Straße'] = excel_file_2['Straße'] + ' ' + excel_file_2['Hausnummer'].fillna('').astype(str)
excel_file_2 = excel_file_2[file_2_needed_columns] # auskommentieren, wenn alle Spalten angezeigt werden sollen

# Anzahl Spalten bestimmen
nb_columns_excel_file_1 = excel_file_1.shape[1]
nb_columns_excel_file_2 = excel_file_2.shape[1]
nb_columns = nb_columns_excel_file_1 + nb_columns_excel_file_2

# Hilfsindex, damit wenn in beiden Tabellen 2 Zeilen mit dem gleichen Sortierschlüssel existieren, keine vier Zeilen entstehen
#excel_file_1['idx'] = excel_file_1.groupby(file_1_sorting_column).cumcount()
#excel_file_2['idx'] = excel_file_2.groupby(file_2_sorting_column).cumcount() 

#excel_file_2_merged = pd.merge(excel_file_2, excel_Netzleitzahlen[['NKZ', 'Stationstyp']], left_on='Kurzname', right_on='NKZ', how='left')

# Datensätze zusammenführen
#merged_files: pd.DataFrame = pd.merge(excel_file_1, excel_file_2_merged, left_on= [file_1_sorting_column, 'idx'], right_on= [file_2_sorting_column, 'idx'], how='outer', indicator=False)
#merged_files = merged_files.drop(columns='idx') 

#merged_files = merged_files[(merged_files['Typ (Anlage)'].isin(['Netzstation', 'Schalthaus', '']) | merged_files['Typ (Anlage)'].isna()) & 
                            #(merged_files['Stationstyp'].isin(['Netzstation', 'Schalthaus', '']) | merged_files['Stationstyp'].isna())]  

merged_files: pd.DataFrame = pd.merge(excel_file_1, excel_file_2, left_on= file_1_sorting_column, right_on= file_2_sorting_column, how='outer', indicator=False)

# Excel-Datei kreiren
file_name = 'Auswertung_Ladesäulen.xlsx'
file_path = path_folder / file_name
sheet_name = 'Vergleich'
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    merged_files.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)

    # Zugriff auf das Workbook und Worksheet
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

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
    rows_different = highlight_differences(file_path, sheet_name, column_1, column_2, header_row=2, valid_combinations=None, empty_is_difference=False)
#header_row = 2  # in dieser Reihe stehen meine Spaltenüberschriften

# # Temporäres Sheet erstellen, wird für die Funktion create_differences_sheet benötigt
# workbook = load_workbook(file_path)
# worksheet = workbook[sheet_name]
# worksheet_copy = workbook.copy_worksheet(worksheet)
# worksheet_copy.title ='Temp'
# workbook.save(file_path)

# #create_differences_sheet(file_path, rows_different, header_row)

# # Temporäres Sheet löschen
# workbook = load_workbook(file_path)
# del workbook['Temp']
# workbook.save(file_path)

print("Datei wurde erfolgreich gespeichert!")

