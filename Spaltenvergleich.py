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

# Mergen von Zeilen zweier beliebiger Excel-Dateien

# !!! Überpüfen: !!!
path_folder = Path(r"C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp\Auslastung ONS")    #Dateipfad überprüfen
file_1_name = 'Daten_Acron'
file_1_sheet_name = 'List1'
file_1_header = 0   # 0 -> Überschrift in erster Zeile

file_2_name = 'Trafoliste_mit_N'
file_2_sheet_name = 'Tabelle1'
file_2_merging_column = 'Nummer NLZ'
file_2_header = 2
file_2_needed_columns = ['Nummer NLZ', 'Anlage', 'ZONE', 'SN [kVA]']

# Excel einlesen
excel_file_1 = pd.read_excel(path_folder / (file_1_name +'.xlsx'), sheet_name=file_1_sheet_name, header=file_1_header)
excel_file_2 = pd.read_excel(path_folder / (file_2_name +'.xlsx'), sheet_name=file_2_sheet_name, header=file_2_header)
excel_file_2[['Nummer NLZ', 'Nummer NKZ']] = (excel_file_2['Nummer'].str.extract(r'.*?(\d{3})\s*/\s*(.+)')) # NLZ und NKZ in eigenen Spalten
excel_file_2 = excel_file_2[file_2_needed_columns]

# Daten individuel bearbeiten
new_df = pd.DataFrame(excel_file_1.iloc[:, 1:].columns, columns=['NLZ'])
new_df['NLZ'] = new_df['NLZ'].str.extract(r'TS\s*(\d+)')

# Maximum
max_indices_list = excel_file_1.iloc[:, 1:].idxmax().tolist()   # [:, 1:] ab zweiter Spalte
new_df['Datum Maximum 2025'] = excel_file_1.iloc[max_indices_list, 0].values
max_values = []
for idx, indices in enumerate(max_indices_list):
    max_values.append(excel_file_1.iloc[indices, idx + 1])
new_df['Maximalwert 2025'] = max_values

# Minimum
min_indices_list = excel_file_1.iloc[:, 1:].idxmin().tolist()
new_df['Datum Minimum 2025'] = excel_file_1.iloc[min_indices_list, 0].values
min_values = []
for idx, indices in enumerate(min_indices_list):
    min_values.append(excel_file_1.iloc[indices, idx + 1])
new_df['Minimalwert 2025'] = min_values

# Datensätze zusammenführen
merged_files: pd.DataFrame = pd.merge(new_df, excel_file_2, left_on= 'NLZ', right_on= file_2_merging_column, how='left', indicator=False)
merged_files = merged_files.drop(columns=['Nummer NLZ'])
merged_files = merged_files.reindex(columns=['NLZ', 'Anlage', 'ZONE', 'SN [kVA]', 'Datum Maximum 2025', 'Maximalwert 2025', 'Datum Minimum 2025', 'Minimalwert 2025'])

# Differenz
merged_files['Differenz Max'] = merged_files['SN [kVA]'] - merged_files['Maximalwert 2025']
merged_files['Differenz Min'] = merged_files['SN [kVA]'] - merged_files['Minimalwert 2025'].abs()
merged_files['Differenz zu SN'] = merged_files[['Differenz Max', 'Differenz Min']].min(axis=1)
merged_files = merged_files.drop(columns=['Differenz Max', 'Differenz Min'])

# Sortieren
merged_files = merged_files.sort_values(by='Differenz zu SN')

# Excel-Datei kreiren
file_name = 'Auswertung_ONS.xlsx'
file_path = path_folder / file_name
sheet_name = 'Auswertung'
header_row = 1 # Excel Zeile wo die Spaltenüberschriften stehen sollen
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    merged_files.to_excel(writer, sheet_name=sheet_name, index=False, startrow=header_row - 1)

    # Zugriff auf das Workbook und Worksheet
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    max_row = worksheet.max_row
    max_column = worksheet.max_column

    for col in range(1, max_column + 1):
        cell = worksheet.cell(row=header_row, column=col)
        cell.fill = PatternFill(start_color="E7E6E6", fill_type="solid")    # graue Formatierung
        cell.font = Font(bold=True)

    # Trennungslinie erzeugen
    for row in range(header_row, max_row + 1):
        worksheet.cell(row=row, column=4).border = Border(right=Side(style='thick'))
        worksheet.cell(row=row, column=6).border = Border(right=Side(style='thick'))
        worksheet.cell(row=row, column=8).border = Border(right=Side(style='thick'))

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

print("Datei wurde erfolgreich gespeichert!")

