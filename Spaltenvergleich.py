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

# Datenvergleich zwischen Netzleitzahlen.xlsm und K3V-Export.xlsx

# !!! Überpüfen: !!!
path_folder = Path(r'C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp')    #Dateipfad eingeben
file_K3V_name = 'Netzstationen_K3V' # Dateiname eingeben
file_K3V_header = 2 # Headergröße überprüfen (2 = Spaltenüberschriften sind in Zeile 3)
column_name_K3V = ['Name', 'Nummer', 'Betriebsstatus', 'Inbetriebnahme', 'Teilnetz'] # Spaltennamen eingeben (Spalten nach dem letzen gebrauchten Element können weggelassen werden)
compare_columns = [('Stationsname', column_name_K3V[0]),    # zu vergleichende Spalten eingeben
                    ('NKZ', 'Nummer NKZ'),
                    ('NLZ', 'Nummer NLZ'),
                    ('Netz', 'Teilnetz')]
valid_combinations = [('Nord', 'Nordnetz'), # gültige Kombinationen eingeben 
                      ('Ost', 'Ostnetz'),
                      ('Süd', 'Südnetz'),
                      ('West', 'Westnetz'),
                      ('Südost', 'Südostnetz')]

# Pfad zu der Netzleitzahlendatei
file_Netzleitzahlen = Path(r'B:\# N-Gemeinsam\Trafostationsliste\Netzleitzahlen_Makro.xlsm')

# Excel einlesen
excel_Netzleitzahlen = pd.read_excel(file_Netzleitzahlen, sheet_name='Stationsliste')
excel_Netzleitzahlen['NLZ'] = pd.to_numeric(excel_Netzleitzahlen['NLZ'], errors='coerce')
excel_K3_Export = pd.read_excel(path_folder/(file_K3V_name +'.xlsx'), sheet_name='Tabelle1', header=file_K3V_header)

# Daten aufbereiten
excel_K3_Export = excel_K3_Export[~excel_K3_Export[column_name_K3V[0]].astype(str).str.contains(r'\(Liegenschaft\)', na=False)] # Zeilen löschen in denen 'Liegenschaft' vorkommt
excel_K3_Export[['Nummer NLZ', 'Nummer NKZ']] = (excel_K3_Export[column_name_K3V[1]].str.extract(r'.*?(\d{3})\s*/\s*(.+)')) # NLZ und NKZ in eigenen Spalten
excel_K3_Export['Nummer NLZ'] = pd.to_numeric(excel_K3_Export['Nummer NLZ'], errors='coerce') # NLZ zu Zahl konvertieren

# Spalten löschen die nicht angezeigt werden sollen
excel_K3_Export = excel_K3_Export.drop(columns=[column_name_K3V[1], 'Strukturelement']) # Überprüfen ob das passt

# Anzahl Spalten bestimmen
nb_columns_excel_K3_Export = excel_K3_Export.shape[1]
nb_columns_excel_Netzleitzahlen = excel_Netzleitzahlen.shape[1]
nb_columns = nb_columns_excel_K3_Export + nb_columns_excel_Netzleitzahlen

# Datensätze zusammenführen
merge_adress: pd.DataFrame = pd.merge(excel_Netzleitzahlen, excel_K3_Export, left_on='Stationsname', right_on=column_name_K3V[0], how='outer', indicator=True)
not_matched_l = merge_adress[merge_adress['_merge']=='left_only'].drop(columns=excel_K3_Export.columns)
not_matched_r = merge_adress[merge_adress['_merge']=='right_only'].drop(columns=excel_Netzleitzahlen.columns)
merge_not_matched = pd.merge(not_matched_l, not_matched_r, left_on='NLZ', right_on='Nummer NLZ', how='outer', indicator=False)
matched = merge_adress[merge_adress['_merge'] == 'both'].drop(columns=['_merge'])
final_data = pd.concat([matched, merge_not_matched], ignore_index=True)

# Sortieren nach Stationsname, wenn nicht vorhanden nach K3V
final_data['_sort_key'] = (final_data['Stationsname'].fillna(final_data[column_name_K3V[0]]).astype(str).str.strip().apply(de_sort_key))
final_data = (final_data.sort_values(by='_sort_key').drop(columns='_sort_key').reset_index(drop=True))
final_data = final_data.loc[:, ~final_data.columns.str.contains('_merge')]  # alle _merge Spalten löschen

# Daten für Netzstationen Statistik extrahieren
copy_merge_final = final_data.copy()
copy_merge_final[column_name_K3V[3]] = pd.to_datetime(copy_merge_final[column_name_K3V[3]], dayfirst=True, errors='coerce')
copy_merge_final[column_name_K3V[3]] = copy_merge_final[column_name_K3V[3]].dt.year.astype('Int64')
only_netzstationen = copy_merge_final[copy_merge_final['Stationstyp'].str.contains('Netz', na=False)]
data_netzstationen_per_year = only_netzstationen.groupby(column_name_K3V[3]).size()

# Diagramm erzeugen
plt.figure(figsize=(15, 7))
data_netzstationen_per_year.plot(kind='bar')
plt.title('Netzstationen pro Jahr')
plt.xlabel('Jahr')
plt.ylabel('Anzahl')
plt.xticks(rotation=90)
#plt.show()

# Diagramm speichern
img_data = io.BytesIO()
plt.savefig(img_data, format='png')
plt.close() # schließt die Figure, um Speicher zu sparen

# Excel-Datei kreiren
file_name = 'Unterschiede_K3V_Netzleitzahlen.xlsx'
file_path = path_folder / 'Ergebnisse' / file_name
sheet_name = 'Vergleich'
with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    final_data.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)

    # Zugriff auf das Workbook und Worksheet
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # Überschrift erstellen und formatieren
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nb_columns_excel_Netzleitzahlen)
    worksheet.merge_cells(start_row=1, start_column=nb_columns_excel_Netzleitzahlen+1, end_row=1, end_column=nb_columns)
    cell_left = worksheet.cell(row=1, column=1)
    cell_right = worksheet.cell(row=1, column=nb_columns_excel_Netzleitzahlen+1)
    cell_left.value = 'Netzleitzahlen'
    cell_right.value = 'K3V'
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

    header_row = 2  # in dieser Reihe stehen meine Spaltenüberschriften

    # Trennungslinie erzeugen
    max_row = worksheet.max_row
    for row in range(1, max_row + 1):
        cell = worksheet.cell(row=row, column=nb_columns_excel_Netzleitzahlen)
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

    # Zweites Sheet für Diagramm
    worksheet = workbook.create_sheet('Diagramm')
    
    # Das Bild aus dem Speicher laden und in das Sheet einfügen
    img_data.seek(0)
    img = Image(img_data)
    worksheet.add_image(img, 'A1')
    

# Temporäres Sheet erstellen, wird für die Funktion create_differences_sheet benötigt
workbook = load_workbook(file_path)
worksheet = workbook[sheet_name]
worksheet_copy = workbook.copy_worksheet(worksheet)
worksheet_copy.title ='Temp'
workbook.save(file_path)

# Unterschiede zwischen den Spalten grafisch hervorheben
all_rows_different = []
for idx, compare_column in enumerate(compare_columns):
    rows_different = highlight_differences(file_path, sheet_name, compare_column[0], compare_column[1], header_row, valid_combinations)
    all_rows_different.append(rows_different)

# Schnittmenge aller Listen finden
common_numbers = set(all_rows_different[0])
for row in all_rows_different[1:]:
    common_numbers &= set(row)
all_rows_different = [[x for x in row if x not in common_numbers] for row in all_rows_different]    # Zahlen aus jeder Liste entfernen
all_rows_different.append(list(common_numbers)) # Gemeinsame Zahlen als neue Liste anhängen

# Sheets mit den Unterschieden erstellen
for idx, compare_column in enumerate(compare_columns):
    create_differences_sheet(file_path, compare_column, all_rows_different[idx], header_row)
create_differences_sheet(file_path, ('leer', 'leer'), all_rows_different[-1], header_row)

# Temporäres Sheet löschen
workbook = load_workbook(file_path)
del workbook['Temp']
workbook.save(file_path)

print('Datei wurde erfolgreich gespeichert!')

