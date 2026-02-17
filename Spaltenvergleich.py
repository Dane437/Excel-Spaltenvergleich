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

# Datenvergleich zwischen Netzleitzahlen.xlsm und GIS-Export.xlsx

# !!! Überpüfen: !!!
path_folder = Path(r"\\estw-01\Bereich-N\NG\NGE\Statistiken ESTW\20-kV-Stationen")  # Dateipfad überprüfen
path_Netzleitzahlen = path_folder / "20260210_Netzleitzahlen.xlsm"  # Dateiname überprüfen
path_GIS_Export = path_folder / "GIS-Export_Station_20260209.xlsx"  # Spaltennamen überpüfen

# Pfad zu der Netzleitzahlendatei
path_Netzleitzahlen = Path(r"B:\# N-Gemeinsam\Trafostationsliste\Netzleitzahlen_Makro.xlsm")

# Excel einlesen
excel_Netzleitzahlen = pd.read_excel(path_Netzleitzahlen, sheet_name="Stationsliste")
excel_Netzleitzahlen['NLZ'] = pd.to_numeric(excel_Netzleitzahlen['NLZ'], errors='coerce')
excel_GIS_Export = pd.read_excel(path_GIS_Export, sheet_name="Station", header=4) #header=4: erste vier Zeilen werden ignoriert, Zeile 5 in Spaltenname

# Daten aufbereiten
excel_GIS_Export['Stations ID'] = pd.to_numeric(excel_GIS_Export['Stations ID'], errors='coerce')
excel_GIS_Export = excel_GIS_Export[['Funktion', 'Teilnetz', 'Kurzname', 'Stations ID', 'Name', 'Baujahr']]
excel_Netzleitzahlen['Netz'] = excel_Netzleitzahlen['Netz'] + "netz" 

# Anzahl Spalten bestimmen
nb_columns_excel_Export = excel_GIS_Export.shape[1]
nb_columns_excel_Netzleitzahlen = excel_Netzleitzahlen.shape[1]
nb_columns = nb_columns_excel_Export + nb_columns_excel_Netzleitzahlen

# Datensätze zusammenführen
merge_nkz: pd.DataFrame = pd.merge(excel_Netzleitzahlen, excel_GIS_Export, left_on='NKZ', right_on='Kurzname', how='outer', indicator=True)
not_matched_l = merge_nkz[merge_nkz['_merge']=='left_only'].drop(columns=excel_GIS_Export.columns)
not_matched_r = merge_nkz[merge_nkz['_merge']=='right_only'].drop(columns=excel_Netzleitzahlen.columns)
merge_not_matched = pd.merge(not_matched_l, not_matched_r, left_on="NLZ", right_on='Stations ID', how='outer', indicator=False)
matched = merge_nkz[merge_nkz['_merge'] == 'both'].drop(columns=['_merge'])
final_data = pd.concat([matched, merge_not_matched], ignore_index=True)

# Sortieren nach Stationsname, wenn nicht vorhanden nach Name
final_data["_sort_key"] = (final_data["Stationsname"].fillna(final_data['Name']).astype(str).str.strip().apply(de_sort_key))
final_data = (final_data.sort_values(by="_sort_key").drop(columns="_sort_key").reset_index(drop=True))
final_data = final_data.loc[:, ~final_data.columns.str.contains("_merge")]  # alle _merge Spalten löschen

# Daten für Netzstationen Statistik extrahieren
only_netzstationen = final_data[final_data["Stationstyp"].str.contains("Netz", na=False)]
only_netzstationen['Baujahr'] = only_netzstationen['Baujahr'].astype('Int64')
data_netzstationen_per_year = only_netzstationen.groupby('Baujahr').size()

# Diagramm erzeugen
plt.figure(figsize=(15, 7))
data_netzstationen_per_year.plot(kind="bar")
plt.title("Netzstationen pro Jahr")
plt.xlabel("Jahr")
plt.ylabel("Anzahl")
plt.xticks(rotation=90)
#plt.show()

# Diagramm speichern
img_data = io.BytesIO()
plt.savefig(img_data, format='png')
plt.close() # Schließt das Diagramm um Speicher zu sparen

# Excel-Datei kreiren
file_name = 'Auswertung_GIS.xlsx'
with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    final_data.to_excel(writer, sheet_name='Vergleich', index=False, startrow=1)

    # Zugriff auf das Workbook und Worksheet
    workbook = writer.book
    worksheet = writer.sheets['Vergleich']

    # Überschrift erstellen und formatieren
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nb_columns_excel_Netzleitzahlen)
    worksheet.merge_cells(start_row=1, start_column=nb_columns_excel_Netzleitzahlen+1, end_row=1, end_column=nb_columns)
    cell_left = worksheet.cell(row=1, column=1)
    cell_right = worksheet.cell(row=1, column=nb_columns_excel_Netzleitzahlen+1)
    cell_left.value = "Netzleitzahlen"
    cell_right.value = "GIS"
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

# Unterschiede zwischen den Spalten grafisch hervorheben
highlight_differences(file_name, 'Vergleich', 'Stationsname', 'Name', header_row=2)
highlight_differences(file_name, 'Vergleich', 'NKZ', 'Kurzname', header_row=2)
highlight_differences(file_name, 'Vergleich', 'NLZ', 'Stations ID', header_row=2)
highlight_differences(file_name, 'Vergleich', 'Netz', 'Teilnetz', header_row=2)
print("Datei wurde erfolgreich gespeichert!")

