import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter
import string
from pathlib import Path
import matplotlib.pyplot as plt

#Datenvergleich zwischen Netzleitzahlen.xlsm und K3V-Export.xlsx

#file paths

path_folder_Netzleitzahlen = Path(r"B:\# N-Gemeinsam\Trafostationsliste")
path_Netzleitzahlen = path_folder_Netzleitzahlen /  "Netzleitzahlen_Makro.xlsm"
path_folder_K3_Export = Path(r"\\estw-01\Bereich-N\NG\NGE\Statistiken ESTW\20-kV-Stationen")
path_K3_Export = path_folder_K3_Export / "K3V-Export_20260213.xlsx"

#Excel einlesen
excel_Netzleitzahlen = pd.read_excel(path_Netzleitzahlen, sheet_name="Stationsliste")
excel_K3_Export = pd.read_excel(path_K3_Export, sheet_name="Tabelle1")

#Datensätze zusammenführen
merge_adress: pd.DataFrame = pd.merge(excel_Netzleitzahlen, excel_K3_Export, left_on='Stationsname', right_on='K3v', how='outer')

copy_merge_adress = merge_adress.copy()
copy_merge_adress["Inbetriebnahme"] = pd.to_datetime(copy_merge_adress["Inbetriebnahme"], dayfirst=True, errors='coerce') 
copy_merge_adress["Inbetriebnahme"] = copy_merge_adress["Inbetriebnahme"].dt.year.astype('Int64')

netzstationen = copy_merge_adress[copy_merge_adress["Stationstyp"] == "Netzstation"]
anzahl_pro_jahr = netzstationen.groupby("Inbetriebnahme").size()


# Unterschiede finden
#only_adress_IBS = merge_adress[['Stationsname', 'K3v', 'Inbetriebnahme']]


# Excel-Datei kreiren
#only_adress_IBS.to_excel('Temp.xlsx', index=False)
with pd.ExcelWriter('merge_adress.xlsx', engine='openpyxl') as writer:
    merge_adress.to_excel(writer, sheet_name='Tabelle 1', index=False, startrow=1)
        # Zugriff auf das Workbook und Worksheet
    workbook = writer.book
    worksheet = writer.sheets['Tabelle 1']

    # Dicke Linie definieren
    thick_border = Side(style='thick')

    # Letzte Zeile bestimmen
    max_row = worksheet.max_row

    # Formatierung Überschrift

    worksheet.merge_cells('A1:F1')
    worksheet.merge_cells('G1:K1')

    cell_left = worksheet['A1']
    cell_right = worksheet['G1']

    cell_left.value = "Netzleitzahlen"
    cell_right.value = "K3V"

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")

    bold_font = Font(bold=True, size=16)
    center_align = Alignment(horizontal="center", vertical="center")

    cell_left.fill = green_fill
    cell_left.font = bold_font
    cell_left.alignment = center_align

    cell_right.fill = blue_fill
    cell_right.font = bold_font
    cell_right.alignment = center_align

    # Zeile zwei grau und fett
    grey_fill = PatternFill(fill_type="solid", start_color="E7E6E6")
    bold_font = Font(bold=True)

    for col in range(1, 12):  # A-K = 1-11
        cell = worksheet.cell(row=2, column=col)
        cell.fill = grey_fill
        cell.font = bold_font

    # Rechte Seite von Spalte F (Spalte 6) dick machen
    for row in range(1, max_row + 1):
        cell = worksheet.cell(row=row, column=6)  # Spalte F
        cell.border = Border(right=thick_border)
    
    for column in range(1,12):
        cell = worksheet.cell(row=1, column=column)
        cell.border = Border(bottom=thick_border)

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
        if adjusted_width > 30 :
            adjusted_width = 30
        worksheet.column_dimensions[column_letter].width = adjusted_width

print("Datei wurde erfolgreich gespeichert!")

# Diagramm erzeugen
#plt.figure(figsize=(14, 8))
#anzahl_pro_jahr.plot(kind="bar")
#plt.title("Netzstationen pro Jahr")
#plt.xlabel("Jahr")
#plt.ylabel("Anzahl")
#plt.xticks(rotation=90)
#plt.tight_layout()
#plt.show()