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

#Datenvergleich zwischen Netzleitzahlen.xlsm und K3V-Export.xlsx

#### !!! Check if correct: !!!
path_folder_K3_Export = Path(r"\\estw-01\Bereich-N\NG\NGE\Statistiken ESTW\20-kV-Stationen")
path_K3_Export = path_folder_K3_Export / "K3V-Export_20260213.xlsx"
K3V_name_column = ["K3v", "Nummer", "Betriebsstatus", "Inbetriebnahme"] #K3V_name_column[0], K3V_name_column[1], ...

#file paths
path_folder_Netzleitzahlen = Path(r"B:\# N-Gemeinsam\Trafostationsliste")
path_Netzleitzahlen = path_folder_Netzleitzahlen /  "Netzleitzahlen_Makro.xlsm"

#Excel einlesen
excel_Netzleitzahlen = pd.read_excel(path_Netzleitzahlen, sheet_name="Stationsliste")
excel_Netzleitzahlen['NLZ'] = pd.to_numeric(excel_Netzleitzahlen['NLZ'], errors='coerce')
excel_K3_Export = pd.read_excel(path_K3_Export, sheet_name="Tabelle1")
# Die Zeilen filtern: Behalte alles, wo " (Liegenschaft)" NICHT vorkommt
excel_K3_Export = excel_K3_Export[~excel_K3_Export[K3V_name_column[0]].astype(str).str.contains(r'\(Liegenschaft\)', na=False)] # na=False -> leere Zellen ignorieren
excel_K3_Export[['Nummer NLZ', 'Nummer NKZ']] = (excel_K3_Export[K3V_name_column[1]].str.extract(r'.*?(\d{3})\s*/\s*(.+)'))
excel_K3_Export['Nummer NLZ'] = pd.to_numeric(excel_K3_Export['Nummer NLZ'], errors='coerce')
nb_columns_excel_K3_Export = excel_K3_Export.shape[1]
nb_columns_excel_Netzleitzahlen = excel_Netzleitzahlen.shape[1]
nb_columns = nb_columns_excel_K3_Export + nb_columns_excel_Netzleitzahlen

#Datensätze zusammenführen
merge_adress: pd.DataFrame = pd.merge(excel_Netzleitzahlen, excel_K3_Export, left_on='Stationsname', right_on=K3V_name_column[0], how='outer', indicator=True)
not_matched_l = merge_adress[merge_adress['_merge']=='left_only'].drop(columns=excel_K3_Export.columns)
not_matched_r = merge_adress[merge_adress['_merge']=='right_only'].drop(columns=excel_Netzleitzahlen.columns)
merge_not_matched = pd.merge(not_matched_l, not_matched_r, left_on="NLZ", right_on='Nummer NLZ', how='outer', indicator=False)
matched = merge_adress[merge_adress['_merge'] == 'both'].drop(columns=['_merge'])
final_data = pd.concat([matched, merge_not_matched], ignore_index=True)

# Sortieren nach Stationsname, wenn nicht vorhanden nach K3v 
final_data["_sort_key"] = (final_data["Stationsname"].fillna(final_data[K3V_name_column[0]]).astype(str).str.strip().apply(de_sort_key))
final_data = (final_data.sort_values(by="_sort_key").drop(columns="_sort_key").reset_index(drop=True))

#final_data = final_data.sort_values(by=["Stationsname", K3V_name_column[0]], ascending=[True, True]);
final_data = final_data.loc[:, ~final_data.columns.str.contains("_merge")]  #alle _merge Spalten löschen

merge_adress.to_excel('merge_adress.xlsx', index=False)
not_matched_l.to_excel('not_matched_l.xlsx', index=False)
not_matched_r.to_excel('not_matched_r.xlsx', index=False)
merge_not_matched.to_excel('fallback.xlsx', index=False)
matched.to_excel('matched.xlsx', index=False)
final_data.to_excel('merge_final.xlsx', index=False)

copy_merge_final = final_data.copy()
copy_merge_final[K3V_name_column[3]] = pd.to_datetime(copy_merge_final[K3V_name_column[3]], dayfirst=True, errors='coerce')
copy_merge_final[K3V_name_column[3]] = copy_merge_final[K3V_name_column[3]].dt.year.astype('Int64')

netzstationen = copy_merge_final[copy_merge_final["Stationstyp"].str.contains("Netz", na=False)]
anzahl_pro_jahr = netzstationen.groupby(K3V_name_column[3]).size()

# Unterschiede finden
#only_adress_IBS = merge_adress[['Stationsname', 'K3v', 'Inbetriebnahme']]

# Diagramm erzeugen
plt.figure(figsize=(15, 7))
anzahl_pro_jahr.plot(kind="bar")
plt.title("Netzstationen pro Jahr")
plt.xlabel("Jahr")
plt.ylabel("Anzahl")
plt.xticks(rotation=90)
#plt.tight_layout()
#plt.show()

img_data = io.BytesIO()
plt.savefig(img_data, format='png')
plt.close() # Schließt die Figure, um Speicher zu sparen

# Excel-Datei kreiren
#merge_name.to_excel('Temp.xlsx', index=False)
with pd.ExcelWriter('Auswertung.xlsx', engine='openpyxl') as writer:
    final_data.to_excel(writer, sheet_name='Vergleich', index=False, startrow=1)
    # Zugriff auf das Workbook und Worksheet
    workbook = writer.book
    worksheet = writer.sheets['Vergleich']

    # Dicke Linie definieren
    thick_border = Side(style='thick')

    # Letzte Zeile bestimmen
    max_row = worksheet.max_row

    # Formatierung Überschrift
    worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nb_columns_excel_Netzleitzahlen)
    worksheet.merge_cells(start_row=1, start_column=nb_columns_excel_Netzleitzahlen+1, end_row=1, end_column=nb_columns)

    cell_left = worksheet.cell(row=1, column=1)  # A1
    cell_right = worksheet.cell(row=1, column=nb_columns_excel_Netzleitzahlen+1)  # G1

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

    for col in range(1, nb_columns + 1):  # A-K = 1-11
        cell = worksheet.cell(row=2, column=col)
        cell.fill = grey_fill
        cell.font = bold_font

    # Rechte Seite von Spalte F (Spalte 6) dick machen
    for row in range(1, max_row + 1):
        cell = worksheet.cell(row=row, column=nb_columns_excel_Netzleitzahlen)  # Spalte F
        cell.border = Border(right=thick_border)

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

    # Zweites Sheet: Der Plot
    # Wir erstellen manuell ein leeres Sheet
    workbook = writer.book
    worksheet = workbook.create_sheet('Diagramm')
    
    # Das Bild aus dem Speicher laden und in das Sheet einfügen
    img_data.seek(0) # Cursor an den Anfang des Speichers setzen
    img = Image(img_data)
    worksheet.add_image(img, 'A1')

highlight_differences('Auswertung.xlsx', 'Vergleich', 'Stationsname', K3V_name_column[0], header_row=2)
highlight_differences('Auswertung.xlsx', 'Vergleich', 'NKZ', 'Nummer NKZ', header_row=2)
highlight_differences('Auswertung.xlsx', 'Vergleich', 'NLZ', 'Nummer NLZ', header_row=2)
print("Datei wurde erfolgreich gespeichert!")

