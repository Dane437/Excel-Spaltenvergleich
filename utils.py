from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
import unicodedata

# Vergleicht die Werte in den angegebenen Spalten und hebt die Unterschiede hervor
def highlight_differences(file_name, sheet_name, col_left, col_right, header_row, valid_combinations):
    wb = load_workbook(file_name)
    ws = wb[sheet_name]

    headers = {cell.value: cell.column for cell in ws[header_row]}

    for row in range(header_row + 1, ws.max_row + 1):
        c1 = ws.cell(row=row, column=headers[col_left])
        c2 = ws.cell(row=row, column=headers[col_right])

        if not valid_combinations:
            if c1.value != c2.value:
                c1.fill = PatternFill(start_color='FFCCCC', fill_type='solid')  # rote Formatierung
                c2.fill = PatternFill(start_color='FFCCCC', fill_type='solid')  # rote Formatierung
        else:
            if (c1.value, c2.value) in valid_combinations:
                continue

            c1.fill = PatternFill(start_color='FFCCCC', fill_type='solid')  # rote Formatierung
            c2.fill = PatternFill(start_color='FFCCCC', fill_type='solid')  # rote Formatierung

    wb.save(file_name)

def create_district_sheets(writer, df_merged_files, district_name):
    if district_name == 'Gesamt':
        df_district = df_merged_files
    else:
        df_district = df_merged_files[df_merged_files['Ortsteil'] == district_name]

    nb_rows_df_district = len(df_district)  # Anzahl Gesamtzeilen
    nb_district_nothing = df_district['Absicherung'].isna().sum().astype('int64')   # Anzahl der Zeilen wo nichts drin ist
    percent_district_nothing = nb_district_nothing / (nb_rows_df_district/100)
    nb_district_nothing_power = df_district['Leistung'].isna().sum().astype('int64') - nb_district_nothing   # Anzahl der Zeilen wo kein Leistungswert zuordbar ist
    percent_district_nothing_power = nb_district_nothing_power / (nb_rows_df_district/100)
    sum_power = df_district['Leistung'].sum()
    df_district.to_excel(writer, sheet_name=district_name, index=False, startrow=0)
    
    ws = writer.sheets[district_name]
    ws.cell(row=1, column=11, value='%')
    ws.cell(row=2, column=9, value='Gesamtzahl Sicherungswerte')
    ws.cell(row=2, column=10, value=nb_rows_df_district)
    ws.cell(row=3, column=9, value='kein Sicherungswert')
    ws.cell(row=3, column=10, value=nb_district_nothing)
    ws.cell(row=3, column=11, value=percent_district_nothing.round(0))
    ws.cell(row=4, column=9, value='kein Leistungswert zuordbar')
    ws.cell(row=4, column=10, value=nb_district_nothing_power)
    ws.cell(row=4, column=11, value=percent_district_nothing_power.round(0))
    ws.cell(row=6, column=9, value='Summe Leistung')
    ws.cell(row=6, column=10, value=sum_power)
    ws.column_dimensions['I'].width = 28
       
# Hilfsfunktion zum Sortieren von Strings mit deutschen Umlauten
def de_sort_key(s):
    return (
        unicodedata.normalize("NFKD", s)
        .encode("ascii", "ignore")
        .decode("ascii")
        .lower()
    )