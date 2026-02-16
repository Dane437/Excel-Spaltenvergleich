from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
import unicodedata

# Vergleicht die Werte in den angegebenen Spalten und hebt die Unterschiede hervor
def highlight_differences(file_name, sheet_name, col_left, col_right, header_row=2):
    wb = load_workbook(file_name)
    ws = wb[sheet_name]

    headers = {cell.value: cell.column for cell in ws[header_row]}

    for row in range(header_row + 1, ws.max_row + 1):
        c1 = ws.cell(row=row, column=headers[col_left])
        c2 = ws.cell(row=row, column=headers[col_right])

        if c1.value != c2.value:
            c1.fill = PatternFill(start_color='FFCCCC', fill_type='solid')  # rote Formatierung
            c2.fill = PatternFill(start_color='FFCCCC', fill_type='solid')  # rote Formatierung

    wb.save(file_name)


# Hilfsfunktion zum Sortieren von Strings mit deutschen Umlauten
def de_sort_key(s):
    return (
        unicodedata.normalize("NFKD", s)
        .encode("ascii", "ignore")
        .decode("ascii")
        .lower()
    )