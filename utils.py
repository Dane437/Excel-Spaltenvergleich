from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
import unicodedata

# Vergleicht die Werte in den angegebenen Spalten und hebt die Unterschiede hervor
def highlight_differences(file_name, sheet_name, col_left, col_right, header_row=2):
    wb = load_workbook(file_name)
    ws = wb[sheet_name]

    red_fill = PatternFill(
        start_color='FFCCCC',
        end_color='FFCCCC',
        fill_type='solid'
    )

    headers = {cell.value: cell.column for cell in ws[header_row]}

    col_station = headers[col_left]
    col_k3v = headers[col_right]

    for row in range(header_row + 1, ws.max_row + 1):
        c1 = ws.cell(row=row, column=col_station)
        c2 = ws.cell(row=row, column=col_k3v)

        if c1.value != c2.value:
            c1.fill = red_fill
            c2.fill = red_fill

    wb.save(file_name)


# Hilfsfunktion zum Sortieren von Strings mit deutschen Umlauten
def de_sort_key(s):
    return (
        unicodedata.normalize("NFKD", s)
        .encode("ascii", "ignore")
        .decode("ascii")
        .lower()
    )