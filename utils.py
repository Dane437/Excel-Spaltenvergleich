from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
import unicodedata

# Vergleicht die Werte in den angegebenen Spalten und hebt die Unterschiede hervor
def highlight_differences(file_name, sheet_name, col_left, col_right, header_row, valid_combinations, empty_is_difference):
    wb = load_workbook(file_name)
    ws = wb[sheet_name]

    headers = {cell.value: cell.column for cell in ws[header_row]}
    rows_different = []

    for row in range(header_row + 1, ws.max_row + 1):
        c1 = ws.cell(row=row, column=headers[col_left])
        c2 = ws.cell(row=row, column=headers[col_right])
        
        if not empty_is_difference and (c1.value is None or c2.value is None):
            #continue
            one_is_empty = True

        if not valid_combinations:
            if c1.value != c2.value:
                rows_different.append(row)
                if one_is_empty:
                    c1.fill = PatternFill(start_color='FFEB9C', fill_type='solid')  # gelbe Formatierung
                    c2.fill = PatternFill(start_color='FFEB9C', fill_type='solid')  # gelbe Formatierung
                    one_is_empty = False
                else:
                    c1.fill = PatternFill(start_color='FFCCCC', fill_type='solid')  # rote Formatierung
                    c2.fill = PatternFill(start_color='FFCCCC', fill_type='solid')  # rote Formatierung
        else:
            if (c1.value, c2.value) in valid_combinations:
                continue

            c1.fill = PatternFill(start_color='FFCCCC', fill_type='solid')  # rote Formatierung
            c2.fill = PatternFill(start_color='FFCCCC', fill_type='solid')  # rote Formatierung

    wb.save(file_name)
    return rows_different


# Hilfsfunktion zum Sortieren von Strings mit deutschen Umlauten
def de_sort_key(s):
    return (
        unicodedata.normalize("NFKD", s)
        .encode("ascii", "ignore")
        .decode("ascii")
        .lower()
    )


def create_differences_sheet(file_name, rows_different, header_row):
    wb = load_workbook(file_name)
    ws = wb['Temp']
    headers = {cell.value: cell.column for cell in ws[header_row]}
    worksheet_copy = wb.copy_worksheet(ws)
    worksheet_copy.title = 'Unterschiede'

    for row in range(worksheet_copy.max_row, header_row, -1):    # zählt rückwärts und stoppt bei header_row
        if row not in rows_different:
             worksheet_copy.delete_rows(row)

    wb.save(file_name)