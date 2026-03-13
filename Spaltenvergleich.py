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

# Spalte einer Excel-Datei filtern nach bestimmten Zeichen

# !!! Überpüfen: !!!
path_folder = Path(r"C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp")    #Dateipfad überprüfen
file_1_name = 'Trafoliste_K3V.xlsx'
file_1_sheet_name = 'Tabelle1'
file_1_header = 0
file_1_filter_column = 'Nummer NKZ'
file_1_filter_character = 't'
file_1_needed_columns = ['Nummer NKZ']

# Excel einlesen
excel_file_1 = pd.read_excel(path_folder / file_1_name, file_1_sheet_name, header=file_1_header)
excel_file_1[['Nummer NLZ', 'Nummer NKZ']] = (excel_file_1['Nummer'].str.extract(r'.*?(\d{3})\s*/\s*(.+)')) # NLZ und NKZ in eigenen Spalten
excel_file_1 = excel_file_1[file_1_needed_columns]  # auskommentieren, wenn alle Spalten angezeigt werden sollen
excel_file_1 = excel_file_1[excel_file_1[file_1_filter_column].astype(str).str.contains(file_1_filter_character, na=False)]

# Excel-Datei kreiren
file_name = 'Filtern_' + file_1_name.split('.')[0] + '.xlsx'
file_path = path_folder / file_name

excel_file_1.to_excel(file_path, index=False)
print("Datei wurde erfolgreich gespeichert!")

