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
path_folder = Path(r"C:\Users\immler.daniel\OneDrive - Erlanger Stadtwerke AG\Dokumente\j-PythonTemp\Temp2\StraßeZuOrtsteil")    #Dateipfad überprüfen
file_1_name = 'Straßen nach Ortsteilen'
file_1_sheet_name = 'Tabelle1'
file_1_merging_column = 'Straße_'
file_1_header = 0
#file_1_needed_columns = ['NLZk', 'NKZk']

# Excel einlesen
excel_file_1 = pd.read_excel(path_folder / (file_1_name +'.xlsx'), sheet_name=file_1_sheet_name, header=file_1_header)



# Daten individuel bearbeiten

# excel_file_1['Inbetriebnahmedatum der Einheit'] = pd.to_datetime(excel_file_1['Inbetriebnahmedatum der Einheit'], dayfirst=True)
# excel_file_1 = excel_file_1[excel_file_1['Inbetriebnahmedatum der Einheit'].dt.year == 2024]

# excel_file_1 = excel_file_1[excel_file_1['Energieträger'] == 'Solare Strahlungsenergie']
# excel_file_1 = excel_file_1[excel_file_1['Betriebsstatus'] == 'In Betrieb']
# excel_file_1 = excel_file_1[excel_file_1['Systemstatus'] == 'Aktiviert']

# excel_file_2['Inbetriebnahmedatum der Einheit'] = pd.to_datetime(excel_file_2['Inbetriebnahmedatum der Einheit'], dayfirst=True)
# excel_file_2 = excel_file_2[excel_file_2['Inbetriebnahmedatum der Einheit'].dt.year == 2024]

# excel_file_2 = excel_file_2[excel_file_2['Energieträger'] == 'Solare Strahlungsenergie']
# excel_file_2 = excel_file_2[excel_file_2['Betriebsstatus'] == 'In Betrieb']
# excel_file_2 = excel_file_2[excel_file_2['Systemstatus'] == 'Aktiviert']

#excel_file_1 = excel_file_1[file_1_needed_columns]

excel_file_1 = excel_file_1.reindex(columns=['Straße', 'Ortsteil', 'Teilnetz'])
excel_file_1 = excel_file_1.drop_duplicates()

# Datensätze zusammenführen
excel_file_1.to_excel('Gesamt.xlsx', index=False)

print("Datei wurde erfolgreich gespeichert!")

