**Kurzzusammenfassung**

Dieses Script führt ein Datenvergleich zwischen der Excel Datei Netzleitzahlen_Makro und der Excel Datei K3v-Export durch.
Das Script generiert eine neue Excel Datei in der die Daten schön übersichtlich angezeigt werden. Dabei werden die Daten intelligent geordent, Unterschiede frablich hervorgehoben, sowie Statistiken erstellt.

**Script starten**

Um das Script zu starten muss wie folgt vorgenangen werden:
In dem Terminal in Visual Studio Code muss folgender Befehl ausgeführt werden:
```bash
pip install -r requirements.txt
```
Dadruch werden die nötigen Bibliotheken für das Script installiert.

Anschließend muss überprüft werden ob im Script Spaltenvergleich.py der Datei-Pfad, der Datei-Name sowie die Spaltenname der K3-Export Datei
übereinstimmen:
```bash
# !!! Check if correct: !!!
path_folder_K3_Export = Path(r"\\estw-01\Bereich-N\NG\NGE\Statistiken ESTW\20-kV-Stationen")
path_K3_Export = path_folder_K3_Export / "K3V-Export_20260213.xlsx"
K3V_name_column = ["K3v", "Nummer", "Betriebsstatus", "Inbetriebnahme"]
```