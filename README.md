**Kurzzusammenfassung**

Dieses Script führt ein Datenvergleich zwischen der Excel Datei Netzleitzahlen_Makro und der Excel Datei K3v-Export durch.
Das Script generiert eine neue Excel Datei in der die Daten schön übersichtlich angezeigt werden. Dabei werden die Daten intelligent geordent, Unterschiede frablich hervorgehoben, sowie Statistiken erstellt.

**Script starten**

Um das Script zu starten muss wie folgt vorgenangen werden:
Nachdem man Visual Studio Code gestartet hat muss der Code des Projekts von GitHub heruntergeladen werden.
Dazu gibt man in dem Terminal von Visual Studio Code folgenden Befehl ein:
```bash
git clone https://github.com/Dane437/Excel-Spaltenvergleich/tree/K3V
```
Durch den folgenden Befehl wird eine virtuelle Umgebung erstellt.
```bash
python -m venv .venv
.\.venv\Scripts\activate
```
Die Nutzung ist optional, aber empfehlenswert, wenn auf dem System bereits andere Projekte oder Bibliotheken installiert sind. So lassen sich Konflikte vermeiden und Abhängigkeiten sauber voneinander trennen.

```bash
pip install -r requirements.txt
```
Dadruch werden die nötigen Bibliotheken für das Script installiert.

Anschließend muss im Script Spaltenvergleich.py überprüft werden, ob der Dateipfad, der Dateiname sowie die Spaltennamen mit der K3-Export Datei
übereinstimmen:
```bash
# Pfad zu der K3V-Export-Datei
# !!! Überpüfen !!!
path_folder_K3_Export = Path(r"\\estw-01\Bereich-N\NG\NGE\Statistiken ESTW\20-kV-Stationen")    #Dateipfad überprüfen
path_K3_Export = path_folder_K3_Export / "K3V-Export_20260213.xlsx" # Dateiname überprüfen
K3V_name_column = ["K3v", "Nummer", "Betriebsstatus", "Inbetriebnahme"] # Spaltennamen überpüfen
```

Jetzt kann das Script gestartet werden. Entweder über das Terminal mit dem Befehl,
```bash
python Spaltenvergleich.py
```
oder alternativ über den folgenden Button:

<img width="380" height="287" alt="image" src="https://github.com/user-attachments/assets/7aa26e5d-a9a1-4942-8079-c92491364a5d" />

Wenn im Terminal steht: "Datei wurde erfolgreich gespeichert!" hat alles funktioniert und die Datei "Unterschiede_K3V_Netzleitzahlen.xlsx" wurde erstellt.
