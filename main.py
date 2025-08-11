import pandas as pd
import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename

Kursanzahl = 23 

print("Dieses Programm erstellt eine CSV-Datei für den Import in school@min. Sie benötigen die Schülerbezogene Kursliste, die Sie aus der LUSD exportiert haben.\n"
"Eine Anleitung, wie Sie die Schülerbezogene Kursliste exportieren, finden Sie unter: https://wiki.medienzentrum-mtk.de/e/de/anleitungen/auth/lusdexport \n"
"Die CSV-Datei wird im selben Verzeichnis wie dieses Programm erstellt und wird wiefolgt benannt: 'Datum'-KNE-Import-SuS.csv\n"
"Bitte beachten Sie, dass die CSV-Datei nur für den Import in school@min geeignet ist.")
print("")
print("Copyright 2025 Luca Dünte luca.duente@baseworks.de BASEWORKS GmbH")
print("")
print("Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), "
"to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, "
"and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:"
"The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n"

"THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, "
"FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, "
"WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.")
print("")
print("Bitte bestätigen Sie, dass Sie die Lizenzbedingungen gelesen und verstanden haben, indem Sie 'Ja' eingeben.")
lizenz = input().strip().lower()
if lizenz != "ja":
    print("Sie müssen die Lizenzbedingungen akzeptieren, um das Script zu verwenden.")
    exit()

kurseEinbeziehen = None
print("Möchten Sie Kurse in die Auswertung einbeziehen? (Ja/Nein)")

while kurseEinbeziehen not in ("Ja", "Nein"):
    antwort = input().strip().lower()
    if antwort == "ja":
        kurseEinbeziehen = "Ja"
    elif antwort == "nein":
        kurseEinbeziehen = "Nein"
    else:
        print("Bitte geben Sie 'Ja' oder 'Nein' ein.")


Tk().withdraw()

print("Bitte wählen Sie die XLSX-Datei mit der Schülerbezogenen Kursliste aus.")
filename = askopenfilename(
    title="Kursliste auswählen",
    filetypes=[("Excel-Dateien", "*.xlsx")]
)

print("Beachten Sie, dass die Erstellung der CSV-Datei einige Minuten in Anspruch nehmen kann, je nach Größe der Datei.")

if not filename:
    print("Es wurde keine Datei ausgewählt.")
    exit()

try:
    df = pd.read_excel(filename, engine='openpyxl')
    print("Datei erfolgreich geladen.")
except FileNotFoundError:
    print("Die Datei wurde nicht gefunden.")
    exit()
except Exception as e:
    print(f"Ein Fehler ist aufgetreten: {e}")
    exit()

try:
    df.replace(r'^\s*$', None, regex=True, inplace=True)
    df['Anzeige_Zeile'] = df['Anzeige_Zeile'].astype(str).str.strip()
    df = df[df['Anzeige_ZeileNr'] == 1]
    df = df.sort_values(by=["Anzeige_Zeile"], ascending=True)

    names = ["SLR_Namenszusatz", "KLA_Klassenlehrer", "Anzeige_ZeileNr", "Anzeige_Gruppe", "Fach"]
    for name in names:
        columnsToDrop = [col for col in df.columns if col.startswith(name)]
        df.drop(columns=columnsToDrop, inplace=True)

    df = df.applymap(lambda x: x.split('/')[0] if isinstance(x, str) else x)
    df = df.applymap(lambda x: x.split('-')[0] if isinstance(x, str) else x)

    kursdatenSpalten = [col for col in df.columns if col.startswith('Kursdaten_')]
    df[kursdatenSpalten] = df[kursdatenSpalten].fillna("")

    csvColumns = [
        'Anmeldename', 'Nachname', 'Vorname', 'Klasse', 'Passwort', 'Gruppe',
        'Beschreibung', 'UserId'
    ]
    for i in range(1, Kursanzahl):
        csvColumns.append(f'Fach{i}')

    csvDf = pd.DataFrame(columns=csvColumns)
    newRows = []

    klassen = df['KLA_Klassennamen'].dropna().unique()

    for klasse in klassen:
        dfKlasse = df[df['KLA_Klassennamen'] == klasse]
        
        uniqueStudentIds = dfKlasse['Anzeige_Zeile'].dropna().unique()
        
        studentCourses = {}
        for student in uniqueStudentIds:
            studentRows = dfKlasse[dfKlasse['Anzeige_Zeile'] == student]
            kurse = set()
            for _, row in studentRows.iterrows():
                for i in range(0, 12): 
                    kursKey = f'Kursdaten_{i}'
                    kursValue = row.get(kursKey, "").strip()
                    if kursValue:
                        kurse.add(kursValue)
            studentCourses[student] = kurse

        if studentCourses:
            gemeinsameKurse = set.intersection(*studentCourses.values())
        else:
            gemeinsameKurse = set()
        
        for student, kurse in studentCourses.items():
            if kurseEinbeziehen == "Ja":
                individuelleKurse = [kurs for kurs in kurse if kurs not in gemeinsameKurse]
            else:
                individuelleKurse = []

            firstRow = dfKlasse[dfKlasse['Anzeige_Zeile'] == student].iloc[0]

            newRow = {
                'Anmeldename': "",
                'Nachname': firstRow.get('SLR_NachName', ""),
                'Vorname': firstRow.get('SLR_VorName', ""),
                'Klasse': klasse,
                'Passwort': "",
                'Gruppe': "Schüler",
                'Beschreibung': "",
                'UserId': ""
            }

            if kurseEinbeziehen:
                for i in range(1, min(len(individuelleKurse) + 1, Kursanzahl)):
                    newRow[f'Fach{i}'] = individuelleKurse[i - 1]

            newRows.append(newRow)

    csvDf = pd.DataFrame(newRows, columns=csvColumns)
    csvDf = csvDf.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    today = datetime.datetime.today().strftime("%d.%m.%Y")

    csvDf.to_csv(f"{today}-KNE-Import-SuS.csv", index=False, sep=';')

    print("Die CSV-Datei wurde als 'KNE-Import.csv' erstellt.")

except Exception as e:
    print(f"Ein Fehler ist aufgetreten: {e}")
    exit()