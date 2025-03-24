import pandas as pd
import datetime

Kursanzahl = 23 

print("Dieses Script erstellt eine CSV-Datei für den Import in school@min. Sie benötigen die Schülerbezogene Kursliste, die Sie aus der LUSD exportiert haben.\n"
      "Bitte installieren Sie die folgenden Python-Bibliotheken, wenn Sie sie noch nicht installiert haben: pandas, openpyxl, datetime.\n"
      "Sie können die Bibliotheken installieren, indem Sie 'pip install -r requirements.txt' in der Kommandozeile eingeben.\n"
"Eine Anleitung, wie Sie die Schülerbezogene Kursliste exportieren, finden Sie hier: https://wiki.medienzentrum-mtk.de/e/de/anleitungen/auth/lusdexport \n"
"Bitte stellen Sie sicher, dass die Schülerbezogene Kursliste im selben Verzeichnis wie dieses Script liegt.\n"
"Die CSV-Datei wird im selben Verzeichnis wie dieses Script erstellt und wird wiefolgt benannt: 'Datum'-KNE-Import-SuS.csv\n"
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
lizenz = input()
if lizenz != "Ja":
    print("Sie müssen die Lizenzbedingungen akzeptieren, um das Script zu verwenden.")
    exit()

print("Bitte geben Sie den Dateinamen der Schülerbezogenen Kursliste ein (ohne Dateiendung).")

filename = input()

try:
    df = pd.read_excel(f"{filename}.xlsx", engine='openpyxl')
except FileNotFoundError:
    print("Die Datei wurde nicht gefunden. Bitte stellen Sie sicher, dass die Datei im selben Verzeichnis wie dieses Script liegt.")
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
        csvColumns.append(f'Fach_{i}')

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
            individuelleKurse = [kurs for kurs in kurse if kurs not in gemeinsameKurse]
            
            firstRow = dfKlasse[dfKlasse['Anzeige_Zeile'] == student].iloc[0]
            
            newRow = {
                'Anmeldename': "",
                'Nachname': firstRow.get('SLR_NachName', ""),
                'Vorname': firstRow.get('SLR_VorName', ""),
                'Klasse': klasse,
                'Passwort': "",
                'Gruppe': "Schüler",
                'Beschreibung': "",
                'UserId': student
            }
            
            for i in range(1, min(len(individuelleKurse) + 1, Kursanzahl)):
                newRow[f'Fach_{i}'] = individuelleKurse[i - 1]
            
            newRows.append(newRow)

    csvDf = pd.DataFrame(newRows, columns=csvColumns)
    csvDf = csvDf.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    today = datetime.datetime.today().strftime("%d.%m.%Y")

    csvDf.to_csv(f"{today}-KNE-Import-SuS.csv", index=False, sep=';')

    print("Die CSV-Datei wurde als 'KNE-Import.csv' erstellt.")

except Exception as e:
    print(f"Ein Fehler ist aufgetreten: {e}")
    exit()