# OneNoteExport2MdFiles
export and convert one note notebooks into md-files

# OneNote-Notizbuch exportieren und mit COSMA verwenden

Praktische Anleitung für Export, Ablage in SharePoint und Nutzung in COSMA.

## Zweck

Ziel des Exports ist es, Informationen aus OneNote für COSMA zugänglich und nutzbar zu machen. So können Notizen über COSMA leichter gefunden und für Antworten herangezogen werden. 
Wo es möglich ist, nennt COSMA außerdem die verwendete Quelle oder verlinkt sie. Damit können Angaben in der Originalnotiz überprüft und bei Bedarf dort weiterrecherchiert werden.

Der Export erstellt eine Kopie der OneNote-Inhalte als Word- und – sofern Pandoc installiert ist – Markdown-Dateien. Das OneNote-Notizbuch selbst wird dabei **nicht verändert**.

## Voraussetzungen

- Ein Windows-Arbeitsplatz mit der OneNote-Desktopanwendung.
- Das gewünschte Notizbuch ist in OneNote verfügbar und synchronisiert.
- Für Word-Dateien ist kein zusätzliches Programm erforderlich.
- Für Markdown-Dateien und das Inhaltsverzeichnis muss **Pandoc** installiert sein. Pandoc bietet Installationspakete für Windows und macOS auf der offiziellen [Pandoc-Downloadseite](https://pandoc.org/installing.html).
- Wenn COSMA später auf die exportierten Dateien zugreifen soll, wird ein SharePoint- oder OneDrive-Ordner benötigt, auf den das Microsoft-Konto Zugriff hat.

## Ausgabeordner und Dateien

Im ausgewählten Ausgabeordner erstellt das Exportprogramm einen Unterordner mit dem Namen des Notizbuchs. Die Dateien werden nach OneNote-Abschnitten geordnet:

```
<Ausgabeordner>\
└── <Name des Notizbuchs>\
    ├── docx\
    │   └── <Abschnitt>\
    │       └── <Seitentitel>.docx
    └── md\                         # wenn Pandoc verfügbar ist
        ├── inhaltsverzeichnis.md
        └── quellen\
            └── <Abschnitt>\
                ├── <Seitentitel>.md
                └── <Seitentitel>_media\
```

- **docx** enthält die Notizseiten als Word-Dateien.
- **md** enthält die Markdown-Dateien sowie das Inhaltsverzeichnis.
- **md/quellen** enthält die einzelnen Markdown-Seiten, sortiert nach Abschnitt.
- **`*_media`-Ordner** enthalten gegebenenfalls Bilder oder andere Medien, die zu einer Markdown-Seite gehören. Diese Ordner beim Verschieben oder Hochladen neben der jeweiligen Markdown-Datei belassen.
- **inhaltsverzeichnis.md** bietet einen Überblick über Abschnitte und Seiten, inklusive kurzer Vorschauen.
- **export-fehler.txt** wird nur angelegt, wenn beim Export Fehler aufgetreten sind.

> **Wichtig:** Wenn Pandoc nicht installiert ist, erstellt das Programm nur die Word-Dateien; Markdown-Dateien und Inhaltsverzeichnis fehlen dann.

## Export durchführen

1. OneNote öffnen und prüfen, ob das gewünschte Notizbuch verfügbar und synchronisiert ist.
2. Das bereitgestellte Exportprogramm starten.
3. Das gewünschte Notizbuch aus der angezeigten Liste auswählen.
4. Den Ausgabeordner angeben. Für die spätere Nutzung mit COSMA kann direkt ein synchronisierter SharePoint- oder OneDrive-Ordner gewählt werden. Alternativ zunächst lokal exportieren und die Dateien später hochladen.
5. Falls für das Notizbuch bereits ein Export vorhanden ist, auswählen, wie damit umgegangen werden soll: Überschreiben, nur neue Dateien erstellen oder abbrechen.
6. Auf die Abschlussmeldung warten. Sie zeigt an, wie viele Seiten exportiert, übersprungen oder nicht verarbeitet wurden.

Bei der Auswahl des Exportmodus gilt:

| Modus | Verhalten |
|---|---|
| **Überschreiben** | Vorhandene Word-Dateien werden durch neu exportierte Dateien ersetzt. |
| **Nur neue Dateien** | Bestehende Word-Dateien bleiben erhalten; fehlende Dateien werden ergänzt. |
| **Abbrechen** | Der Export wird beendet. |

> **Achtung:** Beim Überschreiben können vorhandene Word-Dateien ersetzt werden. Dateien, in denen nach dem Export eigene Änderungen vorgenommen wurden, vorher an einem anderen Ort sichern.

## Markdown-Dateien in SharePoint ablegen

1. Den SharePoint- oder OneDrive-Ordner öffnen, in dem COSMA die Dateien später finden soll. Den Ordner wählen, der zum Notizbuch gehört: `<Name des Notizbuchs>`.
   - Ist dieser Ordner auf dem PC synchronisiert, kann er direkt als Ausgabeordner gewählt werden.
   - Andernfalls zunächst lokal exportieren und anschließend den erzeugten Notizbuchordner einschließlich der Unterordner nach SharePoint hochladen.
2. Die Ordnerstruktur beibehalten – insbesondere `inhaltsverzeichnis.md`, `quellen` und die zugehörigen Medienordner.
3. Warten, bis die Synchronisierung beziehungsweise der Upload abgeschlossen ist.

## SharePoint-Dateien mit COSMA verwenden

COSMA kann SharePoint- und OneDrive-Dateien suchen, einsehen und als Kontext im Chat verwenden. Dafür muss SharePoint in COSMA verbunden sein. Die Verbindung erfolgt unter **Connectors** bei **SharePoint**.

Nach dem Hochladen kann COSMA zum Beispiel wie folgt genutzt werden:

- „Suche im SharePoint-Ordner `<Name des Notizbuchs>` nach Notizen zu PM24.“
- „Finde die Markdown-Dateien zu Projektlager und fasse die wichtigsten Punkte zusammen.“
- „Was steht in der Notiz zum Thema Projektplanung? Bitte nenne Abschnitt und Seitentitel.“

Der Link zum SharePoint-Ordner kann auch direkt in die Frage eingefügt werden. COSMA kann SharePoint-Dateien suchen und Ordner prüfen, sofern das Microsoft-Konto Zugriff darauf hat.

> **Nutzung in COSMA-Projekten:** Bei Verwendung der Notizen im Rahmen eines COSMA-Projekts kann der Link zum SharePoint-Verzeichnis in den Projektanweisungen hinterlegt werden. Dadurch steht dieser Verweis als dauerhafter Kontext für Chats in diesem Projekt zur Verfügung. In der jeweiligen Anfrage sollte weiterhin angegeben werden, welche Themen oder Dateien COSMA berücksichtigen soll. Voraussetzung ist, dass das Microsoft-Konto Zugriff auf die SharePoint-Inhalte hat.

> **Hinweis:** Ob SharePoint Markdown-Dateien über die Suche findet, kann von der jeweiligen SharePoint-Konfiguration abhängen. Findet COSMA eine Datei nicht, sollte der Ordner- oder Dateilink direkt in der Frage angegeben werden.

## Wenn etwas nicht funktioniert

| Problem | Lösung |
|---|---|
| Das Notizbuch wird nicht angezeigt | OneNote öffnen und prüfen, ob das Notizbuch verfügbar und synchronisiert ist. Export danach erneut starten. |
| Es werden nur Word-Dateien erstellt | Pandoc ist möglicherweise nicht installiert. Über die [Pandoc-Downloadseite](https://pandoc.org/installing.html) installieren und Export erneut starten. |
| Dateien fehlen im SharePoint-Ordner | Synchronisierungsstatus prüfen oder warten, bis der Upload abgeschlossen ist. |
| COSMA findet eine Datei nicht | Prüfen, ob SharePoint in COSMA verbunden ist, ob das Konto Zugriff auf den Ordner hat und ob der richtige Ordner- oder Dateilink verwendet wird. |
| Seiten wurden nicht exportiert | Abschlussmeldung und ggf. `export-fehler.txt` im Notizbuch-Exportordner prüfen. |
| Ein OneNote-Link fehlt | Nicht für jede Seite kann ein Link zur Originalseite bereitgestellt werden. Die exportierte Datei kann trotzdem verwendet werden. |

## Hinweis zur Nutzung von Pandoc im Unternehmen

Pandoc wird unter der GNU General Public License (GPL) bereitgestellt und kann grundsätzlich auch in Unternehmen und für kommerzielle Zwecke verwendet werden. Besondere Lizenzpflichten sind insbesondere dann zu beachten, wenn Pandoc verändert oder Kopien davon weitergegeben werden. Für die interne Nutzung einer unveränderten Version ist üblicherweise keine separate kommerzielle Lizenz erforderlich. Zusätzlich sind interne Vorgaben zur Installation und Nutzung von Software zu beachten.

*Dies ist eine allgemeine Einordnung und keine Rechtsberatung.*

Weitere Informationen: [GNU GPL Version 2](https://www.gnu.org/licenses/old-licenses/gpl-2.0.html) · [Pandoc-Downloadseite](https://pandoc.org/installing.html)

## Beispiel für den Export eines Notizbuches

**Vorbereitung:** Neuen Export-Ordner erstellen, in dem die `.docx`- und `.md`-Dateien erzeugt werden sollen.

PowerShell-Skript starten - keine Parameter (hier mit der Windows PowerShell ISE):

```
Pandoc gefunden: C:\Users\<UserName>\AppData\Local\Pandoc\pandoc.exe

Verfügbare Notizbücher:
[1] AAA
[2] BBB
[3] CCC
[4] ...

Nummer des zu exportierenden Notizbuchs: 3
Basis-Ausgabeordner eingeben oder Enter für [C:\tmp\OneNoteExport4Cosma]:

Notizbuch: CCC
Hauptordner: C:\tmp\OneNoteExport4Cosma\CCC
Exportmodus: NewOnly

Exportiert: CCC / ORGA
Exportiert: CCC / Workshop
Exportiert: CCC / Team Viewer
Exportiert: CCC / Schnittstellen & DevOps
Exportiert: CCC / ASSIST
Exportiert: CCC / Planungsarbeitsblatt
Exportiert: CCC / Buchungsregeln
Exportiert: CCC / ...
Papierkorb übersprungen: Gelöschte Seiten

Export abgeschlossen.
DOCX neu erstellt/überschrieben: 13
Markdown neu erstellt/aktualisiert: 13
Index: C:\tmp\OneNoteExport4Cosma\CCC\md\inhaltsverzeichnis.md
Übersprungene Seiten: 1
Fehler: 0
Hauptordner: C:\tmp\OneNoteExport4Cosma\CCC
```

*Das Überspringen der gelöschten Seiten ist gewollt.*

Der Hauptordner im Datei-Explorer sieht anschließend etwa so aus:

```
OneNoteExport4Cosma/
└── CCC/
    ├── docx/
    │   └── CCC/
    └── md/
        └── quellen/
            └── CCC/
                ├── ASSIST_media/
                └── Anyfy_media/
```
