# SAM4K

Dieses Projekt dient der automatisierten Scheibenauswertung mithilfe der **SAM4000 Ring- und Teiler-Messmaschine**.

![Quelle: Knestel](https://github.com/Logogistiks/sam4k/blob/main/sam4000.png)

Hauptgegenstand ist die Datei [SAM_Auswertung.py](https://github.com/Logogistiks/sam4k/blob/main/SAM_Auswertung.py). Die Bedienung erfolgt über eine Konsolenoberfläche mit anschließender Datenspeicherung in einer **Excel-Datei**.

## 📋Inhaltsverzeichnis

- #### [📖Bedienung](#bedienung)
- #### [🛠️Für Entwickler](#für-entwickler-1)

## 📖Bedienung

1. (Maschine per Seriellanschluss mit dem PC verbinden)
1. Programm `SAM_Auswertung.py` starten
1. **Schussanzahl** pro Streifen auswählen
1. **Berechnungsmodus** auswählen
1. Wiederholen bis alle Schützen eingegeben:
    1. **Name** des Schützen eingeben
    1. Scheiben nacheinander in Maschine einlegen
    1. (`n` für neuen Schützen drücken)
1. Wenn fertig, `esc` für Speichern & Beenden drücken
1. Die Excel-Datei öffnet sich automatisch.

## 🛠️Für Entwickler

**Dokumentation** und Infos: siehe [readme_developer.md](https://github.com/Logogistiks/sam4k/blob/main/readme_developer.md)

**Anleitung** als PDF: [SAM4000_Anleitung.pdf](https://github.com/Logogistiks/sam4k/blob/main/SAM4000_Anleitung.pdf)
