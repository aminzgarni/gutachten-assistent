# Gutachten-Assistent

Lokale Hilfe für ärztliche Sachverständige in Betreuungsverfahren (Gutachten nach § 1814 BGB). Läuft im Browser auf dem eigenen Rechner;
Akten, Aufnahmen und Texte werden lokal verarbeitet.

**Funktionen**
- Fälle anlegen (Ordnerstruktur im frei wählbaren Arbeitsordner, auch in OneDrive / Google Drive / Dropbox)
- Gerichts-PDFs lokal zusammenführen und in Text umwandeln, Scans per OCR (Tesseract, Deutsch)
- Pseudonymisierung (Platzhalter für Namen, Orte, Telefon, E-Mail, IBAN, Geburtsdaten) und Rücktausch im Entwurf
- Optional: lokale Transkription von Gesprächen (faster-whisper) als Markdown
- Auftrag und Explorationen (Fremdanamnese, E-Mails ... mit Datum) je Fall erfassen
- Entwurf selbst schreiben oder mit KI: automatisch erzeugter, pseudonymisierter Prompt, Kopieren und Öffnen des gewählten KI-Dienstes (einstellbar)
- Gutachten-Entwurf und Kostenrechnung (JVEG) aus eigenen Word-Vorlagen mit Platzhaltern; automatisch geführte Rechnungsliste je Jahr als Excel-Datei (neue Rechnung = Zeile, gelöschte Rechnung = Zeile entfällt)
- Darstellung hell/dunkel/System
- Einrichtungsassistent, eigene Module und Leistungen

**Start:** siehe `ANLEITUNG.txt` (Windows: `Einrichten-Windows.bat`, Mac: `Einrichten-Mac.command`).

**Voraussetzungen:** Python 3.9+. Pakete: `python-docx openpyxl pymupdf` (Pflicht), `pymupdf4llm faster-whisper av` (optional), Tesseract (optional, für Scans).

## Aufbau
- `programm/gutachten.py` Kernfunktionen (auch per Terminal nutzbar)
- `programm/web.py` lokaler Server (nur 127.0.0.1) und `programm/web/index.html` Oberfläche
- `programm/vorlagen_engine.py` Platzhalter-Engine für Word-Vorlagen, `programm/vorlagen/` Standardvorlagen
- Persönliche Daten (Name, IBAN, Vorlagen, Fälle) liegen **nicht** im Programmordner, sondern im Arbeitsordner (`_einstellungen/`) bzw. im Benutzerordner (`~/.gutachten-assistent`)

## Hinweise
- Keine Rechtsberatung. Honorarsätze, Kilometergeld und Schreibgebühren müssen vom Nutzer gegen das geltende JVEG geprüft werden.
- Datenschutz und Berufsgeheimnis (§ 203 StGB, DSGVO) liegen in der Verantwortung der Nutzerin / des Nutzers: Vor dem Hochladen bei externen Diensten Ausgabe prüfen und Nutzungsbedingungen beachten.
- OCR und Spracherkennung machen Fehler; Ergebnisse immer prüfen.
- Lizenz: GNU GENERAL PUBLIC LICENSE Version 3.
