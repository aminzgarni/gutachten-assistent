#!/bin/bash
cd "$(dirname "$0")"
echo "============================================"
echo "  Gutachten-Assistent - Einrichtung"
echo "============================================"
if ! command -v python3 >/dev/null 2>&1 || ! python3 -c 'import sys;sys.exit(0 if sys.version_info>=(3,9) else 1)' >/dev/null 2>&1; then
  echo "Python 3 fehlt oder ist zu alt. Die Download-Seite wird geoeffnet."
  echo "Bitte den Installer laden und ausfuehren (immer 'Fortfahren'). Danach diese Datei erneut oeffnen."
  open "https://www.python.org/downloads/macos/"
  read -p "Enter zum Beenden" _; exit 1
fi
echo "Python gefunden. Die noetigen Programme werden installiert (einmalig, kann einige Minuten dauern) ..."
python3 -m pip install --upgrade pip >/dev/null 2>&1
if ! python3 -m pip install --upgrade python-docx openpyxl pymupdf; then
  python3 -m pip install --user --break-system-packages --upgrade python-docx openpyxl pymupdf || { echo "Die Installation ist fehlgeschlagen. Bitte Internetverbindung pruefen."; read -p "Enter zum Beenden" _; exit 1; }
fi
echo
echo "Fertig. Der Assistent startet jetzt in deinem Browser."
echo "Weitere Bausteine (Texterkennung, Transkription) richtest du dort im Schritt 'Programme' ein."
nohup python3 programm/web.py >/dev/null 2>&1 &
sleep 2
exit 0
