#!/usr/bin/env bash
# Baut den Ordner  output/upload/  — nur die Dateien für Teilnehmende und
# Gestaltung. Die Skripte, der Master und data/saal-referenz/ laufen lokal und
# gehören nicht mit hierhin. Danach:  output/upload/  auf Google Drive hochladen.
set -euo pipefail
cd "$(dirname "$0")"

ZIEL="output/upload"
rm -rf "$ZIEL"
mkdir -p "$ZIEL/ausgabe"

PY=./.venv/bin/python
[ -x "$PY" ] || PY=python3

# frische Werkmeldung + frische Beispiel-PDFs
"$PY" scripts/build_werkmeldung.py
"$PY" scripts/build_druckdaten.py

cp output/FUB2026_Werkmeldung_VORLAGE.xlsx "$ZIEL/"
cp data/design.json                        "$ZIEL/design_VORLAGE.json"
cp output/HANDBUCH_DESIGNER.pdf             "$ZIEL/"
cp output/ausgabe/*.pdf                     "$ZIEL/ausgabe/"
cp LIESMICH_UPLOAD.txt                      "$ZIEL/LIESMICH.txt"

echo
echo "Fertig: $ZIEL/"
find "$ZIEL" -type f | sort
echo
du -sh "$ZIEL"
