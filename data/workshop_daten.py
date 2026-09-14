# -*- coding: utf-8 -*-
"""
Bearbeitbare Datenbasis für die Workshop-Anmeldeliste von FOTOTAGE MUSTERSTADT 2026.

    ./.venv/bin/python scripts/build_workshop.py   ->   output/WORKSHOP.pdf  (A4 quer)

Nur diese Datei ändern, dann das Script neu laufen lassen. Datumsangaben als
ISO ("2026-09-13"); der Wochentag wird berechnet.

HINWEIS: Namen und Kontaktangaben in dieser Datei sind Beispieldaten für
die öffentliche Demo — keine echten Personen.
"""

# ----------------------------------------------------------------------
VERANSTALTUNG = {
    "titel":      "FOTOTAGE MUSTERSTADT 2026",
    "untertitel": "Ausstellung für zeitgenössische Fotografie",
    "kontakt":    "kontakt@beispiel-verein.de · fototage-musterstadt.de",
}

# ----------------------------------------------------------------------
WORKSHOP = {
    "titel":   "Analoge Fotografie",
    "leitung": "Ada Beispiel",
    "ort":     "Kunstraum Musterstadt · Beispielweg 5, 00000 Musterstadt",
    "plaetze": 10,
    # (ISO-Datum, Uhrzeit, Kurzbeschreibung)
    "termine": [
        ("2026-09-13", "ab 15:00 Uhr", "Fotografieren und Entwickeln"),
        ("2026-09-15", "ab 18:00 Uhr", "Ausbelichten"),
    ],
}

# ----------------------------------------------------------------------
# Schon feststehende Anmeldungen. Leer lassen für eine blanko Liste zum
# Aushängen. Je Eintrag:  {"name": "...", "kontakt": "...", "info": "..."}
# (kontakt = Instagram-Handle, E-Mail oder Telefon)
# "info" (optional) = Ausrüstung / Anmerkung der Person; erscheint auf einem
# eigenen Blatt nach der Anmeldeliste. Ohne "info" taucht die Person dort nicht auf.
ANMELDUNGEN = [
    {"name": "Elin Beispiel", "kontakt": "@elin.beispiel",
     "info": "Bringt eine Analogkamera mit."},
    {"name": "Timo Muster",   "kontakt": "@timomuster"},
]
