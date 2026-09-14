# -*- coding: utf-8 -*-
"""
Bearbeitbare Datenbasis für die Gästeliste von FOTOTAGE MUSTERSTADT 2026.

    ./.venv/bin/python scripts/build_gaesteliste.py   ->   output/GAESTELISTE.pdf  (A4 quer)

Enthält die Zusagen aus Politik und Öffentlichkeit. Nur diese Datei ändern,
dann das Script neu laufen lassen. Datumsangaben als ISO ("2026-09-12");
der Wochentag wird berechnet.

HINWEIS: Namen und Organisationen in dieser Datei sind Beispieldaten für
die öffentliche Demo — keine echten Personen.

Felder je Gast:
    name          Vor- und Nachname
    funktion      Amt / Rolle (leer lassen, wenn keine)
    organisation  Fraktion, Ministerium, Behörde …
    datum         ISO-Datum des Termins
    zeit          Uhrzeit "HH:MM" (leer lassen, wenn offen)
    anlass        "Vernissage" oder "Finissage"
    anmerkung     interne Notiz fürs Empfangsteam
    grusswort     True, wenn die Person ein Grußwort spricht (optional)
"""

# ----------------------------------------------------------------------
VERANSTALTUNG = {
    "titel":      "FOTOTAGE MUSTERSTADT 2026",
    "untertitel": "Ausstellung für zeitgenössische Fotografie",
    "ort":        "Kunstraum Musterstadt · Beispielweg 5, 00000 Musterstadt",
    "laufzeit":   "12.—19. September 2026",
    "kontakt":    "kontakt@beispiel-verein.de · fototage-musterstadt.de",
}

# ----------------------------------------------------------------------
# Anlässe (Datum wird für die Überschriften gebraucht).
ANLAESSE = {
    "Vernissage": "2026-09-12",
    "Finissage":  "2026-09-19",
}

# ----------------------------------------------------------------------
GAESTE = [
    {
        "name":         "Elena Beispiel",
        "funktion":     "",
        "organisation": "Kulturverein Musterstadt",
        "datum":        "2026-09-12",
        "zeit":         "14:00",
        "anlass":       "Vernissage",
        "anmerkung":    "Zusage bestätigt.",
    },
    {
        "name":         "Jan Muster",
        "funktion":     "Referent",
        "organisation": "Amt für Kultur",
        "datum":        "2026-09-12",
        "zeit":         "14:00",
        "anlass":       "Vernissage",
        "anmerkung":    "Bringt Kollegin mit.",
    },
    {
        "name":         "Frida Vogel",
        "funktion":     "Sprecherin",
        "organisation": "Stadtrat Musterstadt",
        "datum":        "2026-09-12",
        "zeit":         "14:00",
        "anlass":       "Vernissage",
        "anmerkung":    "Explizit auf 14:00 Uhr bestätigt.",
    },
    {
        "name":         "Otto Beispiel",
        "funktion":     "Staatssekretär",
        "organisation": "Ministerium für Kultur",
        "datum":        "2026-09-12",
        "zeit":         "14:00",
        "anlass":       "Vernissage",
        "anmerkung":    "Spricht ein Grußwort.",
        "grusswort":    True,
    },
    {
        "name":         "Clara Winter",
        "funktion":     "Bürgermeisterin",
        "organisation": "Stadt Musterstadt",
        "datum":        "2026-09-19",
        "zeit":         "15:00",
        "anlass":       "Finissage",
        "anmerkung":    "Nimmt an der Finissage teil.",
    },
]

# ----------------------------------------------------------------------
# Freie Hinweise fürs Empfangs-/Orga-Team (optional, je ein Absatz).
HINWEISE = [
    "Diese Liste ist intern. Sie dient dem Empfang an der Tür und der "
    "Begrüßung — nicht zum Aushängen.",
    "Alle genannten Personen haben fest zugesagt. Bei kurzfristigen Absagen "
    "oder weiteren Zusagen diese Datei ergänzen und das PDF neu erzeugen.",
]
