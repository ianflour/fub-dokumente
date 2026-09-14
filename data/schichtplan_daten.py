# -*- coding: utf-8 -*-
"""
Bearbeitbare Datenbasis für den Schichtplan von FOTOTAGE MUSTERSTADT 2026.

    ./.venv/bin/python scripts/build_schichtplan.py   ->   output/SCHICHTPLAN.pdf  (A4 quer)

Diese Datei enthält nur die VORGABEN: Rollen, Pools, Programm, den Bedarf je
Zeitfenster (SCHICHTEN) und je Person Vorlieben, mögliche Tage und Blocker.
Wer wann eingeteilt ist, rechnet build_schichtplan.py bei jedem Lauf selbst
aus — Ziel ist ein für alle möglichst ausgeglichener Fairness-Quotient
(nicht Soll-Stunden), unter Beachtung von Pools, Blockern und Vorlieben.
Nur diese Datei ändern, dann das Script neu laufen lassen.

HINWEIS: Namen, Orte und Termine in dieser Datei sind Beispieldaten für die
öffentliche Demo — keine echten Personen.

Datumsangaben als ISO ("2026-09-12"); der Wochentag wird berechnet.
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
# Rollen im Schichtbetrieb.  (Kurzname, Aufgabe, Wann)
ROLLEN = [
    ("Einlass",                  "Tür, Gäste begrüßen und zählen",
     "nur 12.09., bis 20:00 · 2-h-Schichten, 2 Personen"),
    ("Theke / Aufsicht",         "Getränke, Kasse und Verkauf, Auskunft, Blick auf die Werke",
     "alle Tage · am 12.09. bis 22:00, 2-h-Schichten, 2 Personen"),
    ("Social Media",             "Fotos und Storys für Instagram, kurze Clips",
     "nur 12.09., bis 19:30 · 3-h-Schichten, 1 Person"),
    ("Fotografische Begleitung", "Dokumentation der Veranstaltung (Fotos)",
     "nur 12.09., bis 19:30 · 3-h-Schichten, max. 2 Personen"),
    ("Workshop-Leitung",         "Leitet den Analog-Workshop; an dem Tag komplett "
                                 "dafür geblockt, keine weitere Aufgabe",
     "13.09. + 15.09."),
]

# ----------------------------------------------------------------------
# Pools: wer die jeweilige Rolle übernehmen kann.
POOLS = {
    "Social Media": [
        "Noah Muster", "Paul Winter", "Tom Feldmann",
    ],
    "Fotografische Begleitung": [
        "Paul Winter", "Tom Feldmann", "Jonas Wilde",
    ],
}

# ----------------------------------------------------------------------
# Programm (Zeittafel) — nur zur Orientierung.
# (ISO-Datum, Titel, [(Zeit, Programmpunkt), …])
PROGRAMM = [
    ("2026-09-12", "Vernissage", [
        ("14:00",        "Eröffnung"),
        ("14:15",        "Grußwort"),
        ("16:00",        "Artist Talk"),
        ("18:00",        "Kurzvortrag"),
    ]),
    ("2026-09-13", "Workshop Teil 1 / Ausstellung", [
        ("ab 15:00",     "Workshop Analoge Fotografie: Fotografieren und Entwickeln"),
        ("15:00–21:00",  "Ausstellung"),
    ]),
    ("2026-09-14", "Ausstellung", [("18:00–21:00", "Ausstellung")]),
    ("2026-09-15", "Workshop Teil 2 / Ausstellung", [
        ("ab 18:00",     "Workshop Analoge Fotografie: Ausbelichten"),
        ("18:00–21:00",  "Ausstellung"),
    ]),
    ("2026-09-16", "Ausstellung", [("18:00–21:00", "Ausstellung")]),
    ("2026-09-17", "Ausstellung", [("18:00–21:00", "Ausstellung")]),
    ("2026-09-18", "Ausstellung", [("18:00–21:00", "Ausstellung")]),
    ("2026-09-19", "Finissage",   [("15:00 – Open End", "Ausklang")]),
]

# ----------------------------------------------------------------------
# Mögliche Aufgaben für das Feld "vorliebe" (genau so schreiben):
AUFGABEN = ["Einlass", "Theke / Aufsicht", "Social Media", "Fotografische Begleitung"]

# ----------------------------------------------------------------------
# Personen.  Schlüssel = Name exakt wie in data/teilnehmende.txt.
#   vorliebe : bevorzugte Aufgaben, in Reihenfolge der Vorliebe. Werte aus
#              AUFGABEN oben. Leere Liste = keine Vorliebe. Das Script bevorzugt
#              diese Aufgaben, garantiert sie aber nicht.
#   tage     : Liste von ISO-Daten, an denen die Person KANN. [] = alle Tage.
#              (für „kann nur samstags", „kann am 12./16." usw.)
#   blocker  : Liste von (ISO-Datum, "von–bis"[, Grund]) — einzelne Sperrzeiten.
#              "von–bis" darf "ganztags" sein.  [] = keine.
#   notiz    : optionaler Hinweis (erscheint in der Stundenübersicht)
#
# Es gibt keine Soll-Stunden. Das Script verteilt so, dass der Fairness-
# Quotient (Stunden ÷ Schnitt) für alle möglichst nahe bei 1,00 liegt;
# nur die Workshop-Leitung liegt bauartbedingt darüber.
PERSONEN = {
    "Mia Beispiel":     {"vorliebe": [], "blocker": [], "notiz": "Aufgabe egal"},
    "Noah Muster":      {"vorliebe": [], "blocker": [], "notiz": ""},
    "Lina Vogel":       {"vorliebe": ["Einlass"],
                         "tage": ["2026-09-12", "2026-09-13", "2026-09-19"],
                         "blocker": [], "notiz": "Theke eher ungern"},
    "Paul Winter":      {"vorliebe": [], "blocker": [], "notiz": ""},
    "Sara Klein":       {"vorliebe": [],
                         "tage": ["2026-09-15"],
                         "blocker": [], "notiz": "kann höchstens 15.09. 18:00–21:00"},
    "Tom Feldmann":     {"vorliebe": [], "blocker": [], "notiz": ""},
    "Nina Berger":      {"vorliebe": [], "blocker": [], "notiz": ""},
    "Jonas Wilde":      {"vorliebe": [], "blocker": [], "notiz": ""},
    "Studio Nordlicht": {"vorliebe": [], "blocker": [], "notiz": ""},
    "Foto Kollektiv":   {"vorliebe": [], "blocker": [], "notiz": ""},
    "Ada Beispiel":     {"vorliebe": [], "blocker": [],
                         "notiz": "leitet den Analog-Workshop (13. + 15.09.); an "
                                  "diesen Tagen ganztägig geblockt, keine Theke / Aufsicht"},
}

# Tandems: Personen, die nur GEMEINSAM eingeteilt werden — immer dieselbe
# Schicht oder gar nicht. Je Eintrag genau zwei Namen (wie in PERSONEN).
TANDEM = [
    ("Nina Berger", "Jonas Wilde"),
]

# Personen, die BEWUSST nicht in die Schichtplanung aufgenommen werden.
NICHT_EINGEPLANT = {}

# ----------------------------------------------------------------------
# Schichten = der BEDARF je Zeitfenster. Die Namen verteilt das Script
# (build_schichtplan.py) automatisch — hier stehen sie NICHT.
#   datum   : ISO "2026-09-12"
#   zeit    : "14:00–19:30"
#   anlass  : was in dem Fenster passiert
#   bedarf  : {Rolle: Anzahl Personen}
#   fix     : {Rolle: [Name, …]}  — optional. Personen, die in diesem Fenster
#             FEST auf dieser Rolle stehen (z. B. Workshop-Leitung, Wünsche).
#             Zählt in den Bedarf hinein; die restlichen Plätze füllt das Script.
SCHICHTEN = [
    # ===================== Vernissage 12.09. =====================
    {"datum": "2026-09-12", "zeit": "14:00–16:00",
     "anlass": "Vernissage — Eröffnung, Grußwort",
     "bedarf": {"Einlass": 2, "Theke / Aufsicht": 2}},
    {"datum": "2026-09-12", "zeit": "16:00–18:00",
     "anlass": "Vernissage — Artist Talk",
     "bedarf": {"Einlass": 2, "Theke / Aufsicht": 2}},
    {"datum": "2026-09-12", "zeit": "18:00–20:00",
     "anlass": "Vernissage — Einlass bis 20:00",
     "bedarf": {"Einlass": 2}},
    {"datum": "2026-09-12", "zeit": "18:00–20:00",
     "anlass": "Vernissage — Kurzvortrag, offizieller Abschluss (19:00)",
     "bedarf": {"Theke / Aufsicht": 2}},
    {"datum": "2026-09-12", "zeit": "20:00–22:00",
     "anlass": "Vernissage — Ausklang",
     "bedarf": {"Theke / Aufsicht": 2}},
    {"datum": "2026-09-12", "zeit": "14:00–17:00",
     "anlass": "Vernissage — Foto & Social Media (Eröffnung bis Artist Talk)",
     "bedarf": {"Fotografische Begleitung": 2, "Social Media": 1}},
    {"datum": "2026-09-12", "zeit": "17:00–19:30",
     "anlass": "Vernissage — Foto & Social Media (Kurzvortrag)",
     "bedarf": {"Fotografische Begleitung": 2, "Social Media": 1}},

    # ===================== So 13.09. — Workshop I / Ausstellung =====================
    {"datum": "2026-09-13", "zeit": "15:00–21:00",
     "anlass": "Analog-Workshop Teil 1 (Fotografieren & Entwickeln) — Leitung",
     "bedarf": {"Workshop-Leitung": 1},
     "fix": {"Workshop-Leitung": ["Ada Beispiel"]}},
    {"datum": "2026-09-13", "zeit": "14:45–18:00",
     "anlass": "Ausstellung (während Workshop Teil 1)",
     "bedarf": {"Theke / Aufsicht": 1}},
    {"datum": "2026-09-13", "zeit": "18:00–21:15",
     "anlass": "Ausstellung",
     "bedarf": {"Theke / Aufsicht": 1}},

    # ============= Mo–Fr 14.–18.09. — Ausstellung (Di 15.09. auch Workshop II) =============
    {"datum": "2026-09-14", "zeit": "18:00–21:00", "anlass": "Ausstellung",
     "bedarf": {"Theke / Aufsicht": 1}},
    {"datum": "2026-09-15", "zeit": "18:00–21:00",
     "anlass": "Analog-Workshop Teil 2 (Ausbelichten) — Leitung",
     "bedarf": {"Workshop-Leitung": 1},
     "fix": {"Workshop-Leitung": ["Ada Beispiel"]}},
    {"datum": "2026-09-15", "zeit": "18:00–21:00",
     "anlass": "Ausstellung (während Workshop Teil 2)",
     "bedarf": {"Theke / Aufsicht": 1}},
    {"datum": "2026-09-16", "zeit": "18:00–21:00", "anlass": "Ausstellung",
     "bedarf": {"Theke / Aufsicht": 1}},
    {"datum": "2026-09-17", "zeit": "18:00–21:00", "anlass": "Ausstellung",
     "bedarf": {"Theke / Aufsicht": 1}},
    {"datum": "2026-09-18", "zeit": "18:00–21:00", "anlass": "Ausstellung",
     "bedarf": {"Theke / Aufsicht": 1}},

    # ===================== Sa 19.09. — Finissage =====================
    {"datum": "2026-09-19", "zeit": "14:45–19:15",
     "anlass": "Finissage — Eröffnung",
     "bedarf": {"Theke / Aufsicht": 2}},
    {"datum": "2026-09-19", "zeit": "19:15–22:00",
     "anlass": "Finissage — Ausklang / Open End",
     "bedarf": {"Theke / Aufsicht": 1}},
]

# ----------------------------------------------------------------------
# Freie Hinweise — erscheinen unten im PDF. Leer lassen = kein Hinweisblock.
HINWEISE = []
