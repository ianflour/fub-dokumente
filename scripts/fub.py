#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FOTOTAGE MUSTERSTADT 2026 — Haupt-Menü
================================
Ein Startpunkt für alle Werkzeuge. Ruft jedes Script als eigenen Prozess auf,
mit demselben Python wie dieses hier.

    ./.venv/bin/python scripts/fub.py            # Menü anzeigen
    ./.venv/bin/python scripts/fub.py 4          # direkt Punkt 4 (Schichtplan)
    ./.venv/bin/python scripts/fub.py 3 --probe  # Punkt 3, Argumente durchreichen

Pfade sind komplett dynamisch: das Script-Verzeichnis wird aus __file__
bestimmt, alles Weitere über _pfade.py. Egal, aus welchem Ordner gestartet.
"""

import os
import subprocess
import sys

HIER = os.path.dirname(os.path.abspath(__file__))

# (Taste, Titel, [Scripts …], Hinweistext)
AKTIONEN = [
    ("1", "Alles neu bauen",
     ["build_werkmeldung.py", "build_qr_codes.py",
      "build_druckdaten.py", "build_schichtplan.py", "build_gaesteliste.py",
      "build_workshop.py", "build_handbuch_designer.py",
      "build_handbuch_zustaendigkeiten.py"],
     "Werkmeldung → QR-Codes → alle PDFs (Master NICHT neu — Menüpunkt 8)"),
    ("2", "Druckdaten erzeugen", ["build_druckdaten.py"],
     "Wandschilder, Werkliste, Etiketten, Versicherung, Instagram-Wand"),
    ("3", "Werkmeldungen einlesen", ["import_werkmeldungen.py"],
     "aus data/werkmeldungen/ in den Master"),
    ("4", "Schichtplan", ["build_schichtplan.py"],
     "output/SCHICHTPLAN.pdf aus data/schichtplan_daten.py"),
    ("5", "Gästeliste", ["build_gaesteliste.py"],
     "output/GAESTELISTE.pdf aus data/gaesteliste_daten.py"),
    ("6", "Workshop-Anmeldeliste", ["build_workshop.py"],
     "output/WORKSHOP.pdf aus data/workshop_daten.py"),
    ("7", "Handbücher", ["build_handbuch_designer.py",
                         "build_handbuch_zustaendigkeiten.py"],
     "HANDBUCH_DESIGNER.pdf + HANDBUCH_ZUSTAENDIGKEITEN.pdf"),
    ("8", "Master-Tabelle neu", ["build_master.py"],
     "output/FUB2026_Werkdaten_MASTER.xlsx (überschreibt!)"),
    ("9", "Werkmeldung-Vorlage neu", ["build_werkmeldung.py"],
     "output/FUB2026_Werkmeldung_VORLAGE.xlsx"),
    ("10", "QR-Codes", ["build_qr_codes.py"],
     "output/qr-codes/ aus data/instagram.csv"),
    ("11", "Werkmeldung migrieren", ["migrate_werkmeldung.py"],
     "alte Werkmeldung aus data/werkmeldung_alt/ in die neue Vorlage"),
]
NACH_TASTE = {t: a for a in AKTIONEN for t in (a[0],)}


def lauf(script, extra):
    pfad = os.path.join(HIER, script)
    print(f"\n\033[1m▶ {script}\033[0m {' '.join(extra)}".rstrip())
    r = subprocess.run([sys.executable, pfad, *extra])
    return r.returncode


def fuehre_aus(taste, extra):
    aktion = NACH_TASTE.get(taste)
    if not aktion:
        print(f"Unbekannte Auswahl: {taste!r}")
        return 2
    _, titel, scripts, _ = aktion
    mehrere = len(scripts) > 1
    for script in scripts:
        # Extra-Argumente nur an Einzel-Aktionen weiterreichen
        code = lauf(script, [] if mehrere else extra)
        if code != 0:
            print(f"\n\033[31m✗ Abbruch bei {script} (Code {code}).\033[0m")
            return code
    print(f"\n\033[32m✓ {titel} fertig.\033[0m")
    return 0


def menue():
    print("\n\033[1mFOTOTAGE MUSTERSTADT 2026\033[0m\n")
    for taste, titel, _, hinweis in AKTIONEN:
        print(f"  {taste:>2}  {titel:<24} \033[2m{hinweis}\033[0m")
    print(f"  {'0':>2}  Beenden")
    try:
        return input("\nAuswahl: ").strip()
    except (EOFError, KeyboardInterrupt):
        return "0"


def main():
    argv = sys.argv[1:]
    if argv:
        taste, extra = argv[0], argv[1:]
        # "--" trennt optional Menü-Argument von durchgereichten Argumenten
        if extra and extra[0] == "--":
            extra = extra[1:]
        sys.exit(fuehre_aus(taste, extra))

    while True:
        wahl = menue()
        if wahl in ("0", "q", "quit", "exit", ""):
            return
        fuehre_aus(wahl, [])


if __name__ == "__main__":
    main()
