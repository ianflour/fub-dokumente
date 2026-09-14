# -*- coding: utf-8 -*-
"""
Zentrale Pfad-Auflösung für alle Scripts von FOTOTAGE MUSTERSTADT 2026.

Grundsatz: keine absoluten Pfade im Code, keine Annahme über das aktuelle
Arbeitsverzeichnis. Das Projekt-Root wird zur Laufzeit aus __file__ abgeleitet;
alles darunter (data/, assets/, output/) hängt daran. Die Scripts laufen damit
gleich, egal ob aus dem Projekt-Root, aus scripts/ oder per absolutem Aufruf.

    import _pfade as P
    P.data("teilnehmende.txt")      -> <root>/data/teilnehmende.txt
    P.assets("schriften")           -> <root>/assets/schriften
    P.output("SCHICHTPLAN.pdf")     -> <root>/output/SCHICHTPLAN.pdf   (legt output/ an)
    P.output("ausgabe", "x.pdf")    -> <root>/output/ausgabe/x.pdf     (legt ausgabe/ an)

data/ wird an sys.path gehängt, damit `import schichtplan_daten` /
`import gaesteliste_daten` aus den Entry-Points weiter funktioniert.
"""

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "scripts")
DATA = os.path.join(ROOT, "data")
ASSETS = os.path.join(ROOT, "assets")
OUTPUT = os.path.join(ROOT, "output")


def _join(basis, teile):
    return os.path.join(basis, *[str(t) for t in teile])


def root(*teile):
    return _join(ROOT, teile)


def scripts(*teile):
    return _join(SCRIPTS, teile)


def data(*teile):
    return _join(DATA, teile)


def assets(*teile):
    return _join(ASSETS, teile)


def output(*teile):
    """Pfad unter output/ — legt die nötigen Ordner gleich mit an."""
    ziel = _join(OUTPUT, teile)
    ordner = os.path.dirname(ziel) if teile else OUTPUT
    os.makedirs(ordner, exist_ok=True)
    return ziel


def rel(pfad):
    """Pfad relativ zum Projekt-Root — für lesbare Meldungen."""
    try:
        return os.path.relpath(pfad, ROOT)
    except ValueError:
        return str(pfad)


# editierbare Datenmodule (schichtplan_daten.py, gaesteliste_daten.py) liegen
# in data/ — dort auch importierbar machen.
if DATA not in sys.path:
    sys.path.insert(0, DATA)
