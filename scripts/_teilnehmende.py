# -*- coding: utf-8 -*-
"""
Gemeinsame Quelle für die Teilnehmenden von FOTOTAGE MUSTERSTADT 2026.

  data/teilnehmende.txt     — eine Person pro Zeile, Vor- und Nachname
  data/instagram.csv        — Name;Dateiname;Instagram-Handle

Genutzt von build_werkmeldung.py, build_master.py, build_qr_codes.py,
build_druckdaten.py. Ändern → nur in den beiden Quelldateien.
"""

import csv
import os
import re

import _pfade as P

TXT = P.data("teilnehmende.txt")
CSV = P.data("instagram.csv")


def _lies_namen():
    with open(TXT, encoding="utf-8") as f:
        return [z.strip() for z in f if z.strip()]


def slug(name):
    """Vor- und Nachname -> dateisystemtauglicher Bezeichner."""
    s = name.lower().strip()
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss"), ("é", "e"),
                 ("è", "e"), ("ê", "e"), ("á", "a"), ("à", "a"), ("ç", "c")):
        s = s.replace(a, b)
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s


def _lies_instagram():
    """{Name: {'dateiname': slug, 'handle': str, 'url': str|''}}"""
    out = {}
    if not os.path.exists(CSV):
        return out
    with open(CSV, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f, delimiter=";"):
            name = (row.get("Name") or "").strip()
            if not name:
                continue
            handle = (row.get("Instagram") or "").strip().lstrip("@")
            datei = (row.get("Dateiname") or "").strip() or slug(name)
            out[name] = {
                "dateiname": datei,
                "handle": handle,
                "url": f"https://www.instagram.com/{handle}/" if handle else "",
            }
    return out


NAMEN = _lies_namen()
INSTAGRAM = _lies_instagram()


def instagram_url(name):
    return (INSTAGRAM.get(name) or {}).get("url", "")


def fehlende_qr():
    """Namen aus Teilnehmende.txt ohne Instagram-Handle in der CSV."""
    return [n for n in NAMEN if not instagram_url(n)]


if __name__ == "__main__":
    print(f"{len(NAMEN)} Teilnehmende")
    fehlt = fehlende_qr()
    if fehlt:
        print(f"\n{len(fehlt)} ohne Instagram-Link:")
        for n in fehlt:
            print("  -", n)
    extra = [n for n in INSTAGRAM if n not in NAMEN]
    if extra:
        print(f"\n{len(extra)} in instagram.csv, aber nicht in Teilnehmende.txt:")
        for n in extra:
            print("  -", n)
