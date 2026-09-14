#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FOTOTAGE MUSTERSTADT 2026 — QR-Codes für die Instagram-Profile der Teilnehmenden
==========================================================================
Liest  data/instagram.csv  (Name;Dateiname;Instagram-Handle)  und
erzeugt pro Person einen QR-Code als SVG (für die Gestaltung) und als
PNG (für den Druck über build_druckdaten.py, Instagram-Wand).

  ./.venv/bin/python scripts/build_qr_codes.py            # nur fehlende/erzeugte
  ./.venv/bin/python scripts/build_qr_codes.py --alle     # alle neu erzeugen
  ./.venv/bin/python scripts/build_qr_codes.py --probe    # nur anzeigen

Fehlt bei jemandem der Instagram-Handle, wird nichts erzeugt und die
Person am Ende aufgelistet — Handle in instagram.csv nachtragen und
das Skript erneut laufen lassen.

Abhängigkeit: segno  (pip install segno)
"""

import argparse
import os

import segno

import _pfade as P
import _teilnehmende as T

ZIEL = P.output("qr-codes")
PNG_DIR = os.path.join(ZIEL, "png")
os.makedirs(ZIEL, exist_ok=True)


def erzeuge(handle, url, dateiname, png=True):
    qr = segno.make(url, error="m")          # 15 % Fehlerkorrektur, robust genug
    svg_pfad = os.path.join(ZIEL, f"{dateiname}.svg")
    qr.save(svg_pfad, kind="svg", scale=1, border=2, unit="mm",
            svgclass=None, lineclass=None, xmldecl=True, svgns=True)
    if png:
        os.makedirs(PNG_DIR, exist_ok=True)
        qr.save(os.path.join(PNG_DIR, f"{dateiname}.png"),
                kind="png", scale=12, border=2)
    return svg_pfad


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--alle", action="store_true", help="auch vorhandene SVGs neu erzeugen")
    ap.add_argument("--probe", action="store_true", help="nur anzeigen, nichts schreiben")
    ap.add_argument("--ohne-png", action="store_true")
    args = ap.parse_args()

    neu, aktualisiert, fehlt, uebersprungen = [], [], [], []

    for name in T.NAMEN:
        info = T.INSTAGRAM.get(name)
        if not info or not info["handle"]:
            fehlt.append(name)
            continue
        datei = info["dateiname"]
        svg_pfad = os.path.join(ZIEL, f"{datei}.svg")
        existiert = os.path.exists(svg_pfad)
        if existiert and not args.alle:
            uebersprungen.append(name)
            continue
        if args.probe:
            (aktualisiert if existiert else neu).append(f"{name}  →  {info['url']}")
            continue
        erzeuge(info["handle"], info["url"], datei, png=not args.ohne_png)
        (aktualisiert if existiert else neu).append(f"{name}  →  {info['url']}")

    # verwaiste SVGs (Datei da, aber Person nicht mehr in data/teilnehmende.txt)
    bekannt = {i["dateiname"] for i in T.INSTAGRAM.values()}
    verwaist = sorted(
        f for f in os.listdir(ZIEL)
        if f.endswith(".svg") and f[:-4] not in bekannt
    )

    print(f"\nQR-Codes — {len(T.NAMEN)} Teilnehmende")
    if neu:
        print(f"\nNeu erzeugt ({len(neu)}):")
        for z in neu:
            print("  +", z)
    if aktualisiert:
        print(f"\n{'Würde aktualisieren' if args.probe else 'Aktualisiert'} ({len(aktualisiert)}):")
        for z in aktualisiert:
            print("  ~", z)
    if uebersprungen and not args.alle:
        print(f"\nUnverändert ({len(uebersprungen)}) — mit --alle neu erzeugen.")
    if verwaist:
        print(f"\nVerwaiste SVG-Dateien ({len(verwaist)}) — nicht in data/teilnehmende.txt:")
        for f in verwaist:
            print("  ?", f)
    if fehlt:
        print(f"\n>>> {len(fehlt)} Teilnehmende ohne Instagram-Link — bitte in "
              f"data/instagram.csv nachtragen:")
        for n in fehlt:
            print("   ", n)
        print("\nDanach:  ./.venv/bin/python scripts/build_qr_codes.py")
    else:
        print("\nAlle Teilnehmenden haben einen QR-Code.")

    if args.probe:
        print("\nPROBELAUF — nichts geschrieben.")


if __name__ == "__main__":
    main()
