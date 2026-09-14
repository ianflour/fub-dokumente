#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Erzeugt GAESTELISTE.pdf (A4 quer) aus gaesteliste_daten.py.

Zusagen aus Politik und Öffentlichkeit für FOTOTAGE MUSTERSTADT 2026, nach Anlass
gruppiert. Design wie die übrigen Handbücher. Nur gaesteliste_daten.py ändern,
dann dieses Script neu laufen lassen.

Aufbau des PDF: Deckblatt, dann eine Seite mit Überblick (Zusagen je Anlass),
Grußwort-Hinweis und den Tabellen für Vernissage und Finissage.
"""

import glob
import os
from datetime import date

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, Frame, KeepTogether, NextPageTemplate,
                                PageBreak, PageTemplate, Paragraph, Spacer, Table,
                                TableStyle)

import _pfade as P
import gaesteliste_daten as D

OUT = P.output("GAESTELISTE.pdf")

# --- Schriften plattformunabhängig (Liberation Sans / Arial / DejaVu) -------
_FONT_DIRS = [
    P.assets("schriften"),
    "/usr/share/fonts/truetype/liberation", "/usr/share/fonts/truetype/dejavu",
    "/usr/share/fonts", "/Library/Fonts", "/System/Library/Fonts",
    "/System/Library/Fonts/Supplemental", os.path.expanduser("~/Library/Fonts"),
    "C:/Windows/Fonts",
]
_SANS_SETS = [
    ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf", "LiberationSans-Italic.ttf"),
    ("Arial.ttf", "Arial Bold.ttf", "Arial Italic.ttf"),
    ("arial.ttf", "arialbd.ttf", "ariali.ttf"),
    ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf"),
]


def _find(name):
    for d in _FONT_DIRS:
        hits = glob.glob(os.path.join(d, "**", name), recursive=True)
        if hits:
            return hits[0]
    return None


def _alias_builtin(alias, builtin):
    pdfmetrics.registerFont(pdfmetrics.Font(alias, builtin, "WinAnsiEncoding"))


def _register_fonts():
    for reg, bold, ital in _SANS_SETS:
        p_reg, p_bold, p_ital = _find(reg), _find(bold), _find(ital)
        if p_reg and p_bold and p_ital:
            pdfmetrics.registerFont(TTFont("S", p_reg))
            pdfmetrics.registerFont(TTFont("S-B", p_bold))
            pdfmetrics.registerFont(TTFont("S-I", p_ital))
            break
    else:
        _alias_builtin("S", "Helvetica")
        _alias_builtin("S-B", "Helvetica-Bold")
        _alias_builtin("S-I", "Helvetica-Oblique")
    pdfmetrics.registerFontFamily("S", normal="S", bold="S-B", italic="S-I")


_register_fonts()

INK = colors.HexColor("#151515")
GREY = colors.HexColor("#5E5E5E")
LIGHT = colors.HexColor("#9A9A9A")
RULE = colors.HexColor("#D5D5D5")
BOX = colors.HexColor("#F4F2ED")
BOXR = colors.HexColor("#E4DFD3")

W, H = landscape(A4)
ML = MR = 18 * mm
MT = 20 * mm
MB = 16 * mm
VB = W - ML - MR


def st(name, **kw):
    base = dict(fontName="S", fontSize=8.6, leading=12.4, textColor=INK,
                alignment=TA_LEFT, spaceAfter=0)
    base.update(kw)
    return ParagraphStyle(name, **base)


S = {
    "h1": st("h1", fontName="S-B", fontSize=21, leading=25, spaceAfter=3 * mm),
    "h2": st("h2", fontName="S-B", fontSize=13, leading=16, spaceBefore=7 * mm,
             spaceAfter=2 * mm),
    "h3": st("h3", fontName="S-B", fontSize=10.5, leading=14, spaceBefore=4 * mm,
             spaceAfter=1 * mm, keepWithNext=1),
    "p": st("p", fontSize=9.4, leading=13.6, spaceAfter=2.4 * mm),
    "lead": st("lead", fontSize=11, leading=16, textColor=GREY, spaceAfter=4 * mm),
    "tab": st("tab", fontSize=8.8, leading=12.2),
    "tabb": st("tabb", fontName="S-B", fontSize=8.8, leading=12.2),
    "tabh": st("tabh", fontName="S-B", fontSize=8.2, leading=11, textColor=GREY),
    "tabs": st("tabs", fontSize=8.2, leading=11.6, textColor=GREY),
    "boxh": st("boxh", fontName="S-B", fontSize=9.4, leading=13.6, spaceAfter=1.5 * mm),
    "boxp": st("boxp", fontSize=9.0, leading=13.4, spaceAfter=1.2 * mm),
}


def P(t, s="p"):
    return Paragraph(str(t), S[s])


def box(inhalt, farbe=BOX, rand=BOXR):
    t = Table([[inhalt]], colWidths=[VB])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), farbe),
        ("BOX", (0, 0), (-1, -1), 0.6, rand),
        ("LEFTPADDING", (0, 0), (-1, -1), 6 * mm),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 4 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3 * mm),
    ]))
    return t


def hinweis(titel, *absaetze):
    inner = [P(titel, "boxh")] + [P(a, "boxp") for a in absaetze]
    return KeepTogether(box(inner))


def tabelle(daten, breiten, kopf=True, zeilen_stil=None):
    rows = []
    for i, r in enumerate(daten):
        zeile = []
        for c in r:
            if hasattr(c, "wrapOn"):
                zeile.append(c)
            else:
                zeile.append(Paragraph(str(c), S["tabh" if (kopf and i == 0) else "tab"]))
        rows.append(zeile)
    t = Table(rows, colWidths=breiten, repeatRows=1 if kopf else 0)
    stil = [
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 2.1 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.1 * mm),
        ("LINEBELOW", (0, 0), (-1, -2 if kopf else -1), 0.4, RULE),
    ]
    if kopf:
        stil.append(("LINEBELOW", (0, 0), (-1, 0), 0.9, INK))
    for s in (zeilen_stil or []):
        stil.append(s)
    t.setStyle(TableStyle(stil))
    return t


_WT = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]


def tag_lang(datum):
    """'2026-09-12' -> 'Samstag, 12. September 2026'."""
    try:
        d = date.fromisoformat(str(datum))
    except ValueError:
        return str(datum)
    monate = ["", "Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
              "August", "September", "Oktober", "November", "Dezember"]
    return f"{_WT[d.weekday()]}, {d.day}. {monate[d.month]} {d.year}"


# ================================================================= Deckblatt
def _sperr(c, x, y, text, font, size, extra):
    c.setFont(font, size)
    for ch in str(text):
        c.drawString(x, y, ch)
        x += pdfmetrics.stringWidth(ch, font, size) + extra


def deckblatt(c, doc):
    """Ganzseitiges Titelblatt (A4 quer)."""
    c.saveState()
    v = D.VERANSTALTUNG

    y = H - 34 * mm
    c.setFillColor(LIGHT)
    _sperr(c, ML, y, v["titel"].upper(), "S-B", 10.5, 2)
    y -= 6 * mm
    c.setStrokeColor(INK)
    c.setLineWidth(1)
    c.line(ML, y, W - MR, y)

    y -= 30 * mm
    c.setFillColor(INK)
    c.setFont("S-B", 32)
    c.drawString(ML, y, "Gästeliste")
    y -= 11 * mm
    c.setFont("S", 13)
    c.setFillColor(GREY)
    c.drawString(ML, y, "Zusagen aus Politik und Öffentlichkeit")

    by = MB + 4 * mm
    c.setStrokeColor(RULE)
    c.setLineWidth(0.5)
    c.line(ML, by + 15 * mm, W - MR, by + 15 * mm)
    c.setFont("S", 9)
    c.setFillColor(GREY)
    c.drawString(ML, by + 8 * mm, v["untertitel"])
    c.drawString(ML, by + 3 * mm, f"{v['laufzeit']} · {v['ort']}")
    c.setFillColor(LIGHT)
    c.drawString(ML, by - 2.5 * mm,
                 f"{v['kontakt']} · intern · Stand {date.today():%d.%m.%Y}")
    c.restoreState()


def kopf_fuss(c, doc):
    c.saveState()
    seite = doc.page - 1
    if seite >= 1:
        c.setFont("S", 7.6)
        c.setFillColor(LIGHT)
        c.drawString(ML, H - MT + 9 * mm,
                     f"{D.VERANSTALTUNG['titel']} · Gästeliste · intern")
        c.setStrokeColor(RULE)
        c.setLineWidth(0.4)
        c.line(ML, H - MT + 6.5 * mm, W - MR, H - MT + 6.5 * mm)
        c.drawRightString(W - MR, MB - 9 * mm, f"Seite {seite}")
    c.setFont("S", 7.6)
    c.setFillColor(LIGHT)
    c.drawString(ML, MB - 9 * mm, D.VERANSTALTUNG["kontakt"])
    c.restoreState()


# ================================================================= Inhalt
def _sortkey(g):
    return (str(g.get("zeit") or "99:99"), str(g.get("name") or ""))


def inhalt():
    F = []
    a = F.append

    anlaesse = list(D.ANLAESSE.items())
    gaeste_je = {name: sorted((g for g in D.GAESTE if g.get("anlass") == name),
                              key=_sortkey)
                 for name, _ in anlaesse}
    ohne = [g for g in D.GAESTE if g.get("anlass") not in D.ANLAESSE]
    gesamt = len(D.GAESTE)

    a(Spacer(1, 1 * mm))

    # ---------------------------------------------------------- 1 Überblick
    a(P("Überblick", "h2"))
    daten = [["Anlass", "Termin", "Zusagen"]]
    for name, datum in anlaesse:
        daten.append([P(f"<b>{name}</b>", "tab"), P(tag_lang(datum), "tab"),
                      P(f"<b>{len(gaeste_je[name])}</b>", "tab")])
    if ohne:
        daten.append([P("<b>ohne Anlass</b>", "tab"), P("—", "tabs"),
                      P(f"<b>{len(ohne)}</b>", "tab")])
    daten.append([P("<b>Gesamt</b>", "tab"), "", P(f"<b>{gesamt}</b>", "tab")])
    stil = [("LINEABOVE", (0, len(daten) - 1), (-1, len(daten) - 1), 0.9, INK)]
    a(tabelle(daten, [60 * mm, 80 * mm, VB - 140 * mm], zeilen_stil=stil))

    grussw = [g for g in D.GAESTE if g.get("grusswort")]
    if grussw:
        a(Spacer(1, 3 * mm))
        namen = ", ".join(
            g["name"]
            + (" (" + ", ".join(t for t in (g.get("funktion"), g.get("organisation")) if t) + ")"
               if (g.get("funktion") or g.get("organisation")) else "")
            for g in grussw)
        a(hinweis("Grußwort", f"Ein Grußwort spricht: {namen}."))

    # ---------------------------------------------------------- Vernissage + Finissage
    # zusammen auf eine Seite (KeepTogether), getrennt von der Überblicksseite.
    spalten = [62 * mm, 78 * mm, VB - 165 * mm, 25 * mm]
    block = []
    for name, datum in anlaesse:
        gs = gaeste_je[name]
        if not gs:
            continue
        if block:
            block.append(Spacer(1, 6 * mm))
        block.append(P(name, "h2"))
        block.append(P(tag_lang(datum), "tabs"))
        daten = [["Name", "Funktion", "Organisation", "Zeit"]]
        for g in gs:
            daten.append([
                P(f"<b>{g.get('name', '')}</b>", "tab"),
                P(g.get("funktion") or "—", "tab"),
                P(g.get("organisation") or "—", "tab"),
                P(g.get("zeit") or "—", "tab"),
            ])
        block.append(tabelle(daten, spalten))
    if block:
        a(PageBreak())
        a(KeepTogether(block))

    if ohne:
        a(PageBreak())
        a(P("Ohne festen Anlass", "h2"))
        daten = [["Name", "Funktion", "Organisation", "Termin"]]
        for g in sorted(ohne, key=_sortkey):
            daten.append([
                P(f"<b>{g.get('name', '')}</b>", "tab"),
                P(g.get("funktion") or "—", "tab"),
                P(g.get("organisation") or "—", "tab"),
                P(g.get("datum") or "—", "tab"),
            ])
        a(tabelle(daten, spalten))

    return F


def main():
    doc = BaseDocTemplate(OUT, pagesize=landscape(A4),
                          leftMargin=ML, rightMargin=MR, topMargin=MT, bottomMargin=MB,
                          title=f"{D.VERANSTALTUNG['titel']} — Gästeliste",
                          author="KULTURVEREIN MUSTERSTADT")
    frame = Frame(ML, MB, W - ML - MR, H - MT - MB, id="f",
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([
        PageTemplate(id="cover", frames=[Frame(0, 0, W, H, id="cv")], onPage=deckblatt),
        PageTemplate(id="main", frames=[frame], onPage=kopf_fuss),
    ])
    doc.build([NextPageTemplate("main"), Spacer(1, 2), PageBreak()] + inhalt())

    print(f"{OUT} geschrieben.")
    for name, datum in D.ANLAESSE.items():
        gs = [g for g in D.GAESTE if g.get("anlass") == name]
        print(f"  {name:12s} {len(gs):>2} Zusage(n)")
    print(f"  {'Gesamt':12s} {len(D.GAESTE):>2}")


if __name__ == "__main__":
    main()
