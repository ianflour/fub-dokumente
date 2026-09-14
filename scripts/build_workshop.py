#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Erzeugt WORKSHOP.pdf (A4 quer) aus workshop_daten.py.

Anmeldeliste für den Workshop Analoge Fotografie: Deckblatt + eine Seite mit
den Terminen und einer Liste zum Eintragen. Nur workshop_daten.py ändern.
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
from reportlab.platypus import (BaseDocTemplate, Frame, NextPageTemplate,
                                PageBreak, PageTemplate, Paragraph, Spacer, Table,
                                TableStyle)

import _pfade as P
import workshop_daten as D

OUT = P.output("WORKSHOP.pdf")

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

W, H = landscape(A4)
ML = MR = 18 * mm
MT = 20 * mm
MB = 16 * mm
VB = W - ML - MR

_WT_KURZ = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
_WT_LANG = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]
_MON = ["", "Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
        "August", "September", "Oktober", "November", "Dezember"]


def _d(datum):
    try:
        return date.fromisoformat(str(datum))
    except ValueError:
        return None


def tag_kurz(datum):
    d = _d(datum)
    return f"{_WT_KURZ[d.weekday()]} {d:%d.%m.}" if d else str(datum)


def tag_lang(datum):
    d = _d(datum)
    return f"{_WT_LANG[d.weekday()]}, {d.day}. {_MON[d.month]} {d.year}" if d else str(datum)


def st(name, **kw):
    base = dict(fontName="S", fontSize=8.6, leading=12.4, textColor=INK,
                alignment=TA_LEFT, spaceAfter=0)
    base.update(kw)
    return ParagraphStyle(name, **base)


S = {
    "h2": st("h2", fontName="S-B", fontSize=13, leading=16, spaceBefore=7 * mm,
             spaceAfter=2 * mm),
    "p": st("p", fontSize=9.4, leading=13.6, spaceAfter=2.4 * mm),
    "tab": st("tab", fontSize=9.2, leading=13),
    "tabh": st("tabh", fontName="S-B", fontSize=8.2, leading=11, textColor=GREY),
    "tabc": st("tabc", fontSize=9.2, leading=13, alignment=1),
}


def Par(t, s="p"):
    return Paragraph(str(t), S[s])


# ================================================================= Deckblatt
def _sperr(c, x, y, text, font, size, extra):
    c.setFont(font, size)
    for ch in str(text):
        c.drawString(x, y, ch)
        x += pdfmetrics.stringWidth(ch, font, size) + extra


def deckblatt(c, doc):
    c.saveState()
    v, ws = D.VERANSTALTUNG, D.WORKSHOP

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
    c.drawString(ML, y, "Workshop")
    y -= 13 * mm
    c.setFont("S-B", 20)
    c.setFillColor(INK)
    c.drawString(ML, y, ws["titel"])
    y -= 10 * mm
    c.setFont("S", 12)
    c.setFillColor(GREY)
    c.drawString(ML, y, f"Anmeldeliste · Leitung {ws['leitung']}")

    y -= 12 * mm
    c.setFont("S", 11.5)
    c.setFillColor(INK)
    for dt, zeit, was in ws["termine"]:
        c.drawString(ML, y, f"{tag_lang(dt)} · {zeit} — {was}")
        y -= 6.5 * mm

    by = MB + 4 * mm
    c.setStrokeColor(RULE)
    c.setLineWidth(0.5)
    c.line(ML, by + 12 * mm, W - MR, by + 12 * mm)
    c.setFont("S", 9)
    c.setFillColor(GREY)
    c.drawString(ML, by + 5 * mm, ws["ort"])
    c.setFillColor(LIGHT)
    c.drawString(ML, by, f"Stand {date.today():%d.%m.%Y}")
    c.restoreState()


def kopf_fuss(c, doc):
    c.saveState()
    seite = doc.page - 1
    if seite >= 1:
        c.setFont("S", 7.6)
        c.setFillColor(LIGHT)
        c.drawString(ML, H - MT + 9 * mm,
                     f"{D.VERANSTALTUNG['titel']} · Workshop {D.WORKSHOP['titel']}")
        c.setStrokeColor(RULE)
        c.setLineWidth(0.4)
        c.line(ML, H - MT + 6.5 * mm, W - MR, H - MT + 6.5 * mm)
    c.setFont("S", 7.6)
    c.setFillColor(LIGHT)
    c.drawString(ML, MB - 9 * mm, D.VERANSTALTUNG["kontakt"])
    c.restoreState()


# ================================================================= Inhalt
def tabelle(daten, breiten, hoehe=None):
    rows = []
    for i, r in enumerate(daten):
        rows.append([c if hasattr(c, "wrapOn")
                     else Paragraph(str(c), S["tabh" if i == 0 else "tab"]) for c in r])
    t = Table(rows, colWidths=breiten, rowHeights=hoehe, repeatRows=1)
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("VALIGN", (0, 0), (-1, 0), "BOTTOM"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING", (1, 0), (-1, -1), 2 * mm),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2 * mm),
        ("TOPPADDING", (0, 0), (-1, 0), 2 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 2 * mm),
        ("LINEBELOW", (0, 0), (-1, 0), 0.9, INK),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, RULE),
        ("LINEAFTER", (0, 0), (-2, -1), 0.4, RULE),
        ("BACKGROUND", (0, 0), (-1, 0), colors.white),
    ]))
    return t


def inhalt():
    F = []
    a = F.append
    ws = D.WORKSHOP

    a(Spacer(1, 1 * mm))
    a(Par("Anmeldung", "h2"))
    a(Spacer(1, 2 * mm))

    daten = [["Nr.", "Name", "Kontakt (Instagram / E-Mail / Telefon)"]]
    eintraege = list(D.ANMELDUNGEN)
    for i in range(1, max(ws["plaetze"], len(eintraege)) + 1):
        e = eintraege[i - 1] if i <= len(eintraege) else {}
        daten.append([Paragraph(str(i), S["tabc"]),
                      e.get("name", ""),
                      e.get("kontakt", "")])

    breiten = [12 * mm, 78 * mm, VB - 12 * mm - 78 * mm]
    a(tabelle(daten, breiten, hoehe=[9 * mm] + [12 * mm] * (len(daten) - 1)))

    infos = [e for e in D.ANMELDUNGEN if e.get("info")]
    if infos:
        a(PageBreak())
        a(Par("Ausrüstung der Teilnehmenden", "h2"))
        a(Spacer(1, 2 * mm))
        idaten = [["Name", "Kamera / Anmerkung"]]
        for e in infos:
            idaten.append([e.get("name", ""), e["info"]])
        ibreiten = [50 * mm, VB - 50 * mm]
        a(tabelle(idaten, ibreiten,
                  hoehe=[9 * mm] + [14 * mm] * (len(idaten) - 1)))

    return F


def main():
    doc = BaseDocTemplate(OUT, pagesize=landscape(A4),
                          leftMargin=ML, rightMargin=MR, topMargin=MT, bottomMargin=MB,
                          title=f"{D.VERANSTALTUNG['titel']} — Workshop {D.WORKSHOP['titel']}",
                          author="KULTURVEREIN MUSTERSTADT")
    frame = Frame(ML, MB, W - ML - MR, H - MT - MB, id="f",
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([
        PageTemplate(id="cover", frames=[Frame(0, 0, W, H, id="cv")], onPage=deckblatt),
        PageTemplate(id="main", frames=[frame], onPage=kopf_fuss),
    ])
    doc.build([NextPageTemplate("main"), Spacer(1, 2), PageBreak()] + inhalt())
    print(f"{OUT} geschrieben. {D.WORKSHOP['plaetze']} Plätze, "
          f"{len(D.ANMELDUNGEN)} schon eingetragen.")


if __name__ == "__main__":
    main()
