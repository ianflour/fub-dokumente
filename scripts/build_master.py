#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Erzeugt FUB2026_Werkdaten_MASTER.xlsx
FOTOTAGE MUSTERSTADT 2026 - Konzept A (Wandschild + Werkliste)

ACHTUNG: Baut die Datei komplett neu (nur die 14 Beispielwerke) — sobald echte
Werke eingelesen sind (scripts/import_werkmeldungen.py), würde ein Neubau sie
löschen. Das Script verweigert sich deshalb, wenn schon echte Werke im Master
stehen; nur mit --force wird trotzdem überschrieben.
"""

import argparse
import os
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import _pfade as P
import _saal_listen as SL
import _teilnehmende as T

OUT = P.output("FUB2026_Werkdaten_MASTER.xlsx")


def _hat_echte_werke(pfad):
    """True, wenn in pfad -> Blatt "Werke" mindestens eine Zeile ohne
    Beispiel-Markierung steht (also schon eingelesene, echte Werkmeldungen)."""
    if not os.path.exists(pfad):
        return False
    try:
        wb = load_workbook(pfad, data_only=True)
    except Exception:
        return False
    if "Werke" not in wb.sheetnames:
        return False
    ws = wb["Werke"]
    header = {c.value: c.column for c in ws[1] if c.value}
    if "Künstler:in" not in header:
        return False
    sp_beispiel = header.get("Beispiel")
    for r in range(2, ws.max_row + 1):
        if not ws.cell(row=r, column=header["Künstler:in"]).value:
            continue
        bsp = ws.cell(row=r, column=sp_beispiel).value if sp_beispiel else None
        if str(bsp or "").strip().lower() not in ("ja", "yes", "true", "1", "x"):
            return True
    return False

FONT = "Arial"
C_HEAD = "1A1A1A"
C_INPUT = "0000FF"      # blau = einzutragen
C_FORMULA = "000000"    # schwarz = Formel, nicht anfassen
FILL_HEAD = PatternFill("solid", fgColor="D9D9D9")
FILL_KEY = PatternFill("solid", fgColor="FFF2CC")
FILL_LOCK = PatternFill("solid", fgColor="F2F2F2")

thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# ---------------------------------------------------------------- Spalten
# (Header, Breite, Typ)  Typ: in = Eingabe, fx = Formel, opt = optional
COLS = [
    ("Nr",                     6,  "in"),
    ("Wand",                   7,  "in"),
    ("Position",               9,  "in"),
    ("Künstler:in",           22,  "in"),
    ("Titel",                 30,  "in"),
    ("Ort",                   16,  "in"),
    ("Land",                  13,  "in"),
    ("Jahr",                   7,  "in"),
    ("Druckverfahren",        16,  "in"),
    ("Papier",                34,  "in"),
    ("Bildmaß H cm",          12,  "in"),
    ("Bildmaß B cm",          12,  "in"),
    ("Blattmaß H cm",         13,  "in"),
    ("Blattmaß B cm",         13,  "in"),
    ("Maßangabe",             20,  "fx"),
    ("Auflage",                8,  "in"),
    ("AP",                     6,  "in"),
    ("Exemplar",              10,  "in"),
    ("Auflageangabe",         16,  "fx"),
    ("Rahmenmodell",          24,  "in"),
    ("Rahmenmaß",             12,  "in"),
    ("Passepartout",          20,  "in"),
    ("Preis EUR",             11,  "in"),
    ("Rahmen im Preis",       16,  "in"),
    ("Versicherungswert EUR", 20,  "in"),
    ("Status",                14,  "in"),
    ("Käufer:in",             20,  "opt"),
    ("Notiz",                 28,  "opt"),
    ("Beispiel",              10,  "opt"),
]

# ---------------------------------------------------------------- Beispieldaten
# 14 Beispielwerke, 5 Künstler:innen - realistische Formate für Fotografie
WERKE = [
    ("001","A","1","Lena Hoffmann","Auto eines alten Fischers","Bari","Italien",2021,
     "C-Print","Fujicolor Crystal Archive DPII lustre",27,40,30,45,5,1,"2/5",
     "Boesner Uno 40×50 cm, schwarz","40×50 cm","Dorée, 3 mm Weißkern",75,"ja",120,"verfügbar","",""),
    ("002","A","2","Lena Hoffmann","Hafenbecken, morgens","Bari","Italien",2021,
     "C-Print","Fujicolor Crystal Archive DPII lustre",27,40,30,45,5,1,"1/5",
     "Boesner Uno 40×50 cm, schwarz","40×50 cm","Dorée, 3 mm Weißkern",75,"ja",120,"verfügbar","",""),
    ("003","A","3","Lena Hoffmann","Netzflicker","Molfetta","Italien",2022,
     "C-Print","Fujicolor Crystal Archive DPII lustre",40,27,45,30,5,1,"1/5",
     "Boesner Uno 40×50 cm, schwarz","40×50 cm","Dorée, 3 mm Weißkern",75,"ja",120,"reserviert","",
     "Anfrage per Mail 14.08."),
    ("004","B","1","Tomas Weiler","Zwischenraum I","Musterstadt","Deutschland",2024,
     "Pigmentdruck","Hahnemühle Photo Rag 308 g",45,60,50,65,3,1,"1/3",
     "Nielsen Alpha 50×70 cm, alu roh","50×70 cm","ohne",240,"ja",380,"verfügbar","",""),
    ("005","B","2","Tomas Weiler","Zwischenraum II","Musterstadt","Deutschland",2024,
     "Pigmentdruck","Hahnemühle Photo Rag 308 g",45,60,50,65,3,1,"1/3",
     "Nielsen Alpha 50×70 cm, alu roh","50×70 cm","ohne",240,"ja",380,"verkauft","M. Beispiel",
     "Abholung nach Finissage"),
    ("006","B","3","Tomas Weiler","Zwischenraum III","Völklingen","Deutschland",2024,
     "Pigmentdruck","Hahnemühle Photo Rag 308 g",45,60,50,65,3,1,"1/3",
     "Nielsen Alpha 50×70 cm, alu roh","50×70 cm","ohne",240,"ja",380,"verfügbar","",""),
    ("007","B","4","Tomas Weiler","Halde, Abend","Völklingen","Deutschland",2023,
     "Pigmentdruck","Hahnemühle Photo Rag 308 g",30,45,35,50,3,1,"2/3",
     "Nielsen Alpha 40×60 cm, alu roh","40×60 cm","ohne",180,"ja",280,"verfügbar","",""),
    ("008","C","1","Aylin Demir","Ohne Titel (Küche)","Istanbul","Türkei",2023,
     "Silbergelatine, Handabzug","Ilford Multigrade RC Deluxe",24,30,30,40,8,2,"3/8",
     "Halbe Magnetrahmen 30×40 cm","30×40 cm","weiß, 2 mm",140,"ja",200,"verfügbar","",""),
    ("009","C","2","Aylin Demir","Ohne Titel (Fenster)","Istanbul","Türkei",2023,
     "Silbergelatine, Handabzug","Ilford Multigrade RC Deluxe",24,30,30,40,8,2,"2/8",
     "Halbe Magnetrahmen 30×40 cm","30×40 cm","weiß, 2 mm",140,"ja",200,"verfügbar","",""),
    ("010","C","3","Aylin Demir","Ohne Titel (Treppe)","Ankara","Türkei",2024,
     "Silbergelatine, Handabzug","Ilford Multigrade RC Deluxe",30,24,40,30,8,2,"1/8",
     "Halbe Magnetrahmen 30×40 cm","30×40 cm","weiß, 2 mm",None,"",200,"nicht verkäuflich","",
     "Leihgabe der Künstlerin"),
    ("011","D","1","Marek Sobota","Nachtschicht","Neunkirchen","Deutschland",2025,
     "Inkjet","Canson Baryta Photographique II 310 g",50,75,55,80,5,1,"1/5",
     "Rahmenwerkstatt Sonderanfertigung, Eiche","55×80 cm","ohne",420,"ja",600,"verfügbar","",""),
    ("012","D","2","Marek Sobota","Nachtschicht II","Neunkirchen","Deutschland",2025,
     "Inkjet","Canson Baryta Photographique II 310 g",50,75,55,80,5,1,"1/5",
     "Rahmenwerkstatt Sonderanfertigung, Eiche","55×80 cm","ohne",420,"ja",600,"verfügbar","",""),
    ("013","E","1","Sophie Braun","Wildcard: Grenzverlauf","Perl","Deutschland",2026,
     "Pigmentdruck","Hahnemühle Photo Rag Baryta 315 g",30,30,40,40,3,1,"1/3",
     "Boesner Uno 40×40 cm, weiß","40×40 cm","ohne",160,"ja",240,"verfügbar","",
     "HBKsaar-Wildcard"),
    ("014","E","2","Sophie Braun","Wildcard: Grenzverlauf II","Schengen","Luxemburg",2026,
     "Pigmentdruck","Hahnemühle Photo Rag Baryta 315 g",30,30,40,40,3,1,"1/3",
     "Boesner Uno 40×40 cm, weiß","40×40 cm","ohne",160,"ja",240,"verfügbar","",
     "HBKsaar-Wildcard"),
]
# Beispiel-Flag anhängen: markiert alle 14 Zeilen oben als Demodaten, damit
# der Generator sie automatisch weglässt, sobald echte Werke importiert sind.
WERKE = [t + ("ja",) for t in WERKE]

# Erst fünf Beispiel-Künstler:innen (gehören zu den 14 Beispielwerken oben —
# beim Start der echten Arbeit mitsamt den Beispielwerken löschen), dann die
# echten Teilnehmenden aus Teilnehmende.txt inkl. Instagram-Handle.
_DEMO = [
    ("Lena Hoffmann","lena.hoffmann@example.org","+49 170 0000001","@lenahoffmann.foto","Musterstadt",50,"ja"),
    ("Tomas Weiler","t.weiler@example.org","+49 170 0000002","@tomasweiler","Saarlouis",50,"ja"),
    ("Aylin Demir","aylin.demir@example.org","+49 170 0000003","@aylin.d.photo","Musterstadt",50,"ja"),
    ("Marek Sobota","m.sobota@example.org","+49 170 0000004","@sobota.works","Neunkirchen",50,"nein"),
    ("Sophie Braun","s.braun@example.org","+49 170 0000005","@sophie.br","Musterstadt",0,"ja"),
]
_ECHT = [
    (n, "", "",
     ("@" + T.INSTAGRAM[n]["handle"]) if (T.INSTAGRAM.get(n) or {}).get("handle") else "",
     "", None, "")
    for n in T.NAMEN
]
KUENSTLER = _DEMO + _ECHT


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT, size=9, bold=True, color=C_HEAD)
        cell.fill = FILL_HEAD
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 30


def build():
    wb = Workbook()

    # =============================================================== ANLEITUNG
    ws = wb.active
    ws.title = "ANLEITUNG"
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 105

    lines = [
        ("FOTOTAGE MUSTERSTADT 2026 — Werkdaten-Master", ""),
        ("", ""),
        ("Zweck", "Diese Datei ist die EINZIGE Quelle für alle Werkdaten. Aus ihr entstehen: "
                  "Wandschilder, Werkliste (Print), Rückseitenetiketten, Versicherungsliste."),
        ("", "Wird ein Preis geändert, wird er NUR hier geändert. Danach Generator neu laufen lassen."),
        ("", ""),
        ("Blätter", ""),
        ("Werke", "Alle Exponate, eine Zeile pro Werk. Hier arbeitest du zu 95 %."),
        ("Teilnehmende", "Stammdaten der Teilnehmenden, Kontakt und Provision."),
        ("Übersicht", "Automatische Auswertung. Nichts eintragen."),
        ("Listen", "Quelle für die Dropdowns. Nur erweitern, nicht umbenennen."),
        ("", ""),
        ("Farbcode", ""),
        ("blaue Schrift", "Eingabefeld — hier trägst du ein."),
        ("schwarze Schrift", "Formel — nicht überschreiben."),
        ("gelbe Füllung", "Pflichtfeld für den Druck. Bleibt es leer, fehlt es auf dem Schild."),
        ("graue Füllung", "Automatisch berechnet."),
        ("", ""),
        ("Pflichtfelder Wandschild", "Nr · Künstler:in · Titel · Ort · Land · Jahr · Druckverfahren · Papier · "
                                     "Bildmaß H/B · Auflage"),
        ("Pflichtfelder Werkliste", "zusätzlich: Preis EUR · Rahmen im Preis · Status"),
        ("Nur intern", "Rahmenmodell · Rahmenmaß · Passepartout · Versicherungswert · Käufer:in · Notiz"),
        ("", "Diese Felder erscheinen NICHT auf dem Wandschild."),
        ("", ""),
        ("Schreibregeln", ""),
        ("Nr", "Immer dreistellig mit führenden Nullen: 001, 002 … 100. Als Text formatiert."),
        ("Künstler:in", "Vorname Nachname, genau wie im Blatt Künstler:innen geschrieben."),
        ("Titel", "Ohne Anführungszeichen. Unbetitelt = „Ohne Titel“ oder „Ohne Titel (Küche)“."),
        ("Maße", "Immer Höhe × Breite in cm, als Zahl (nicht „30 cm“). Bildmaß = sichtbares Bild "
                 "ohne Weißrand. Blattmaß = Papier inkl. Weißrand."),
        ("Auflage / AP", "Auflage = Editionsgröße (Zahl). AP = Anzahl Artist Proofs (Zahl, sonst 0)."),
        ("Exemplar", "Welches Exemplar hängt, z. B. 2/5. Erscheint NICHT auf dem Schild, nur intern."),
        ("Preis EUR", "Endpreis in Euro als Zahl, ohne € und ohne Punkt. Unverkäuflich = leer lassen "
                      "und Status auf „nicht verkäuflich“."),
        ("Rahmen im Preis", "ja / nein. Steht so in der Werkliste. Häufigste Rückfrage im Verkaufsgespräch."),
        ("Status", "verfügbar · reserviert · verkauft · nicht verkäuflich"),
        ("", ""),
        ("Beispieldaten", "Die 14 Werke von 5 Beispiel-Künstler:innen ganz oben sind Demodaten "
                          "(Spalte „Beispiel“ = ja) und dienen nur zum Testen. Sobald echte Werke "
                          "importiert wurden, lässt der Generator sie automatisch aus dem Druck "
                          "weg — von Hand löschen musst du sie nicht, kannst es aber jederzeit."),
        ("", ""),
        ("Nächster Schritt", "./.venv/bin/python scripts/fub.py  (Menüpunkt 2: Druckdaten erzeugen)"),
        ("Hinweis MwSt.", "Preisangaben sind Endpreise. Bei Kleinunternehmerregelung nach § 19 UStG "
                          "keine USt. ausweisen — Hinweis steht im Fuß der Werkliste und ist im "
                          "Generator anpassbar (--ust-hinweis)."),
        ("Hinweis Recht", "Kunstgegenstände sind nach § 9 Abs. 7 PAngV von der Preisauszeichnungspflicht "
                          "ausgenommen. Die ausliegende Werkliste ist ausreichend."),
    ]
    for i, (a, b) in enumerate(lines, start=1):
        ws.cell(row=i, column=1, value=a).font = Font(
            name=FONT, size=10, bold=(b == "" and a != ""), color=C_HEAD)
        ws.cell(row=i, column=2, value=b).font = Font(name=FONT, size=10)
        ws.cell(row=i, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row=1, column=1).font = Font(name=FONT, size=14, bold=True)
    ws.sheet_view.showGridLines = False

    # =============================================================== LISTEN
    wl = wb.create_sheet("Listen")
    listen = {
        "A": ("Status", ["verfügbar", "reserviert", "verkauft", "nicht verkäuflich"]),
        "B": ("Rahmen im Preis", list(SL.JA_NEIN)),
        "C": ("Druckverfahren", list(SL.DRUCKVERFAHREN)),
        "D": ("Wand", ["A", "B", "C", "D", "E", "F", "G", "H"]),
        "E": ("Papier", list(SL.PAPIERE)),
        "F": ("Maße cm", list(SL.MASSE_CM)),
    }
    listen_bereich = {}
    for col, (title, vals) in listen.items():
        wl[f"{col}1"] = title
        wl[f"{col}1"].font = Font(name=FONT, size=9, bold=True)
        wl[f"{col}1"].fill = FILL_HEAD
        for i, v in enumerate(vals, start=2):
            wl[f"{col}{i}"] = v
            wl[f"{col}{i}"].font = Font(name=FONT, size=10)
        wl.column_dimensions[col].width = 44 if col == "E" else 26
        listen_bereich[title] = f"Listen!${col}$2:${col}${len(vals) + 1}"
    wl.sheet_view.showGridLines = False

    # =============================================================== WERKE
    w = wb.create_sheet("Werke")
    headers = [c[0] for c in COLS]
    for i, h in enumerate(headers, start=1):
        w.cell(row=1, column=i, value=h)
        w.column_dimensions[get_column_letter(i)].width = COLS[i - 1][1]
    style_header(w, 1, len(COLS))

    # Pflichtfelder gelb markieren (Header)
    pflicht = {"Nr", "Künstler:in", "Titel", "Ort", "Land", "Jahr", "Druckverfahren", "Papier",
               "Bildmaß H cm", "Bildmaß B cm", "Auflage", "Preis EUR", "Rahmen im Preis", "Status"}
    for i, h in enumerate(headers, start=1):
        if h in pflicht:
            w.cell(row=1, column=i).fill = FILL_KEY

    idx = {h: i + 1 for i, h in enumerate(headers)}

    def put(r, name, value, kind="in"):
        c = w.cell(row=r, column=idx[name], value=value)
        c.font = Font(name=FONT, size=10,
                      color=C_INPUT if kind == "in" else C_FORMULA)
        c.alignment = Alignment(vertical="center", wrap_text=(name in ("Papier", "Rahmenmodell", "Notiz", "Titel")))
        c.border = BORDER
        if kind == "fx":
            c.fill = FILL_LOCK
        return c

    ROWS = 120  # Platz für 100+ Werke
    for n in range(ROWS):
        r = n + 2
        data = WERKE[n] if n < len(WERKE) else None
        vals = dict(zip(
            ["Nr", "Wand", "Position", "Künstler:in", "Titel", "Ort", "Land", "Jahr",
             "Druckverfahren", "Papier", "Bildmaß H cm", "Bildmaß B cm", "Blattmaß H cm",
             "Blattmaß B cm", "Auflage", "AP", "Exemplar", "Rahmenmodell", "Rahmenmaß",
             "Passepartout", "Preis EUR", "Rahmen im Preis", "Versicherungswert EUR",
             "Status", "Käufer:in", "Notiz", "Beispiel"],
            data)) if data else {}

        for h, _wdt, kind in COLS:
            if kind == "fx":
                continue
            put(r, h, vals.get(h, None), "in")

        # Formeln: Maßangabe / Auflageangabe
        hb = f"{get_column_letter(idx['Bildmaß H cm'])}{r}"
        bb = f"{get_column_letter(idx['Bildmaß B cm'])}{r}"
        au = f"{get_column_letter(idx['Auflage'])}{r}"
        ap = f"{get_column_letter(idx['AP'])}{r}"
        put(r, "Maßangabe",
            f'=IF({hb}="","",{hb}&" × "&{bb}&" cm")', "fx")
        put(r, "Auflageangabe",
            f'=IF({au}="","",IF(OR({ap}="",{ap}=0),"Auflage "&{au},'
            f'"Auflage "&{au}&" (+"&{ap}&" AP)"))', "fx")

        w.row_dimensions[r].height = 28

    # Zahlenformate
    for name, fmt in [("Preis EUR", '#,##0 "€"'), ("Versicherungswert EUR", '#,##0 "€"'),
                      ("Bildmaß H cm", "0.#"), ("Bildmaß B cm", "0.#"),
                      ("Blattmaß H cm", "0.#"), ("Blattmaß B cm", "0.#"),
                      ("Jahr", "0"), ("Auflage", "0"), ("AP", "0")]:
        col = get_column_letter(idx[name])
        for r in range(2, ROWS + 2):
            w[f"{col}{r}"].number_format = fmt
    col_nr = get_column_letter(idx["Nr"])
    for r in range(2, ROWS + 2):
        w[f"{col_nr}{r}"].number_format = "@"

    # Dropdowns
    dvs = [
        ("Status", listen_bereich["Status"], True),
        ("Rahmen im Preis", listen_bereich["Rahmen im Preis"], True),
        ("Druckverfahren", listen_bereich["Druckverfahren"], False),
        ("Wand", listen_bereich["Wand"], True),
        ("Papier", listen_bereich["Papier"], False),
        ("Bildmaß H cm", listen_bereich["Maße cm"], False),
        ("Bildmaß B cm", listen_bereich["Maße cm"], False),
        ("Blattmaß H cm", listen_bereich["Maße cm"], False),
        ("Blattmaß B cm", listen_bereich["Maße cm"], False),
    ]
    for name, src, streng in dvs:
        # Bereichsbezug ohne führendes "=" (Excel repariert es still, andere
        # Programme wie Numbers oder Google Sheets verwerfen die Regel).
        dv = DataValidation(type="list", formula1=src, allow_blank=True,
                            showDropDown=False, showErrorMessage=streng)
        w.add_data_validation(dv)
        col = get_column_letter(idx[name])
        dv.add(f"{col}2:{col}{ROWS + 1}")

    w.freeze_panes = "E2"
    w.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{ROWS + 1}"

    # =============================================================== KÜNSTLER
    k = wb.create_sheet("Teilnehmende")
    kh = ["Künstler:in", "E-Mail", "Telefon", "Instagram", "Wohnort",
          "Provision %", "Werkliste erhalten", "Anzahl Werke", "Summe Preise EUR"]
    for i, h in enumerate(kh, start=1):
        k.cell(row=1, column=i, value=h)
        k.column_dimensions[get_column_letter(i)].width = [24, 30, 18, 22, 16, 12, 18, 14, 18][i - 1]
    style_header(k, 1, len(kh))

    wcol_k = get_column_letter(idx["Künstler:in"])
    wcol_p = get_column_letter(idx["Preis EUR"])
    for n in range(40):
        r = n + 2
        d = KUENSTLER[n] if n < len(KUENSTLER) else ("", "", "", "", "", None, "")
        for i, v in enumerate(d, start=1):
            c = k.cell(row=r, column=i, value=v if v != "" else None)
            c.font = Font(name=FONT, size=10, color=C_INPUT)
            c.border = BORDER
        k.cell(row=r, column=8,
               value=f'=IF(A{r}="","",COUNTIF(Werke!${wcol_k}$2:${wcol_k}${ROWS+1},A{r}))')
        k.cell(row=r, column=9,
               value=f'=IF(A{r}="","",SUMIF(Werke!${wcol_k}$2:${wcol_k}${ROWS+1},A{r},'
                     f'Werke!${wcol_p}$2:${wcol_p}${ROWS+1}))')
        for i in (8, 9):
            c = k.cell(row=r, column=i)
            c.font = Font(name=FONT, size=10, color=C_FORMULA)
            c.fill = FILL_LOCK
            c.border = BORDER
        k.cell(row=r, column=9).number_format = '#,##0 "€"'
        k.cell(row=r, column=6).number_format = '0"%"'
    k.freeze_panes = "A2"

    # =============================================================== ÜBERSICHT
    u = wb.create_sheet("Übersicht")
    u.column_dimensions["A"].width = 34
    u.column_dimensions["B"].width = 18
    u.column_dimensions["C"].width = 60
    u["A1"] = "Übersicht — automatisch"
    u["A1"].font = Font(name=FONT, size=14, bold=True)

    scol = get_column_letter(idx["Status"])
    vcol = get_column_letter(idx["Versicherungswert EUR"])
    tcol = get_column_letter(idx["Titel"])

    ue = [
        ("Werke erfasst",            f'=COUNTA(Werke!${tcol}$2:${tcol}${ROWS+1})', "0",
         "Zeilen mit ausgefülltem Titel"),
        ("davon verfügbar",          f'=COUNTIF(Werke!${scol}$2:${scol}${ROWS+1},"verfügbar")', "0", ""),
        ("davon reserviert",         f'=COUNTIF(Werke!${scol}$2:${scol}${ROWS+1},"reserviert")', "0", ""),
        ("davon verkauft",           f'=COUNTIF(Werke!${scol}$2:${scol}${ROWS+1},"verkauft")', "0", ""),
        ("davon nicht verkäuflich",  f'=COUNTIF(Werke!${scol}$2:${scol}${ROWS+1},"nicht verkäuflich")', "0", ""),
        ("Künstler:innen",           '=COUNTA(Teilnehmende!$A$2:$A$41)', "0", ""),
        ("", "", "", ""),
        ("Summe Verkaufspreise",     f'=SUM(Werke!${wcol_p}$2:${wcol_p}${ROWS+1})', '#,##0 "€"',
         "alle erfassten Werke"),
        ("Umsatz verkauft",          f'=SUMIF(Werke!${scol}$2:${scol}${ROWS+1},"verkauft",'
                                     f'Werke!${wcol_p}$2:${wcol_p}${ROWS+1})', '#,##0 "€"', ""),
        ("Versicherungswert gesamt", f'=SUM(Werke!${vcol}$2:${vcol}${ROWS+1})', '#,##0 "€"',
         "Wert für die Ausstellungsversicherung — an Räumchen e.V. melden"),
        ("", "", "", ""),
        ("Druckbögen Wandschilder A7", f'=ROUNDUP(COUNTA(Werke!${tcol}$2:${tcol}${ROWS+1})/8,0)', "0",
         "8 Schilder pro A4-Bogen, ohne Reserve"),
        ("Bögen inkl. 15 % Reserve", f'=ROUNDUP(COUNTA(Werke!${tcol}$2:${tcol}${ROWS+1})*1.15/8,0)', "0",
         "so viel Karton 300 g bestellen"),
    ]
    r = 3
    for label, formula, fmt, note in ue:
        if label == "":
            r += 1
            continue
        u.cell(row=r, column=1, value=label).font = Font(name=FONT, size=10, bold=True)
        c = u.cell(row=r, column=2, value=formula)
        c.font = Font(name=FONT, size=10)
        c.number_format = fmt
        c.fill = FILL_LOCK
        c.alignment = Alignment(horizontal="right")
        u.cell(row=r, column=3, value=note).font = Font(name=FONT, size=9, color="666666")
        r += 1
    u.sheet_view.showGridLines = False

    wb.save(OUT)
    print("geschrieben:", OUT)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true",
                    help="überschreiben, auch wenn schon echte Werke im Master stehen")
    args = ap.parse_args()
    if not args.force and _hat_echte_werke(OUT):
        sys.exit(
            f"FEHLER: {P.rel(OUT)} enthält schon echte Werke (nicht nur Beispiele) — "
            "ein Neubau würde sie löschen.\n"
            "Neue Meldungen kommen über 'Werkmeldungen einlesen' (fub.py 3) rein, "
            "ohne Bestehendes zu überschreiben. Nur mit --force trotzdem neu bauen."
        )
    build()
