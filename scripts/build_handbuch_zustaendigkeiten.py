#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Erzeugt output/HANDBUCH_ZUSTAENDIGKEITEN.pdf — wer öffnet was womit (macOS)."""

import glob
import os
from datetime import date

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, Frame, NextPageTemplate, PageBreak,
                                PageTemplate, Paragraph, Spacer, Table, TableStyle,
                                KeepTogether)

import _pfade as P

OUT = P.output("HANDBUCH_ZUSTAENDIGKEITEN.pdf")

# --- Schriften plattformunabhängig suchen (Liberation Sans / Arial / DejaVu),
#     sonst eingebaute PDF-Standardschriften Helvetica / Courier -------------
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
_MONO_NAMES = ["LiberationMono-Regular.ttf", "DejaVuSansMono.ttf", "Courier New.ttf",
               "cour.ttf"]


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

    p_mono = _find("LiberationMono-Regular.ttf") or next(
        (q for q in (_find(n) for n in _MONO_NAMES) if q), None)
    if p_mono and p_mono.lower().endswith((".ttf", ".otf")):
        pdfmetrics.registerFont(TTFont("M", p_mono))
    else:
        _alias_builtin("M", "Courier")

    pdfmetrics.registerFontFamily("S", normal="S", bold="S-B", italic="S-I")


_register_fonts()

INK = colors.HexColor("#151515")
GREY = colors.HexColor("#5E5E5E")
LIGHT = colors.HexColor("#9A9A9A")
RULE = colors.HexColor("#D5D5D5")
BOX = colors.HexColor("#F4F2ED")
BOXR = colors.HexColor("#E4DFD3")
WARN = colors.HexColor("#FDF3F1")
WARNR = colors.HexColor("#E8C9C2")
GO = colors.HexColor("#F1F5EF")
GOR = colors.HexColor("#CBD9C6")

W, H = A4
ML = MR = 20 * mm
MT = 24 * mm
MB = 22 * mm


def st(name, **kw):
    base = dict(fontName="S", fontSize=9.6, leading=14.4, textColor=INK,
                alignment=TA_LEFT, spaceAfter=0)
    base.update(kw)
    return ParagraphStyle(name, **base)


S = {
    "h1": st("h1", fontName="S-B", fontSize=21, leading=25, spaceAfter=3 * mm),
    "h2": st("h2", fontName="S-B", fontSize=14, leading=18, spaceBefore=8 * mm,
             spaceAfter=2.5 * mm, keepWithNext=1),
    "h3": st("h3", fontName="S-B", fontSize=10.6, leading=15, spaceBefore=5 * mm,
             spaceAfter=1.2 * mm, keepWithNext=1),
    "p": st("p", spaceAfter=2.6 * mm),
    "lead": st("lead", fontSize=11, leading=16.5, textColor=GREY, spaceAfter=4 * mm),
    "li": st("li", leftIndent=6.5 * mm, bulletIndent=1.5 * mm, spaceAfter=1.6 * mm),
    "step": st("step", leftIndent=9 * mm, bulletIndent=0, spaceAfter=2.4 * mm,
               bulletFontName="S-B", bulletFontSize=9.6),
    "code": st("code", fontName="M", fontSize=8.5, leading=13,
               textColor=colors.HexColor("#2A2A2A"), spaceAfter=1 * mm),
    "boxh": st("boxh", fontName="S-B", fontSize=9.6, leading=14, spaceAfter=1.5 * mm),
    "boxp": st("boxp", fontSize=9.2, leading=13.8, spaceAfter=1.5 * mm),
    "tab": st("tab", fontSize=8.5, leading=12.2),
    "tabh": st("tabh", fontName="S-B", fontSize=8.5, leading=12.2),
    "tabs": st("tabs", fontSize=8.0, leading=11.6, textColor=GREY),
}


def P(t, s="p"):
    return Paragraph(t, S[s])


def LI(t, b="—"):
    return Paragraph(t, S["li"], bulletText=b)


def STEP(n, t):
    return Paragraph(t, S["step"], bulletText=n)


def box(inhalt, farbe=BOX, rand=BOXR):
    t = Table([[inhalt]], colWidths=[W - ML - MR])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), farbe),
        ("BOX", (0, 0), (-1, -1), 0.6, rand),
        ("LEFTPADDING", (0, 0), (-1, -1), 6 * mm),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 4.5 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3 * mm),
    ]))
    return t


def hinweis(titel, *absaetze, art="neutral"):
    farben = {"neutral": (BOX, BOXR), "warn": (WARN, WARNR), "ok": (GO, GOR)}
    f, r = farben[art]
    inner = [P(titel, "boxh")] + [P(a, "boxp") for a in absaetze]
    return KeepTogether(box(inner, f, r))


def codebox(*zeilen):
    inner = [P(z.replace(" ", "&nbsp;"), "code") for z in zeilen]
    t = Table([[inner]], colWidths=[W - ML - MR])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F0F0EE")),
        ("BOX", (0, 0), (-1, -1), 0.5, RULE),
        ("LEFTPADDING", (0, 0), (-1, -1), 5 * mm),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5 * mm),
    ]))
    return KeepTogether([t, Spacer(1, 3 * mm)])


def tabelle(daten, breiten, kopf=True, zebra=False):
    rows = [[Paragraph(c, S["tabh" if (kopf and i == 0) else "tab"]) for c in r]
            for i, r in enumerate(daten)]
    t = Table(rows, colWidths=breiten, repeatRows=1 if kopf else 0)
    stil = [
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3.5 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 2.2 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2 * mm),
        ("LINEBELOW", (0, 0), (-1, -2), 0.4, RULE),
    ]
    if kopf:
        stil.append(("LINEBELOW", (0, 0), (-1, 0), 0.9, INK))
    t.setStyle(TableStyle(stil))
    return t


def kopf_fuss(c, doc):
    c.saveState()
    seite = doc.page - 1                       # Deckblatt zählt nicht mit
    if seite > 1:
        c.setFont("S", 7.6)
        c.setFillColor(LIGHT)
        c.drawString(ML, H - MT + 9 * mm,
                     "FOTOTAGE MUSTERSTADT 2026 · Dateien, Zuständigkeiten, Programme")
        c.setStrokeColor(RULE)
        c.setLineWidth(0.4)
        c.line(ML, H - MT + 6.5 * mm, W - MR, H - MT + 6.5 * mm)
    c.setFont("S", 7.6)
    c.setFillColor(LIGHT)
    c.drawString(ML, MB - 10 * mm, "kontakt@beispiel-verein.de · fototage-musterstadt.de")
    c.drawRightString(W - MR, MB - 10 * mm, f"Seite {seite}")
    c.restoreState()


def _sperr(c, x, y, text, font, size, extra):
    c.setFont(font, size)
    for ch in str(text):
        c.drawString(x, y, ch)
        x += pdfmetrics.stringWidth(ch, font, size) + extra


def _wrap_zeilen(text, font, size, maxw):
    zeilen, cur = [], ""
    for wort in str(text).split():
        probe = (cur + " " + wort).strip()
        if pdfmetrics.stringWidth(probe, font, size) <= maxw or not cur:
            cur = probe
        else:
            zeilen.append(cur)
            cur = wort
    if cur:
        zeilen.append(cur)
    return zeilen


def deckblatt(c, doc):
    """Ganzseitiges Titelblatt (Seite 1)."""
    c.saveState()
    tw = W - ML - MR

    y = H - 46 * mm
    c.setFillColor(LIGHT)
    _sperr(c, ML, y, "FOTOTAGE MUSTERSTADT 2026", "S-B", 10.5, 2)
    y -= 6 * mm
    c.setStrokeColor(INK)
    c.setLineWidth(1)
    c.line(ML, y, W - MR, y)

    y -= 30 * mm
    c.setFillColor(INK)
    c.setFont("S-B", 27)
    for ln in _wrap_zeilen("Dateien, Zuständigkeiten & Programme", "S-B", 27, tw):
        c.drawString(ML, y, ln)
        y -= 32
    y -= 3 * mm
    c.setFont("S", 13)
    c.setFillColor(GREY)
    c.drawString(ML, y, "Wer öffnet was — und womit")

    y -= 16 * mm
    c.setFont("S", 10.3)
    c.setFillColor(GREY)
    for ln in _wrap_zeilen(
            "Fünf Personen arbeiten mit derselben Handvoll Dateien. Dieses Blatt ordnet "
            "jeder Datei eine Rolle, ein Programm und eine zuständige Person zu — damit "
            "keine Datei mit dem falschen Programm geöffnet wird. Alles Genannte ist auf "
            "einem Mac vorhanden oder kostenlos von Apple.", "S", 10.3, tw):
        c.drawString(ML, y, ln)
        y -= 15.5

    by = MB + 6 * mm
    c.setStrokeColor(RULE)
    c.setLineWidth(0.5)
    c.line(ML, by + 16 * mm, W - MR, by + 16 * mm)
    c.setFont("S", 9)
    c.setFillColor(GREY)
    c.drawString(ML, by + 7 * mm, "12.–19. September 2026 · Kunstraum Musterstadt")
    c.setFillColor(LIGHT)
    c.drawString(ML, by + 1.5 * mm,
                 f"kontakt@beispiel-verein.de · fototage-musterstadt.de · Stand {date.today():%d.%m.%Y}")
    c.restoreState()


# ================================================================= Inhalt
def inhalt():
    F = []
    a = F.append
    VB = W - ML - MR

    a(Spacer(1, 6 * mm))
    a(P("Wer öffnet was — und womit", "h1"))
    a(P("Dateien, Zuständigkeiten und Programme<br/>"
        "FOTOTAGE MUSTERSTADT 2026 · Beschilderung und Werkliste · alles mit macOS-Bordmitteln",
        "lead"))

    a(hinweis(
        "Warum es dieses Blatt gibt",
        "An dem Projekt arbeiten fünf verschiedene Personen mit derselben Handvoll Dateien. "
        "Die meisten Pannen entstehen nicht dadurch, dass jemand etwas Falsches tut, "
        "sondern dadurch, dass eine Datei mit dem falschen Programm geöffnet wird.",
        "Alles hier Genannte ist auf einem Mac vorhanden oder kostenlos von Apple. "
        "Es muss nichts gekauft und außer den Xcode-Befehlszeilenwerkzeugen "
        "nichts installiert werden."))

    # ------------------------------------------------------------ Rollen
    a(P("Die fünf Rollen", "h2"))
    a(tabelle([
        ["Kürzel", "Wer", "Aufgabe"],
        ["<b>ORG</b>", "Florian / Kernteam", "pflegt die Daten, erzeugt und druckt alle PDFs"],
        ["<b>FOTO</b>", "Teilnehmende Fotograf:innen",
         "tragen ihre Werkdaten in die gemeinsame Google-Tabelle ein"],
        ["<b>DES</b>", "Gestaltung", "gestaltet ein Musterschild, liefert Werte und Schriften"],
        ["<b>PRINT</b>", "Copyshop oder Bürodrucker", "druckt die fertigen PDFs"],
        ["<b>THEKE</b>", "Aufsicht während der Ausstellung",
         "arbeitet mit der ausgedruckten Werkliste"],
    ], [18 * mm, 46 * mm, VB - 64 * mm]))

    # ------------------------------------------------------------ Haupttabelle
    a(P("Die Zuordnung im Überblick", "h2"))
    a(P("Diese Tabelle ist der Kern des Blattes. Wenn du nur eine Seite liest, dann diese.",
        "p"))

    a(tabelle([
        ["Datei", "Rolle", "Programm auf dem Mac", "Was damit passiert"],

        ["<b>output/FUB2026_Werkdaten_MASTER.xlsx</b>", "ORG", "<b>Numbers</b>",
         "Die einzige Datenquelle. Werke eintragen, Preise und Status pflegen, "
         "Wand und Position vergeben."],

        ["<b>output/FUB2026_Werkmeldung_VORLAGE.xlsx</b>", "ORG → FOTO",
         "<b>Google Drive</b>",
         "ORG lädt sie als Google Tabelle hoch und teilt den Link. FOTO trägt pro Werk "
         "eine Zeile ein und wählt den eigenen Namen. Blätter: Werke, Kontakte, Preishilfe "
         "(Preisrechner), Saal Digital, Glossar. Kein Zurückschicken."],

        ["<b>data/teilnehmende.txt · data/instagram.csv</b>", "ORG", "<b>TextEdit</b>",
         "Namensliste und Instagram-Handles. Bei Änderungen scripts/build_werkmeldung.py bzw. "
         "scripts/build_qr_codes.py neu laufen lassen."],

        ["<b>data/werkmeldungen/*.xlsx</b>", "ORG", "<b>Finder</b>",
         "Die aus Google Drive geladene Excel-Fassung. Nur ablegen, nicht öffnen — "
         "das Importskript liest sie."],

        ["<b>data/design.json</b>", "DES", "<b>TextEdit</b>",
         "Schriftgrößen, Farben, Maße und QR-Größe eintragen. Nur als reiner Text — "
         "siehe Abschnitt „TextEdit“."],

        ["<b>output/qr-codes/*.svg</b>", "DES", "<b>Illustrator / Vorschau</b>",
         "QR-Codes als Vektor, falls DES die Schilder selbst setzt."],

        ["<b>scripts/build_druckdaten.py</b>", "ORG", "<b>Terminal</b>",
         "Wird ausgeführt, nicht geöffnet. Erzeugt alle Druck-PDFs inkl. QR-Codes."],

        ["<b>scripts/import_werkmeldungen.py</b>", "ORG", "<b>Terminal</b>",
         "Wird ausgeführt. Liest die Werkmeldung in den Master."],

        ["<b>scripts/migrate_werkmeldung.py · data/werkmeldung_alt/</b>", "ORG", "<b>Terminal</b>",
         "Nur nötig, wenn die Vorlage NACH dem Teilen geändert wird: ausgefüllte Datei "
         "in data/werkmeldung_alt/ legen, Skript überträgt sie in die neue Vorlage."],

        ["<b>build_qr_codes.py</b>", "ORG", "<b>Terminal</b>",
         "Erzeugt fehlende Instagram-QR-Codes, meldet fehlende Links."],

        ["<b>output/ausgabe/01_Wandschilder.pdf</b>", "ORG → PRINT", "<b>Vorschau</b>",
         "Prüfen, dann drucken. Skalierung zwingend 100 % — siehe Abschnitt „Vorschau“."],

        ["<b>output/ausgabe/02_Werkliste.pdf</b>", "ORG → PRINT → THEKE", "<b>Vorschau</b>",
         "Erst nach der endgültigen Hängung drucken. Mindestens drei Exemplare."],

        ["<b>output/ausgabe/03_Rueckseitenetiketten.pdf</b>", "ORG", "<b>Vorschau</b>",
         "Vor dem Aufhängen drucken und auf die Rahmenrückseiten kleben."],

        ["<b>output/ausgabe/04_Versicherungsliste.pdf</b>", "ORG", "<b>Vorschau</b>",
         "Intern. Einmal ausdrucken, Summe an Räumchen e. V. melden. Nicht auslegen."],

        ["<b>output/ausgabe/05_Instagram_Wand.pdf</b>", "ORG → PRINT", "<b>Vorschau</b>",
         "Aushang mit den Instagram-QR-Codes aller Teilnehmenden. Optional."],

        ["<b>output/ausgabe/Adobe/</b> (3 Dateien)", "DES", "<b>Illustrator / InDesign</b>",
         "Nur, wenn DES selbst produziert. Direkt in Adobe laden, vorher nirgendwo öffnen."],

        ["<b>output/ausgabe/Werkdaten_Serienbrief.csv</b>", "ORG", "<b>—</b>",
         "Notfallweg für einen Word-Serienbrief. Im Normalbetrieb liegen lassen."],

        ["<b>output/HANDBUCH_DESIGNER.pdf</b>", "ORG → DES", "<b>Vorschau</b>",
         "Wird an die Gestaltung weitergegeben."],

        ["<b>HANDBUCH_BESCHILDERUNG.md</b>", "ORG", "<b>TextEdit</b>",
         "Internes Ablaufhandbuch. Nur lesen."],
    ], [53 * mm, 20 * mm, 30 * mm, VB - 103 * mm]))

    # ------------------------------------------------------------ Verbote
    a(P("Die vier Regeln, an denen es hängt", "h2"))

    a(hinweis(
        "1 — Die Dateien im Ordner Adobe/ niemals in Numbers oder Excel öffnen",
        "Das betrifft <b>InDesign_Datenzusammenfuehrung.txt</b> und "
        "<b>Illustrator_Variablen.csv</b>. Numbers wandelt beim Speichern die Textkodierung "
        "um. Aus jedem <b>ü</b> wird dann <b>Ã¼</b>, aus <b>Völklingen</b> wird "
        "<b>VÃ¶lklingen</b> — sichtbar erst auf dem gedruckten Schild.",
        "Wenn du nur hineinschauen willst: Datei im Finder einmal anklicken und "
        "<b>Leertaste</b> drücken. Das ist die Übersicht und verändert nichts.",
        art="warn"))

    a(Spacer(1, 3 * mm))
    a(hinweis(
        "2 — die design.json nur als reinen Text bearbeiten",
        "TextEdit legt neue Dokumente standardmäßig als formatierten Text an und ersetzt "
        "gerade Anführungszeichen durch typografische. Beides macht die Datei unlesbar. "
        "Wie du das abstellst, steht im Abschnitt „TextEdit“.",
        art="warn"))

    a(Spacer(1, 3 * mm))
    a(hinweis(
        "3 — PDFs immer mit Skalierung 100 % drucken",
        "Die Schilder sind exakt 100 × 70 mm. Druckt jemand mit „An Seite anpassen“, "
        "werden daraus 96 × 67 mm — die Schnittmarken stimmen dann nicht mehr und alle "
        "Schilder haben unterschiedliche Ränder.",
        art="warn"))

    a(Spacer(1, 3 * mm))
    a(hinweis(
        "4 — Der Master ist die einzige Wahrheit",
        "Preis, Titel oder Status ändern sich ausschließlich in "
        "<b>output/FUB2026_Werkdaten_MASTER.xlsx</b>. Danach den Generator neu laufen lassen. "
        "Niemals im fertigen PDF nachbessern und niemals eine Kopie der Tabelle anlegen, "
        "in der „nur schnell etwas“ geändert wird.",
        art="warn"))

    # ------------------------------------------------------------ Numbers
    a(P("Numbers — für die beiden Tabellen", "h2"))
    a(P("Numbers ist auf neuen Macs vorinstalliert; sonst kostenlos im App Store. "
        "Es öffnet .xlsx-Dateien direkt per Doppelklick.", "p"))

    a(P("Master bearbeiten und wieder abspeichern", "h3"))
    for n, t in [
        ("1.", "Doppelklick auf <b>output/FUB2026_Werkdaten_MASTER.xlsx</b>. Numbers zeigt beim "
               "Öffnen eine Warnung über Änderungen beim Import — das ist normal, wegklicken."),
        ("2.", "Im Blatt <b>Werke</b> arbeiten. Blaue Schrift bedeutet Eingabefeld, "
               "graue Felder sind Formeln und bleiben unberührt."),
        ("3.", "Zum Sichern: <b>Ablage &gt; Exportieren &gt; Excel</b>. Als Dateiname wieder "
               "genau <b>output/FUB2026_Werkdaten_MASTER.xlsx</b> wählen und die alte Datei ersetzen."),
    ]:
        a(STEP(n, t))

    a(hinweis(
        "Der eine Haken bei Numbers",
        "Numbers kann die Auswahlmenüs (Status, Rahmen im Preis, Künstler:in, Druckverfahren, "
        "Papier, Maße) beim Import verlieren. Die Tabelle funktioniert weiter, du musst die "
        "Werte dann nur von Hand tippen statt auszuwählen.",
        "Status exakt so: <b>verfügbar</b> · <b>reserviert</b> · <b>verkauft</b> · "
        "<b>nicht verkäuflich</b>, dazu <b>ja</b> / <b>nein</b>. Für Druckverfahren und "
        "Papier steht die genaue Schreibweise im Blatt <b>Listen</b> des Masters bzw. im "
        "Blatt <b>Saal Digital</b> der Werkmeldung — von dort kopieren. Der Generator prüft "
        "Status und ja/nein vor dem Druck und meckert, wenn etwas anderes dasteht."))

    a(Spacer(1, 3 * mm))
    a(hinweis(
        "Blattschutz in der Werkmeldung",
        "Kopfzeilen, Dropdown-Quellen und die Nachschlage-Blätter der Werkmeldung sind "
        "gesperrt (Passwort <b>fub2026</b>, dokumentiert in <b>scripts/build_werkmeldung.py</b>). "
        "Das ist ein Schutz gegen versehentliches Verändern in der von allen gemeinsam "
        "bearbeiteten Datei, kein Sicherheitsmerkmal.",
        "Öffnet ORG die Datei einmal in Google Tabellen, sollte unter <b>Daten &gt; Blätter "
        "und Bereiche schützen</b> geprüft werden, ob die Sperren korrekt übernommen wurden — "
        "die Übersetzung aus dem Excel-Format ist nicht immer exakt. Dabei auch kurz testen, "
        "ob die Dropdowns in einer leeren „Werke“-Zeile erscheinen: Das Quellblatt <b>Listen</b> "
        "ist bewusst sichtbar (ganz hinten), weil Google Sheets Dropdowns von ausgeblendeten "
        "Blättern beim Import verwirft.",
        art="ok"))

    a(Spacer(1, 3 * mm))
    a(hinweis(
        "Beispieldaten verschwinden von selbst",
        "Die 14 Beispielwerke im Master (Spalte <b>Beispiel</b> = ja) druckt der Generator "
        "automatisch nicht mehr mit, sobald mindestens ein echtes Werk importiert wurde. "
        "Kein manuelles Löschen nötig. <font name='M' size='8'>--mit-beispielen</font> "
        "erzwingt sie trotzdem, z. B. für einen Layout-Test mit vollem Datensatz."))

    a(Spacer(1, 3 * mm))
    a(hinweis(
        "Werknummern: automatisch, alphabetisch",
        "<b>scripts/import_werkmeldungen.py</b> vergibt die Werknummer bei jedem Lauf neu — sortiert "
        "nach Künstler:in (A-Z), unabhängig von Melde- oder Zeilenreihenfolge. Nicht nur die "
        "gerade importierten Werke, alle echten Werke im Master werden dabei neu durchnummeriert.",
        "Die Nummer ist also erst nach dem letzten Import vor dem Druck endgültig. Wandschilder "
        "aus einem Zwischenstand können danach eine andere Nummer zeigen als der finale Master — "
        "erst final drucken, wenn wirklich alle Meldungen eingesammelt sind.",
        art="warn"))

    a(P("Ein Sicherheitsnetz, das nichts kostet", "h3"))
    a(P("Leg den Projektordner in <b>iCloud Drive</b> ab. Dann hat der Finder für jede Datei "
        "eine Versionsgeschichte: Rechtsklick auf die Datei, <b>Informationen</b>, oder in "
        "Numbers <b>Ablage &gt; Zurücksetzen auf &gt; Alle Versionen durchsuchen</b>. "
        "Damit bekommst du jede frühere Fassung des Masters zurück.", "p"))

    # ------------------------------------------------------------ TextEdit
    a(P("TextEdit — für die design.json", "h2"))
    a(P("TextEdit liegt im Ordner <b>Programme</b>. Es kann JSON-Dateien bearbeiten, aber nur "
        "mit zwei Einstellungen, die man einmal umstellen muss. Ohne sie zerstört TextEdit "
        "die Datei zuverlässig.", "p"))

    a(P("Einmalig einrichten", "h3"))
    for n, t in [
        ("1.", "TextEdit öffnen. Menü <b>TextEdit &gt; Einstellungen</b> "
               "(auf neueren Systemen <b>TextEdit &gt; Einstellungen…</b>)."),
        ("2.", "Reiter <b>Neues Dokument</b>: oben <b>Reiner Text</b> auswählen."),
        ("3.", "Reiter <b>Neues Dokument</b>, Bereich <b>Optionen</b>: die Häkchen bei "
               "<b>Intelligente Anführungszeichen</b> und <b>Intelligente Bindestriche</b> "
               "<b>entfernen</b>. Das ist der wichtigste Schritt überhaupt."),
        ("4.", "Fenster schließen. Die Einstellung gilt ab sofort dauerhaft."),
    ]:
        a(STEP(n, t))

    a(P("Warum das so wichtig ist", "h3"))
    a(P("JSON verlangt gerade Anführungszeichen. TextEdit macht daraus ungefragt "
        "typografische:", "p"))
    a(codebox(
        'richtig:  "titel": 13',
        'falsch:   „titel“: 13',
    ))
    a(P("Das sieht fast gleich aus, ist aber der häufigste Grund, warum der Generator später "
        "meldet, die Datei sei unlesbar. Falls eine bestehende Datei schon im falschen Modus "
        "geöffnet wurde: <b>Format &gt; In reinen Text umwandeln</b> (Umschalt-Befehl-T).", "p"))

    a(P("Eine geöffnete Datei prüfen", "h3"))
    a(P("Sieht die Datei in TextEdit aus wie Fließtext mit Schriftarten und Formatierung, "
        "ist etwas falsch. Reiner Text erscheint in einer nichtproportionalen Schrift, "
        "alle Zeilen linksbündig, keine Formatierungsleiste oben.", "p"))

    # ------------------------------------------------------------ Terminal
    a(P("Terminal — für die beiden Skripte", "h2"))
    a(P("Terminal liegt in <b>Programme &gt; Dienstprogramme</b>. Es sieht unfreundlicher aus, "
        "als es ist: Du tippst hier genau zwei bis drei Befehle, immer dieselben.", "p"))

    a(P("Einmalige Einrichtung", "h3"))
    for n, t in [
        ("1.", "Terminal öffnen und eingeben:"),
    ]:
        a(STEP(n, t))
    a(codebox("python3 --version"))
    a(P("Kommt eine Versionsnummer, ist alles da. Öffnet sich stattdessen ein Fenster mit der "
        "Frage nach den <b>Befehlszeilenwerkzeugen</b>, klick auf <b>Installieren</b> und "
        "warte, bis es fertig ist. Das ist Apple-eigene Software, kein Fremdprogramm.", "p"))
    a(STEP("2.", "Danach die benötigten Bausteine holen:"))
    a(codebox("pip3 install --user reportlab openpyxl segno"))
    a(P("Falls das mit einer Meldung über eine „extern verwaltete Umgebung“ abbricht, "
        "stattdessen:", "p"))
    a(codebox("pip3 install --user --break-system-packages reportlab openpyxl segno"))

    a(P("Der Alltagsbefehl", "h3"))
    for n, t in [
        ("1.", "Terminal öffnen. <b>cd</b> tippen, dann ein Leerzeichen — und jetzt den "
               "Projektordner aus dem Finder direkt ins Terminalfenster ziehen. "
               "Der Pfad wird automatisch eingesetzt. Return drücken."),
        ("2.", "Den eigentlichen Befehl eingeben:"),
    ]:
        a(STEP(n, t))
    a(codebox("./.venv/bin/python scripts/build_druckdaten.py"))
    a(P("Der Generator prüft die Daten und schreibt danach alle PDFs in den Ordner "
        "<b>ausgabe</b>. Findet er einen Fehler — ein fehlender Preis, eine doppelte "
        "Werknummer — bricht er ab und nennt Werknummer und Titel. Dann im Master "
        "korrigieren und den Befehl wiederholen.", "p"))

    a(P("Mit der Pfeiltaste nach oben holst du den letzten Befehl zurück. Nach dem ersten Mal "
        "besteht der ganze Vorgang also aus: Terminal öffnen, Pfeil hoch, Return.", "p"))

    a(P("Die weiteren Befehle", "h3"))
    a(tabelle([
        ["Wofür", "Befehl"],
        ["Werkmeldungen einlesen, erst zur Ansicht",
         "<font name='M' size='8'>./.venv/bin/python scripts/import_werkmeldungen.py "
         "--probe</font>"],
        ["Werkmeldungen wirklich einlesen",
         "<font name='M' size='8'>./.venv/bin/python scripts/import_werkmeldungen.py</font>"],
        ["Nur die Werkliste neu erzeugen",
         "<font name='M' size='8'>./.venv/bin/python scripts/build_druckdaten.py --nur werkliste</font>"],
        ["Mit der Typografie der Gestaltung",
         "<font name='M' size='8'>./.venv/bin/python scripts/build_druckdaten.py --design data/design.json</font>"],
        ["Datendateien für Adobe mitschreiben",
         "<font name='M' size='8'>./.venv/bin/python scripts/build_druckdaten.py --adobe</font>"],
    ], [58 * mm, VB - 58 * mm]))

    a(Spacer(1, 4 * mm))
    a(hinweis(
        "Optional: einen Doppelklick daraus machen",
        "Wenn du das Terminal nie wieder sehen willst, kannst du dir mit <b>Automator</b> "
        "(liegt in Programme) ein kleines Programm bauen: <b>Neu &gt; Programm</b>, links "
        "<b>Shell-Skript ausführen</b> in die Fläche ziehen, dort die zwei Zeilen eintragen "
        "und als <b>Schilder erzeugen.app</b> in den Projektordner sichern.",
        "Der Inhalt des Shell-Skripts:",
        art="ok"))
    a(codebox('cd "/Pfad/zum/Projektordner"',
              './.venv/bin/python scripts/build_druckdaten.py'))
    a(P("Den Pfad bekommst du, indem du den Ordner ins Automator-Fenster ziehst. "
        "Danach genügt ein Doppelklick auf die App.", "p"))

    # ------------------------------------------------------------ Vorschau
    a(P("Vorschau — für alle PDFs", "h2"))
    a(P("Vorschau öffnet PDFs per Doppelklick und ist zum Drucken völlig ausreichend. "
        "Entscheidend sind drei Einstellungen im Druckdialog.", "p"))

    a(P("So druckst du die Wandschilder", "h3"))
    for n, t in [
        ("1.", "<b>01_Wandschilder.pdf</b> öffnen, <b>Ablage &gt; Drucken</b> (Befehl-P)."),
        ("2.", "Unten links auf <b>Details einblenden</b> klicken, falls der Dialog klein ist."),
        ("3.", "<b>Papierformat: A4.</b>"),
        ("4.", "<b>Skalierung: 100 %.</b> Nicht „An Papierformat anpassen“, nicht "
               "„Automatisch skalieren“. Das ist der Punkt, an dem am häufigsten "
               "etwas schiefgeht."),
        ("5.", "<b>Automatisch drehen und zentrieren</b> ausschalten."),
        ("6.", "Papierzufuhr auf den Schacht mit dem 300-g-Karton stellen, "
               "beim Bürodrucker meist der manuelle Einzug."),
        ("7.", "<b>Erst einen einzelnen Bogen drucken.</b> Mit dem Lineal nachmessen: "
               "Ein Schild muss exakt 100 mm breit und 70 mm hoch sein. Stimmt das, "
               "den Rest drucken."),
    ]:
        a(STEP(n, t))

    a(Spacer(1, 3 * mm))
    a(hinweis(
        "Die Probe, die zehn Sekunden kostet",
        "Miss auf dem ersten gedruckten Bogen den Abstand zwischen zwei Schnittmarken. "
        "100 mm heißt: alles richtig. 96 oder 97 mm heißt: Die Skalierung stand nicht auf "
        "100 %. Wenn du das erst nach 15 Bogen merkst, hast du 15 Bogen Karton verloren.",
        art="ok"))

    a(P("Was Vorschau sonst noch kann", "h3"))
    a(LI("<b>Seitenzahl prüfen:</b> Sidebar einblenden mit Wahl-Befehl-2 — bei 100 Werken müssen "
         "13 Bogen erscheinen."))
    a(LI("<b>Suchen:</b> Befehl-F, um zu prüfen, ob ein bestimmter Titel wirklich enthalten ist."))
    a(LI("<b>Werkliste beidseitig:</b> Im Druckdialog <b>Beidseitig &gt; Bindung an langer "
         "Kante</b>."))
    a(LI("<b>Zum Copyshop schicken:</b> Die PDFs unverändert weitergeben. Nichts exportieren, "
         "nichts neu sichern — sonst gehen die Schriften verloren."))

    # ------------------------------------------------------------ Finder
    a(P("Finder und Mail — der Rest", "h2"))

    a(P("Ordner verschicken", "h3"))
    a(P("Für die Gestaltung: den Ordner <b>upload/</b> auf Google Drive hochladen (oder "
        "rechtsklick &gt; <b>Komprimieren</b> und die .zip per Mail). "
        "Für die Teilnehmenden: <b>output/FUB2026_Werkmeldung_VORLAGE.xlsx</b> als Google Tabelle "
        "öffnen und den Link mit „Jeder mit Link – Bearbeiter“ teilen — nicht die Datei "
        "selbst verschicken.", "p"))

    a(P("Standardprogramm dauerhaft festlegen", "h3"))
    a(P("Damit .json-Dateien künftig nicht versehentlich in einem anderen Programm landen: "
        "Datei anklicken, <b>Befehl-I</b> für Informationen, unter <b>Öffnen mit</b> "
        "<b>TextEdit</b> auswählen und auf <b>Alle ändern…</b> klicken.", "p"))

    a(P("Schnell hineinschauen ohne zu öffnen", "h3"))
    a(P("Datei anklicken, <b>Leertaste</b>. Die Übersicht zeigt PDFs, Tabellen und Textdateien "
        "an, ohne sie zu verändern. Für alles, was du nur kurz kontrollieren willst, ist das "
        "der sichere Weg.", "p"))

    a(P("Schrift prüfen — für die Gestaltung", "h3"))
    a(P("Die Maßangaben enthalten ein echtes Multiplikationszeichen <b>×</b>, nicht den "
        "Buchstaben x. Ob eine Schrift dieses Zeichen hat, zeigt die <b>Schriftsammlung</b> "
        "(Programme): Schrift auswählen, oben auf <b>Repertoire</b> umschalten und im "
        "Suchfeld <b>00D7</b> eingeben.", "p"))

    # ------------------------------------------------------------ Ablauf
    a(P("Der typische Tag, in Reihenfolge", "h2"))
    a(tabelle([
        ["Wann", "Wer", "Was", "Womit"],
        ["Vorbereitung", "ORG", "Werkmeldung auf Google Drive teilen, QR-Codes prüfen",
         "Google Drive, Terminal"],
        ["laufend", "FOTO", "Werke in die gemeinsame Tabelle eintragen", "Google Tabellen"],
        ["nach Frist", "ORG", "Tabelle als Excel laden, in data/werkmeldungen legen", "Google Drive, Finder"],
        ["nach Frist", "ORG", "Import zur Probe, dann scharf", "Terminal"],
        ["nach Frist", "ORG", "Daten sichten, Wand und Position vergeben", "Numbers"],
        ["parallel", "DES", "Musterschild gestalten, data/design.json ausfüllen", "TextEdit"],
        ["nach Hängeplan", "ORG", "Alle PDFs erzeugen", "Terminal"],
        ["vor Aufbau", "ORG", "Schilder und Etiketten drucken, Probe messen", "Vorschau"],
        ["Aufbautag", "ORG", "Werkliste final drucken", "Vorschau"],
        ["Ausstellung", "THEKE", "Verkäufe notieren", "Papier"],
        ["Ausstellung", "ORG", "Status pflegen, Werkliste nachdrucken",
         "Numbers, Terminal, Vorschau"],
    ], [26 * mm, 18 * mm, VB - 88 * mm, 44 * mm]))

    # ------------------------------------------------------------ Störungen
    a(P("Wenn etwas nicht stimmt", "h2"))
    a(tabelle([
        ["Symptom", "Ursache", "Lösung"],
        ["Umlaute erscheinen als <b>Ã¼</b>",
         "Eine Datei aus Adobe/ wurde in Numbers geöffnet und gesichert",
         "Generator mit <b>--adobe</b> neu laufen lassen, frische Datei nehmen, "
         "diesmal nicht öffnen"],
        ["„Designdatei nicht lesbar“",
         "TextEdit hat typografische Anführungszeichen eingesetzt",
         "Format &gt; In reinen Text umwandeln, Anführungszeichen ersetzen, "
         "Einstellung dauerhaft ändern"],
        ["Schilder sind zu klein gedruckt",
         "Skalierung stand nicht auf 100 %",
         "Druckdialog prüfen, einen Testbogen messen"],
        ["<b>command not found: python3</b>",
         "Befehlszeilenwerkzeuge fehlen",
         "<font name='M' size='8'>xcode-select --install</font> eingeben und "
         "installieren lassen"],
        ["<b>No such file or directory</b>",
         "Terminal steht im falschen Ordner",
         "<b>cd</b> tippen, Leerzeichen, Ordner aus dem Finder hineinziehen, Return"],
        ["Generator bricht mit Datenfehlern ab",
         "Pflichtangaben fehlen im Master",
         "Genau das ist der Zweck. Werknummern in der Meldung nachschlagen und "
         "in Numbers ergänzen"],
        ["Auswahlmenüs im Master weg",
         "Numbers hat sie beim Import verloren",
         "Werte von Hand tippen, exakt in der Schreibweise aus dem Abschnitt „Numbers“"],
    ], [40 * mm, 46 * mm, VB - 86 * mm]))

    a(Spacer(1, 7 * mm))
    a(box([
        P("<font name='S-B' size='11'>Im Zweifel</font>", "boxp"),
        P("Nichts überschreiben, nichts löschen, kurz melden. Alle Skripte legen vor dem "
          "Schreiben eine Sicherung des Masters an, und die PDFs lassen sich jederzeit neu "
          "erzeugen — solange die Tabelle heil ist.", "boxp"),
        P("<b>kontakt@beispiel-verein.de</b> · fototage-musterstadt.de · "
          "instagram.com/fototage.musterstadt", "boxp"),
        P("Team / Kulturverein Musterstadt | FOTOTAGE MUSTERSTADT 2026", "boxp"),
    ]))

    return F


def main():
    doc = BaseDocTemplate(OUT, pagesize=A4, leftMargin=ML, rightMargin=MR,
                          topMargin=MT, bottomMargin=MB,
                          title="FOTOTAGE MUSTERSTADT 2026 — Dateien, Zuständigkeiten, Programme",
                          author="KULTURVEREIN MUSTERSTADT")
    frame = Frame(ML, MB, W - ML - MR, H - MT - MB, id="f",
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([
        PageTemplate(id="cover", frames=[Frame(0, 0, W, H, id="cv")], onPage=deckblatt),
        PageTemplate(id="main", frames=[frame], onPage=kopf_fuss),
    ])
    doc.build([NextPageTemplate("main"), Spacer(1, 2), PageBreak()] + inhalt())
    print("geschrieben:", OUT)


if __name__ == "__main__":
    main()
