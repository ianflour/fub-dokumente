#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Erzeugt output/HANDBUCH_DESIGNER.pdf — Anleitung zum Mitgeben."""

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

OUT = P.output("HANDBUCH_DESIGNER.pdf")

# --- Schriften plattformunabhängig suchen -----------------------------------
# Gesucht wird eine Arial-metrische Sans (Liberation Sans / Arial / DejaVu) und
# eine Monospace. Wird nichts gefunden, nehmen wir die eingebauten PDF-Standard-
# schriften Helvetica / Courier — das Handbuch sieht dann minimal anders aus,
# baut aber trotzdem.
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
               "cour.ttf", "Menlo.ttc"]


def _find(name):
    for d in _FONT_DIRS:
        hits = glob.glob(os.path.join(d, "**", name), recursive=True)
        if hits:
            return hits[0]
    return None


def _alias_builtin(alias, builtin):
    """Registriert einen der PDF-Standardfonts unter unserem Kurznamen."""
    pdfmetrics.registerFont(pdfmetrics.Font(alias, builtin, "WinAnsiEncoding"))


def _register_fonts():
    """Belegt die festen Namen S / S-B / S-I / M, egal welche Schrift da ist."""
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

    p_mono = next((q for q in (_find(n) for n in _MONO_NAMES) if q), None)
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
ACC = colors.HexColor("#8A3324")

W, H = A4
ML = MR = 22 * mm
MT = 24 * mm
MB = 22 * mm


def st(name, **kw):
    base = dict(fontName="S", fontSize=9.6, leading=14.6, textColor=INK,
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
    "small": st("small", fontSize=8.4, leading=12.6, textColor=GREY, spaceAfter=2 * mm),
    "li": st("li", leftIndent=6.5 * mm, bulletIndent=1.5 * mm, spaceAfter=1.6 * mm),
    "step": st("step", leftIndent=9 * mm, bulletIndent=0, spaceAfter=2.4 * mm,
               bulletFontName="S-B", bulletFontSize=9.6),
    "code": st("code", fontName="M", fontSize=8.3, leading=12.4,
               textColor=colors.HexColor("#2A2A2A"), spaceAfter=1 * mm),
    "boxh": st("boxh", fontName="S-B", fontSize=9.6, leading=14, spaceAfter=1.5 * mm),
    "boxp": st("boxp", fontSize=9.2, leading=13.8, spaceAfter=1.5 * mm),
    "tab": st("tab", fontSize=8.8, leading=12.6),
    "tabh": st("tabh", fontName="S-B", fontSize=8.8, leading=12.6),
}


def P(t, s="p"):
    return Paragraph(t, S[s])


def LI(t):
    return Paragraph(t, S["li"], bulletText="—")


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


def hinweis(titel, *absaetze, warnung=False):
    inner = [P(titel, "boxh")] + [P(a, "boxp") for a in absaetze]
    return KeepTogether(box(inner, WARN if warnung else BOX, WARNR if warnung else BOXR))


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


def tabelle(daten, breiten, kopf=True):
    rows = []
    for i, r in enumerate(daten):
        rows.append([Paragraph(c, S["tabh" if (kopf and i == 0) else "tab"]) for c in r])
    t = Table(rows, colWidths=breiten, repeatRows=1 if kopf else 0)
    stil = [
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4 * mm),
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
        c.drawString(ML, H - MT + 9 * mm, "FOTOTAGE MUSTERSTADT 2026 · Wandschilder — Anleitung")
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
    for ln in _wrap_zeilen("Wandschilder — Anleitung für die Gestaltung", "S-B", 27, tw):
        c.drawString(ML, y, ln)
        y -= 32
    y -= 3 * mm
    c.setFont("S", 13)
    c.setFillColor(GREY)
    c.drawString(ML, y, "Wie aus einem gestalteten Schild einhundert werden")

    y -= 16 * mm
    c.setFont("S", 10.3)
    c.setFillColor(GREY)
    for ln in _wrap_zeilen(
            "Diese Anleitung richtet sich an die Person, die das Erscheinungsbild der "
            "Wandschilder und der Werkliste festlegt. Sie beschreibt drei Wege von der "
            "Gestaltung eines einzigen Schildes zur fertigen Produktion aller Schilder "
            "und empfiehlt einen davon.", "S", 10.3, tw):
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

    # ---------------------------------------------------------- Titel
    a(Spacer(1, 6 * mm))
    a(P("Wandschilder für 100 Fotografien", "h1"))
    a(P("Anleitung für die Gestaltung und die Produktion<br/>"
        "FOTOTAGE MUSTERSTADT 2026 · 12.–19. September 2026 · Kunstraum Musterstadt", "lead"))

    a(hinweis(
        "Worum es geht",
        "Für jedes der rund 100 ausgestellten Fotos hängt ein kleines Schild an der Wand. "
        "Darauf stehen: Name, Titel, Ort und Jahr, Druckverfahren und Papier, Bildmaß, Auflage. "
        "Alle Schilder sehen gleich aus — nur der Text ist jeweils ein anderer.",
        "Deine Aufgabe ist die <b>Gestaltung eines einzigen Schildes</b>. "
        "Die 100 Varianten entstehen danach automatisch. "
        "Diese Anleitung erklärt drei Wege dorthin und empfiehlt einen davon."))

    a(Spacer(1, 5 * mm))
    a(P("Was in dem Ordner liegt, den du bekommen hast", "h3"))
    a(tabelle([
        ["Datei", "Wofür"],
        ["<b>design_VORLAGE.json</b>", "Eine kleine Textdatei mit allen Schriftgrößen, Farben "
                                       "und Maßen. Die füllst du aus. Mehr dazu im Abschnitt „Weg A“."],
        ["<b>ausgabe/01_Wandschilder.pdf</b>", "So sehen die Schilder gerade aus — der "
                                               "Zwischenstand, den du gestalterisch ablöst."],
        ["<b>ausgabe/Adobe/</b>", "Drei Datendateien für InDesign und Illustrator. "
                                  "Brauchst du nur bei Weg B oder C."],
        ["<b>ausgabe/02_Werkliste.pdf</b>", "Die Liste, die in der Ausstellung ausliegt. "
                                            "Gehört zum selben Gestaltungssystem."],
    ], [46 * mm, W - ML - MR - 46 * mm]))

    a(Spacer(1, 6 * mm))
    a(hinweis(
        "Eine Sache vorweg, damit nichts schiefgeht",
        "Die Texte sind unterschiedlich lang. Der kürzeste Titel hat 9 Zeichen, der längste 34. "
        "Bei „Papier“ reicht die Spanne von <i>Ilford Multigrade RC Deluxe</i> bis "
        "<i>Canson Baryta Photographique II 310 g</i>.",
        "Ein Layout, das nur mit dem kurzen Beispiel getestet wurde, bricht bei Werk 63 "
        "auseinander. Bitte prüfe deinen Entwurf immer mit dem <b>längsten</b> Text, "
        "nicht mit dem schönsten. Die längsten Beispiele stehen im Abschnitt „Gestaltungsvorgaben“."))

    a(Spacer(1, 4 * mm))

    # ---------------------------------------------------------- Entscheidung
    a(P("Drei Wege — und welcher der richtige ist", "h2"))
    a(P("Alle drei führen zu 100 fertigen Schildern. Sie unterscheiden sich darin, wie viel "
        "Technik bei dir landet.", "p"))

    a(tabelle([
        ["", "Weg A — empfohlen", "Weg B", "Weg C"],
        ["<b>Du machst</b>",
         "Ein Schild gestalten, Werte in eine Textdatei eintragen, Schriften mitgeben",
         "Ein Schild in InDesign bauen und die Daten selbst einsetzen",
         "Ein Schild in Illustrator bauen und die Daten selbst einsetzen"],
        ["<b>Technikaufwand</b>", "sehr gering", "mittel", "hoch"],
        ["<b>Umbruch bei langen Texten</b>", "passiert automatisch",
         "musst du prüfen", "musst du prüfen"],
        ["<b>Nachträgliche Änderung</b>", "Knopfdruck", "erneut zusammenführen",
         "erneut zusammenführen"],
        ["<b>Druckbogen mit 8 Schildern</b>", "kommt fertig raus", "kommt fertig raus",
         "musst du selbst bauen"],
        ["<b>Zeit für 100 Stück</b>", "unter einer Minute", "ca. 20 Minuten",
         "ca. eine Stunde"],
    ], [26 * mm, 46 * mm, 45 * mm, 49 * mm]))

    a(Spacer(1, 5 * mm))
    a(hinweis(
        "Die Empfehlung: Weg A",
        "Du gestaltest ein Schild so, wie du es immer tust — in Illustrator, auf Papier, "
        "wo du willst. Danach trägst du die fertigen Werte (Schriftgrößen, Farben, Abstände) "
        "in eine vorbereitete Textdatei ein und legst die Schriftdateien dazu. "
        "Den Rest übernimmt ein kleines Programm bei uns.",
        "Der Grund ist nicht Bequemlichkeit, sondern Verlässlichkeit: Das Programm misst jeden "
        "Text und bricht ihn um. Kein Titel läuft aus dem Schild. Und wenn zwei Tage vor "
        "der Eröffnung noch ein Preis oder ein Titel geändert wird, kostet das keinen Nachmittag, "
        "sondern einen Klick."))

    a(Spacer(1, 4 * mm))
    a(P("Wenn du lieber selbst produzieren möchtest, ist <b>Weg B (InDesign)</b> der zweitbeste. "
        "<b>Weg C (Illustrator)</b> beschreiben wir vollständig, weil Illustrator dein "
        "Hauptwerkzeug ist — aber lies bitte den Kasten am Ende dieses Abschnitts, bevor du dich dafür "
        "entscheidest.", "p"))

    a(Spacer(1, 4 * mm))

    # ---------------------------------------------------------- Weg A
    a(P("Weg A — Du gestaltest, wir produzieren", "h2"))
    a(P("Der ganze Weg besteht aus vier Schritten. Du brauchst dafür keine besondere Software "
        "und musst nichts installieren.", "p"))

    a(P("Schritt 1 — Ein Schild gestalten", "h3"))
    a(P("Format: <b>100 × 70 mm quer</b>. Lege in Illustrator eine Zeichenfläche in dieser "
        "Größe an und gestalte darin ein einzelnes Schild mit den Testtexten aus dem Abschnitt „Gestaltungsvorgaben“. "
        "Die inhaltliche Reihenfolge steht fest und darf nicht geändert werden — "
        "die Gestaltung darin ist deine.", "p"))
    a(P("Falls dir 100 × 70 mm zu klein ist, gibt es eine zweite, barrierearme Variante mit "
        "<b>140 × 100 mm</b>. Entscheide dich für eine der beiden, nicht für beide.", "p"))

    a(P("Schritt 2 — Werte ablesen", "h3"))
    a(P("Öffne die Datei <b>design_VORLAGE.json</b> mit einem einfachen Texteditor "
        "(TextEdit auf dem Mac, Editor unter Windows — nicht mit Word). Sie sieht so aus:", "p"))
    a(codebox(
        '"farben": {',
        '  "text":      "#000000",',
        '  "sekundaer": "#616161",',
        '  "tertiaer":  "#999999",',
        '  "linie":     "#C7C7C7"',
        '},',
        '"schriftgroessen_pt": {',
        '  "nummer":      7,',
        '  "kuenstlerin": 11,',
        '  "titel":       13,',
        '  "ortjahr":     9.5,',
        '  "technik":     8.3,',
        '  "hinweis":     7.2',
        '}'))
    a(P("Du überschreibst nur die Zahlen und die Farbwerte rechts. Alles andere — die "
        "Anführungszeichen, die Doppelpunkte, die Kommas — bleibt exakt stehen. "
        "Ein fehlendes Komma macht die Datei unlesbar.", "p"))

    a(hinweis(
        "Die zwei Regeln für diese Datei",
        "<b>1.</b> Jede Zeile endet mit einem Komma — außer der letzten vor einer geschweiften "
        "Klammer <b>}</b>. Genau so, wie es jetzt schon dasteht.",
        "<b>2.</b> Farben immer als Hex-Wert mit Rautezeichen und sechs Stellen: "
        "<b>#B31F1F</b>. Illustrator zeigt dir diesen Wert im Farbwähler unten rechts an.",
        "Wenn du unsicher bist: Schick uns die Datei einfach, wir prüfen sie in zwei Minuten."))

    a(P("Schritt 3 — Schriften beilegen", "h3"))
    a(P("Wir brauchen zwei Schnitte deiner Schrift als Dateien: <b>Regular und Bold</b>. "
        "Endung .ttf oder .otf. Leg sie einfach in einen Unterordner <b>schriften/</b> — "
        "der Dateiname ist egal, wir erkennen Bold automatisch daran, dass „bold“ oder "
        "„fett“ im Namen steht. Nichts in der JSON eintragen. Es werden bewusst nur diese "
        "zwei Schnitte verwendet, auch dort, wo sonst kursiv stünde.", "p"))
    a(P("Variable Fonts funktionieren nicht — davon brauchen wir statische Schnitte. "
        "Wenn deine Schrift eine Lizenz hat, die die Weitergabe verbietet, sag früh Bescheid, "
        "dann suchen wir gemeinsam eine Alternative.", "p"))

    a(P("Schritt 4 — Abgeben", "h3"))
    a(P("Ein Ordner, gezippt, an kontakt@beispiel-verein.de. Was drin sein muss, steht in der "
        "Checkliste am Ende dieser Anleitung. Wir schicken dir am selben Tag ein Muster-PDF mit acht "
        "echten Schildern zurück, damit du prüfen kannst, ob das Ergebnis stimmt.", "p"))

    a(Spacer(1, 4 * mm))

    # ---------------------------------------------------------- Weg B
    a(P("Weg B — InDesign, Datenzusammenführung", "h2"))
    a(P("Wenn du selbst produzieren willst, ist das der solide Weg. InDesign hat die Funktion "
        "seit vielen Jahren eingebaut, und sie ist für genau diesen Fall gemacht: viele gleiche "
        "Layouts, unterschiedlicher Text.", "p"))

    a(P("Die Datei, die du brauchst", "h3"))
    a(P("<b>ausgabe/Adobe/InDesign_Datenzusammenfuehrung.txt</b> — eine Tabelle mit einer Zeile "
        "pro Werk. Nicht in Excel öffnen und neu speichern, das zerstört die Umlaute. "
        "Einfach so lassen, wie sie ist.", "p"))

    a(P("Ablauf", "h3"))
    for n, t in [
        ("1.", "Neues Dokument anlegen: <b>Datei &gt; Neu &gt; Dokument</b>. Seitengröße A4, "
               "keine Musterseiten nötig."),
        ("2.", "Menü <b>Fenster &gt; Hilfsprogramme &gt; Datenzusammenführung</b> "
               "(englisch: Window &gt; Utilities &gt; Data Merge)."),
        ("3.", "Im Bedienfeld oben rechts auf das kleine Menü klicken, "
               "<b>Datenquelle auswählen</b>. Die .txt-Datei auswählen. Danach stehen im "
               "Bedienfeld die Feldnamen: Nr, Kuenstlerin, Titel, OrtJahr, Technik, Masse, "
               "Auflage, Hinweis."),
        ("4.", "Ein Textrahmen von 100 × 70 mm zeichnen und darin dein Schild aufbauen. "
               "Statt echten Text zu tippen, klickst du im Bedienfeld auf den Feldnamen — "
               "InDesign setzt einen Platzhalter wie <b>&lt;&lt;Titel&gt;&gt;</b> ein."),
        ("5.", "Absatzformate anlegen und den Platzhaltern zuweisen. Das ist wichtig: "
               "Nur was als Absatzformat definiert ist, sieht bei allen 100 Schildern gleich aus."),
        ("6.", "Im Bedienfeldmenü <b>Zusammengeführtes Dokument erstellen</b>. Im Reiter "
               "<b>Datensätze</b> auf <b>Mehrere Datensätze</b> stellen — dann kommen mehrere "
               "Schilder auf eine Seite. Im Reiter <b>Layout für mehrere Datensätze</b> die "
               "Ränder und Abstände so setzen, dass 2 Spalten × 4 Zeilen entstehen."),
        ("7.", "<b>Vorschau</b> anklicken und einmal durch alle Datensätze blättern. "
               "Achte auf das rote Übersatz-Kreuz am Textrahmen — das bedeutet, "
               "der Text passt nicht."),
        ("8.", "Als PDF exportieren, Vorgabe <b>Druckausgabequalität</b>, "
               "Schnittmarken einschalten."),
    ]:
        a(STEP(n, t))

    a(Spacer(1, 3 * mm))
    a(hinweis(
        "Der Stolperstein bei Weg B",
        "InDesign bricht Text nicht automatisch kleiner, wenn er zu lang ist. Bei Werk 63 mit "
        "dem langen Papiernamen erscheint stattdessen ein rotes Kreuz und der Text ist "
        "unsichtbar — im PDF fehlt dann einfach eine Zeile, ohne Warnung.",
        "Gegenmittel: Klick den Textrahmen an, dann <b>Objekt &gt; Textrahmenoptionen &gt; "
        "Automatische Größenanpassung</b>, oder gib der betroffenen Zeile im Absatzformat "
        "unter <b>Grundlegende Zeichenformate</b> genug Platz für zwei Zeilen. "
        "Und blättere vor dem Export wirklich durch alle Datensätze."))

    a(Spacer(1, 4 * mm))

    # ---------------------------------------------------------- Weg C
    a(P("Weg C — Illustrator, Variablenbedienfeld", "h2"))
    a(P("Illustrator kann das auch. Die Funktion heißt <b>Variablen</b> und ist etwas "
        "versteckter als das InDesign-Gegenstück, funktioniert aber nach demselben Prinzip: "
        "Du markierst Textobjekte als veränderlich und lädst dann eine Tabelle dazu.", "p"))

    a(P("Ablauf", "h3"))
    for n, t in [
        ("1.", "Neues Dokument, eine Zeichenfläche 100 × 70 mm. Dein Schild aufbauen, "
               "mit den Testtexten aus dem Abschnitt „Gestaltungsvorgaben“."),
        ("2.", "Menü <b>Fenster &gt; Variablen</b>."),
        ("3.", "Ein Textobjekt anklicken — zum Beispiel den Titel. Dann im Variablen-Bedienfeld "
               "unten auf den Knopf <b>Objekt dynamisch machen</b>. In der Liste erscheint eine "
               "Variable. Doppelklick darauf und exakt umbenennen in <b>Titel</b>."),
        ("4.", "Das für alle acht Textobjekte wiederholen. Die Namen müssen "
               "<b>zeichengenau</b> stimmen, sonst findet Illustrator die Daten nicht: "
               "Nr, Kuenstlerin, Titel, OrtJahr, Technik, Masse, Auflage, Hinweis. "
               "Kein Umlaut, kein Leerzeichen, Groß- und Kleinschreibung genau so."),
        ("5.", "Im Menü des Variablen-Bedienfeldes <b>Variablenbibliothek laden</b> "
               "beziehungsweise die Datenquelle importieren. Nimm "
               "<b>ausgabe/Adobe/Illustrator_Variablen.csv</b>."),
        ("6.", "Oben im Bedienfeld kannst du jetzt mit den Pfeilen durch die Datensätze "
               "blättern. Das Schild auf der Zeichenfläche ändert sich mit. "
               "Blättere einmal komplett durch."),
        ("7.", "Datenzusammenführung ausführen. Illustrator legt für jeden Datensatz eine "
               "eigene Zeichenfläche an — bei 100 Werken also 100 Zeichenflächen."),
        ("8.", "Danach musst du den Druckbogen selbst bauen: acht Schilder auf ein A4. "
               "Das ist Handarbeit oder ein Ausschieß-Schritt in InDesign."),
    ]:
        a(STEP(n, t))

    a(Spacer(1, 3 * mm))
    a(hinweis(
        "Warum wir Weg C nicht empfehlen, obwohl du Illustrator bevorzugst",
        "Drei konkrete Gründe, keine Geschmacksfrage:",
        "<b>Erstens</b> — 100 Zeichenflächen sind noch kein Druckbogen. Die Ausschießerei "
        "auf 8-pro-A4 machst du am Ende von Hand.",
        "<b>Zweitens</b> — Illustrator ist bei CSV-Dateien wählerisch. Adobe nennt "
        "ausdrücklich, dass kommagetrennte Dateien im Macintosh-Format nicht unterstützt "
        "werden. Unsere Datei ist im richtigen Format, aber wenn sie einmal in Excel geöffnet "
        "und neu gespeichert wird, ist sie es womöglich nicht mehr.",
        "<b>Drittens</b> — Text, der zu lang ist, wird in Illustrator einfach abgeschnitten "
        "oder läuft über den Rand hinaus. Beides fällt in einer 100-Zeichenflächen-Datei "
        "leicht durch.",
        "Falls die CSV nicht angenommen wird, liegt als Notfallvariante noch "
        "<b>Illustrator_Variablenbibliothek.xml</b> im selben Ordner. Wenn auch die zickt: "
        "kurz anrufen statt lange suchen."))

    a(Spacer(1, 4 * mm))

    # ---------------------------------------------------------- Gestaltung
    a(P("Gestaltungsvorgaben", "h2"))
    a(P("Innerhalb dieser Vorgaben ist alles deine Entscheidung: Schrift, Größen, Farben, "
        "Abstände, Ausrichtung, Linien.", "p"))

    a(P("Was fest ist", "h3"))
    a(tabelle([
        ["Vorgabe", "Wert", "Warum"],
        ["Format klein", "100 × 70 mm quer", "8 Stück auf A4, Schnittmarken passen"],
        ["Format groß", "140 × 100 mm quer", "4 Stück auf A4 quer, barrierearm"],
        ["Farbmodus", "CMYK oder Graustufen", "wird auf einem Bürodrucker ausgegeben"],
        ["Papier", "300 g, matt gestrichen, weiß", "Glanz spiegelt unter Ausstellungslicht"],
        ["Reihenfolge", "Nummer, Name, Titel, Ort/Jahr, Technik, Maß + Auflage, Hinweis",
         "Museumskonvention, Besucher:innen lesen von oben"],
        ["Randabstand", "mindestens 5 mm", "Schnittkante"],
        ["Kleinste Schrift", "nicht unter 7 pt", "sonst aus Lesedistanz nicht mehr lesbar"],
        ["Instagram-QR-Code", "ca. 12 mm, Ecke unten rechts",
         "Profil der/des Ausstellenden, vom Generator automatisch platziert"],
    ], [30 * mm, 46 * mm, W - ML - MR - 76 * mm]))

    a(Spacer(1, 3 * mm))
    a(P("Jedes Wandschild trägt unten rechts einen kleinen QR-Code zum Instagram-Profil "
        "der/des Ausstellenden. Der Generator setzt ihn selbst — halte die Ecke unten "
        "rechts frei (Kantenlänge über <b>qr_mm</b> in der design.json steuerbar, "
        "Standard 12 mm klein / 17 mm groß). Wer die Schilder selbst in Illustrator/InDesign "
        "setzt, findet die Codes als Vektor in <b>qr-codes/*.svg</b>.", "p"))

    a(Spacer(1, 5 * mm))
    a(P("Die Hierarchie", "h3"))
    a(P("Diese Abstufung ist Standard in Ausstellungen und sollte erkennbar bleiben — "
        "der Name ist das größte Element, der Titel folgt, alles Technische tritt zurück:", "p"))
    a(LI("<b>Ebene 1</b> — Name der Künstlerin oder des Künstlers. Am stärksten."))
    a(LI("<b>Ebene 2</b> — Werktitel. Kursiv ist üblich, aber nicht Pflicht."))
    a(LI("<b>Ebene 3</b> — Ort, Land, Jahr."))
    a(LI("<b>Ebene 4</b> — Druckverfahren, Papier, Bildmaß, Auflage. "
         "Zurückgenommen, gern grau."))
    a(LI("<b>Ebene 5</b> — Werknummer und die Zeile „Preis siehe Werkliste“. "
         "Am leisesten, klein und hell."))
    a(LI("<b>Nebenelement</b> — QR-Code unten rechts, ca. 12 mm. Ruhig, ohne Rahmen, "
         "nicht größer als die Werknummer wirkt."))

    a(Spacer(1, 4 * mm))
    a(P("Testtexte — bitte damit prüfen", "h3"))
    a(P("Das ist kein zufälliges Beispiel, sondern der jeweils längste Wert aus dem "
        "echten Datenbestand. Wenn dein Entwurf das hier aushält, hält er alles aus.", "p"))

    a(box([
        P("<font name='S-B'>Kürzester Fall</font>", "boxp"),
        P("008 · Aylin Demir · <i>Ohne Titel (Küche)</i> · Istanbul, Türkei, 2023 · "
          "Silbergelatine, Handabzug, Ilford Multigrade RC Deluxe · 24 × 30 cm · "
          "Auflage 8 (+2 AP)", "boxp"),
        Spacer(1, 2 * mm),
        P("<font name='S-B'>Längster Fall — hieran scheitern schlechte Layouts</font>", "boxp"),
        P("014 · Sophie Braun · <i>Wildcard: Grenzverlauf II</i> · Schengen, Luxemburg, 2026 · "
          "Pigmentdruck, Hahnemühle Photo Rag Baryta 315 g · 30 × 30 cm · Auflage 3 (+1 AP)",
          "boxp"),
        Spacer(1, 2 * mm),
        P("<font name='S-B'>Sonderfälle, die vorkommen</font>", "boxp"),
        P("— Statt der Preiszeile kann dort auch <b>unverkäuflich</b> stehen.<br/>"
          "— Namen mit Umlauten, Bindestrichen und Akzenten: Völklingen, Türkei, Dorée.<br/>"
          "— Das Multiplikationszeichen ist <b>×</b> (U+00D7), nicht der Buchstabe x. "
          "Prüf bitte, dass deine Schrift dieses Zeichen hat.", "boxp"),
    ]))

    a(Spacer(1, 4 * mm))

    # ---------------------------------------------------------- Werkliste
    a(P("Das zweite Stück: die Werkliste", "h2"))
    a(P("Neben den Schildern gehört eine gedruckte Werkliste zum System. Sie liegt an der "
        "Theke aus und enthält alles, was nicht an die Wand kommt: Preis, Rahmung, Blattmaß, "
        "Verfügbarkeit. Rund fünf Seiten A4, zweispaltig.", "p"))
    a(P("Für dich heißt das: Die Typografie der Werkliste soll erkennbar aus derselben Familie "
        "stammen wie die Schilder. Gleiche Schrift, gleiche Farbabstufungen, gleiche Logik der "
        "Hierarchie. Ein Vorschlag liegt als <b>ausgabe/02_Werkliste.pdf</b> bei.", "p"))
    a(P("Wenn du Weg A gehst, übernimmt der Generator deine Schrift und deine Farben "
        "automatisch auch für die Werkliste. Du musst dafür nichts Zusätzliches tun.", "p"))

    a(P("Und noch ein drittes, kleines Stück", "h3"))
    a(P("Rückseitenetiketten, 90 × 45 mm, die vor dem Aufhängen auf die Rahmenrückseite "
        "geklebt werden. Reine Gebrauchsgrafik, sieht kein Besucher. Die kannst du gestalten, "
        "musst du aber nicht — sag einfach Bescheid.", "p"))

    a(Spacer(1, 4 * mm))

    # ---------------------------------------------------------- Checkliste
    a(P("Checkliste für die Abgabe", "h2"))

    a(P("Weg A", "h3"))
    for t in ["<b>design.json</b> — ausgefüllt, alle Zahlen und Farbwerte eingetragen",
              "<b>schriften/</b> — zwei Dateien: Regular und Bold (.ttf oder .otf, Dateiname egal)",
              "<b>Entwurf.pdf</b> — dein gestaltetes Musterschild, damit wir vergleichen können",
              "<b>Entwurf.ai</b> — die offene Datei, falls wir etwas nachmessen müssen",
              "kurze Notiz, ob 100 × 70 mm oder 140 × 100 mm gilt"]:
        a(LI(t))

    a(P("Weg B", "h3"))
    for t in ["die <b>.indd</b>-Datei mit den angelegten Absatzformaten",
              "das <b>zusammengeführte PDF</b> aller Schilder mit Schnittmarken",
              "die verwendeten <b>Schriftdateien</b>",
              "Bestätigung, dass du einmal durch alle Datensätze geblättert hast "
              "und nirgends ein Übersatz-Kreuz stand"]:
        a(LI(t))

    a(P("Weg C", "h3"))
    for t in ["die <b>.ai</b>-Datei mit den benannten Variablen",
              "das fertig <b>ausgeschossene PDF</b> mit 8 Schildern pro A4 und Schnittmarken",
              "die verwendeten <b>Schriftdateien</b>",
              "Bestätigung, dass du alle Datensätze durchgeblättert hast"]:
        a(LI(t))

    a(Spacer(1, 5 * mm))
    a(hinweis(
        "Bitte nicht",
        "— Keine Schriften in Pfade umwandeln, bevor wir das Muster freigegeben haben. "
        "Danach gern.<br/>"
        "— Keine Platzhalter-Blindtexte im abgegebenen PDF. Wir haben schon einmal ein "
        "Schild mit „Lorem ipsum“ an einer Wand hängen sehen.<br/>"
        "— Die Datendateien im Ordner <b>Adobe/</b> nicht in Excel öffnen und speichern. "
        "Das macht aus jedem ü ein Ã¼."))

    a(Spacer(1, 6 * mm))
    a(P("Wenn etwas nicht klappt", "h2"))
    a(tabelle([
        ["Symptom", "Ursache und Lösung"],
        ["Umlaute erscheinen als <b>Ã¼</b> oder <b>?</b>",
         "Die Datendatei wurde in Excel geöffnet und neu gespeichert. Nimm eine frische "
         "Kopie aus dem Ordner Adobe/ und öffne sie gar nicht erst."],
        ["Illustrator nimmt die CSV nicht an",
         "Datei nicht verändern, nicht umbenennen, direkt laden. Falls es weiter hakt, nimm "
         "die .xml aus demselben Ordner. Falls auch das nicht geht: anrufen."],
        ["InDesign zeigt ein rotes Kreuz am Textrahmen",
         "Der Text ist länger als der Rahmen. Rahmen vergrößern, Schrift verkleinern oder "
         "im Absatzformat zwei Zeilen zulassen."],
        ["Im zusammengeführten PDF fehlt eine Zeile",
         "Dasselbe Problem, nur unbemerkt geblieben. Zurück in die Vorschau und alle "
         "Datensätze durchblättern."],
        ["Das × wird als Kästchen dargestellt",
         "Deine Schrift hat das Multiplikationszeichen nicht. Anderen Schnitt wählen "
         "oder uns Bescheid geben."],
        ["Der Generator meldet „Schriftdateien nicht gefunden“",
         "Die Pfade in der design.json stimmen nicht mit den tatsächlichen Dateinamen "
         "überein. Groß- und Kleinschreibung zählt."],
    ], [46 * mm, W - ML - MR - 46 * mm]))

    a(Spacer(1, 7 * mm))
    a(box([
        P("<font name='S-B' size='11'>Fragen?</font>", "boxp"),
        P("Ruf lieber einmal zu früh an als einmal zu spät. Eine Rückfrage kostet fünf Minuten, "
          "ein Nachdruck von 100 Schildern kostet einen Abend.", "boxp"),
        P("<b>kontakt@beispiel-verein.de</b> · fototage-musterstadt.de · "
          "instagram.com/fototage.musterstadt", "boxp"),
        P("Team / Kulturverein Musterstadt | FOTOTAGE MUSTERSTADT 2026", "boxp"),
    ]))

    return F


def main():
    doc = BaseDocTemplate(OUT, pagesize=A4, leftMargin=ML, rightMargin=MR,
                          topMargin=MT, bottomMargin=MB,
                          title="FOTOTAGE MUSTERSTADT 2026 — Wandschilder, Anleitung für Designer",
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
