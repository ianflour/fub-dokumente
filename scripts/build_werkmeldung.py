#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Erzeugt Werkmeldung_VORLAGE.xlsx.

Eine gemeinsame Datei für ALLE Teilnehmenden — gedacht als gemeinsam
bearbeitete Tabelle, in die alle gleichzeitig ihre Werke eintragen. Jede Zeile
gehört über die Spalte „Künstler:in" (Dropdown) zu einer Person; die
Kontaktdaten stehen einmalig im Blatt „Kontakte".
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import _pfade as P
import _saal_listen as SL
import _teilnehmende as T

OUT = P.output("Werkmeldung_VORLAGE.xlsx")
FONT = "Arial"
FILL_HEAD = PatternFill("solid", fgColor="D9D9D9")
FILL_KEY = PatternFill("solid", fgColor="FFF2CC")
FILL_SUB = PatternFill("solid", fgColor="EDEDED")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

EINGABE_ZEILEN = 200

# Blattschutz: sperrt Kopfzeilen/Formeln/Referenzblätter gegen Verändern.
# Kein Sicherheitsmerkmal (das Passwort steht im Handbuch) — nur ein Schutz
# gegen versehentliches Löschen/Überschreiben in der gemeinsam bearbeiteten
# Datei. Freigeschaltete (gelbe/blaue) Zellen bleiben für alle editierbar.
PASSWORT = "werkmeldung"


def schuetzen(ws, autofilter=False):
    ws.protection.sheet = True
    ws.protection.password = PASSWORT
    ws.protection.formatCells = False
    if autofilter:
        ws.protection.autoFilter = False


def frei(cell):
    cell.protection = Protection(locked=False)


def text_zelle(c, wert):
    """Setzt einen Wert und erzwingt Text, falls er mit = + @ beginnt —
    sonst liest das Tabellenprogramm ihn als (kaputte) Formel und zeigt
    #NAME? / #ERROR! an."""
    c.value = wert
    if isinstance(wert, str) and wert[:1] in ("=", "+", "@"):
        c.data_type = "s"
    return c

# (Spaltenname, Breite, Hinweistext)  — Pflichtfelder fürs Schild sind gelb.
SPALTEN = [
    ("Künstler:in", 22, "Dropdown. Dein Name genau so, wie er auf dem Schild stehen soll. "
                        "In jeder deiner Zeilen wählen."),
    ("Titel", 28, "Werktitel. Unbetitelt: „Ohne Titel“ oder „Ohne Titel (Küche)“."),
    ("Ort", 15, "Aufnahmeort, z. B. Bari"),
    ("Land", 13, "z. B. Italien"),
    ("Jahr", 7, "Aufnahmejahr, vierstellig"),
    ("Druckverfahren", 18, "Dropdown. Saal-Foto → C-Print, Saal-FineArt → Pigmentdruck, "
                           "Alu-Dibond/Acryl → Direktdruck. Details im Blatt „Saal Digital“."),
    ("Papier", 40, "Dropdown mit den Saal-Papieren/Trägern. Anderes Labor: „anderes Papier …“ "
                   "wählen und Hersteller + Typ + Grammatur in die Zelle schreiben."),
    ("Bildmaß H cm", 12, "Höhe des sichtbaren Bildes OHNE Weißrand, nur Zahl. Dropdown = Eingabehilfe."),
    ("Bildmaß B cm", 12, "Breite des sichtbaren Bildes OHNE Weißrand, nur Zahl."),
    ("Blattmaß H cm", 12, "Höhe des Papiers INKL. Weißrand, nur Zahl."),
    ("Blattmaß B cm", 12, "Breite des Papiers INKL. Weißrand, nur Zahl."),
    ("Auflage", 8, "Editionsgröße als Zahl, z. B. 5. Unikat = 1"),
    ("AP", 6, "Anzahl Artist Proofs als Zahl, sonst 0"),
    ("Exemplar", 10, "welches Exemplar hängt, z. B. 2/5"),
    ("Rahmenmodell", 24, "Hersteller, Modell, Farbe, z. B. Nielsen Alpha 50×70 cm, alu roh. "
                         "„50x70“ tippen ist ok — wird beim Einlesen zu „50×70“."),
    ("Rahmenmaß", 12, "Außenmaß des Rahmens, z. B. 50×70 cm. „50x70“ genügt — das „x“ "
                      "wird automatisch zu „×“."),
    ("Passepartout", 18, "Farbe/Typ oder „ohne“"),
    ("Preis EUR", 11, "Endpreis in Euro, nur Zahl. Unverkäuflich: leer lassen"),
    ("Rahmen im Preis", 14, "ja / nein"),
    ("Versicherungswert EUR", 18, "Wiederbeschaffungswert, nur Zahl"),
    ("Verkäuflich", 12, "ja / nein"),
]
PFLICHT = {"Künstler:in", "Titel", "Ort", "Land", "Jahr", "Druckverfahren", "Papier",
           "Bildmaß H cm", "Bildmaß B cm", "Auflage"}

BEISPIEL = ["(Beispiel — bitte stehen lassen)", "Auto eines alten Fischers", "Bari", "Italien",
            2021, "C-Print", "Fujifilm Crystal Archive DP II Silk Portrait 232 g/m²",
            27, 40, 30, 45, 5, 1, "2/5", "Boesner Uno 40×50 cm, schwarz", "40×50 cm",
            "Dorée, 3 mm Weißkern", 75, "ja", 120, "ja"]

GLOSSAR = [
    ("Bildmaß", "Höhe × Breite des sichtbaren Bildes, OHNE den Weißrand. Kommt so aufs Wandschild."),
    ("Blattmaß", "Höhe × Breite des ganzen Papiers, INKLUSIVE Weißrand. Nur für Werkliste und Rückseite."),
    ("Rahmenmaß", "Außenmaß des fertigen Rahmens."),
    ("Weißrand / Rand", "Unbedruckter Papierrand rund ums Bild."),
    ("Auflage / Edition", "Gesamtzahl der Abzüge, die von einem Motiv verkauft werden. „Auflage 5“ = "
                          "fünf nummerierte Abzüge insgesamt."),
    ("Unikat", "Einzelstück, Auflage 1."),
    ("AP (Artist Proof, e. a.)", "Künstlerexemplare zusätzlich zur Auflage, meist 1–3. Bleiben bei "
                                 "der Künstlerin und werden separat gezählt."),
    ("Exemplar", "Welcher Abzug hängt, z. B. „2/5“ = zweiter von fünf. Nur intern, nicht aufs Schild."),
    ("C-Print (C-Type)", "Fotografische Belichtung auf lichtempfindliches Farbpapier, nass entwickelt. "
                         "Bei Saal Digital: Fujifilm Crystal Archive."),
    ("Pigmentdruck / FineArt / Giclée", "Tintenstrahldruck mit pigmentierten Tinten auf Künstlerpapier "
                                        "(z. B. Hahnemühle). Sehr lichtecht und langlebig."),
    ("Digitaldruck", "Hochwertiger Digitaldruck für Poster und Kunstdrucke."),
    ("Direktdruck", "Die Farbe wird direkt auf eine Trägerplatte gedruckt (Alu-Dibond, Acrylglas, "
                    "Hartschaum)."),
    ("Latexdruck", "Druck mit Latextinten, u. a. auf Leinwand und Digitalvlies."),
    ("UV-Druck", "Digitaldruck mit sofort UV-gehärteter Tinte, u. a. auf Digitalvlies und Platten."),
    ("Solventdruck / Öko-Solvent", "Großformat-Digitaldruck mit lösemittelhaltiger Tinte, die ins "
                                   "Medium einzieht — wisch- und wetterfest ohne Laminat. Öko-Solvent "
                                   "ist die geruchsärmere Variante. Typisch für Plakate (Blueback) und "
                                   "Folien."),
    ("Digitalvlies / Vliestapete", "PVC-freies, dimensionsstabiles Vlies als Bildträger für "
                                   "großformatige Wandbilder, im Latex-/UV-Druck bedruckt "
                                   "(z. B. ERFURT DV150/DV195/DV6xx). Glatt oder mit geprägter "
                                   "Struktur. Wird wie Tapete an die Wand gebracht."),
    ("Affichenpapier / Blueback / Blaurücken", "Einseitig gestrichenes Plakatpapier (115–135 g/m²) "
                                   "mit blaugrauem, ungestrichenem Rücken: das alte Plakat darunter "
                                   "scheint nicht durch, und Kleister haftet gut. Holzfrei, nassfest, "
                                   "wetterfest ca. 4–6 Monate. Offset-, Digital-, Solvent-, Latex- oder "
                                   "UV-Druck. Für großformatige Plakate, nicht für den Innenraum-Abzug."),
    ("Whiteback / City-Light-Papier", "Plakatpapier mit ungefärbtem, lichtdurchlässigem Rücken für "
                                   "hinterleuchtete Leuchtkästen (City-Light-Plakate an Haltestellen, "
                                   "in Passagen). Sonst wie Blueback. Nicht zu verwechseln mit "
                                   "Backlit-Folie."),
    ("Inkjet", "Sammelbegriff für Tintenstrahldruck."),
    ("Silbergelatine, Handabzug", "Klassischer Schwarzweiß-Abzug im Fotolabor auf Barytpapier, von Hand."),
    ("Cyanotypie", "Historisches Edeldruckverfahren mit preußischblauen Bildern (Blaudruck). "
                   "Eisensalz-Beschichtung, Kontaktbelichtung mit UV/Sonne, Auswässerung mit Wasser. "
                   "Klassisch = Herschel-Formel (1842)."),
    ("Cyanotypie (New Cyanotype)", "Variante nach Mike Ware: eine stabilere Einbad-Lösung, "
                                   "kürzere Belichtung, sattere Blautöne, weniger Ausbluten."),
    ("Cyanotypie, getont", "Fertige Cyanotypie, nachträglich in Tannin, Tee, Kaffee oder Rotwein "
                           "umgetont — von Violett über Braun bis Schwarzblau."),
    ("Nass-Cyanotypie (Wet Cyanotype)", "Cyanotypie, die noch feucht belichtet wird, oft mit Zusätzen "
                                        "(Seife, Essig, Salz) — fleckige, unvorhersehbare Verläufe."),
    ("Kollodium-Nassplatte (Wet Plate)", "Historisches Verfahren (ab 1851): Glas- oder Metallplatte wird "
                                         "mit Kollodium beschichtet, im Silberbad sensibilisiert und muss "
                                         "nass — innerhalb von Minuten — belichtet und entwickelt werden. "
                                         "Jede Platte ist ein Unikat."),
    ("Ambrotypie", "Kollodium-Nassplatten-Positiv auf Glas (dunkel hinterlegt). Unikat."),
    ("Ferrotypie (Tintype / Melainotype)", "Kollodium-Nassplatten-Positiv auf schwarz lackiertem "
                                           "Metallblech. Unikat."),
    ("Pannotypie", "Kollodium-Nassplatten-Positiv auf Wachstuch, Leder oder schwarzem Textil. Unikat."),
    ("Kollodium-Negativ", "Kollodium-Nassplatte als Glasnegativ, von dem Kontaktabzüge (z. B. Salzpapier, "
                          "Albumin) gemacht werden."),
    ("Risografie", "Schablonendruck (Riso): körnige Farbflächen, kleine Auflagen."),
    ("Offsetdruck", "Industrieller Flachdruck für hohe Auflagen."),
    ("Alu-Dibond / Alu-Verbund", "Verbundplatte aus zwei Alublechen mit Kunststoffkern. Stabil, leicht, "
                                 "rahmenlos hängbar."),
    ("Acrylglas, hinterklebt", "Foto rückseitig hinter Acrylglas kaschiert; glänzende Tiefenwirkung."),
    ("GalleryPrint", "Saal-Bezeichnung: Fotoabzug rückseitig auf Acrylglas kaschiert."),
    ("Hartschaumplatte", "Leichte Kunststoffplatte (Forex/PVC), günstige Direktdruck-Variante."),
    ("Hahnemühle", "Deutscher Hersteller von FineArt-Papieren (Photo Rag, Baryta, Museum Etching …)."),
    ("Fujifilm Crystal Archive", "Farbfotopapier von Fujifilm für C-Prints. Oberflächen bei Saal: "
                                 "Glänzend, Matt, Silk (Silk Portrait). „SoftTouch matt“ ist derselbe "
                                 "Abzug mit zusätzlicher matter Laminierfolie."),
    ("Art Line / Kunstdruck (Saal)", "Fotoabzug-Alternativen bei Saal, die per Digitaldruck — nicht "
                                     "fotografisch — auf strukturiertes Papier gedruckt werden. Art Line "
                                     "ist FSC-zertifiziertes Recyclingpapier. Aufs Schild dann "
                                     "„Digitaldruck“, nicht „C-Print“."),
    ("FSC-zertifiziert", "Papier aus verantwortungsvoll bewirtschafteter Forstwirtschaft (Forest "
                         "Stewardship Council). Bei Saal-Fotoabzügen: Art Line und SoftTouch matt."),
    ("Grammatur (g/m²)", "Flächengewicht des Papiers — grob ein Maß für die Dicke/Stärke."),
    ("Passepartout", "Kartonrahmen zwischen Bild und Glas, hält das Bild auf Abstand zum Glas."),
    ("Verkäuflich (ja/nein)", "Ob die Arbeit während der Ausstellung verkauft werden darf."),
    ("Rahmen im Preis (ja/nein)", "Ob der Rahmen im genannten Preis enthalten ist („inkl. Rahmen“) "
                                  "oder extra kostet („zzgl. Rahmen“). Häufigste Rückfrage im Verkauf."),
    ("Endpreis", "Der Betrag, den ein:e Besucher:in tatsächlich zahlt — nicht dein Nettoanteil."),
    ("Provision / Ausstellungsanteil", "Anteil am Verkaufspreis, der an die Ausstellungsorganisation geht. "
                                       "Im Blatt „Preishilfe“ einstellbar."),
    ("Versicherungswert", "Wiederbeschaffungswert: was es kosten würde, die Arbeit bei Verlust oder "
                          "Schaden neu herzustellen. Für die Ausstellungsversicherung."),
    ("Kleinunternehmer (§ 19 UStG)", "Wer darunter fällt, weist keine Umsatzsteuer aus. Kunstverkäufe "
                                     "sind zudem oft von der Preisauszeichnungspflicht befreit (§ 9 PAngV)."),
    ("Werknummer", "Fortlaufende Nummer, die das Kernteam vergibt. Nicht selbst eintragen."),
    ("Wand / Position", "Hängeort im Raum. Vergibt das Kernteam nach dem Hängeplan."),
    ("Wandschild", "Kleines Schild neben dem Bild: Name, Titel, Ort/Jahr, Technik, Maß, Auflage, QR-Code."),
    ("Werkliste", "Die an der Theke ausliegende Liste mit Preisen, Rahmung und Verfügbarkeit."),
    ("Vernissage / Finissage", "Eröffnung bzw. Abschlussveranstaltung der Ausstellung."),
    ("Handle (Instagram)", "Der Profilname nach dem @ — bei instagram.com/dein.profil ist „dein.profil“ "
                           "der Handle."),
]


def blatt_preishilfe(wb):
    """Interaktiver Rechner für einen Verkaufspreis-Vorschlag (unverbindlich)."""
    ps = wb.create_sheet("Preishilfe")
    ps.sheet_view.showGridLines = False
    ps.column_dimensions["A"].width = 36
    ps.column_dimensions["B"].width = 13
    ps.column_dimensions["C"].width = 70

    def kopf(r, t):
        c = text_zelle(ps.cell(row=r, column=1), t)
        c.font = Font(name=FONT, size=10, bold=True)
        c.fill = FILL_HEAD
        for col in (2, 3):
            ps.cell(row=r, column=col).fill = FILL_HEAD

    def L(r, a, note=""):
        c = text_zelle(ps.cell(row=r, column=1), a)
        c.font = Font(name=FONT, size=10)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        n = text_zelle(ps.cell(row=r, column=3), note)
        n.font = Font(name=FONT, size=9, italic=True, color="808080")
        n.alignment = Alignment(wrap_text=True, vertical="center")

    def IN(r, val, fmt='#,##0.00'):
        c = ps.cell(row=r, column=2, value=val)
        c.font = Font(name=FONT, size=10, color="0000FF")
        c.fill = FILL_KEY
        c.border = BORDER
        c.number_format = fmt
        frei(c)
        return f"B{r}"

    def FX(r, formel, fmt='#,##0.00 "€"', gross=False):
        c = ps.cell(row=r, column=2, value=formel)
        c.font = Font(name=FONT, size=12 if gross else 10, bold=True)
        c.fill = FILL_SUB
        c.border = BORDER
        c.number_format = fmt
        return f"B{r}"

    ps.cell(row=1, column=1, value="Preishilfe — Verkaufspreis abschätzen").font = \
        Font(name=FONT, size=13, bold=True)
    L(3, "Trag die gelben Felder für EIN Werk aus. Der Vorschlag unten ist ein Richtwert, "
         "kein Muss — du kannst jeden Preis eintragen.")
    L(4, "Danach die Zahl (oder deine eigene) im Blatt „Werke“ unter „Preis EUR“ eintragen.")

    kopf(6, "1 · Kosten je Abzug")
    b_druck = IN(7, 0);  L(7, "Druck + Papier (Labor)", "z. B. deine Saal-Bestellung für dieses Format")
    IN(8, 0);            L(8, "Rahmen", "Leiste, Zuschnitt, Rückwand")
    IN(9, 0);            L(9, "Passepartout + Glas / Acryl", "oder 0")
    b_sonst = IN(10, 0); L(10, "Sonstiges", "Versand, Aufhängung, Verpackung")
    b_selbst = FX(11, f"=SUM({b_druck}:{b_sonst})")   # Zeilen 7–10 (Rahmen/PP inbegriffen)
    L(11, "→ Selbstkosten je Abzug")

    kopf(13, "2 · Deine Arbeit")
    b_zeit = IN(14, 0, "0.0");  L(14, "Arbeitszeit je Abzug (Std)", "Proof, Einrahmen, Signieren, Verpacken")
    b_vor  = IN(15, 0, "0.0");  L(15, "Einmalige Vorarbeit je Motiv (Std)", "Bildbearbeitung, Testdrucke — nur einmal")
    b_aufl = IN(16, 5, "0");    L(16, "Auflage (Editionsgröße)", "auf so viele Abzüge verteilt sich die Vorarbeit")
    b_satz = IN(17, 45, '#,##0'); L(17, "Stundensatz (EUR)", "Vorschlag 40–60")
    b_arb  = FX(18, f"={b_satz}*({b_zeit}+{b_vor}/MAX({b_aufl},1))")
    L(18, "→ Arbeitswert je Abzug")

    kopf(20, "3 · Aufschlag und Provision")
    b_fakt = IN(21, 2.5, "0.0"); L(21, "Faktor Künstlerhonorar", "übliche Spanne 2–4 auf (Kosten + Arbeit)")
    b_prov = IN(22, 40, '0');    L(22, "Ausstellungsprovision (%)", "Anteil KULTURVEREIN MUSTERSTADT / Räumchen e. V.")
    b_mwst = ps.cell(row=23, column=2, value="nein")
    b_mwst.font = Font(name=FONT, size=10, color="0000FF"); b_mwst.fill = FILL_KEY; b_mwst.border = BORDER
    frei(b_mwst)
    L(23, "MwSt (19 %) aufschlagen?", "Kleinunternehmer nach § 19 UStG: nein")
    dv = DataValidation(type="list", formula1='"ja,nein"', allow_blank=True, showDropDown=False)
    ps.add_data_validation(dv); dv.add("B23")

    b_netto = FX(24, f'=({b_selbst}+{b_arb})*{b_fakt}*IF(LOWER(B23)="ja",1.19,1)')
    L(24, "→ Netto (vor Provision)")
    # IFERROR: falls Provision versehentlich auf 100 (%) steht -> Division durch 0
    b_end = FX(25, f"=IFERROR({b_netto}/(1-{b_prov}/100),0)")
    L(25, "→ Endpreis inkl. Provision")

    FX(27, f"=IFERROR(CEILING({b_end},10),0)", '#,##0 "€"', gross=True)
    ps.cell(row=27, column=1, value="► EMPFOHLENER VERKAUFSPREIS").font = Font(name=FONT, size=12, bold=True)
    L(28, "", "auf 10 € aufgerundet")

    kopf(30, "Gegenprobe nach Bildfläche")
    b_h = IN(31, 0, "0.#"); L(31, "Bildmaß Höhe (cm)")
    b_b = IN(32, 0, "0.#"); L(32, "Bildmaß Breite (cm)")
    b_dm = IN(33, 20, '#,##0'); L(33, "Preis pro dm²", "Fotografie grob 12–30 €")
    b_soc = IN(34, 120, '#,##0'); L(34, "Sockelbetrag", "Grundpreis unabhängig von der Größe")
    FX(35, f"=({b_h}/10)*({b_b}/10)*{b_dm}+{b_soc}", '#,##0 "€"')
    L(35, "→ Richtwert nach Fläche", "Fläche in dm². Weicht der Wert stark vom Vorschlag oben ab, nochmal prüfen.")

    L(37, "")
    c = ps.cell(row=38, column=1,
                value="Alle Werte sind unverbindliche Richtwerte. Du entscheidest über deinen Preis.")
    c.font = Font(name=FONT, size=9, italic=True, color="808080")
    ps.merge_cells("A38:C38")

    ps.freeze_panes = "A6"
    schuetzen(ps)


def blatt_glossar(wb):
    gl = wb.create_sheet("Glossar")
    gl.sheet_view.showGridLines = False
    gl.column_dimensions["A"].width = 28
    gl.column_dimensions["B"].width = 98
    gl.cell(row=1, column=1, value="Glossar — die Begriffe in dieser Datei").font = \
        Font(name=FONT, size=13, bold=True)
    for r, (begriff, erkl) in enumerate(GLOSSAR, start=3):
        a = gl.cell(row=r, column=1, value=begriff)
        a.font = Font(name=FONT, size=10, bold=True)
        a.alignment = Alignment(wrap_text=True, vertical="top")
        b = gl.cell(row=r, column=2, value=erkl)
        b.font = Font(name=FONT, size=10)
        b.alignment = Alignment(wrap_text=True, vertical="top")
    gl.freeze_panes = "A3"
    schuetzen(gl)


def build():
    wb = Workbook()

    # ---------------------------------------------------------- Deckblatt
    ws = wb.active
    ws.title = "Bitte lesen"
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 104
    ws.sheet_view.showGridLines = False

    text = [
        ("FOTOTAGE MUSTERSTADT 2026 — Werkmeldung", ""),
        ("", ""),
        ("", "Das ist eine GEMEINSAME Datei für alle Teilnehmenden. Bitte trag hier die "
             "Angaben zu deinen ausgestellten Arbeiten ein — daraus entstehen die "
             "Wandschilder und die ausliegende Werkliste."),
        ("", ""),
        ("So geht's", "1. Blatt „Werke“ öffnen."),
        ("", "2. Pro ausgestelltem Werk EINE Zeile ausfüllen — auch bei Serien: "
             "jedes Bild bekommt ein eigenes Schild."),
        ("", "3. In jeder deiner Zeilen links unter „Künstler:in“ deinen Namen aus dem "
             "Dropdown wählen. Immer dieselbe Schreibweise — sonst landen deine Werke "
             "unter zwei Namen."),
        ("", "4. Einmalig im Blatt „Kontakte“ deine E-Mail und Telefonnummer eintragen "
             "(dein Name steht schon da)."),
        ("", "5. Nichts weiter tun — wir sammeln die Datei zum Stichtag ein."),
        ("", ""),
        ("Dropdowns", "Künstler:in, Druckverfahren, Papier und die vier Maßfelder haben eine "
                      "Auswahlliste. Die Papierliste ist die von Saal Digital — druckst du "
                      "woanders, wähle „anderes Papier …“ und trag Hersteller, Typ und "
                      "Grammatur von Hand ein. Bei den Maßen darfst du jederzeit eigene "
                      "Zahlen tippen (auch 27,5)."),
        ("Blatt „Saal Digital“", "Übersicht: welches Saal-Papier / welcher Bildträger zu welchem "
                                 "Druckverfahren gehört und welche Formate es gibt. Nur zum "
                                 "Nachschlagen."),
        ("Blatt „Preishilfe“", "Rechner für einen Verkaufspreis-Vorschlag: die gelben Felder "
                               "ausfüllen, unten steht ein Richtwert. Unverbindlich."),
        ("Blatt „Glossar“", "Alle Begriffe dieser Datei kurz erklärt — Bildmaß, Auflage, AP, "
                            "C-Print, Passepartout, Provision und so weiter."),
        ("Blatt „Listen“", "Ganz hinten, gesperrt. Speist die Auswahllisten für Künstler:in, "
                           "Druckverfahren und Papier. Bitte nicht bearbeiten."),
        ("", ""),
        ("Gesperrte Bereiche", "Kopfzeilen, Formeln und die Nachschlage-Blätter (Saal Digital, "
                               "Glossar, Preishilfe-Formeln) sind gegen Verändern geschützt — "
                               "damit in der gemeinsamen Datei nichts kaputtgeht. Schreiben kannst "
                               "du überall dort, wo die Schrift blau ist bzw. die Zelle gelb "
                               "hinterlegt ist: deine Werk-Zeilen, deine Kontaktzeile, die gelben "
                               "Felder in „Preishilfe“."),
        ("", ""),
        ("Frist", "Bitte bis spätestens 22. August 2026. Danach gehen die Schilder in den "
                  "Druck; spätere Änderungen kosten einen kompletten Nachdruck."),
        ("", ""),
        ("Was aufs Schild kommt", "Name · Titel · Ort, Land, Jahr · Druckverfahren und Papier · "
                                  "Bildmaß · Auflage. Sonst nichts."),
        ("Was NICHT aufs Schild kommt", "Rahmen, Passepartout, Preis, Blattmaß. Das brauchen wir "
                                        "trotzdem — es steht in der ausliegenden Werkliste "
                                        "bzw. auf der Rahmenrückseite."),
        ("Instagram", "Dein Instagram-Profil kommt als QR-Code aufs Wandschild und auf eine "
                      "Übersichtswand. Der Code liegt schon vor. Falls nicht (siehe „Kontakte“, "
                      "Spalte Instagram ist leer): Handle dort eintragen."),
        ("", ""),
        ("Zwei häufige Fehler", ""),
        ("Maße", "Wir brauchen zwei Maßpaare: Bildmaß = das sichtbare Bild ohne Weißrand. "
                 "Blattmaß = das Papier inklusive Weißrand. Immer Höhe × Breite, "
                 "immer nur die Zahl (45, nicht „45 cm“)."),
        ("Preis", "Bitte den Endpreis, den ein:e Besucher:in zahlt — nicht deinen Nettoanteil. "
                  "Und gib an, ob der Rahmen im Preis enthalten ist. Das ist die häufigste "
                  "Frage im Verkaufsgespräch. Unsicher? Das Blatt „Preishilfe“ rechnet dir "
                  "einen Vorschlag aus."),
        ("", ""),
        ("Beispielzeile", "Die graue Zeile 3 im Blatt „Werke“ ist ein ausgefülltes Beispiel. "
                          "Bitte einfach stehen lassen — wir filtern sie beim Import heraus."),
        ("", ""),
        ("Fragen", "kontakt@beispiel-verein.de"),
        ("", "Team / Kulturverein Musterstadt | FOTOTAGE MUSTERSTADT 2026"),
    ]
    for i, (a, b) in enumerate(text, start=1):
        ws.cell(row=i, column=1, value=a).font = Font(name=FONT, size=10, bold=True)
        c = ws.cell(row=i, column=2, value=b)
        c.font = Font(name=FONT, size=10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row=1, column=1).font = Font(name=FONT, size=14, bold=True)
    schuetzen(ws)

    # ---------------------------------------------------------- Listen (Quelle für Dropdowns)
    wl = wb.create_sheet("Listen")
    wl.sheet_view.showGridLines = False
    spalten_listen = [
        ("A", "Künstler:in", list(T.NAMEN)),
        ("B", "Ja/Nein", SL.JA_NEIN),
        ("C", "Druckverfahren", SL.DRUCKVERFAHREN),
        ("D", "Papier", SL.PAPIERE),
        ("E", "Maße cm", SL.MASSE_CM),
    ]
    bereich = {}
    for col, titel, werte in spalten_listen:
        wl[f"{col}1"] = titel
        wl[f"{col}1"].font = Font(name=FONT, size=9, bold=True)
        wl[f"{col}1"].fill = FILL_HEAD
        for i, v in enumerate(werte, start=2):
            wl[f"{col}{i}"] = v
            wl[f"{col}{i}"].font = Font(name=FONT, size=10)
        wl.column_dimensions[col].width = 42 if col == "D" else 22
        # Bereichsbezug OHNE führendes "=" — manche Tabellenprogramme können die
        # Dropdown-Regel sonst beim xlsx-Import nicht lesen und lässt sie ganz weg.
        bereich[titel] = f"Listen!${col}$2:${col}${len(werte) + 1}"
    wl["G1"] = ("Bitte nicht bearbeiten — dieses Blatt füllt die Auswahllisten "
                "im Blatt „Werke“.")
    wl["G1"].font = Font(name=FONT, size=10, bold=True, color="C00000")
    wl.column_dimensions["G"].width = 70
    # Blatt bleibt SICHTBAR: manche Tabellenprogramme importieren Dropdown-
    # Quellen von ausgeblendeten Blättern nicht zuverlässig. Wird zum Schluss ans Ende
    # der Blattreihe geschoben, damit es nicht im Weg ist.
    schuetzen(wl)

    # ---------------------------------------------------------- Werke
    w = wb.create_sheet("Werke")
    head = 1
    for i, (name, breite, hinweis) in enumerate(SPALTEN, start=1):
        col = get_column_letter(i)
        w.column_dimensions[col].width = breite
        c = w.cell(row=head, column=i, value=name)
        c.font = Font(name=FONT, size=9, bold=True)
        c.fill = FILL_KEY if name in PFLICHT else FILL_HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = BORDER
        h = w.cell(row=head + 1, column=i,
                   value="(Hinweis — nicht ausfüllen)" if i == 1 else hinweis)
        h.font = Font(name=FONT, size=8, italic=True, color="808080")
        h.alignment = Alignment(wrap_text=True, vertical="top")
        h.border = BORDER
    w.row_dimensions[head].height = 30
    w.row_dimensions[head + 1].height = 66

    for i, v in enumerate(BEISPIEL, start=1):
        c = w.cell(row=head + 2, column=i, value=v)
        c.font = Font(name=FONT, size=9, italic=True, color="A6A6A6")
        c.border = BORDER

    first, last = head + 3, head + 2 + EINGABE_ZEILEN
    for r in range(first, last + 1):
        for i in range(1, len(SPALTEN) + 1):
            c = w.cell(row=r, column=i)
            c.font = Font(name=FONT, size=10, color="0000FF")
            c.border = BORDER
            frei(c)
        w.row_dimensions[r].height = 18

    for name, fmt in [("Preis EUR", '#,##0'), ("Versicherungswert EUR", '#,##0'),
                      ("Jahr", "0"), ("Auflage", "0"), ("AP", "0"),
                      ("Bildmaß H cm", "0.#"), ("Bildmaß B cm", "0.#"),
                      ("Blattmaß H cm", "0.#"), ("Blattmaß B cm", "0.#")]:
        ci = [s[0] for s in SPALTEN].index(name) + 1
        for r in range(first, last + 1):
            w.cell(row=r, column=ci).number_format = fmt

    # ---------------------------------------------------------- Dropdowns
    # inline=True  -> kommagetrennte Literalliste ("a,b,c"), von den meisten
    #                 Tabellenprogrammen zuverlässig übernommen; nur für kurze Listen ohne
    #                 Komma im Wert und unter 255 Zeichen.
    # inline=False -> Bereichsbezug auf das Blatt „Listen" (ohne führendes "=").
    def add_dv(spaltenname, quelle, inline=False):
        i = [s[0] for s in SPALTEN].index(spaltenname) + 1
        col = get_column_letter(i)
        f1 = f'"{quelle}"' if inline else quelle
        dv = DataValidation(type="list", formula1=f1, allow_blank=True,
                            showDropDown=False, showErrorMessage=False)
        w.add_data_validation(dv)
        dv.add(f"{col}{first}:{col}{last}")

    masse_inline = ",".join(str(x) for x in SL.MASSE_CM)
    add_dv("Künstler:in", bereich["Künstler:in"])
    add_dv("Druckverfahren", bereich["Druckverfahren"])
    add_dv("Papier", bereich["Papier"])
    for feld in ("Bildmaß H cm", "Bildmaß B cm", "Blattmaß H cm", "Blattmaß B cm"):
        add_dv(feld, masse_inline, inline=True)
    add_dv("Rahmen im Preis", "ja,nein", inline=True)
    add_dv("Verkäuflich", "ja,nein", inline=True)

    w.freeze_panes = f"B{head + 3}"
    w.auto_filter.ref = f"A{head}:{get_column_letter(len(SPALTEN))}{last}"
    schuetzen(w, autofilter=True)

    # ---------------------------------------------------------- Kontakte
    ko = wb.create_sheet("Kontakte")
    ko.sheet_view.showGridLines = False
    kh = [("Künstler:in", 24), ("E-Mail", 32), ("Telefon", 20), ("Instagram (Handle)", 26)]
    for i, (h, br) in enumerate(kh, start=1):
        c = ko.cell(row=1, column=i, value=h)
        c.font = Font(name=FONT, size=9, bold=True)
        c.fill = FILL_HEAD
        c.border = BORDER
        ko.column_dimensions[get_column_letter(i)].width = br
    ko.cell(row=1, column=6, value="Dein Name steht schon da. Bitte E-Mail und Telefon "
            "ergänzen. Instagram-Handle nur eintragen, wenn die Spalte bei dir leer ist "
            "(ohne @, z. B. dein.profil).").font = Font(name=FONT, size=9, italic=True, color="808080")
    for r, name in enumerate(T.NAMEN, start=2):
        info = T.INSTAGRAM.get(name) or {}
        werte = [name, None, None, info.get("handle") or None]
        for i, v in enumerate(werte, start=1):
            c = ko.cell(row=r, column=i, value=v)
            c.font = Font(name=FONT, size=10,
                          color="000000" if i == 1 else "0000FF")
            c.border = BORDER
            if i > 1:                # Name (Spalte A) bleibt gesperrt, Kontaktfelder frei
                frei(c)
    ko.freeze_panes = "A2"
    schuetzen(ko)

    # ---------------------------------------------------------- Preishilfe
    blatt_preishilfe(wb)

    # ---------------------------------------------------------- Referenzblatt Saal Digital
    sd = wb.create_sheet("Saal Digital")
    sd.sheet_view.showGridLines = False
    sd.column_dimensions["A"].width = 46
    sd.column_dimensions["B"].width = 16
    sd.column_dimensions["C"].width = 62

    def zeile(r, a, b="", c="", bold=False, fill=None, size=10):
        for col, val in (("A", a), ("B", b), ("C", c)):
            cell = text_zelle(sd[f"{col}{r}"], val)
            cell.font = Font(name=FONT, size=size, bold=bold)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if fill:
                cell.fill = fill

    r = 1
    zeile(r, "Saal Digital — Papiere, Bildträger und Druckverfahren", size=13, bold=True); r += 2
    zeile(r, "Nur zum Nachschlagen. Quelle: saal-digital.de, Stand September 2026. "
             "Vollständiger Katalog: saal-referenz/saal_digital_katalog.xlsx.", size=9); r += 2

    zeile(r, "Welches Saal-Produkt → welches Druckverfahren aufs Schild",
          bold=True, fill=FILL_HEAD); r += 1
    zeile(r, "Saal-Produkt", "aufs Schild", "Erläuterung", bold=True, fill=FILL_SUB); r += 1
    for prod, verf, erkl in SL.SAAL_VERFAHREN_HILFE:
        zeile(r, prod, verf, erkl); r += 1
    r += 1

    zeile(r, "Papier- und Trägerliste (so steht sie im Dropdown)", bold=True, fill=FILL_HEAD); r += 1
    zeile(r, "Papier / Bildträger", "Druckverfahren", "Saal-Kategorie", bold=True, fill=FILL_SUB); r += 1
    for pap, verf, kat in SL.SAAL_PAPIERE:
        zeile(r, pap, verf, kat); r += 1
    r += 1

    zeile(r, "Gängige Saal-Formate für Wandbilder (cm, „B × H“ — je nach Hoch-/Querformat tauschen)",
          bold=True, fill=FILL_HEAD); r += 1
    formate = SL.SAAL_FORMATE
    for k in range(0, len(formate), 3):
        zeile(r, "     ".join(formate[k:k + 3])); r += 1
    r += 1

    zeile(r, "Fotoabzüge — Oberflächen / Papierarten", bold=True, fill=FILL_HEAD); r += 1
    zeile(r, "Oberfläche", "Grammatur", "Material · FSC · Beschreibung", bold=True, fill=FILL_SUB); r += 1
    for ob, gr, mat, fsc, besch in SL.FOTO_OBERFLAECHEN:
        zeile(r, ob, f"{gr} g/m²", f"{mat} · FSC: {fsc} · {besch}"); r += 1
    r += 1

    zeile(r, "Fotoabzüge — verfügbare Formate (cm, „B × H“, gilt für alle Oberflächen)",
          bold=True, fill=FILL_HEAD); r += 1
    zeile(r, "Seitenverhältnis", "", "Formate", bold=True, fill=FILL_SUB); r += 1
    for sv, fmts in SL.FOTO_FORMATE:
        zeile(r, sv, "", fmts); r += 1
    r += 1

    zeile(r, "Weitere Druckmedien — Digitalvlies (ERFURT), NICHT Saal Digital",
          bold=True, fill=FILL_HEAD); r += 1
    zeile(r, "Medium", "aufs Schild", "Beschreibung", bold=True, fill=FILL_SUB); r += 1
    for name, verf, beschr in SL.VLIESTAPETEN:
        zeile(r, name, verf, beschr); r += 1
    r += 1

    zeile(r, "Weitere Druckmedien — Affichenpapier / Blueback, NICHT Saal Digital",
          bold=True, fill=FILL_HEAD); r += 1
    zeile(r, "Medium", "aufs Schild", "Beschreibung", bold=True, fill=FILL_SUB); r += 1
    for name, verf, beschr in SL.AFFICHENPAPIERE:
        zeile(r, name, verf, beschr); r += 1
    schuetzen(sd)

    # ---------------------------------------------------------- Glossar
    blatt_glossar(wb)

    # „Listen" ans Ende der Blattreihe (sichtbar, aber aus dem Weg).
    idx_listen = wb.sheetnames.index("Listen")
    wb.move_sheet("Listen", offset=len(wb.sheetnames) - 1 - idx_listen)

    # Blattstruktur sperren: niemand kann Blätter löschen, umbenennen oder
    # verschieben. „Listen" bleibt sichtbar (siehe oben).
    wb.security.lockStructure = True
    wb.security.workbookPassword = PASSWORT

    wb.save(OUT)
    print("geschrieben:", OUT)
    fehlt = T.fehlende_qr()
    if fehlt:
        print(f"Hinweis: {len(fehlt)} Teilnehmende ohne Instagram-Handle "
              f"({', '.join(fehlt)}) — in data/instagram.csv nachtragen.")


if __name__ == "__main__":
    build()
