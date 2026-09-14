# -*- coding: utf-8 -*-
"""
Gemeinsame Auswahllisten für die Dropdowns in
  build_werkmeldung.py  (FUB2026_Werkmeldung_VORLAGE.xlsx)
  build_master.py       (FUB2026_Werkdaten_MASTER.xlsx)

Damit beide Dateien exakt dieselben Optionen anbieten. Nur hier ändern.

Quelle der Papier-, Verfahrens- und Formatangaben:
  data/saal-referenz/saal_digital_katalog.xlsx  (aus saal-digital.de aufbereitet,
  Bereiche „Wandbilder" und „Poster / FineArt", Stand September 2026).
  data/saal-referenz/fotoabzuege_infos.csv      (Fotoabzüge: Oberflächen + Formate).
Die Listen unten sind aus diesen Quellen abgeleitet — schildtaugliche
Kurzbegriffe statt der internen Saal-Bezeichnungen. Katalog neu aufbereiten:
  ./.venv/bin/python data/saal-referenz/saal_katalog_aufbereiten.py
"""

import re

# --- Malzeichen in Maßangaben ------------------------------------------
# „40x50", „40 x 50 cm", „40X50" -> „40×50" / „40 × 50 cm".
# Ersetzt nur ein x/X, das zwischen zwei Ziffern steht (mit optionalem
# Leerraum). Ein x in Wörtern („Box", „exakt") bleibt unangetastet.
# Wirkt auch auf Ketten wie „120x80x5".
_MASS_X = re.compile(r"(\d\s*)[xX](\s*\d)")


def mal_zeichen(text):
    """Gibt text mit echtem Malzeichen × zurück (None/'' bleiben unverändert)."""
    if text is None:
        return text
    return _MASS_X.sub(r"\1×\2", str(text))


# --- Ja/Nein -------------------------------------------------------------
JA_NEIN = ["ja", "nein"]

# --- Druckverfahren ----------------------------------------------------
# Kurze, schildtaugliche Begriffe. Auf dem Wandschild erscheint später
# „<Druckverfahren>, <Papier>". Die ersten fünf decken den Saal-Katalog ab
# (Verfahrensgruppen: Fotografisch, Pigment/FineArt, Direktdruck, Latexdruck),
# der Rest sind gängige Handabzugs- und Druckverfahren außerhalb von Saal.
DRUCKVERFAHREN = [
    "C-Print",                       # Fotobelichtung auf Fujifilm Crystal Archive
    "Pigmentdruck",                  # FineArt-Inkjet auf Hahnemühle
    "Digitaldruck",                  # Poster / Kunstdruck
    "Direktdruck",                   # auf Alu-Dibond, Acrylglas, Hartschaum
    "Latexdruck",                    # auf Leinwand und Digitalvlies
    "UV-Druck",                      # auf Digitalvlies / Trägerplatten
    "Solventdruck",                  # Öko-/Solvent, Plakat/Blueback, Folien
    "Inkjet",
    "Silbergelatine, Handabzug",
    # --- Cyanotypie (Blaudruck) und Varianten ---
    "Cyanotypie",                    # klassisch (Herschel-Formel)
    "Cyanotypie (New Cyanotype)",    # Ware-Formel, ein Bad
    "Cyanotypie, getont",            # nachträglich getont (Tannin/Tee/Kaffee)
    "Nass-Cyanotypie",               # wet cyanotype, nass belichtet
    # --- Kollodium-Nassplatte (wet plate) und Varianten ---
    "Kollodium-Nassplatte",          # Sammelbegriff
    "Ambrotypie",                    # Nassplatten-Positiv auf Glas
    "Ferrotypie (Tintype)",          # Nassplatten-Positiv auf Metall
    "Pannotypie",                    # Nassplatten-Positiv auf Textil/Leder/Wachstuch
    "Kollodium-Negativ",             # Nassplatten-Glasnegativ (für Kontaktabzüge)
    "Risografie",
    "Offsetdruck",
    "anderes",
]

# --- Papiere / Bildträger --------------------------------------------
# (Anzeigename,  schildtaugliches Druckverfahren,  Saal-Kategorie)
# Reihenfolge = so steht es im Dropdown.
SAAL_PAPIERE = [
    # Fotoabzug — Fujifilm Crystal Archive (fotografische Belichtung)
    ("Fujifilm Crystal Archive DP II Glänzend 250 g/m²",      "C-Print",      "Fotoabzug"),
    ("Fujifilm Crystal Archive DP II Matt 234 g/m²",          "C-Print",      "Fotoabzug"),
    ("Fujifilm Crystal Archive DP II Silk Portrait 232 g/m²", "C-Print",      "Fotoabzug"),
    ("SoftTouch matt 300 g/m² (Fotoabzug, laminiert)",        "C-Print",      "Fotoabzug"),
    # Fotoabzug — Digitaldruck auf Spezialpapier (kein fotografischer Abzug)
    ("Art Line 250 g/m² (Digitaldruck, Recyclingpapier FSC)", "Digitaldruck", "Fotoabzug"),
    ("Kunstdruck / Art Print 250 g/m² (Digitaldruck)",        "Digitaldruck", "Fotoabzug"),
    # FineArt — Hahnemühle (Pigment-/FineArt-Inkjet)
    ("Hahnemühle Photo Rag 308 g/m²",                    "Pigmentdruck", "FineArt"),
    ("Hahnemühle FineArt Baryta 325 g/m²",               "Pigmentdruck", "FineArt"),
    ("Hahnemühle Museum Etching 350 g/m²",               "Pigmentdruck", "FineArt"),
    ("Hahnemühle Bamboo Natural Line 290 g/m²",          "Pigmentdruck", "FineArt"),
    ("Hahnemühle Hemp Natural Line 290 g/m²",            "Pigmentdruck", "FineArt"),
    # Wandbild — Direktdruck auf Trägerplatte
    ("Alu-Dibond (Direktdruck)",                         "Direktdruck",  "Wandbild"),
    ("Alu-Dibond gebürstet, Gold",                       "Direktdruck",  "Wandbild"),
    ("Alu-Dibond gebürstet, Silber",                     "Direktdruck",  "Wandbild"),
    ("Acrylglas, hinterklebt",                           "Direktdruck",  "Wandbild"),
    ("Hartschaumplatte",                                 "Direktdruck",  "Wandbild"),
    ("GalleryPrint (Fotoabzug unter Acrylglas)",         "C-Print",      "Wandbild"),
    # Leinwand
    ("Fotoleinwand",                                     "Latexdruck",   "Leinwand"),
    ("Baumwoll-Leinwand",                                "Latexdruck",   "Leinwand"),
]

# --- Weitere Druckmedien: Digitalvlies (ERFURT) ----------------------
# Digital bedruckbare Vliestapete für groß­formatige Wandbilder (Latex-/UV-,
# teils Öko-Solvent-Druck). Kein Saal-Digital-Produkt. Recherche: erfurt.com
# und erfurtspezialpapiere.com, Stand September 2026.
# (Anzeigename,  schildtaugliches Druckverfahren,  Kurzbeschreibung)
VLIESTAPETEN = [
    ("ERFURT Digitalvlies DV150 PRO 150 g/m² (glatt)", "Latexdruck",
     "PVC-freies Digitalvlies, glatt, unbeschichtet. Latex-/UV-Druck."),
    ("ERFURT Digitalvlies DV150 ECO PRO 150 g/m² (glatt, Recycling)", "Latexdruck",
     "wie DV150, aus 100 % Recyclingfasern. Latex-/UV-Druck."),
    ("ERFURT Digitalvlies DV195 PRO 195 g/m² (glatt, beschichtet)", "Latexdruck",
     "glatt, spezialbeschichtet, diffusionsoffen. Öko-Solvent-/Latex-/UV-Druck."),
    ("ERFURT Digitalvlies DV615 PRO 160 g/m² (Struktur: putzartig)", "Latexdruck",
     "geprägte, putzartige Oberfläche. Latex-/UV-Druck."),
    ("ERFURT Digitalvlies DV617 PRO 160 g/m² (Struktur: leinenartig)", "Latexdruck",
     "geprägte, leinenartige Oberfläche. Latex-/UV-Druck."),
    ("ERFURT Digitalvlies DV635 PRO 160 g/m² (feine Faserstruktur)", "Latexdruck",
     "geprägte, feine Faserstruktur. Latex-/UV-Druck."),
    ("ERFURT Digitalvlies DV638 PRO 160 g/m² (feinkörnige Struktur)", "Latexdruck",
     "geprägte, feinkörnige Struktur. Latex-/UV-Druck."),
    ("ERFURT Variovlies T150 AIRLESS 150 g/m² (Malervlies)", "anderes",
     "eigentlich ein Malervlies für den Airless-Dispersionsanstrich, kein "
     "Digitaldruckmedium — nur wählen, wenn wirklich darauf gedruckt oder gemalt wurde."),
]

# --- Weitere Druckmedien: Affichenpapier (Blueback / Whiteback) -----
# Plakatpapier für großformatige Plakate; Offset-, Digital-, Öko-Solvent-,
# Latex- oder UV-Druck. Kein Saal-Digital-Produkt.
#   Blueback  = einseitig gestrichen, blaugrauer ungestrichener Rücken →
#               hohe Opazität (altes Plakat scheint nicht durch), guter
#               Kleisterhalt; holzfrei, nassfest, wetterfest ca. 4–6 Monate.
#               Standardgrammaturen 115 / 120 / 135 g/m², Bildseite matt
#               oder glänzend.
#   Whiteback = ungefärbter, lichtdurchlässiger Rücken → für hinterleuchtete
#               Leuchtkästen (City-Light-Plakate). Sonst wie Blueback.
# Recherche: typolexikon.de, de.wikipedia.org, igepa.de, posterprint-online.ch,
# Stand September 2026.
# (Anzeigename,  schildtaugliches Druckverfahren,  Kurzbeschreibung)
AFFICHENPAPIERE = [
    ("Affichenpapier / Blueback 115 g/m² (Blaurücken)", "Digitaldruck",
     "einseitig gestrichenes Plakatpapier, Bildseite matt oder glänzend, blaugrauer "
     "ungestrichener Rücken für hohe Opazität und guten Kleisterhalt. Offset-, "
     "Digital-, Öko-Solvent-, Latex- oder UV-Druck."),
    ("Affichenpapier / Blueback 120 g/m² (Blaurücken)", "Digitaldruck",
     "wie 115 g/m², etwas steifer und deckender. Offset-, Digital-, Öko-Solvent-, "
     "Latex- oder UV-Druck."),
    ("Affichenpapier / Blueback 135 g/m² (Blaurücken)", "Digitaldruck",
     "schwerere Blueback-Variante für längere Standzeiten und mehr Stabilität. "
     "Offset-, Digital-, Öko-Solvent-, Latex- oder UV-Druck."),
    ("Affichenpapier / Whiteback 115–120 g/m² (City-Light, hinterleuchtbar)", "Digitaldruck",
     "ungefärbter, lichtdurchlässiger Rücken — für hinterleuchtete Leuchtkästen "
     "(City-Light-Plakate). Sonst wie Blueback. Öko-Solvent-, Latex-, UV- oder Digitaldruck."),
]

PAPIERE = ([p[0] for p in SAAL_PAPIERE]
           + [v[0] for v in VLIESTAPETEN]
           + [a[0] for a in AFFICHENPAPIERE]
           + ["anderes Papier — bitte in die Zelle schreiben"])

# --- Maße (cm) ------------------------------------------------------
# Eingabehilfe für die vier Maßfelder. Kein Zwang: eigene Zahlen
# (auch mit Komma, z. B. 27,5) bleiben jederzeit erlaubt.
# Kantenlängen aus dem Saal-Katalog (bis 120 cm) plus die kleinen
# Fotoabzug-Kanten (16/17/19/22/25) aus fotoabzuege_infos.csv.
MASSE_CM = [9, 10, 13, 15, 16, 17, 18, 19, 20, 21, 22, 24, 25, 27, 30, 33, 40, 42, 45,
            50, 60, 70, 75, 80, 90, 100, 105, 110, 120]

# --- Gängige Saal-Formate (für das Referenzblatt) --------------------
# Kurzschreibweise „B × H"; je nach Hoch-/Querformat tauschen.
# Aus dem Katalog, nach Fläche sortiert, auf die für eine Ausstellung
# üblichen Größen eingekürzt.
SAAL_FORMATE = [
    "10 × 15", "13 × 18", "15 × 20", "15 × 21 (A5)", "18 × 24", "20 × 30",
    "21 × 30 (A4)", "24 × 30", "20 × 40", "30 × 30", "30 × 40", "30 × 42 (A3)",
    "30 × 45", "40 × 40", "30 × 60", "40 × 50", "40 × 60", "42 × 60 (A2)",
    "50 × 50", "45 × 60", "40 × 70", "40 × 80", "50 × 70", "60 × 60",
    "50 × 75", "60 × 80", "70 × 70", "60 × 84 (A1)", "50 × 100", "60 × 90",
    "70 × 100", "70 × 105", "75 × 100", "80 × 100", "70 × 120", "80 × 120",
]

# --- Welches Saal-Produkt -> welches Druckverfahren aufs Schild -------
SAAL_VERFAHREN_HILFE = [
    ("Fotoabzug — Glänzend / Matt / Silk / SoftTouch matt", "C-Print",
     "fotografische Belichtung auf Fujifilm Crystal Archive (SoftTouch zusätzlich laminiert)"),
    ("Fotoabzug — Art Line / Kunstdruck (Art Print)", "Digitaldruck",
     "Digitaldruck auf Spezialpapier, KEIN fotografischer Abzug"),
    ("FineArt Print (Hahnemühle-Papiere)", "Pigmentdruck",
     "pigmentierter Inkjetdruck, Hahnemühle Certified Studio"),
    ("Poster / Kunstdruck", "Digitaldruck",
     "hochwertiger Digitaldruck auf Posterpapier"),
    ("Wandbild auf Alu-Dibond / Acrylglas / Hartschaum", "Direktdruck",
     "UV-Direktdruck auf die Trägerplatte"),
    ("GalleryPrint (Foto unter Acrylglas)", "C-Print",
     "Fotoabzug, rückseitig auf Acrylglas kaschiert"),
    ("Leinwand (Foto-, Baumwoll-, Akustik-Leinwand)", "Latexdruck",
     "Latex-/Digitaldruck auf Leinwand"),
]

# --- Fotoabzüge: Oberflächen und Formate -----------------------------
# Quelle: saal-referenz/fotoabzuege_infos.csv
# (Oberfläche, Grammatur g/m², Material/Verfahren, FSC, Kurzbeschreibung)
FOTO_OBERFLAECHEN = [
    ("Glänzend", 250, "Fujifilm Crystal Archive DP II Professional", "nein",
     "Brillante Farben, hoher Kontrast, tiefe Schwarztöne. Silberhalogenid, sehr langlebig."),
    ("Matt", 234, "Fujifilm Crystal Archive DP II Professional", "nein",
     "Natürlich und elegant, minimale Reflexionen, kaum sichtbare Fingerabdrücke."),
    ("Silk (Silk Portrait)", 232, "Fujifilm Crystal Archive DP II Professional", "nein",
     "Feine Seidenrasterstruktur, weiche Zeichnung und natürliche Schatten — für Porträts."),
    ("Art Line", 250, "Digitaldruck auf Art Line Papier", "ja",
     "Leicht raue, samtige Textur. 100 % Sekundärfasern, FSC-zertifiziert."),
    ("Kunstdruck", 250, "Digitaldruck auf Art Print Papier", "nein",
     "Fein strukturierte Oberfläche, elegantes und künstlerisches Finish."),
    ("SoftTouch matt", 300, "Klassisches Fotopapier mit Laminierfolie", "ja",
     "Sehr glatte, sehr matte Oberfläche (Haptik wie Pfirsichhaut), hohe Farbintensität."),
]

# Fotoabzug-Formate in cm („B × H"), gruppiert nach Seitenverhältnis.
FOTO_FORMATE = [
    ("2:3", "9 × 13 · 10 × 15 · 13 × 18 · 13 × 19 · 15 × 21 · 15 × 22 · "
            "20 × 30 · 21 × 30 · 30 × 45"),
    ("3:4", "12 × 16 · 13 × 17 · 15 × 20 · 18 × 24 · 30 × 40"),
    ("3:5", "17 × 30"),
    ("4:5", "10 × 13 · 20 × 25 · 24 × 30"),
    ("Panorama", "10 × 20"),
    ("Quadratisch", "10 × 10 · 13 × 13 · 15 × 15 · 20 × 20 · 30 × 30"),
]
