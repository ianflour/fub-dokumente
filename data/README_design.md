# design.json — Erklärung

`data/design.json` legt fest, **wie die Wandschilder aussehen**: Format, Schriftgrößen,
Farben. Sie ist optional — ohne die Datei rechnet der Generator alles selbst.

```bash
./.venv/bin/python scripts/build_druckdaten.py --design data/design.json --schildformat a7
```

`--design` sagt, *welche Datei* gilt; `--schildformat` sagt, *welches Format* daraus
gebaut wird (a0 … a10, Standard a8).

## Regeln für die Datei

- Reiner Text im JSON-Format (in TextEdit: **Format > In reinen Text umwandeln**).
- Zahlen mit Punkt: `7.2`, nicht `7,2`. Farben und Wörter in `"Anführungszeichen"`.
- Kein Komma nach dem letzten Eintrag einer Klammer, keine Kommentare.
- **Schlüssel mit Unterstrich** (`_hinweis`, `_farben_hinweis` …) sind Kommentare und werden
  ignoriert.
- **Alles ist optional.** Fehlt ein Wert, nimmt der Generator den automatisch berechneten.
  Nicht benötigte Format-Blöcke dürfen gelöscht werden.
- Die Datei enthält **keine Schriftpfade**: Regular- und Bold-Schnitt (`.ttf`/`.otf`) einfach in
  `assets/schriften/` legen.

## Aufbau

```
{
  "farben":     { … },                 Farben aller PDFs
  "wandschild": {
    "anpassung": { … },                automatische Anpassung, gilt für alle Formate
    "a0": { … }, "a1": { … }, … "a10": { … }    ein Block je DIN-A-Format
  }
}
```

### farben

Hex-Werte (`"#RRGGBB"`), gelten für **alle** PDFs.

| Schlüssel | wofür |
|---|---|
| `text` | Name und Titel |
| `sekundaer` | technische Angaben, Werknummer auf dem Schild |
| `tertiaer` | Fußzeilen und Ähnliches |
| `linie` | Trennstrich |
| `signal` | „verkauft“/„reserviert“ (Werkliste, Statuspunkt) |

### wandschild → aX (ein Block je Format)

`a0` bis `a10` sind die DIN-A-Formate, **quer** gedruckt (lange Seite = Breite).

| Schlüssel | Einheit | Bedeutung |
|---|---|---|
| `breite_mm`, `hoehe_mm` | mm | Größe des Schilds. Ändert sich eines der beiden, skalieren Schrift, Rand und QR mit und das Raster wird neu berechnet |
| `rand_mm` | mm | Abstand des Inhalts zum Schildrand |
| `qr_mm` | mm | Kantenlänge des QR-Codes (unten rechts) |
| `spalten`, `zeilen` | Anzahl | Schilder nebeneinander/untereinander auf dem Bogen |
| `bogen` | — | `"A4"` (hoch), `"A4quer"`, `"A3"` … `"A3quer"`, `"auto"` oder `"karte"` — siehe unten |
| `schriftgroessen_pt` | pt | `nummer`, `kuenstlerin`, `titel`, `ortjahr`, `technik`, `hinweis` |

`bogen`:

| Wert | Wirkung |
|---|---|
| fehlt / `"auto"` | A4, hoch oder quer — was mehr Schilder ergibt. Passt kein Schild (samt Rand) auf A4, gilt `"karte"`; ein Schild pro Blatt auf A4 quer |
| `"A4"`, `"A3"`, `"A2"` … | genau dieser Bogen, hochkant; mit Zusatz `quer` (`"A3quer"`) im Querformat. Schnittmarken werden gedruckt |
| `"karte"` | **Blatt = Karte**: jede Seite der PDF ist genau so groß wie das Schild, eine Karte pro Seite, keine Schnittmarken |

Die mitgelieferten Werte (aus dem Code berechnet, so entstehen sie ohne jede Angabe):

| Format | Schild | Anordnung | Name / Titel | QR |
|---|---|---|---|---|
| `a0` | 1189 × 841 mm | 1 Karte pro Blatt (Blatt = Karte) | 136.6 / 160.7 pt | 145 mm |
| `a1` | 841 × 594 mm | 1 Karte pro Blatt (Blatt = Karte) | 96.6 / 113.6 pt | 102 mm |
| `a2` | 594 × 420 mm | 1 Karte pro Blatt (Blatt = Karte) | 68.2 / 80.3 pt | 72 mm |
| `a3` | 420 × 297 mm | 1 Karte pro Blatt (Blatt = Karte) | 48.2 / 56.8 pt | 51 mm |
| `a4` | 297 × 210 mm | 1 Karte pro Blatt (Blatt = Karte) | 34.1 / 40.1 pt | 36 mm |
| `a5` | 210 × 148 mm | 1 × 1 = 1 je A4 quer | 24.1 / 28.4 pt | 26 mm |
| `a6` | 140 × 100 mm | 2 × 2 = 4 je A4 quer | 15.0 / 18.0 pt | 17 mm |
| `a7` | 100 × 70 mm | 2 × 4 = 8 je A4 hoch | 11.0 / 13.0 pt | 12 mm |
| `a8` | 74 × 52 mm | 2 × 5 = 10 je A4 hoch | 8.5 / 10.0 pt | 9 mm |
| `a9` | 52 × 37 mm | 5 × 5 = 25 je A4 quer | 6.0 / 7.0 pt | 6 mm |
| `a10` | 37 × 26 mm | 5 × 10 = 50 je A4 hoch | 4.2 / 5.0 pt | 4 mm |

`a7` und `a6` sind **bewusst kleiner als das DIN-Maß** (100 × 70 und 140 × 100 mm),
damit 8 bzw. 4 Schilder samt Schnittmarken auf einen A4-Bogen passen. Alle anderen Formate
sind exakt DIN. Ein exaktes A7-Schild (105 × 74 mm) gibt es mit
`"a7": { "breite_mm": 105, "hoehe_mm": 74 }` — dann passen 4 auf einen A4-Bogen (quer, 2 × 2).

### wandschild → anpassung

| Schlüssel | Standard | Bedeutung |
|---|---|---|
| `aktiv` | `true` | `false` schaltet die Anpassung komplett ab: Schriften bleiben, wie eingetragen, lange Texte werden nur mit „…“ gekürzt |
| `text_min_faktor` | `0.7` | eine zu lange Zeile darf bis auf diesen Anteil ihrer Schriftgröße schrumpfen |
| `inhalt_min_faktor` | `0.4` | so weit darf das ganze Schild höchstens verkleinert werden |
| `bogenrand_mm` | `8` | Rand rund um den Bogen für Schnittmarken beim automatischen Anordnen |

## Wie sich der Inhalt anpasst

1. **Mit der Schildgröße.** Ohne eigene Angabe werden Schriftgrößen, Rand und QR-Code
   linear vom Bezugsformat A8 (74 × 52 mm) abgeleitet. Ein A3-Schild hat also
   proportional größere Schrift als ein A8-Schild. Gleiches gilt, wenn `breite_mm`
   oder `hoehe_mm` geändert werden. Eigene `schriftgroessen_pt` gewinnen immer.
2. **Mit dem Schild.** Die Angaben stehen in einem festen Raster (Nummer, Name, zwei
   Zeilen Titel, Ort, Strich, zwei Zeilen Technik, Maß/Auflage, Hinweis). Passt das
   Raster nicht in die Schildhöhe — etwa weil die Schrift zu groß eingetragen ist —,
   wird **alles im gleichen Verhältnis** verkleinert, höchstens bis `inhalt_min_faktor`.
3. **Mit dem Text.** Ist eine einzelne Zeile zu lang, schrumpft nur ihre Schrift, bis
   `text_min_faktor`. Reicht das nicht, wird sie mit „…“ gekürzt.

Bei jedem Lauf meldet der Generator unter dem Format, was passiert ist, z. B.:

```
01_Wandschilder.pdf        14 Schilder auf 2 Bogen (A8), QR auf 14
    74 × 52 mm, 2 × 5 = 10 je Bogen A4 hoch
    Inhalt auf 47 % verkleinert, damit das Raster ins Schild passt
```

`ACHTUNG: kleinste Schrift nur … pt` erscheint, wenn die kleinste Zeile unter 5 pt fällt
(typisch ab A9) — dann nur bei sehr guter Druckauflösung lesbar.

## Beispiele

**Nur die Schrift auf A7 leicht vergrößern** (alles andere bleibt automatisch; wird es zu groß, verkleinert die Anpassung von selbst):

```json
{
  "wandschild": {
    "a7": { "schriftgroessen_pt": { "kuenstlerin": 12, "titel": 14 } }
  }
}
```

**Eigenes Format, 200 × 60 mm** (Schrift, Rand und QR skalieren mit, das Raster
ergibt sich von selbst):

```json
{ "wandschild": { "a5": { "breite_mm": 200, "hoehe_mm": 60 } } }
```

**Ein A2-Schild pro Blatt** für ein sehr großes Wandbild:

```bash
./.venv/bin/python scripts/build_druckdaten.py --einzelkarte 070 --schildformat a2
```

**A8-Schilder auf A3-Bögen statt A4** (21 statt 10 pro Bogen, der Rand für Schnittmarken bleibt erhalten):

```json
{ "wandschild": { "a8": { "bogen": "A3" } } }
```

**Anpassung abschalten:**

```json
{ "wandschild": { "anpassung": { "aktiv": false } } }
```

## Wenn etwas nicht klappt

| Meldung | Ursache |
|---|---|
| `Designdatei … nicht gefunden` | Pfad hinter `--design` stimmt nicht |
| JSON-Fehler beim Start | Komma vergessen oder zu viel, Anführungszeichen fehlen, Komma statt Punkt in Zahlen |
| `… ist kein DIN-A-Format … ignoriert` | Blockname nicht `a0` … `a10` |
| `… ist kein DIN-A-Bogen` | `bogen` ist keiner der Werte `A0` … `A10`, mit/ohne `quer`, `auto`, `karte` |
| `ACHTUNG: Raster ist größer als der Bogen` | `spalten`/`zeilen` passen nicht zu Schild- und Bogengröße |
