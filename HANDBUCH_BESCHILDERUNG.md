# Beschilderung & Werkliste — Handbuch

**FOTOTAGE MUSTERSTADT 2026** · 12.–19. September 2026 · Kunstraum Musterstadt
KULTURVEREIN MUSTERSTADT in Kooperation mit Partnerverein Musterstadt

Dieses Handbuch beschreibt Konzept A: **reduziertes Wandschild am Bild + vollständige Werkliste zum Auslegen.** Es gilt für 70–100 Werke von rund 27 Teilnehmenden (Liste: `data/teilnehmende.txt`).

---

## 1. Das Prinzip in drei Sätzen

Am Bild hängt nur, was Besucher:innen beim Schauen brauchen: Name, Titel, Ort/Jahr, Technik, Maß, Auflage. Alles Kaufmännische — Preis, Rahmung, Verfügbarkeit — steht in der ausliegenden Werkliste. Beide entstehen automatisch aus **einer** Tabelle, damit sich nichts widerspricht.

Kunstgegenstände sind nach § 9 Abs. 7 PAngV von der Preisauszeichnungspflicht ausgenommen. Die ausliegende Werkliste ist rechtlich vollkommen ausreichend.

---

## 2. Die Dateien

| Datei | Rolle | Wer arbeitet damit |
|---|---|---|
| `output/FUB2026_Werkdaten_MASTER.xlsx` | **Einzige Datenquelle.** Alle Werkdaten. | Kernteam |
| `output/FUB2026_Werkmeldung_VORLAGE.xlsx` | **gemeinsames** Formular auf Google Drive (Blätter: Werke · Kontakte · Preishilfe · Saal Digital · Glossar) | geht als Link raus |
| `data/teilnehmende.txt` | Namensliste aller Ausstellenden — speist das „Künstler:in“-Dropdown | Kernteam |
| `data/instagram.csv` | Name → Instagram-Handle; Quelle für die QR-Codes | Kernteam |
| `output/qr-codes/*.svg` + `output/qr-codes/png/` | fertige QR-Codes je Person (SVG für Design, PNG für Druck) | Design / Kernteam |
| `scripts/import_werkmeldungen.py` | liest die heruntergeladene Werkmeldung in den Master | Kernteam |
| `scripts/migrate_werkmeldung.py` | überträgt eine schon ausgefüllte Werkmeldung in eine neu gebaute Vorlage (bei Änderungen an der Vorlage) | Kernteam |
| `scripts/build_druckdaten.py` | erzeugt alle Druck-PDFs aus dem Master | Kernteam |
| `scripts/build_qr_codes.py` | erzeugt fehlende QR-Codes aus `data/instagram.csv` | Kernteam |
| `data/design.json` | Schriften, Größen, Farben — füllt der/die Designer:in aus | Design |
| `output/HANDBUCH_DESIGNER.pdf` | Anleitung zum Mitgeben an die Gestaltung | geht raus |
| `scripts/build_master.py` / `scripts/build_werkmeldung.py` / `scripts/build_handbuch_designer.py` | erzeugen die Vorlagen und das Designer-Handbuch neu, falls sie kaputtgehen | Notfall |
| `scripts/_saal_listen.py` | Auswahllisten (Druckverfahren, Saal-Papiere/Träger, Maße) für die Dropdowns, aus `data/saal-referenz/` abgeleitet | Notfall |
| `scripts/_teilnehmende.py` | liest `data/teilnehmende.txt` + `data/instagram.csv` für die anderen Skripte | Notfall |
| `output/ausgabe/` | die fertigen PDFs | Druckerei / Copyshop |
| `output/ausgabe/Adobe/` | Datendateien für InDesign und Illustrator | Designer:in |

**Goldene Regel:** Ändert sich ein Preis, ein Titel, ein Status — nur im Master ändern und den Generator neu laufen lassen. Niemals im PDF nachbessern.

---

## 3. Einmalige Einrichtung

Python 3 und drei Pakete (am einfachsten über `./setup.sh`, das die Umgebung `.venv/` anlegt):

```bash
pip install openpyxl reportlab segno
```

Test, ob alles läuft:

```bash
./.venv/bin/python scripts/build_druckdaten.py
```

Wenn die vier PDFs in `output/ausgabe/` landen, ist alles bereit. Der Generator sucht automatisch nach Arial oder einer metrisch identischen Schrift; findet er keine, nimmt er Helvetica. Beides sieht praktisch gleich aus.

---

## 4. Ablauf

### Schritt 1 — Werkmeldung teilen *(Frist bis 22. August)*

`output/FUB2026_Werkmeldung_VORLAGE.xlsx` auf Google Drive hochladen (bzw. als Google Tabelle öffnen) und den Link mit **„Jeder mit Link – Bearbeiter“** an alle Teilnehmenden geben. Textbaustein für die Mail steht in Abschnitt 9.

Es ist **eine gemeinsame Datei**: Alle tragen im Blatt „Werke“ ihre Werke ein — eine Zeile pro Bild, links unter „Künstler:in“ den eigenen Namen aus dem Dropdown. Kontaktdaten einmalig im Blatt „Kontakte“ (Name steht schon da).

**Blattschutz.** Kopfzeilen, das Quellblatt „Listen“ (sichtbar, ganz hinten) und die Nachschlage-Blätter (Saal Digital, Glossar, die Formeln in Preishilfe) sind gesperrt — nur die Eingabezellen (blaue Schrift / gelbe Felder) bleiben editierbar. Das ist kein Sicherheitsmerkmal, sondern ein Schutz gegen versehentliches Löschen in der von 27 Personen gleichzeitig bearbeiteten Datei. Passwort zum Aufheben (nur falls im Kernteam wirklich mal eine Struktur geändert werden muss): `fub2026` (in `scripts/build_werkmeldung.py`, Konstante `PASSWORT`).

**Dropdowns.** Die Auswahllisten für Künstler:in, Druckverfahren und Papier ziehen ihre Werte aus dem Blatt „Listen“ — deshalb ist das Blatt **sichtbar** (nicht ausgeblendet): Google Sheets übernimmt Dropdowns, die auf ein verstecktes Blatt zeigen, beim Import nicht zuverlässig. Prüfe nach dem Öffnen in Google Tabellen kurz, ob in einer leeren „Werke“-Zeile die Dropdowns erscheinen (Künstler:in, Druckverfahren, Papier, Maße, Rahmen im Preis, Verkäuflich). Fehlt einer, in Google unter **Daten > Datenvalidierung** neu setzen — die Werte stehen im Blatt „Listen“.

Öffne die Datei einmal selbst in Google Tabellen und prüfe unter **Daten > Blätter und Bereiche schützen**, ob die Sperren korrekt übernommen wurden — Google Sheets übersetzt den Excel-Blattschutz beim Import, aber nicht immer 1:1. Passe dort bei Bedarf nach (z. B. „Nur du kannst diesen Bereich bearbeiten“ für die gesperrten Bereiche, alle anderen Teilnehmenden bleiben Bearbeiter der restlichen Tabelle).

Vor dem Teilen prüfen, dass alle QR-Codes da sind:

```bash
./.venv/bin/python scripts/build_qr_codes.py            # meldet fehlende Instagram-Links
```

Setze eine **harte Frist drei Wochen vor Aufbau**. Erfahrungsgemäß liefert ein Drittel pünktlich, ein Drittel nach der ersten Erinnerung, ein Drittel erst, wenn du anrufst. Plane die Erinnerung als Kalendereintrag ein, nicht als guten Vorsatz.

**Wenn die Vorlage nach dem Teilen noch geändert werden muss** (neue Dropdown-Werte, ein behobener Fehler …) und in der geteilten Datei schon Werke drinstehen: nicht die Spalten von Hand nachziehen. Stattdessen die ausgefüllte Datei von Google Drive laden (`Datei > Herunterladen > Microsoft Excel`), in den Ordner **`data/werkmeldung_alt/`** legen, dann

```bash
./.venv/bin/python scripts/build_werkmeldung.py            # frische Vorlage
./.venv/bin/python scripts/migrate_werkmeldung.py --probe  # erst zeigen
./.venv/bin/python scripts/migrate_werkmeldung.py          # dann schreiben
```

Das schreibt `output/FUB2026_Werkmeldung_UEBERTRAGEN.xlsx`: alle ausgefüllten Zeilen aus „Werke“ und die E-Mail/Telefon-Angaben aus „Kontakte“, Spalten **nach Name** zugeordnet (umbenannte oder neue Spalten stören nicht). Dropdowns, Formeln und Blattschutz kommen aus der frischen Vorlage. Kurz gegenprüfen, dann auf Google Drive hochladen und die alte Datei ersetzen. **Wichtig:** Während des Umstiegs darf niemand mehr in der alten Datei tippen.

### Schritt 2 — Werkmeldung einlesen

Zum Stichtag die Tabelle als **Excel (.xlsx)** herunterladen (`Datei > Herunterladen > Microsoft Excel`) und in den Ordner `data/werkmeldungen/` legen. Dann:

```bash
./.venv/bin/python scripts/import_werkmeldungen.py --probe   # erst gucken
./.venv/bin/python scripts/import_werkmeldungen.py           # dann schreiben
```

Das Skript vergibt die Werknummern automatisch — **alphabetisch nach Künstler:in (A–Z)**, nicht in Meldereihenfolge —, legt neue Teilnehmende automatisch an (bzw. ergänzt leere Kontaktfelder), überspringt bereits importierte Werke und legt vor jedem Schreiben eine Sicherung des Masters an. Du kannst es also gefahrlos mehrfach laufen lassen, wenn Nachzügler eintrudeln — jeweils den aktuellen Export nachlegen. Dabei werden bei **jedem** Lauf alle echten Werke neu durchnummeriert (nicht nur die gerade importierten), damit die Nummer immer zur aktuellen Namensliste passt. Das heißt auch: Solange noch Meldungen fehlen, ist die Nummer eines Werks nicht endgültig — erst der letzte Import vor dem Druck vergibt die finalen Nummern. Wandschilder, die vorher schon gedruckt wurden, können danach eine falsche Nummer tragen.

Sobald hier das erste echte Werk steht, druckt der Generator die 14 Beispielwerke ganz oben im Master **automatisch nicht mehr mit** (Spalte „Beispiel“ = ja). Von Hand löschen musst du sie nicht — kannst es aber jederzeit, das ändert am Ergebnis nichts.

In `Rahmenmaß` und `Rahmenmodell` wird ein zwischen zwei Zahlen getipptes `x` bzw. `X` beim Einlesen automatisch zum echten Malzeichen `×` (also `40x50` → `40×50`). Der Generator macht dasselbe nochmal beim Drucken, falls du im Master von Hand ein `x` eintippst. Ein `x` in Wörtern bleibt unberührt.

### Schritt 3 — Hängung festlegen

Die Werknummer hat mit der Hängung nichts mehr zu tun — sie kommt automatisch aus Schritt 2 (alphabetisch nach Künstler:in) und bleibt stabil, egal wie und wann gehängt wird. Trage im Master trotzdem `Wand` (A, B, C …) und `Position` (1, 2, 3 … von links) ein, sobald der Hängeplan steht: Diese beiden Felder steuern nur, in welcher Reihenfolge `--gruppierung Wand` die Werkliste abschnittet, und stehen auf dem Rückseitenetikett.

Die Werkliste ist standardmäßig eine **fortlaufende Liste nach Werknummer** (also alphabetisch nach Künstler:in). Wer die Wände als Abschnitte in der Liste haben will, erzeugt sie mit `--gruppierung Wand` — das geht jederzeit, auch wenn `Wand`/`Position` erst nach dem letzten Import eingetragen werden. Anders als früher muss dafür nichts neu nummeriert werden.

### Schritt 4 — Daten prüfen

```bash
./.venv/bin/python scripts/build_druckdaten.py
```

Der Generator prüft vor dem Druck und bricht ab, wenn etwas fehlt:

- fehlende Pflichtfelder pro Werk
- Werke mit Status „verfügbar“, aber ohne Preis
- „Rahmen im Preis“ weder ja noch nein
- doppelt vergebene Werknummern

Fehler ausbessern, erneut laufen lassen. Mit `--trotz-fehler` druckt er trotzdem — nur sinnvoll für einen Zwischenstand, nie für den Enddruck.

### Schritt 5 — Drucken

```bash
./.venv/bin/python scripts/build_druckdaten.py
```

Erzeugt:

| PDF | Inhalt | Druck |
|---|---|---|
| `01_Wandschilder.pdf` | 8 Schilder à 100 × 70 mm pro A4-Bogen, mit Schnittmarken und Instagram-QR-Code | 300 g Karton, matt, weiß |
| `02_Werkliste.pdf` | zweispaltige Werk- und Preisliste | 120 g, A4, beidseitig |
| `03_Rueckseitenetiketten.pdf` | 12 Etiketten à 90 × 45 mm pro Bogen | Haftpapier oder 120 g + Klebestift |
| `04_Versicherungsliste.pdf` | interne Tabelle mit Werten und Summe | 1× für die Mappe |
| `05_Instagram_Wand.pdf` | Aushang mit den Instagram-QR-Codes aller Teilnehmenden | A4, nach Belieben |
| `Werkdaten_Serienbrief.csv` | Fallback für Word-Serienbrief | — |

Der QR-Code kommt aus dem Master-Blatt „Teilnehmende“ (Spalte Instagram), ersatzweise aus `data/instagram.csv`. Fehlt der Handle, bleibt das Schild einfach ohne Code.

Nützliche Schalter:

```bash
--schildformat a6          # große Schilder 140 × 100 mm, 4 pro A4 quer, größere Schrift
--gruppierung Wand         # Werkliste nach Wänden gruppieren (Standard: fortlaufend nach Nr)
--gruppierung Künstler:in  # Werkliste nach Person gruppieren
--spalten 1                # einspaltige Werkliste, größere Schrift, mehr Seiten
--ohne-preishinweis        # Zeile „Preis siehe Werkliste“ weglassen
--ohne-qr                  # Wandschilder ohne Instagram-QR-Code
--mit-beispielen           # die 14 Beispielwerke immer mitdrucken (Layout-Test)
--statuspunkte             # verkauft/reserviert als farbigen Punkt mitdrucken
--nur werkliste            # nur ein einzelnes PDF neu erzeugen (auch: --nur instagram)
--design data/design.json       # Typografie der Designerin/des Designers übernehmen
--adobe                    # zusätzlich Datendateien für InDesign und Illustrator
```

---

## 5a. Zusammenarbeit mit der Gestaltung

Konzept A funktioniert auch ohne Designer:in — die mitgelieferten PDFs sind druckfertig. Wenn
aber jemand die Schilder gestaltet, gibt es zwei Wege, und nur einer davon skaliert sauber auf
100 Stück.

**Empfohlen: die Gestaltung fließt in den Generator zurück.**
Der oder die Designer:in gestaltet ein einzelnes Schild, trägt die fertigen Werte in
`data/design.json` ein und legt Regular- und Bold-Schnitt in `schriften/` ab — der Generator
erkennt sie automatisch am Dateinamen, nichts in der JSON einzutragen. Danach:

```bash
./.venv/bin/python scripts/build_druckdaten.py --design data/design.json
```

Alle Schilder und die Werkliste erscheinen ab sofort in der gestalteten Typografie. Der
entscheidende Vorteil: Der Generator misst jeden Text und bricht ihn um. Kein Titel läuft aus
dem Schild — auch nicht bei dem Werk mit dem längsten Papiernamen.

**Alternative: die Gestaltung produziert selbst.**
Dann braucht sie die Dateien aus `output/ausgabe/Adobe/`:

| Datei | Für | Format |
|---|---|---|
| `InDesign_Datenzusammenfuehrung.txt` | InDesign-Datenzusammenführung | tabgetrennt, UTF-16 |
| `Illustrator_Variablen.csv` | Illustrator-Variablenbedienfeld | kommagetrennt, UTF-8 |
| `Illustrator_Variablenbibliothek.xml` | Notfallweg für Illustrator | XML |

Die Spaltennamen sind bewusst ohne Umlaute und Leerzeichen geschrieben (`Kuenstlerin`,
`OrtJahr`, `Masse`), weil Illustrator Variablennamen zeichengenau abgleicht.

Zwei Punkte, die in der Praxis schiefgehen: Diese Dateien dürfen **nicht in Excel geöffnet und
gespeichert** werden, das zerstört die Umlautkodierung. Und Illustrator legt bei der
Datenzusammenführung pro Datensatz eine eigene Zeichenfläche an — bei 100 Werken also 100
Zeichenflächen, aus denen der Druckbogen mit 8 Schildern erst noch von Hand entstehen muss.
InDesign kann das in einem Schritt.

`output/HANDBUCH_DESIGNER.pdf` erklärt all das auf neun Seiten ohne Technikvokabular. Das ist die
Datei, die du weitergibst.

### Schritt 6 — Schneiden und montieren

Siehe Abschnitt 6.

### Schritt 7 — Während der Ausstellung

Verkauft sich etwas: Status im Master auf `verkauft` setzen, Käufer:in eintragen, dann

```bash
./.venv/bin/python scripts/build_druckdaten.py --nur werkliste
```

und die Werkliste nachdrucken. Am Wandschild wird **nichts** geändert — dort klebt ein roter Punkt.

---

## 5. Was steht wo

| Angabe | Wandschild | Werkliste | Rückseite | intern |
|---|:--:|:--:|:--:|:--:|
| Instagram-QR-Code | ● | — | — | — |
| Werknummer | ● | ● | ● | ● |
| Künstler:in | ● | ● | ● | ● |
| Titel | ● | ● | ● | ● |
| Ort, Land, Jahr | ● | ● | ● | ● |
| Druckverfahren, Papier | ● | ● | ● | ● |
| Bildmaß | ● | ● | ● | ● |
| Blattmaß | — | ● | — | ● |
| Auflage | ● | ● | — | ● |
| Exemplar (2/5) | — | — | ● | ● |
| Preis | — | ● | — | ● |
| Rahmen im Preis | — | ● | — | ● |
| Rahmenmodell, Passepartout | — | — | ● | ● |
| Status | — | ● | — | ● |
| Versicherungswert | — | — | ● | ● |
| Käufer:in | — | — | — | ● |

Warum Rahmenmodell und Passepartout **nicht** aufs Wandschild kommen: Das ist Werkstattinfo. Besucher:innen sehen den Rahmen ja. Wer ihn genau wissen will, fragt — und dann steht es in der Werkliste bzw. auf der Rückseite.

Warum „Rahmen im Preis: ja/nein“ so wichtig ist: Es ist die meistgestellte Frage im Verkaufsgespräch und die häufigste Quelle für Ärger nach dem Kauf.

---

## 6. Material, Schnitt und Montage

**Karton.** 300 g/m², matt gestrichen, weiß oder gebrochen weiß. Kein glänzendes Papier — Ausstellungsstrahler erzeugen darauf Reflexe, die das Schild unlesbar machen. Empfohlene Beleuchtung im Raum: 100–300 Lux, blendfrei.

**Menge.** 8 Schilder pro A4-Bogen. Bei 100 Werken also 13 Bogen, plus 15 % Reserve für Tippfehler und Nachzügler = **15 Bogen**. Das Blatt „Übersicht“ im Master rechnet das automatisch aus.

**Schneiden.** Schnittmarken liegen außen an. Mit Stapelschneider oder Cutter und Stahllineal an den Marken durchschneiden. Nicht mit der Schere — die Kanten sieht man aus zwei Metern.

**Montage.** Doppelseitige Klebepads (z. B. Tesa Powerstrips Small oder Montagepads), zwei pro Schild, an den oberen Ecken. **Kein einfaches Tesa-Röllchen:** Erfahrungsberichte aus der Praxis sind eindeutig — nach drei Tagen hängen die ersten Schilder schief und zwei sind heruntergefallen. Bei einer achttägigen Ausstellung ist das garantiert.

**Position.** Rechts vom Werk, 5–8 cm Abstand zum Rahmen, Schildmitte auf **140–145 cm** über dem Boden. Nicht tiefer als 100 cm anbringen, sonst ist es für einen Teil des Publikums unlesbar. Und: im ganzen Raum konsequent gleich — links oder rechts, aber nicht gemischt. Ein Wechsel des Prinzips innerhalb eines Wandabschnitts zerstört das geschlossene Bild.

**Schriftgrößen.** Auf dem A7-Schild: Name 11 pt fett, Titel 13 pt kursiv, Ort/Jahr 9,5 pt, Technik 8,3 pt. Das liegt unter dem, was die Smithsonian Guidelines for Accessible Exhibition Design als Minimum nennen (16–20 pt) — auf 100 × 70 mm passt mehr schlicht nicht. Wenn Barrierefreiheit Priorität hat, nimm `--schildformat a6`: 140 × 100 mm mit Name 15 pt und Titel 18 pt. Das kostet 25 statt 13 Bogen und wirkt neben kleinen Arbeiten etwas groß, ist aber deutlich besser lesbar. Entscheide das einmal für die ganze Ausstellung, nicht pro Wand.

**Werkliste.** Mindestens drei Exemplare: Theke, Eingang, ein Wandabschnitt in der Mitte. Kalkuliere Schwund ein — Werklisten verschwinden. Bei 100 Werken werden es rund 5 Seiten zweispaltig; getackert oder in einer einfachen Klemmmappe.

---

## 7. Der Aufbautag

- [ ] Werkliste-PDF **erst nach der endgültigen Hängung** final drucken
- [ ] Schilder vorsortiert nach Wand mitbringen (Briefumschlag pro Wand beschriften)
- [ ] Ersatzkarton, Cutter, Lineal, Klebepads, Wasserwaage im Werkzeugkoffer
- [ ] Zollstock: Schildmitte 140–145 cm, Abstand 5–8 cm — einmal ausmessen, dann Pappschablone bauen
- [ ] Rote Klebepunkte (Ø 8 mm) für Verkäufe bereitlegen
- [ ] Rückseitenetiketten schon **vor** dem Aufhängen auf die Rahmenrückseiten kleben
- [ ] Versicherungsliste ausgedruckt in die Ausstellungsmappe, Summe an Partnerverein Musterstadt melden
- [ ] Ein Laptop mit Master + Generator vor Ort, für den unvermeidlichen Nachdruck

---

## 8. Häufige Fehler und wie man sie vermeidet

**Nummern vor der Hängung vergeben.** Führt garantiert zu einer Werkliste, die nicht zum Raum passt. Erst Wand und Position festlegen.

**Zwei Wahrheiten.** Sobald ein Preis in einer Mail, einem Instagram-Post und der Werkliste unterschiedlich steht, verliert ihr Vertrauen. Nur der Master zählt.

**Titel in Anführungszeichen.** Auf dem Schild steht der Titel kursiv, ohne Zusatzzeichen. Kein „…“, kein »…«.

**Maße als Text.** In der Tabelle gehören Zahlen in die Maßspalten, nicht „30 cm“. Sonst rechnet nichts mehr.

**Vergessene Reserve.** Am Aufbautag fällt immer auf, dass jemand einen Titel geändert hat. Ohne Ersatzkarton und funktionierenden Drucker vor Ort wird das ein Problem.

**Preis ohne Rahmenangabe.** Siehe oben. Immer „inkl. Rahmen“ oder „zzgl. Rahmen“.

---

## 9. Textbaustein für die Werkmeldung

> **Betreff: FOTOTAGE MUSTERSTADT 2026 — Werkdaten bis 22. August**
>
> Hallo [Name],
>
> für die Wandschilder und die ausliegende Werkliste brauchen wir von dir noch die Angaben zu deinen ausgestellten Arbeiten. Alles läuft über **eine gemeinsame Tabelle**, die alle zusammen ausfüllen:
>
> **[LINK zur Google-Tabelle]**
>
> Im Blatt „Werke“ pro ausgestelltem Bild eine Zeile — auch bei Serien jedes Bild einzeln. Wähle links unter „Künstler:in“ deinen Namen aus dem Dropdown (immer dieselbe Schreibweise). Und trag im Blatt „Kontakte“ einmalig deine E-Mail und Telefonnummer ein, dein Name steht dort schon. Dauert etwa zehn Minuten.
>
> Für Druckverfahren, Papier und die Maße gibt es Auswahllisten — überwiegend die Papiere von Saal Digital, dazu die ERFURT-Digitalvlies-Medien und Affichenpapier/Blueback für großformatige Wandbilder bzw. Plakate. Druckst du auf etwas anderem, wähle „anderes Papier …“ und trag es von Hand ein. Das Blatt „Saal Digital“ erklärt, welches Papier / welcher Bildträger zu welchem Druckverfahren gehört; das Blatt „Glossar“ erklärt alle Begriffe. Unsicher beim Preis? Das Blatt „Preishilfe“ rechnet dir einen Vorschlag aus.
>
> Zwei Dinge, bei denen wir oft nachfragen müssen:
>
> **Maße.** Wir brauchen beides: Bildmaß (das sichtbare Bild, ohne Weißrand) und Blattmaß (das Papier inklusive Weißrand). Immer Höhe × Breite, immer nur die Zahl.
>
> **Preis.** Bitte den Endpreis, den Besucher:innen zahlen — nicht deinen Nettoanteil. Und gib bitte an, ob der Rahmen im Preis enthalten ist. Das ist die häufigste Frage vor Ort.
>
> Auf dem Wandschild erscheinen später nur: dein Name, Titel, Ort/Land/Jahr, Druckverfahren und Papier, Bildmaß und Auflage — plus ein kleiner QR-Code zu deinem Instagram-Profil. Preis und Rahmung stehen ausschließlich in der Werkliste, die an der Theke ausliegt.
>
> Bitte bis **22. August** eintragen. Danach gehen die Schilder in den Druck.
>
> Danke dir!
>
> Team
> KULTURVEREIN MUSTERSTADT | FOTOTAGE MUSTERSTADT 2026
> fototage-musterstadt.de · instagram.com/fototage.musterstadt

---

## 10. Fallback ohne Python

Falls das Skript partout nicht laufen will: `output/ausgabe/Werkdaten_Serienbrief.csv` ist ein Semikolon-CSV mit UTF-8-BOM, das Word und Excel direkt öffnen. Damit funktioniert ein klassischer Word-Serienbrief:

1. Word → *Sendungen* → *Etiketten* → benutzerdefiniertes Etikett 100 × 70 mm, 2 Spalten × 4 Zeilen
2. *Empfänger auswählen* → *Vorhandene Liste verwenden* → die CSV
3. Seriendruckfelder einsetzen: `Kuenstlerin`, `Titel`, `OrtLandJahr`, `Technik`, `Masse`, `Auflage`
4. *Fertig stellen und zusammenführen* → *Einzelne Dokumente bearbeiten*

Für die Werkliste eignet sich derselbe Datensatz als Serienbrief vom Typ *Verzeichnis*.

---

## 11. Zeitplan

| Wann | Was |
|---|---|
| jetzt | Werkmeldung auf Google Drive teilen (Link an alle Teilnehmenden) |
| jetzt | `scripts/build_qr_codes.py` laufen lassen, fehlende Instagram-Links einsammeln |
| jetzt | `output/HANDBUCH_DESIGNER.pdf` an die Gestaltung geben |
| 22. August | Frist Werkmeldung, danach Tabelle als Excel herunterladen |
| 24. August | Erinnerung an Nachzügler, dann anrufen |
| ~1. September | Import abgeschlossen, Daten geprüft |
| ~5. September | Hängeplan steht → Wand und Position eintragen |
| ~8. September | Schilder und Rückseitenetiketten drucken und schneiden |
| Aufbautag | Hängung, dann Werkliste final drucken |
| 12. September | Vernissage |
| laufend | Verkäufe im Master pflegen, Werkliste nachdrucken |
| 19. September | Finissage |

---

*Team / Kulturverein Musterstadt | FOTOTAGE MUSTERSTADT 2026*
*fototage-musterstadt.de · kontakt@beispiel-verein.de · instagram.com/fototage.musterstadt*
