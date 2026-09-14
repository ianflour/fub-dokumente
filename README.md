# FOTOTAGE MUSTERSTADT 2026 — Beschilderung, Werkliste & Orga

Werkzeugkasten, der aus **einer** Tabelle alle Druckdaten für die Ausstellung
erzeugt: Wandschilder, ausliegende Werkliste, Rückseitenetiketten,
Versicherungsliste, Instagram-Wand — plus Schichtplan, Gästeliste und die
Anleitungs-PDFs.

> **Über dieses Repo:** Veranstaltung, Ort, Personen, Werke und Gästeliste in
> `data/` sind frei erfundene Beispieldaten für die öffentliche Demo — keine
> echten Namen oder Kontaktdaten. `output/` (alles Erzeugte) ist nicht
> mitversioniert; wer die Werkzeuge ausprobieren will, baut es lokal aus den
> Beispieldaten neu (siehe Schnellstart).

> Vollständige Anleitung: **[HANDBUCH_BESCHILDERUNG.md](HANDBUCH_BESCHILDERUNG.md)**
> Umgebung einrichten: **[SETUP.md](SETUP.md)**

---

## Ordnerstruktur

| Ordner | Inhalt |
|---|---|
| `scripts/` | alle Werkzeuge. `fub.py` ist das Menü; `_pfade.py`, `_teilnehmende.py`, `_saal_listen.py` sind geteilte Hilfsmodule (nicht direkt aufrufen). |
| `data/` | alles von Hand Gepflegte: `teilnehmende.txt`, `instagram.csv`, `design.json`, die editierbaren Datenbasen `schichtplan_daten.py` / `gaesteliste_daten.py` / `workshop_daten.py`, die Ablageordner `werkmeldungen/` + `werkmeldung_alt/`, das Nachschlagewerk `saal-referenz/`. |
| `assets/` | `schriften/` — hier legt die Gestaltung Regular- und Bold-Schnitt ab (wird automatisch erkannt). |
| `output/` | **alles Erzeugte.** Master, Werkmeldung-Vorlage, alle PDFs, `ausgabe/`, `qr-codes/`, Backups, `upload/`. Jederzeit neu baubar — nichts hier von Hand ändern. |

Pfade sind komplett dynamisch aus dem Script-Verzeichnis abgeleitet — die
Werkzeuge laufen gleich, egal aus welchem Ordner gestartet.

> **Hinweis Master/Vorlage:** `output/FUB2026_Werkdaten_MASTER.xlsx` liegt in
> `output/`, weil `build_master.py` sie erzeugt. Sobald echte Werkdaten drin
> stehen (statt der 14 Beispielwerke), ist sie **nicht** mehr wegwerfbar —
> dann regelmäßig sichern. `import_werkmeldungen.py` legt vor jedem Schreiben
> automatisch eine Kopie an.

---

## Schnellstart

```bash
./setup.sh                          # einmalig: Python-Umgebung
./.venv/bin/python scripts/fub.py   # Menü
```

Menü:

```
  1  Alles neu bauen          Werkmeldung → QR-Codes → alle PDFs (Master NICHT neu)
  2  Druckdaten erzeugen      Wandschilder, Werkliste, Etiketten, …
  3  Werkmeldungen einlesen   aus data/werkmeldungen/ in den Master
  4  Schichtplan              output/SCHICHTPLAN.pdf
  5  Gästeliste               output/GAESTELISTE.pdf
  6  Workshop-Anmeldeliste    output/WORKSHOP.pdf
  7  Handbücher               HANDBUCH_DESIGNER + HANDBUCH_ZUSTAENDIGKEITEN
  8  Master-Tabelle neu       output/FUB2026_Werkdaten_MASTER.xlsx
  9  Werkmeldung-Vorlage neu  output/FUB2026_Werkmeldung_VORLAGE.xlsx
 10  QR-Codes                 output/qr-codes/ aus data/instagram.csv
 11  Werkmeldung migrieren    data/werkmeldung_alt/ → neue Vorlage
```

Direkt ohne Menü: `./.venv/bin/python scripts/fub.py 2`
(Argumente durchreichen: `… scripts/fub.py 3 -- --probe`).

Einzelnes Werkzeug geht auch weiterhin direkt:
`./.venv/bin/python scripts/build_druckdaten.py --nur werkliste`

Preis/Titel/Status geändert? Nur im Master, dann Menüpunkt 2 —
**nie** im PDF nachbessern.

---

## Auf Google Drive teilen

`./bau_upload.sh` baut **`output/upload/`** mit nur dem, was Teilnehmende und
Gestaltung brauchen (Werkmeldung-Vorlage, `design_VORLAGE.json`,
Designer-Handbuch, Beispiel-PDFs). Diesen Ordner hochladen — siehe
`output/upload/LIESMICH.txt`.
