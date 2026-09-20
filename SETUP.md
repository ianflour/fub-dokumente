# Einrichtung der Arbeitsumgebung

Kurzfassung für den Rechner, auf dem die PDFs und Tabellen erzeugt werden.

## Voraussetzung

Python 3.9 oder neuer. Prüfen:

```bash
python3 --version
```

## Einrichten

Ein Aufruf:

```bash
./setup.sh
```

Das legt eine gekapselte Umgebung `.venv/` an und installiert die beiden
benötigten Pakete aus `requirements.txt`:

| Paket | wofür |
|---|---|
| `openpyxl` | liest und schreibt die `.xlsx`-Dateien (Master, Werkmeldung) |
| `reportlab` | erzeugt die Druck-PDFs (Wandschilder, Werkliste, Etiketten, Versicherungsliste) |
| `segno` | erzeugt die Instagram-QR-Codes (`scripts/build_qr_codes.py`, Wandschilder, Instagram-Wand) |

`et-xmlfile` (für openpyxl) und `pillow` (für reportlab) kommen automatisch mit;
`segno` hat keine Abhängigkeiten.

### Von Hand statt `setup.sh`

```bash
python3 -m venv .venv
./.venv/bin/pip install -r requirements.txt
```

`requirements.lock.txt` hält die exakt getesteten Versionen fest
(`pip install -r requirements.lock.txt` für eine 1:1-Reproduktion).

Hinter einem Firmen-/Sandbox-Proxy kann pip am TLS-Zertifikat scheitern
(`OSStatus -26276`). Dann:

```bash
./.venv/bin/pip install --use-deprecated=legacy-certs -r requirements.txt
```

(`setup.sh` versucht das automatisch als zweiten Anlauf.)

## Benutzen

Immer den Python-Interpreter aus `.venv` nehmen. Alles läuft über das Menü:

```bash
./.venv/bin/python scripts/menue.py              # Menü
./.venv/bin/python scripts/menue.py 2            # direkt: alle Druck-PDFs
./.venv/bin/python scripts/menue.py 3 -- --probe # Argumente durchreichen
```

Einzelne Werkzeuge gehen auch direkt:

```bash
./.venv/bin/python scripts/build_druckdaten.py       # Druck-PDFs -> output/ausgabe/
./.venv/bin/python scripts/import_werkmeldungen.py --help
./.venv/bin/python scripts/build_qr_codes.py         # fehlende Instagram-QR-Codes
./.venv/bin/python scripts/build_master.py           # Vorlagen neu bauen (Notfall)
./.venv/bin/python scripts/build_werkmeldung.py
```

Oder die Umgebung für die Sitzung aktivieren:

```bash
source .venv/bin/activate
python scripts/menue.py
```

## Schriften

Der Generator sucht selbst nach Arial bzw. einer metrisch identischen
Schrift und nimmt sonst Helvetica — das reicht für einen Testlauf.

Für die finale Produktion legt die Gestaltung zwei Schnitte in
`assets/schriften/` ab (Regular + Bold, Dateiname egal — werden automatisch
erkannt) und trägt Größen/Farben in `data/design.json` ein. Aufruf dann:

```bash
./.venv/bin/python scripts/build_druckdaten.py --design data/design.json
```

Hinweis: `build_handbuch_designer.py` und `build_handbuch_zustaendigkeiten.py`
suchen die Schrift plattformübergreifend (Liberation Sans / Arial / DejaVu)
und fallen sonst auf Helvetica/Courier zurück — sie laufen also auch ohne
installierte Schriften.

## Offline-Rechner ohne Zugang zu pypi.org

Auf einem anderen Rechner mit Internet einmal die Pakete herunterladen …

```bash
pip download -r requirements.txt -d wheelhaus --only-binary=:all:
```

… den Ordner `wheelhaus/` mitkopieren und hier installieren:

```bash
python3 -m venv .venv
./.venv/bin/pip install --no-index --find-links wheelhaus -r requirements.txt
```
