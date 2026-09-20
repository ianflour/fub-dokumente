# data/ — was die Skripte hier lesen

Alles in diesem Ordner wird von Hand gepflegt. Die mitgelieferten Dateien sind
**frei erfundene Beispieldaten**; wer eigene Daten verwenden will, ersetzt sie
in derselben Form.

| Datei / Ordner | gelesen von | Inhalt | Beispiel vorhanden |
|---|---|---|---|
| `teilnehmende.txt` | `build_werkmeldung.py`, `build_master.py`, `build_qr_codes.py` | ein Name pro Zeile; speist das Dropdown „Künstler:in“ | ja |
| `instagram.csv` | `build_qr_codes.py`, `build_werkmeldung.py`, `build_druckdaten.py` | `Name;Dateiname;Instagram` — Quelle der QR-Codes | ja |
| `design.json` | `build_druckdaten.py --design` | alle DIN-A-Schildformate (a0 … a10), Schriftgrößen, Farben, automatische Anpassung — Erklärung: [README_design.md](README_design.md) | ja |
| `schichtplan_daten.py` | `build_schichtplan.py` | Datenbasis für den Schichtplan | ja |
| `gaesteliste_daten.py` | `build_gaesteliste.py` | Datenbasis für die Gästeliste | ja |
| `workshop_daten.py` | `build_workshop.py` | Datenbasis für die Workshop-Anmeldeliste | ja |
| `saal-referenz/` | `_saal_listen.py`, `build_werkmeldung.py` | Papier-/Formatkatalog als Nachschlagewerk | ja |
| `werkmeldungen/` | `import_werkmeldungen.py` | Ablage für ausgefüllte Werkmeldung(en) als `.xlsx` | Ordner leer (echte Meldungen enthalten Kontaktdaten und sind nicht versioniert) |
| `werkmeldung_alt/` | `migrate_werkmeldung.py` | Ablage für eine ältere, schon ausgefüllte Werkmeldung | Ordner leer, s. o. |
| `beispiele/Werkmeldung_BEISPIEL.xlsx` | — (nur zum Ausprobieren) | fertig ausgefüllte Werkmeldung mit 10 erfundenen Werken | ja |

Außerhalb von `data/` gibt es noch `assets/schriften/` — dort legt die Gestaltung
Regular- und Bold-Schnitt ab. Ohne Schriftdateien nimmt der Generator Arial /
Liberation Sans / Helvetica; Schriften sind lizenziert und daher nicht mitgeliefert.

## Beispiel-Werkmeldung ausprobieren

`beispiele/Werkmeldung_BEISPIEL.xlsx` ersetzt die echte Ablage, damit sich die
Skripte ohne eigene Daten testen lassen. Sie liegt bewusst **nicht** in
`werkmeldungen/`, sonst würde sie beim nächsten Einlesen als echte Meldung
importiert.

```bash
# Einlesen in den Master, nur anzeigen
./.venv/bin/python scripts/import_werkmeldungen.py --ordner data/beispiele --probe

# In eine frische Vorlage übertragen, nur anzeigen
./.venv/bin/python scripts/build_werkmeldung.py
./.venv/bin/python scripts/migrate_werkmeldung.py --alt data/beispiele/Werkmeldung_BEISPIEL.xlsx --probe
```

Neu erzeugen (z. B. nach Änderungen an der Vorlage oder an `teilnehmende.txt`):

```bash
./.venv/bin/python scripts/build_beispiel_werkmeldung.py     # oder Menüpunkt 12
```
