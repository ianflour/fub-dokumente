# Saal-Digital-Referenz

Nachschlagematerial zu den Druckoptionen von **Saal Digital**
(saal-digital.de, Bereiche *Wandbilder* und *Poster / FineArt*, Stand September 2026).
Gehört **nicht** zum Beschilderungs-Workflow — nur zur Orientierung für
Teilnehmende und Gestaltung. Speist die Auswahllisten in `../saal_listen.py`.

| Datei | Inhalt |
|---|---|
| `saal_digital_wandbild_poster_fineart.csv` | Rohdaten: eine Zeile je Produkt × Material × Ausrichtung × Format (3308 Zeilen, 8 Spalten, `;`-getrennt) |
| `saal_digital_katalog.csv` | dieselben Daten, in 22 einzelne, saubere Spalten zerlegt und kategorisiert |
| `saal_digital_katalog.xlsx` | Katalog + drei Auswertungsblätter (Zusammenfassung, Produktlinien, Materialien) |
| `fotoabzuege_infos.csv` | Fotoabzüge: 6 Oberflächen (Grammatur, Material, FSC, Beschreibung) + verfügbare Formate nach Seitenverhältnis. Fließt in `../saal_listen.py` (`FOTO_OBERFLAECHEN`, `FOTO_FORMATE`, Papier-Dropdown). |
| `saal_katalog_aufbereiten.py` | erzeugt `katalog.csv` + `katalog.xlsx` aus der Rohdatei neu (nutzt `fotoabzuege_infos.csv` nicht — das ist ein separater Auszug) |

## Neu aufbereiten

Wenn die Rohdatei aktualisiert wird:

```bash
./.venv/bin/python saal-referenz/saal_katalog_aufbereiten.py
```

## Spalten in `saal_digital_katalog.csv`

`Kategorie · Produktlinie · Teile · Anordnung · Materialgruppe · Material ·
Oberflaeche · Marke · Druckverfahren · Verfahrensgruppe · Ausrichtung ·
Seitenverhaeltnis · Breite_cm · Hoehe_cm · Format · DIN · Gesamtformat ·
Aufhaengung · Rahmenfarben · Staerke · Form · Hinweis`
