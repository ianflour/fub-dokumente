#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Erzeugt SCHICHTPLAN.pdf (A4 quer) aus schichtplan_daten.py.

Das Script VERTEILT die Personen selbst auf den Bedarf in SCHICHTEN — die
Namen stehen nicht in den Daten. Ziel: ein für alle möglichst ausgeglichener
Fairness-Quotient (annähernd gleiche Stundenzahl), unter Beachtung von Pools,
Blockern, möglichen Tagen und Vorlieben. Keine Soll-Stunden. Höchstens eine
Schicht pro Person und Tag.

Vorgaben aus schichtplan_daten.py: Rollen, Pools, Bedarf je Zeitfenster,
sowie je Person Vorlieben, Blocker und mögliche Tage.

Aufbau des PDF (Design wie die übrigen Handbücher):
  1. Programm            — Zeittafel
  2. Rollen              — Aufgaben + Pools
  3. Wer wann eingeteilt ist — der berechnete Plan, Tag für Tag
  4. Stundenübersicht    — Stunden + Fairness-Quotient je Person

Namen kommen aus data/teilnehmende.txt (über _teilnehmende.py). Nur
schichtplan_daten.py ändern, dann dieses Script neu laufen lassen.
"""

import glob
import os
import re
from datetime import date

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, Frame, KeepTogether, NextPageTemplate,
                                PageBreak, PageTemplate, Paragraph, Spacer, Table,
                                TableStyle)

import _pfade as P
import schichtplan_daten as D

try:
    import _teilnehmende as T
    NAMEN = list(T.NAMEN)
except Exception:                       # pragma: no cover
    NAMEN = list(D.PERSONEN)

OUT = P.output("SCHICHTPLAN.pdf")

# --- Schriften plattformunabhängig (Liberation Sans / Arial / DejaVu) -------
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
OFFEN = "#B24A3A"                       # „offen“-Markierung

W, H = landscape(A4)
ML = MR = 18 * mm
MT = 20 * mm
MB = 16 * mm
VB = W - ML - MR


def st(name, **kw):
    base = dict(fontName="S", fontSize=8.6, leading=12.4, textColor=INK,
                alignment=TA_LEFT, spaceAfter=0)
    base.update(kw)
    return ParagraphStyle(name, **base)


S = {
    "h1": st("h1", fontName="S-B", fontSize=21, leading=25, spaceAfter=3 * mm),
    "h2": st("h2", fontName="S-B", fontSize=13, leading=16, spaceBefore=7 * mm,
             spaceAfter=2 * mm),
    "h3": st("h3", fontName="S-B", fontSize=10.5, leading=14, spaceBefore=4 * mm,
             spaceAfter=1 * mm, keepWithNext=1),
    "p": st("p", fontSize=9.4, leading=13.6, spaceAfter=2.4 * mm),
    "lead": st("lead", fontSize=11, leading=16, textColor=GREY, spaceAfter=4 * mm),
    "tab": st("tab", fontSize=8.6, leading=11.8),
    "tabb": st("tabb", fontName="S-B", fontSize=8.6, leading=11.8),
    "tabh": st("tabh", fontName="S-B", fontSize=8.2, leading=11, textColor=GREY),
    "tabs": st("tabs", fontSize=8.0, leading=11, textColor=GREY),
    "boxh": st("boxh", fontName="S-B", fontSize=9.4, leading=13.6, spaceAfter=1.5 * mm),
    "boxp": st("boxp", fontSize=9.0, leading=13.4, spaceAfter=1.2 * mm),
}


def P(t, s="p"):
    return Paragraph(str(t), S[s])


def box(inhalt, farbe=BOX, rand=BOXR):
    t = Table([[inhalt]], colWidths=[VB])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), farbe),
        ("BOX", (0, 0), (-1, -1), 0.6, rand),
        ("LEFTPADDING", (0, 0), (-1, -1), 6 * mm),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 4 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3 * mm),
    ]))
    return t


def hinweis(titel, *absaetze, warnung=False):
    inner = [P(titel, "boxh")] + [P(a, "boxp") for a in absaetze]
    return KeepTogether(box(inner, WARN if warnung else BOX, WARNR if warnung else BOXR))


def tabelle(daten, breiten, kopf=True, zeilen_stil=None):
    """daten: Liste von Zeilen; jede Zelle ist ein Flowable oder str."""
    rows = []
    for i, r in enumerate(daten):
        zeile = []
        for c in r:
            if hasattr(c, "wrapOn"):
                zeile.append(c)
            else:
                zeile.append(Paragraph(str(c), S["tabh" if (kopf and i == 0) else "tab"]))
        rows.append(zeile)
    t = Table(rows, colWidths=breiten, repeatRows=1 if kopf else 0)
    stil = [
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3 * mm),
        ("TOPPADDING", (0, 0), (-1, -1), 1.9 * mm),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.9 * mm),
        ("LINEBELOW", (0, 0), (-1, -2 if kopf else -1), 0.4, RULE),
    ]
    if kopf:
        stil.append(("LINEBELOW", (0, 0), (-1, 0), 0.9, INK))
    for s in (zeilen_stil or []):
        stil.append(s)
    t.setStyle(TableStyle(stil))
    return t


# ================================================================= Verteilung
def _span(zeit):
    """'14:00–19:30' -> (840, 1170) Minuten ab Mitternacht.
    'ganztags' / leer -> (0, 1440).  Unlesbar -> None."""
    s = str(zeit or "").strip().lower()
    if not s or s in ("ganztags", "ganzer tag", "immer", "–", "-"):
        return (0, 1440)
    teile = re.split(r"\s*[–—-]\s*", s)
    if len(teile) != 2:
        return None
    try:
        a = [int(x) for x in teile[0].split(":")]
        b = [int(x) for x in teile[1].split(":")]
        return (a[0] * 60 + a[1], b[0] * 60 + b[1])
    except (ValueError, IndexError):
        return None


def _stunden(zeit):
    sp = _span(zeit)
    return round((sp[1] - sp[0]) / 60.0, 2) if sp else 0.0


def _ueberlappt(a, b):
    return a and b and a[0] < b[1] and b[0] < a[1]


# Höchstens eine Schicht pro Person und Tag. Wer an einem Tag zwei Schichten
# machen soll, muss dann auf False gesetzt werden.
EINE_SCHICHT_PRO_TAG = True

# Mindestpause zwischen zwei Schichten derselben Person am selben Tag (Minuten).
# Greift nur, wenn EINE_SCHICHT_PRO_TAG = False ist.
MIN_PAUSE = 30


def _abstand(a, b):
    """Lücke zwischen zwei Zeitspannen in Minuten. Überlappung -> negativ."""
    if not a or not b:
        return 9999
    if a[0] < b[1] and b[0] < a[1]:
        return -1
    return b[0] - a[1] if a[1] <= b[0] else a[0] - b[1]


def _zu_dicht(a, b):
    """True, wenn a und b sich überschneiden oder weniger als MIN_PAUSE
    Minuten auseinanderliegen (= „hintereinander")."""
    return _abstand(a, b) < MIN_PAUSE


def blocker_text(info):
    """Verfügbarkeits-Einschränkungen (tage + blocker) -> lesbarer Text."""
    info = info or {}
    out = []
    tage = info.get("tage") or []
    if tage:
        out.append("nur " + ", ".join(tag_label(t) for t in tage))
    for e in info.get("blocker", []) or []:
        if isinstance(e, (list, tuple)) and e:
            tag = tag_label(e[0])
            zeit = e[1] if len(e) > 1 else "ganztags"
            grund = f" ({e[2]})" if len(e) > 2 and e[2] else ""
            if str(zeit).strip().lower() in ("ganztags", "", "ganzer tag"):
                out.append(f"nicht {tag}{grund}")
            else:
                out.append(f"nicht {tag} {zeit}{grund}")
    return " · ".join(out)


def _rollen_reihenfolge():
    return [r[0] for r in D.ROLLEN]


def verteile():
    """Verteilt die Personen auf den Bedarf in D.SCHICHTEN.

    Ziel: ein für alle möglichst ausgeglichener Fairness-Quotient — jede
    Person bekommt ungefähr gleich viele Stunden. Keine Soll-Stunden, keine
    Vorzugsbehandlung (auch kein „Orga-Team"). Unter Beachtung von Pools
    (Rollen mit fester Kandidatenliste), Blockern (Sperrzeiten), möglichen
    Tagen und Vorlieben. Höchstens eine Schicht pro Person und Tag
    (EINE_SCHICHT_PRO_TAG). Bleibt ein Platz unbesetzbar, erscheint er als
    „offen". Deterministisch (Gleichstand: nach Name).

    -> (shifts, plan, warnungen)
       shifts: Liste dict(idx, datum, zeit, anlass, span, stunden, bedarf)
       plan:   {idx: {rolle: [name | ""]}}   ("" = Platz blieb offen)
    """
    warnungen = []
    rollen = _rollen_reihenfolge()
    pools = getattr(D, "POOLS", {}) or {}
    gueltige = set(getattr(D, "AUFGABEN", []) or rollen)

    aktive = [n for n in list(NAMEN) + list(D.PERSONEN)
              if n not in D.NICHT_EINGEPLANT]
    seen = set()
    aktive = [n for n in aktive if not (n in seen or seen.add(n))]

    info = {n: (D.PERSONEN.get(n, {}) or {}) for n in aktive}
    vorliebe = {n: [v for v in (info[n].get("vorliebe") or [])] for n in aktive}
    for n in aktive:
        for v in vorliebe[n]:
            if v not in gueltige:
                warnungen.append(f"{n}: Vorliebe „{v}“ ist keine gültige Aufgabe")

    blocker = {}
    for n in aktive:
        bl = []
        for e in info[n].get("blocker", []) or []:
            if isinstance(e, (list, tuple)) and e:
                sp = _span(e[1] if len(e) > 1 else "ganztags")
                if sp:
                    bl.append((str(e[0]), sp))
        blocker[n] = bl
    tage = {n: set(info[n].get("tage") or []) for n in aktive}   # leer = alle Tage

    # Tandems: Personen, die nur gemeinsam eingeteilt werden — immer dieselbe
    # Schicht oder gar nicht. partner[a] = b und partner[b] = a.
    partner = {}
    for grp in getattr(D, "TANDEM", []) or []:
        grp = [n for n in grp if isinstance(n, str)]
        if len(grp) != 2:
            warnungen.append(f"TANDEM {grp}: bitte genau zwei Namen")
            continue
        a, b = grp
        fehlt = [n for n in grp if n not in aktive]
        if fehlt:
            warnungen.append(f"TANDEM {a} / {b}: {', '.join(fehlt)} nicht aktiv")
            continue
        partner[a], partner[b] = b, a

    shifts = []
    for i, sch in enumerate(D.SCHICHTEN):
        sp = _span(sch["zeit"])
        shifts.append(dict(idx=i, datum=sch["datum"], zeit=sch["zeit"],
                           anlass=sch["anlass"], span=sp,
                           stunden=round((sp[1] - sp[0]) / 60.0, 2) if sp else 0.0,
                           bedarf=dict(sch.get("bedarf", {})),
                           fix=dict(sch.get("fix", {}))))

    ist = {n: 0.0 for n in aktive}
    plan = {sh["idx"]: {} for sh in shifts}
    belegt = {n: [] for n in aktive}          # name -> [(datum, span)]

    def frei(n, sh):
        if tage[n] and sh["datum"] not in tage[n]:
            return False
        if sh["span"] is None:
            return True
        for (d, s) in belegt[n]:
            if d != sh["datum"]:
                continue
            if EINE_SCHICHT_PRO_TAG:
                return False               # höchstens eine Schicht pro Tag
            if _zu_dicht(s, sh["span"]):
                return False               # Überschneidung oder direkt hintereinander
        for (d, s) in blocker[n]:
            if d == sh["datum"] and _ueberlappt(s, sh["span"]):
                return False
        return True

    def _setze(n, sh, rolle):
        plan[sh["idx"]].setdefault(rolle, []).append(n)
        ist[n] += sh["stunden"]
        belegt[n].append((sh["datum"], sh["span"]))

    def zuweisen(n, sh, rolle):
        _setze(n, sh, rolle)
        p = partner.get(n)
        if (p and p not in plan[sh["idx"]].get(rolle, []) and frei(p, sh)
                and (rolle not in pools or p in pools[rolle])
                and len(plan[sh["idx"]][rolle]) < sh["bedarf"].get(rolle, 0)):
            _setze(p, sh, rolle)

    def kann(n, sh, rolle):
        if rolle in pools and n not in pools[rolle]:
            return False
        if n in plan[sh["idx"]].get(rolle, []):
            return False
        if not frei(n, sh):
            return False
        p = partner.get(n)
        if p and p not in plan[sh["idx"]].get(rolle, []):
            # Tandem: der Partner muss dieselbe Schicht/Rolle auch übernehmen
            # können und es müssen noch zwei Plätze frei sein.
            if rolle in pools and p not in pools[rolle]:
                return False
            if not frei(p, sh) or kapazitaet(sh, rolle) < 2:
                return False
        return True

    # 1) feste Zuweisungen
    for sh in shifts:
        for rolle, namen in sh["fix"].items():
            for n in namen:
                if n not in aktive:
                    warnungen.append(f"{tag_label(sh['datum'])} {sh['zeit']}: "
                                     f"fixe Person „{n}“ ist nicht aktiv")
                    continue
                if frei(n, sh):
                    zuweisen(n, sh, rolle)
                else:
                    warnungen.append(f"{tag_label(sh['datum'])} {sh['zeit']} · {rolle}: "
                                     f"„{n}“ fest gesetzt, aber blockiert/doppelt")

    fix_namen = set()
    for sh in shifts:
        for ns in sh["fix"].values():
            fix_namen.update(ns)

    def kapazitaet(sh, rolle):
        return sh["bedarf"].get(rolle, 0) - len(plan[sh["idx"]].get(rolle, []))

    def eff_last(n):
        return ist[n]

    def last_key(n, sh, rolle):
        """kleiner = besser. Ausgeglichener Fairness-Quotient: wenig belastet
        zuerst, dann wenige Schichten. Vorliebe zieht auf die Wunsch-Rolle vor."""
        vb = 0.0
        if rolle in vorliebe[n]:
            vb = 2.5 - 0.4 * vorliebe[n].index(rolle)
        return (round(ist[n] - vb, 3), len(belegt[n]), n)

    # 2) alle offenen Plätze füllen — knappste Rolle zuerst (Pools), dann die
    #    langen Schichten (solange noch viele Leute frei sind), dann kurze;
    #    je Platz die am wenigsten belastete passende Person.
    offen = []
    for sh in shifts:
        for rolle in sh["bedarf"]:
            n_kand = len(pools[rolle]) if rolle in pools else len(aktive)
            offen += [(n_kand, -sh["stunden"], sh["idx"], rolle)] * max(0, kapazitaet(sh, rolle))
    offen.sort(key=lambda o: (o[0], o[1], o[2]))

    for _, _, idx, rolle in offen:
        sh = shifts[idx]
        if kapazitaet(sh, rolle) <= 0:
            continue                       # schon voll (z. B. durch Tandem-Partner)
        kand = [n for n in aktive if kann(n, sh, rolle)]
        if not kand:
            plan[sh["idx"]].setdefault(rolle, []).append("")
            continue
        zuweisen(min(kand, key=lambda n: last_key(n, sh, rolle)), sh, rolle)

    for a, b in {tuple(sorted((n, p))) for n, p in partner.items()}:
        if not belegt[a] and not belegt[b]:
            warnungen.append(f"Tandem {a} / {b}: keine gemeinsame Schicht gefunden")

    # 3) Ausgleich: eine stärker belastete Person gibt eine Schicht an eine
    #    weniger belastete ab, sobald das den Stundenabstand zwischen beiden
    #    verkleinert (frei, passende Rolle, kein Zeitkonflikt). Läuft, bis
    #    kein solcher Tausch mehr möglich ist — die Stunden liegen dann so
    #    dicht beieinander, wie es die Schichtlängen zulassen.
    #    Fest gesetzte Personen (fix, z. B. Workshop-Leitung) zählen nicht
    #    in den Schnitt und werden nicht getauscht.
    frei_von_fix = [n for n in aktive if n not in fix_namen]
    tauschbar = [n for n in frei_von_fix if n not in partner]   # Tandems nicht einzeln tauschen
    for _ in range(2000):
        basis = frei_von_fix or aktive
        schnitt = sum(ist[n] for n in basis) / max(len(basis), 1)
        getauscht = False
        for a_n in sorted((n for n in tauschbar if ist[n] > schnitt + 1e-6),
                          key=lambda n: -ist[n]):
            for sh in shifts:
                for rolle, leute in list(plan[sh["idx"]].items()):
                    if a_n not in leute or any(a_n in v for v in sh["fix"].values()):
                        continue
                    for b_n in sorted(tauschbar, key=lambda n: ist[n]):
                        if b_n == a_n or b_n in leute:
                            continue
                        # nur tauschen, wenn der Abstand danach kleiner ist
                        if sh["stunden"] >= ist[a_n] - ist[b_n] - 1e-9:
                            continue
                        if rolle in pools and b_n not in pools[rolle]:
                            continue
                        if tage[b_n] and sh["datum"] not in tage[b_n]:
                            continue
                        if EINE_SCHICHT_PRO_TAG:
                            if any(d == sh["datum"] for (d, s) in belegt[b_n]):
                                continue
                        elif any(d == sh["datum"] and _zu_dicht(s, sh["span"])
                                 for (d, s) in belegt[b_n]):
                            continue
                        if any(d == sh["datum"] and _ueberlappt(s, sh["span"])
                               for (d, s) in blocker[b_n]):
                            continue
                        leute.remove(a_n)
                        leute.append(b_n)
                        ist[a_n] -= sh["stunden"]
                        ist[b_n] += sh["stunden"]
                        belegt[a_n].remove((sh["datum"], sh["span"]))
                        belegt[b_n].append((sh["datum"], sh["span"]))
                        getauscht = True
                        break
                    if getauscht:
                        break
                if getauscht:
                    break
            if getauscht:
                break
        if not getauscht:
            break

    return shifts, plan, warnungen


def einsaetze_von(shifts, plan):
    e = {n: [] for n in NAMEN}
    for sh in shifts:
        for rolle, namen in plan[sh["idx"]].items():
            for n in namen:
                if n and str(n).strip():
                    e.setdefault(n, []).append(dict(
                        datum=sh["datum"], zeit=sh["zeit"], anlass=sh["anlass"],
                        rolle=rolle, stunden=sh["stunden"]))
    return e


def fmt_std(x):
    if x is None:
        return "—"
    return (f"{x:.0f}" if float(x).is_integer() else f"{x:.2f}".rstrip("0").rstrip(".")
            ).replace(".", ",")


_WT = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]


def tag_label(datum):
    """'2026-09-12' -> 'Sa 12.09.'  ·  unbekanntes Format bleibt unverändert."""
    try:
        d = date.fromisoformat(str(datum))
    except ValueError:
        return str(datum)
    return f"{_WT[d.weekday()]} {d:%d.%m.}"


# ================================================================= Deckblatt
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
    """Ganzseitiges Titelblatt (A4 quer)."""
    c.saveState()
    v = D.VERANSTALTUNG

    y = H - 34 * mm
    c.setFillColor(LIGHT)
    _sperr(c, ML, y, v["titel"].upper(), "S-B", 10.5, 2)
    y -= 6 * mm
    c.setStrokeColor(INK)
    c.setLineWidth(1)
    c.line(ML, y, W - MR, y)

    y -= 30 * mm
    c.setFillColor(INK)
    c.setFont("S-B", 32)
    c.drawString(ML, y, "Schichtplan")

    by = MB + 4 * mm
    c.setStrokeColor(RULE)
    c.setLineWidth(0.5)
    c.line(ML, by + 15 * mm, W - MR, by + 15 * mm)
    c.setFont("S", 9)
    c.setFillColor(GREY)
    c.drawString(ML, by + 8 * mm, v["untertitel"])
    c.drawString(ML, by + 3 * mm, f"{v['laufzeit']} · {v['ort']}")
    c.setFillColor(LIGHT)
    c.drawString(ML, by - 2.5 * mm, f"{v['kontakt']} · Stand {date.today():%d.%m.%Y}")
    c.restoreState()


def kopf_fuss(c, doc):
    c.saveState()
    seite = doc.page - 1
    if seite >= 1:
        c.setFont("S", 7.6)
        c.setFillColor(LIGHT)
        c.drawString(ML, H - MT + 9 * mm, f"{D.VERANSTALTUNG['titel']} · Schichtplan")
        c.setStrokeColor(RULE)
        c.setLineWidth(0.4)
        c.line(ML, H - MT + 6.5 * mm, W - MR, H - MT + 6.5 * mm)
        c.drawRightString(W - MR, MB - 9 * mm, f"Seite {seite}")
    c.setFont("S", 7.6)
    c.setFillColor(LIGHT)
    c.drawString(ML, MB - 9 * mm, D.VERANSTALTUNG["kontakt"])
    c.restoreState()


# ================================================================= Inhalt
def inhalt(shifts, plan, einsaetze, warnungen):
    F = []
    a = F.append
    rollen_namen = [r[0] for r in D.ROLLEN]

    a(Spacer(1, 1 * mm))

    # ---------------------------------------------------------- 1 Programm
    if getattr(D, "PROGRAMM", None):
        a(P("Programm", "h2"))
        daten = [["Tag", "", "Programm"]]
        for datum, titel, punkte in D.PROGRAMM:
            pp = "<br/>".join(f"<b>{z}</b>&nbsp;&nbsp;{w}" for z, w in punkte)
            daten.append([P(f"<b>{tag_label(datum)}</b>", "tab"),
                          P(f'<font color="#5E5E5E">{titel}</font>', "tabs"),
                          P(pp, "tab")])
        a(tabelle(daten, [22 * mm, 46 * mm, VB - 68 * mm]))

    # ---------------------------------------------------------- 2 Rollen
    pools = getattr(D, "POOLS", None) or {}
    daten = [["Rolle", "Aufgabe", "Wann"]]
    for eintrag in D.ROLLEN:
        name, beschr = eintrag[0], eintrag[1]
        wann = eintrag[2] if len(eintrag) > 2 else ""
        if name in pools:
            beschr = (f"{beschr}<br/><font color='#5E5E5E'>Pool: "
                      f"{', '.join(pools[name])}</font>")
        daten.append([P(f"<b>{name}</b>", "tab"), P(beschr, "tab"), P(wann, "tabs")])
    inner = [P("Rollen", "h2"),
             tabelle(daten, [46 * mm, VB - 46 * mm - 44 * mm, 44 * mm])]
    a(KeepTogether(inner))

    # ---------------------------------------------------------- 3 Schichtplan
    a(PageBreak())
    a(P("Wer wann eingeteilt ist", "h2"))
    a(P("Nach Uhrzeit sortiert. Je Zeitfenster eine Zeile; die Zahl in Klammern "
        "ist die benötigte Personenzahl, „offen“ (rot) = Platz noch frei. "
        "Jede Person hat pro Tag höchstens eine Schicht.", "tabs"))

    tage = []
    for sh in shifts:
        if not tage or tage[-1][0] != sh["datum"]:
            tage.append((sh["datum"], []))
        tage[-1][1].append(sh)

    def _gruppen(schichten):
        """Schichten chronologisch; identische Zeitfenster in eine Gruppe."""
        g = []
        for sh in sorted(schichten, key=lambda s: (s["span"] or (0, 0), s["zeit"])):
            if g and g[-1][0] == sh["zeit"]:
                g[-1][1].append(sh)
            else:
                g.append((sh["zeit"], [sh]))
        return g

    def _besetzung(shs):
        """(Anlass-Text, Besetzungs-Text) für eine Zeitfenster-Gruppe."""
        anlaesse, bes = [], []
        for sh in shs:
            if sh["anlass"] not in anlaesse:
                anlaesse.append(sh["anlass"])
            zuord = plan.get(sh["idx"], {})
            rollen_hier = ([r for r in rollen_namen if r in sh["bedarf"] or r in zuord]
                           + [r for r in zuord if r not in rollen_namen])
            for rolle in rollen_hier:
                leute = list(zuord.get(rolle, []))
                soll_n = sh["bedarf"].get(rolle, len(leute))
                while len(leute) < soll_n:
                    leute.append("")
                teile = [(str(x).strip() if str(x).strip()
                          else f'<font color="{OFFEN}">offen</font>') for x in leute]
                bes.append(f"<b>{rolle}</b> <font color='#5E5E5E'>({soll_n})</font>  "
                           + ", ".join(teile))
        return " · ".join(anlaesse), "<br/>".join(bes)

    # --- Vernissage (12.09.) auf einer eigenen Seite ---
    if tage:
        datum, schichten = tage[0]
        a(P(tag_label(datum), "h3"))
        daten = [["Zeit", "Anlass", "Besetzung"]]
        for zeit, shs in _gruppen(schichten):
            anl, bes = _besetzung(shs)
            daten.append([P(f"<b>{zeit}</b>", "tab"), P(anl, "tabs"), P(bes, "tab")])
        a(tabelle(daten, [24 * mm, 62 * mm, VB - 86 * mm]))

    # --- 13.–19.09. kompakt: eine Tabelle, Tag in Spalte 1, auf ein Blatt ---
    rest = [["Tag", "Zeit", "Anlass", "Besetzung"]]
    tag_start = []                         # Zeilen, in denen ein neuer Tag beginnt
    for datum, schichten in tage[1:]:
        tag_start.append(len(rest))
        for gi, (zeit, shs) in enumerate(_gruppen(schichten)):
            anl, bes = _besetzung(shs)
            rest.append([P(f"<b>{tag_label(datum)}</b>" if gi == 0 else "", "tab"),
                         P(f"<b>{zeit}</b>", "tab"), P(anl, "tabs"), P(bes, "tab")])
    if len(rest) > 1:
        stil = [("LINEABOVE", (0, i), (-1, i), 0.5, RULE) for i in tag_start[1:]]
        a(PageBreak())
        a(P("So 13.09. – Sa 19.09.", "h3"))
        a(tabelle(rest, [20 * mm, 24 * mm, 60 * mm, VB - 104 * mm], zeilen_stil=stil))

    # ---------------------------------------------------------- 4 Stundenübersicht
    a(PageBreak())
    a(P("Stundenübersicht &amp; Fairness", "h2"))

    wl_fix = set()
    for sh in shifts:
        for r, ns in sh["fix"].items():
            if r == "Workshop-Leitung":
                wl_fix.update(ns)

    zeilen = []
    for n in NAMEN:
        if n in D.NICHT_EINGEPLANT:
            continue
        info = D.PERSONEN.get(n, {}) or {}
        e = einsaetze.get(n, [])
        ist = sum(x["stunden"] for x in e)
        vorliebe = ", ".join(info.get("vorliebe", []) or [])
        slots = " · ".join(f"{tag_label(x['datum'])} {x['zeit']} · {x['rolle']}" for x in e)
        hinweis_txt = " · ".join(t for t in (blocker_text(info),
                                             info.get("notiz", "")) if t)
        rolle_tag = "Workshop-Leitung" if n in wl_fix else ""
        zeilen.append((ist, n, vorliebe, slots, hinweis_txt, rolle_tag))
    zeilen.sort(key=lambda z: (-z[0], z[1]))

    ist_alle = [z[0] for z in zeilen]
    normal = [z[0] for z in zeilen if z[5] != "Workshop-Leitung"]   # alle außer WL
    ref = (sum(normal) / len(normal)) if normal else (sum(ist_alle) / max(len(ist_alle), 1))

    def quot(h):
        return h / ref if ref else 0.0

    def fq(q):
        return f"{q:.2f}".replace(".", ",")

    a(P(f"„Fairness-Quotient“ = Stunden ÷ {fmt_std(round(ref, 2))} h (Ø aller "
        f"Personen ohne die Workshop-Leitung). <b>1,00</b> = genau dieser Schnitt. "
        "Ziel ist ein für alle möglichst gleicher Wert; nur die Workshop-Leitung "
        "liegt bauartbedingt darüber.", "tabs"))

    sp = [42 * mm, 33 * mm, 12 * mm, 18 * mm, 102 * mm]
    sp.append(VB - sum(sp))
    daten = [["Name", "Vorliebe", "Std.", "Quotient", "Eingeteilt für", "Hinweis"]]
    for ist, n, vorliebe, slots, hinweis_txt, rolle_tag in zeilen:
        nm = f"<b>{n}</b>" + (f" <font size=7 color='#5E5E5E'>· {rolle_tag}</font>"
                              if rolle_tag else "")
        daten.append([P(nm, "tab"), P(vorliebe or "—", "tabs"),
                      P(f"<b>{fmt_std(ist)}</b>", "tab"), P(fq(quot(ist)), "tab"),
                      P(slots or "—", "tabs"), P(hinweis_txt or "", "tabs")])
    ist_sum = sum(ist_alle)
    daten.append([P("<b>Summe</b>", "tab"), "",
                  P(f"<b>{fmt_std(round(ist_sum, 2))}</b>", "tab"), "",
                  P(f"{len(zeilen)} Personen · Ø {fmt_std(round(ist_sum / max(len(zeilen), 1), 2))} h",
                    "tabs"), ""])
    stil = [("LINEABOVE", (0, len(daten) - 1), (-1, len(daten) - 1), 0.9, INK)]
    a(tabelle(daten, sp, zeilen_stil=stil))

    # --- Fairness-Kennzahlen ------------------------------------------------
    s1, s2 = sum(ist_alle), sum(h * h for h in ist_alle)
    jain = (s1 * s1) / (len(ist_alle) * s2) if s2 else 1.0
    nq = [quot(h) for h in normal]
    a(Spacer(1, 2.5 * mm))
    a(P(f"<b>Fairness-Index</b> (Jain, alle {len(ist_alle)} Personen): "
        f"<b>{fq(jain)}</b> von 1,00 — 1,00 wäre völlig gleiche Stundenzahl.", "tabs"))
    if nq:
        a(P(f"<b>Ohne Workshop-Leitung</b> ({len(nq)} Personen): "
            f"Quotient {fq(min(nq))} bis {fq(max(nq))}, Ø 1,00.", "tabs"))

    return F


def main():
    doc = BaseDocTemplate(OUT, pagesize=landscape(A4),
                          leftMargin=ML, rightMargin=MR, topMargin=MT, bottomMargin=MB,
                          title=f"{D.VERANSTALTUNG['titel']} — Schichtplan",
                          author="KULTURVEREIN MUSTERSTADT")
    frame = Frame(ML, MB, W - ML - MR, H - MT - MB, id="f",
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([
        PageTemplate(id="cover", frames=[Frame(0, 0, W, H, id="cv")], onPage=deckblatt),
        PageTemplate(id="main", frames=[frame], onPage=kopf_fuss),
    ])
    shifts, plan, warnungen = verteile()
    einsaetze = einsaetze_von(shifts, plan)
    doc.build([NextPageTemplate("main"), Spacer(1, 2), PageBreak()]
              + inhalt(shifts, plan, einsaetze, warnungen))

    # kurze Zusammenfassung ins Terminal
    aktiv = [n for n in NAMEN if n not in D.NICHT_EINGEPLANT]
    print(f"{OUT} geschrieben.")
    for n in sorted(aktiv, key=lambda n: (-(sum(x['stunden'] for x in einsaetze.get(n, []))), n)):
        e = einsaetze.get(n, [])
        ist = sum(x["stunden"] for x in e)
        print(f"  {n:22s} {ist:>4} h  ({len(e)} Schicht(en))")
    offen = sum(1 for sh in shifts for r, ls in plan[sh['idx']].items() for x in ls if not str(x).strip())
    if offen:
        print(f"  --> {offen} Platz/Plätze noch offen")
    for w in warnungen:
        print("  WARNUNG:", w)


if __name__ == "__main__":
    main()
