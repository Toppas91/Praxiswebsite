#!/usr/bin/env python3
# coding: utf-8
"""
aktualisieren.py  –  Praxis-Editor
────────────────────────────────────
Erster Aufruf (Excel fehlt):
    Erstellt praxis-editor.xlsx aus praxis-daten.json.
    Bitte dann die Excel-Datei bearbeiten und das Script erneut starten.

Normaler Aufruf (Excel vorhanden):
    Liest praxis-editor.xlsx und schreibt Änderungen in praxis-daten.json.

Aufruf mit --neu:
    Erstellt praxis-editor.xlsx neu aus der aktuellen praxis-daten.json
    (z. B. nach manuellem Hinzufügen eines Vertretungsarztes).
"""
import json
import sys
from datetime import datetime, time as dtime
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    print("✗ openpyxl nicht gefunden.")
    print("  Bitte ausführen:  pip install openpyxl")
    sys.exit(1)

BASE        = Path(__file__).parent
JSON_FILE   = BASE / "praxis-daten.json"
EXCEL_FILE  = next(
    (BASE / n for n in ("praxis-editor.xlsm", "praxis-editor.xlsx")
     if (BASE / n).exists()),
    BASE / "praxis-editor.xlsm"
)
WOCHENTAGE  = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag"]


# ═══════════════════════════════════════════════════════════════
#  HILFSFUNKTIONEN
# ═══════════════════════════════════════════════════════════════

def als_zeit(val):
    """Zellwert → 'HH:MM' oder None."""
    if val is None:
        return None
    if isinstance(val, float):
        mins = round(val * 1440)
        return f"{mins // 60:02d}:{mins % 60:02d}"
    if isinstance(val, dtime):
        return val.strftime("%H:%M")
    if isinstance(val, datetime):
        return val.strftime("%H:%M")
    s = str(val).strip()
    if not s or s in ("-", "–", "—"):
        return None
    try:
        datetime.strptime(s, "%H:%M")
        return s
    except ValueError:
        return None

def als_datum(val):
    """Zellwert → 'YYYY-MM-DD' oder None."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if not s:
        return None
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None

def als_int(val, default=7):
    if val is None:
        return default
    try:
        return int(float(str(val)))
    except (ValueError, TypeError):
        return default

def get_sheet(wb, token):
    """Findet ein Sheet anhand eines Teil-Namens (umlaut-tolerant)."""
    for name in wb.sheetnames:
        if token.lower() in name.lower():
            return wb[name]
    raise KeyError(f"Sheet mit '{token}' nicht gefunden. Vorhanden: {wb.sheetnames}")


# ═══════════════════════════════════════════════════════════════
#  JSON LESEN (EXCEL → JSON)
# ═══════════════════════════════════════════════════════════════

def lese_oeffnungszeiten(ws):
    tage = []
    for i, name in enumerate(WOCHENTAGE):
        r = 5 + i
        def z(col, _r=r): return als_zeit(ws.cell(row=_r, column=col).value)
        tage.append({
            "tag": name,
            "gesund": {"von": z(2), "bis": z(3), "nachmittag_von": z(4), "nachmittag_bis": z(5)},
            "krank":  {"von": z(6), "bis": z(7), "nachmittag_von": z(8), "nachmittag_bis": z(9)},
        })
    return tage

def lese_schliesstage(ws):
    ergebnis = []
    for row in range(4, 20):
        von     = als_datum(ws.cell(row=row, column=1).value)
        bis     = als_datum(ws.cell(row=row, column=2).value)
        grund   = str(ws.cell(row=row, column=3).value or "").strip()
        vorlauf = als_int(ws.cell(row=row, column=4).value, default=7)
        if von and bis:
            ergebnis.append({"von": von, "bis": bis, "grund": grund, "vorlauf_tage": vorlauf})
    return ergebnis

def lese_vertretung(ws, bestehende):
    id_map = {a["id"]: a for a in bestehende}
    neu = []
    for row in range(4, 30):
        id_val    = ws.cell(row=row, column=1).value
        aktiv_val = ws.cell(row=row, column=2).value
        if id_val is None:
            break
        arzt_id = str(id_val).strip()
        if arzt_id in id_map:
            arzt = dict(id_map[arzt_id])
            arzt["aktiv"] = str(aktiv_val or "").strip().upper() == "JA"
            neu.append(arzt)
    return neu or bestehende

def git_push():
    """Führt git add → commit → push aus. Gibt True bei Erfolg zurück."""
    import subprocess

    def git(args, **kw):
        return subprocess.run(
            ["git"] + args,
            cwd=BASE,
            capture_output=True,
            text=True,
            **kw
        )

    # Prüfen ob überhaupt ein git-Repo vorliegt
    check = git(["rev-parse", "--is-inside-work-tree"])
    if check.returncode != 0:
        print("⚠ Kein Git-Repository gefunden – Push übersprungen.")
        return False

    # Gibt es überhaupt Änderungen?
    status = git(["status", "--porcelain", str(JSON_FILE)])
    if not status.stdout.strip():
        print("  Keine Änderungen in praxis-daten.json – kein Commit nötig.")
        return True

    zeitstempel = datetime.now().strftime("%d.%m.%Y %H:%M")

    print("Git: Änderungen werden hochgeladen ...")
    schritte = [
        (["add", str(JSON_FILE)],            "add"),
        (["commit", "-m",
          f"Praxisdaten aktualisiert ({zeitstempel})"], "commit"),
        (["push"],                            "push"),
    ]

    for args, name in schritte:
        ergebnis = git(args)
        if ergebnis.returncode != 0:
            print(f"✗ Git-{name} fehlgeschlagen:")
            print(ergebnis.stderr.strip() or ergebnis.stdout.strip())
            print("  Bitte manuell pushen.")
            return False
        print(f"  ✓ git {name}")

    print("✓ Website wird in ~60 Sekunden aktuell sein.")
    return True


def aktualisiere():
    if not EXCEL_FILE.exists():
        print(f"✗ {EXCEL_FILE.name} nicht gefunden.")
        print("  Starte Script erneut ohne Argumente, um sie zu erstellen.")
        return False

    print(f"Lese {EXCEL_FILE.name} ...")
    wb = load_workbook(EXCEL_FILE, data_only=True)

    with open(JSON_FILE, encoding="utf-8") as f:
        daten = json.load(f)

    daten["oeffnungszeiten"]["tage"] = lese_oeffnungszeiten(get_sheet(wb, "ffnungszeiten"))
    daten["schliesstage"]            = lese_schliesstage(get_sheet(wb, "chlie"))
    daten["vertretung"]["aerzte"]    = lese_vertretung(get_sheet(wb, "ertretung"), daten["vertretung"]["aerzte"])

    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(daten, f, ensure_ascii=False, indent=2)

    print(f"✓ {JSON_FILE.name} wurde erfolgreich aktualisiert.")
    git_push()
    return True


# ═══════════════════════════════════════════════════════════════
#  EXCEL ERSTELLEN (JSON → EXCEL)
# ═══════════════════════════════════════════════════════════════

def _font(bold=False, size=10, color="000000", italic=False):
    return Font(name="Arial", bold=bold, size=size, color=color, italic=italic)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

C  = Alignment(horizontal="center", vertical="center")
CW = Alignment(horizontal="center", vertical="center", wrap_text=True)
L  = Alignment(horizontal="left",   vertical="center")

def titel_zeile(ws, text, ncols, bg):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=text)
    c.font      = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    c.fill      = _fill(bg)
    c.alignment = C
    ws.row_dimensions[1].height = 30

def hinweis_zeile(ws, text, ncols):
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(row=2, column=1, value=text)
    c.font      = _font(italic=True, size=9, color="888888")
    c.alignment = L
    ws.row_dimensions[2].height = 16

def kopf(ws, row, col, text):
    c = ws.cell(row=row, column=col, value=text)
    c.font      = _font(bold=True)
    c.fill      = _fill("EEEEEE")
    c.alignment = CW
    c.border    = _border()
    return c

def eingabe(ws, row, col, val, fmt=None, readonly=False):
    c = ws.cell(row=row, column=col, value=val)
    c.font      = _font(size=10)
    c.fill      = _fill("F5F5F5") if readonly else _fill("FFFDE7")
    c.alignment = C
    c.border    = _border()
    if fmt:
        c.number_format = fmt
    return c

def erstelle_excel(daten):
    wb = Workbook()

    # ── Sheet 1: Öffnungszeiten ──────────────────────────────────
    ws1 = wb.active
    ws1.title = "Öffnungszeiten"

    titel_zeile(ws1, "ÖFFNUNGSZEITEN", 9, "2E7D32")
    hinweis_zeile(ws1,
        "Zeiten im Format HH:MM  ·  Kein Nachmittag? → Zelle leer lassen", 9)

    # Abschnittsköpfe (Zeile 3)
    ws1.merge_cells(start_row=3, start_column=2, end_row=3, end_column=5)
    c = ws1.cell(row=3, column=2, value="Gesundsprechstunde")
    c.font = _font(bold=True, color="1B5E20"); c.fill = _fill("C8E6C9")
    c.alignment = C; c.border = _border()

    ws1.merge_cells(start_row=3, start_column=6, end_row=3, end_column=9)
    c = ws1.cell(row=3, column=6, value="Kranksprechstunde")
    c.font = _font(bold=True, color="B71C1C"); c.fill = _fill("FFCDD2")
    c.alignment = C; c.border = _border()

    c = ws1.cell(row=3, column=1)
    c.fill = _fill("EEEEEE"); c.border = _border()
    ws1.row_dimensions[3].height = 20

    # Spaltenköpfe (Zeile 4)
    for col, h in enumerate(["Wochentag",
            "von", "bis", "Nachmittag von", "Nachmittag bis",
            "von", "bis", "Nachmittag von", "Nachmittag bis"], 1):
        kopf(ws1, 4, col, h)
    ws1.row_dimensions[4].height = 20

    # Datenzeilen (5–9)
    tage = daten["oeffnungszeiten"]["tage"]
    for i, tag in enumerate(tage):
        row = 5 + i
        g, k = tag["gesund"], tag["krank"]
        c = ws1.cell(row=row, column=1, value=tag["tag"])
        c.font = _font(bold=True); c.fill = _fill("F5F5F5")
        c.alignment = C; c.border = _border()
        for col, val in enumerate([
            g.get("von") or "", g.get("bis") or "",
            g.get("nachmittag_von") or "", g.get("nachmittag_bis") or "",
            k.get("von") or "", k.get("bis") or "",
            k.get("nachmittag_von") or "", k.get("nachmittag_bis") or "",
        ], 2):
            eingabe(ws1, row, col, val, fmt="@")
        ws1.row_dimensions[row].height = 22

    ws1.column_dimensions["A"].width = 14
    for col in list("BCDEFGHI"):
        ws1.column_dimensions[col].width = 15
    ws1.freeze_panes = "B5"

    # ── Sheet 2: Schließzeiten ───────────────────────────────────
    ws2 = wb.create_sheet("Schließzeiten")
    titel_zeile(ws2, "SCHLIESS- & FERIENZEITEN", 4, "C62828")
    hinweis_zeile(ws2,
        "Datum: TT.MM.JJJJ  ·  Vorlauf = Tage, ab denen der Hinweis-Banner erscheint  ·  Leere Zeilen werden ignoriert",
        4)

    for col, h in enumerate(["Von", "Bis", "Grund", "Vorlauf (Tage)"], 1):
        kopf(ws2, 3, col, h)
    ws2.row_dimensions[3].height = 20

    schliesstage = daten.get("schliesstage", [])
    for i in range(15):
        row = 4 + i
        if i < len(schliesstage):
            st = schliesstage[i]
            try:
                von_dt = datetime.strptime(st["von"], "%Y-%m-%d")
                bis_dt = datetime.strptime(st["bis"], "%Y-%m-%d")
            except Exception:
                von_dt, bis_dt = "", ""
            vals = [von_dt, bis_dt, st.get("grund", ""), st.get("vorlauf_tage", 7)]
            fmts = ["DD.MM.YYYY", "DD.MM.YYYY", "@", "0"]
        else:
            vals = ["", "", "", ""]
            fmts = ["DD.MM.YYYY", "DD.MM.YYYY", "@", "0"]
        for col, (val, fmt) in enumerate(zip(vals, fmts), 1):
            eingabe(ws2, row, col, val, fmt=fmt)
        ws2.row_dimensions[row].height = 22

    ws2.column_dimensions["A"].width = 16
    ws2.column_dimensions["B"].width = 16
    ws2.column_dimensions["C"].width = 34
    ws2.column_dimensions["D"].width = 16
    ws2.freeze_panes = "A4"

    # ── Sheet 3: Vertretung ──────────────────────────────────────
    ws3 = wb.create_sheet("Vertretung")
    titel_zeile(ws3, "VERTRETUNGSÄRZTE", 6, "1565C0")
    hinweis_zeile(ws3,
        'Aktiv → "JA": Arzt erscheint auf der Website  ·  "NEIN": ausgeblendet  ·  Alle anderen Felder nur zur Ansicht',
        6)

    for col, h in enumerate(["ID", "Aktiv", "Name", "Fachrichtung", "Adresse", "Telefon"], 1):
        kopf(ws3, 3, col, h)
    ws3.row_dimensions[3].height = 20

    # Dropdown-Validierung für Aktiv-Spalte
    dv = DataValidation(type="list", formula1='"JA,NEIN"', allow_blank=False)
    dv.error      = 'Bitte "JA" oder "NEIN" wählen.'
    dv.errorTitle = "Ungültige Eingabe"
    ws3.add_data_validation(dv)

    aerzte = daten["vertretung"]["aerzte"]
    for i, arzt in enumerate(aerzte):
        row = 4 + i
        aktiv = "JA" if arzt.get("aktiv", True) else "NEIN"

        # Versteckte ID-Spalte
        c = ws3.cell(row=row, column=1, value=arzt["id"])
        c.font = _font(size=9, color="AAAAAA")
        c.fill = _fill("F0F0F0"); c.border = _border()

        # Aktiv-Zelle mit Dropdown
        c = ws3.cell(row=row, column=2, value=aktiv)
        c.font      = _font(bold=True, color="1B5E20" if aktiv == "JA" else "B71C1C")
        c.fill      = _fill("E8F5E9") if aktiv == "JA" else _fill("FFEBEE")
        c.alignment = C; c.border = _border()
        dv.add(c)

        # Read-only Infospalten
        for col, val in enumerate([
            arzt["name"], arzt.get("fachrichtung",""),
            arzt.get("adresse",""), arzt.get("telefon","")
        ], 3):
            eingabe(ws3, row, col, val, readonly=True).alignment = L
        ws3.row_dimensions[row].height = 22

    ws3.column_dimensions["A"].width = 12
    ws3.column_dimensions["B"].width = 10
    ws3.column_dimensions["C"].width = 28
    ws3.column_dimensions["D"].width = 36
    ws3.column_dimensions["E"].width = 40
    ws3.column_dimensions["F"].width = 18
    ws3.column_dimensions["A"].hidden = True
    ws3.freeze_panes = "B4"

    wb.save(EXCEL_FILE)
    print(f"✓ {EXCEL_FILE.name} wurde erstellt.")


# ═══════════════════════════════════════════════════════════════
#  EINSTIEGSPUNKT
# ═══════════════════════════════════════════════════════════════

def main():
    neu_erstellen = "--neu" in sys.argv

    with open(JSON_FILE, encoding="utf-8") as f:
        daten = json.load(f)

    if neu_erstellen:
        print("Erstelle praxis-editor.xlsx neu aus praxis-daten.json ...")
        erstelle_excel(daten)
        print("Fertig. Bitte Excel öffnen, prüfen und danach Script normal starten.")
        return True

    if not EXCEL_FILE.exists():
        print(f"{EXCEL_FILE.name} nicht gefunden – wird jetzt erstellt ...")
        erstelle_excel(daten)
        print()
        print("Bitte praxis-editor.xlsx jetzt bearbeiten und das Script")
        print("danach erneut starten, um die JSON zu aktualisieren.")
        return True

    return aktualisiere()


if __name__ == "__main__":
    try:
        ok = main()
    except Exception as e:
        print(f"✗ Fehler: {e}")
        ok = False
    sys.exit(0 if ok else 1)
