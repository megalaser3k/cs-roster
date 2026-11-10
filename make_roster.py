"""Generates a formula-driven Excel roster planner for German teams.

Sheets:
  Config     - year input, FO strength, employee list with Bundesland, legend
  Feiertage  - German public holidays for the configured year, per Bundesland
  Urlaub     - vacation periods per employee (start/end dates)
  Calc       - hidden helper: FO assignees, bases, per-day absence matrix
  Roster     - names row + Excel table: one row per working day (Mon-Fri),
               one column per employee; FO/BO/frei/Urlaub/holiday by formula
  Auswertung - per-employee totals and hours
"""
import datetime
import re
import shutil
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo, TableColumn
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.workbook.defined_name import DefinedName

STATES = ["BW", "BY", "BE", "BB", "HB", "HH", "HE", "MV",
          "NI", "NW", "RP", "SL", "SN", "ST", "SH", "TH"]

MAX_EMP = 60
FIRST_EMP_ROW = 6
LAST_EMP_ROW = FIRST_EMP_ROW + MAX_EMP - 1
ROSTER_ROWS = 263          # max Mon-Fri days in a year is 262; one spare row
DATA_ROW0 = 3              # first data row on Roster (row 1 = names, row 2 = header)
URLAUB_ROWS = 200          # max vacation entries
MAX_FO = 8                 # max simultaneous front-office persons (Config!E4)
FO_DEPTH = 8               # absence fallback chain length

ALL = set(STATES)
HOLIDAYS = [
    # (label, date formula, short name, applicable states)
    ("Neujahr",                 "=DATE(Config!$B$3,1,1)",     "Neujahr",        ALL),
    ("Heilige Drei Könige",     "=DATE(Config!$B$3,1,6)",     "Dreikönige",     {"BW", "BY", "ST"}),
    ("Internationaler Frauentag", "=DATE(Config!$B$3,3,8)",   "Frauentag",      {"BE", "MV"}),
    ("Karfreitag",              "=$U$13-2",                   "Karfreitag",     ALL),
    ("Ostermontag",             "=$U$13+1",                   "Ostermontag",    ALL),
    ("Tag der Arbeit",          "=DATE(Config!$B$3,5,1)",     "1. Mai",         ALL),
    ("Christi Himmelfahrt",     "=$U$13+39",                  "Himmelfahrt",    ALL),
    ("Pfingstmontag",           "=$U$13+50",                  "Pfingstmontag",  ALL),
    ("Fronleichnam",            "=$U$13+60",                  "Fronleichnam",   {"BW", "BY", "HE", "NW", "RP", "SL"}),
    ("Mariä Himmelfahrt",       "=DATE(Config!$B$3,8,15)",    "Mariä Himmelfahrt", {"BY", "SL"}),
    ("Tag der Deutschen Einheit", "=DATE(Config!$B$3,10,3)",  "Dt. Einheit",    ALL),
    ("Reformationstag",         "=DATE(Config!$B$3,10,31)",   "Reformationstag", {"BB", "HB", "HH", "MV", "NI", "SN", "ST", "SH", "TH"}),
    ("Allerheiligen",           "=DATE(Config!$B$3,11,1)",    "Allerheiligen",  {"BW", "BY", "NW", "RP", "SL"}),
    ("Buß- und Bettag",         "=DATE(Config!$B$3,11,22)-WEEKDAY(DATE(Config!$B$3,11,22),3)+2", "Buß-/Bettag", {"SN"}),
    ("1. Weihnachtstag",        "=DATE(Config!$B$3,12,25)",   "Weihnachten",    ALL),
    ("2. Weihnachtstag",        "=DATE(Config!$B$3,12,26)",   "2. Weihnachtstag", ALL),
]

HEADER_FILL = PatternFill("solid", fgColor="FF1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14)
THIN = Border(*[Side(style="thin", color="FFBFBFBF")] * 4)

# dxf fills: Excel reads the solid fill colour from bgColor, alpha must be FF
FILL_FO = PatternFill("solid", fgColor="FFBDD7EE", bgColor="FFBDD7EE")       # blue
FILL_BO = PatternFill("solid", fgColor="FFC6EFCE", bgColor="FFC6EFCE")       # green
FILL_FREI = PatternFill("solid", fgColor="FFD9D9D9", bgColor="FFD9D9D9")     # grey
FILL_URLAUB = PatternFill("solid", fgColor="FFFFEB9C", bgColor="FFFFEB9C")   # yellow
FILL_HOLIDAY = PatternFill("solid", fgColor="FFFFC7CE", bgColor="FFFFC7CE")  # red
FILL_WEEK = PatternFill("solid", fgColor="FFEDEDED", bgColor="FFEDEDED")     # week band
FILL_INPUT = PatternFill("solid", fgColor="FFFFF2CC", bgColor="FFFFF2CC")    # editable input cells


def unlock(ws, ref):
    """Mark input cells as editable (unlocked) and colour them yellow."""
    cells = ws[ref]
    if not isinstance(cells, tuple):
        cells = ((cells,),)
    for row in cells:
        for c in row:
            c.protection = Protection(locked=False)
            c.fill = FILL_INPUT


def protect(ws):
    """Lock the sheet (formulas read-only) but keep table filter/sort usable."""
    ws.protection.sheet = True
    ws.protection.autoFilter = False
    ws.protection.sort = False


def build(path):
    wb = Workbook()

    # ---------------- Config ----------------
    cfg = wb.active
    cfg.title = "Config"
    cfg["A1"] = "Roster-Konfiguration"
    cfg["A1"].font = TITLE_FONT
    cfg["A3"] = "Jahr"
    cfg["B3"] = datetime.date.today().year + 1
    cfg["B3"].number_format = "0"
    cfg["A3"].font = Font(bold=True)
    cfg["D3"] = "Anzahl Mitarbeiter"
    cfg["D3"].font = Font(bold=True)
    cfg["E3"] = f"=COUNTA(A{FIRST_EMP_ROW}:A{LAST_EMP_ROW})"
    cfg["D4"] = "Personen im Front-Office"
    cfg["D4"].font = Font(bold=True)
    cfg["E4"] = 1
    cfg["E4"].number_format = "0"
    dv_fo = DataValidation(type="whole", operator="between",
                           formula1="1", formula2=str(MAX_FO), allow_blank=False)
    cfg.add_data_validation(dv_fo)
    dv_fo.add("E4")
    cfg["D5"] = "FO-Plausibilität"
    cfg["D5"].font = Font(bold=True)
    cfg["E5"] = ('=IF(E3=0,"-",IF(2*E4>E3,'
                 '"HINWEIS: 2x FO-Staerke > Mitarbeiter - FO-Abdeckung evtl. reduziert",'
                 '"OK"))')

    cfg["A5"] = "Name"
    cfg["B5"] = "Bundesland"
    for c in ("A5", "B5"):
        cfg[c].fill = HEADER_FILL
        cfg[c].font = HEADER_FONT
    dv = DataValidation(type="list", formula1='"' + ",".join(STATES) + '"',
                        allow_blank=True)
    cfg.add_data_validation(dv)
    dv.add(f"B{FIRST_EMP_ROW}:B{LAST_EMP_ROW}")

    wb.defined_names["MitarbeiterNamen"] = DefinedName(
        "MitarbeiterNamen", attr_text=f"Config!$A${FIRST_EMP_ROW}:$A${LAST_EMP_ROW}")

    legend = [
        "Legende:",
        "FO  = Front-Office (Mo-Do je 10 h, Fr 8 h) - blau",
        "BO  = Back-Office (8 h) - grün",
        "frei = Ausgleichs-Freitag nach kompletter FO-Woche (Mo-Do) - grau",
        "(eine reine Freitag-Vertretung gibt keinen frei-Tag),",
        "Urlaub = Urlaubstag (Zeiträume auf Blatt 'Urlaub' eintragen) - gelb",
        "Feiertagsname = arbeitsfrei (Feiertag im eigenen Bundesland) - rot",
        "Jede zweite Woche ist im Roster hellgrau unterlegt (Wochenwechsel).",
        "",
        "Jedes/r Mitarbeiter/in ist alle N/K Wochen im Front-Office",
        "(N = Anzahl Mitarbeiter, K = FO-Staerke in E4): pro Woche eine",
        "Gruppe von K Personen - Freitag der Start, dann Mo-Do der",
        "Folgewoche, der folgende Freitag ist frei (10-h-Tage Mo-Do).",
        "Niemand arbeitet zwei FO-Wochen hintereinander: die Vertretung",
        "(Urlaub/Feiertag) geht an die naechste freie Person, aber nie an",
        "jemanden, der in der Vorwoche schon FO hatte. Ist niemand frei,",
        "bleibt FO an dem Tag unbesetzt.",
        "",
        "Das Roster-Blatt ist eine Tabelle: Filtern über die Kopfzeile",
        "(Zeile 2). Zeile 1 zeigt die Namen aus dieser Config-Tabelle;",
        "Spalte D = 1. Person aus Config, Spalte E = 2. Person usw.",
        "",
        "Front-Office wird automatisch auf die naechste verfuegbare Person",
        "verschoben, wenn die geplante Person Urlaub/Feiertag hat.",
        "",
        "Zukunftssicher: Zum Jahreswechsel nur das Jahr in B3 ändern.",
        "",
        "Alle Blätter sind schreibgeschützt; nur die gelb markierten",
        "Eingabefelder (Jahr, FO-Staerke, Namen/Bundesländer, Urlaub)",
        "sind bearbeitbar. Schutz ohne Passwort (Blatt entsperren via",
        "Überprüfen > Blattschutz aufheben, falls nötig).",
    ]
    for i, line in enumerate(legend, start=7):
        cfg[f"D{i}"] = line
    cfg.column_dimensions["A"].width = 28
    cfg.column_dimensions["B"].width = 14
    cfg.column_dimensions["D"].width = 62
    unlock(cfg, "B3")
    unlock(cfg, "E4")
    unlock(cfg, f"A{FIRST_EMP_ROW}:B{LAST_EMP_ROW}")
    protect(cfg)

    # ---------------- Feiertage ----------------
    ft = wb.create_sheet("Feiertage")
    ft["A1"] = "Datum"
    ft["R1"] = "Feiertag"
    for j, st in enumerate(STATES):
        c = ft.cell(row=1, column=2 + j, value=st)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")
    for c in (ft["A1"], ft["R1"]):
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
    for i, (label, formula, short, states) in enumerate(HOLIDAYS):
        r = 2 + i
        a = ft.cell(row=r, column=1, value=formula)
        a.number_format = "DD.MM.YYYY"
        ft.cell(row=r, column=18, value=label)
        for j, st in enumerate(STATES):
            cell = ft.cell(row=r, column=2 + j,
                           value=short if st in states else None)
            cell.alignment = Alignment(horizontal="center")
    ft.column_dimensions["A"].width = 12
    ft.column_dimensions["R"].width = 26
    for j in range(len(STATES)):
        ft.column_dimensions[get_column_letter(2 + j)].width = 13

    # Easter helpers (Gauss algorithm) in column T/U
    helpers = [
        ("Oster-Hilfsberechnung (Gauss)", None),
        ("Jahr", "=Config!$B$3"),
        ("a = Jahr mod 19", "=MOD(U2,19)"),
        ("b = Jahr mod 4", "=MOD(U2,4)"),
        ("c = Jahr mod 7", "=MOD(U2,7)"),
        ("k = Jahr div 100", "=INT(U2/100)"),
        ("p", "=INT((13+8*U6)/25)"),
        ("q", "=INT(U6/4)"),
        ("M", "=MOD(15-U7+U6-U8,30)"),
        ("N", "=MOD(4+U6-U8,7)"),
        ("d", "=MOD(19*U3+U9,30)"),
        ("e", "=MOD(2*U4+4*U5+6*U11+U10,7)"),
        ("Ostersonntag", "=DATE(U2,3,22)+U11+U12"),
    ]
    for i, (label, formula) in enumerate(helpers, start=1):
        ft.cell(row=i, column=20, value=label)
        if formula:
            cell = ft.cell(row=i, column=21, value=formula)
            if label == "Ostersonntag":
                cell.number_format = "DD.MM.YYYY"
    ft.column_dimensions["T"].width = 26
    ft.freeze_panes = "B2"
    protect(ft)

    # ---------------- Urlaub ----------------
    ur = wb.create_sheet("Urlaub")
    for j, h in enumerate(["Mitarbeiter", "Von", "Bis"], start=1):
        c = ur.cell(row=1, column=j, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
    for r in range(2, 2 + URLAUB_ROWS):
        ur.cell(row=r, column=2).number_format = "DD.MM.YYYY"
        ur.cell(row=r, column=3).number_format = "DD.MM.YYYY"
    dv2 = DataValidation(type="list", formula1="MitarbeiterNamen", allow_blank=True)
    ur.add_data_validation(dv2)
    dv2.add(f"A2:A{1 + URLAUB_ROWS}")
    ur["E1"] = "Mehrere Urlaube pro Person: einfach mehrere Zeilen mit demselben Namen eintragen."
    ur["E1"].font = Font(italic=True)
    ur.column_dimensions["A"].width = 28
    ur.column_dimensions["B"].width = 12
    ur.column_dimensions["C"].width = 12
    ur.freeze_panes = "A2"
    unlock(ur, f"A2:C{1 + URLAUB_ROWS}")
    protect(ur)

    # ------------- Calc (hidden: FO assignees, bases, absence/blocked matrix) -------------
    # Week anchor DATE(2000,1,3) is a Monday in both Excel's and LibreOffice's
    # serial systems, so INT((date-anchor)/7) increments exactly on Mondays.
    # Calc row r (2..264) mirrors Roster data row r+1 (3..265).
    # Columns: A..H = FO assignee 1..MAX_FO, I = FO base,
    # K..BR = absence flags (vacation/holiday), BS.. = blocked flags
    # (absent OR wasFOprev), EA = MonRow helper, EB.. = wasFOprev flags
    # (FO in previous duty block), then MT.. = wasFO Mon-Thu of current week
    # (used for the off-Friday: only a full FO week earns the Friday off)
    ca = wb.create_sheet("Calc")
    ABSENT0 = 11            # column K
    BLOCK0 = ABSENT0 + MAX_EMP  # column BS
    MONROW = BLOCK0 + MAX_EMP   # column EA
    WASFO0 = MONROW + 1          # column EB
    MT0 = WASFO0 + MAX_EMP       # was-FO-Mon-Thu matrix
    absent_last = get_column_letter(ABSENT0 + MAX_EMP - 1)
    block_col = get_column_letter(BLOCK0)
    block_last = get_column_letter(BLOCK0 + MAX_EMP - 1)
    wasfo_col = get_column_letter(WASFO0)
    fo_cols = [get_column_letter(1 + j) for j in range(MAX_FO)]
    for r in range(2, 2 + ROSTER_ROWS):
        rr = r + 1  # roster data row
        for i in range(MAX_EMP):
            cr = FIRST_EMP_ROW + i
            ca.cell(row=r, column=ABSENT0 + i, value=(
                f'=IF(Roster!$A{rr}="","",OR('
                f'COUNTIFS(Urlaub!$A$2:$A${1 + URLAUB_ROWS},Config!$A${cr},'
                f'Urlaub!$B$2:$B${1 + URLAUB_ROWS},"<="&Roster!$A{rr},'
                f'Urlaub!$C$2:$C${1 + URLAUB_ROWS},">="&Roster!$A{rr})>0,'
                f'IFERROR(INDEX(Feiertage!$B$2:$Q$17,MATCH(Roster!$A{rr},Feiertage!$A$2:$A$17,0),'
                f'MATCH(Config!$B${cr},Feiertage!$B$1:$Q$1,0)),"")<>""))'))
        w = f'INT((Roster!$A{rr}-DATE(2000,1,3))/7)'
        blockw = f'({w}-IF(WEEKDAY(Roster!$A{rr},2)=5,0,1))'
        fobase = f'MOD({blockw}*Config!$E$4,Config!$E$3)'
        ca.cell(row=r, column=9, value=f'=IF(Roster!$A{rr}="","",{fobase})')
        mc = f'${get_column_letter(MONROW)}{r}'
        ca.cell(row=r, column=MONROW,
                value=f'=IF(Roster!$A{rr}="","",ROW()-WEEKDAY(Roster!$A{rr},2)+1)')
        # blocked = absent OR actually assigned FO on any day of the previous
        # duty block. Previous block rows (Calc rows): on a Friday -> r-5..r-1;
        # on Mon-Thu -> MonRow-6..MonRow-2 (via INDEX, clamped to valid rows).
        wd = f'WEEKDAY(Roster!$A{rr},2)'
        for i in range(MAX_EMP):
            terms = []
            for col in fo_cols:
                # INDEX range ends at r-1 so the formula never references its
                # own row (avoids circular-reference detection in Excel)
                fri = [f'INDEX(${col}$2:${col}${r - 1},{p - 1})={i}'
                       for p in (r - 5, r - 4, r - 3, r - 2, r - 1) if p >= 2]
                mon = [f'IF(AND({mc}-{d}-1>=1,{mc}-{d}-1<={r - 2}),'
                       f'INDEX(${col}$2:${col}${r - 1},{mc}-{d}-1)={i},FALSE)'
                       for d in (6, 5, 4, 3, 2)] if r >= 3 else []
                if fri and mon:
                    terms.append(f'IF({wd}=5,OR({",".join(fri)}),OR({",".join(mon)}))')
                elif fri:
                    terms.append(f'IF({wd}=5,OR({",".join(fri)}),FALSE)')
                elif mon:
                    terms.append(f'IF({wd}<5,OR({",".join(mon)}),FALSE)')
            was_prev = "OR(" + ",".join(terms) + ")" if terms else "FALSE"
            ca.cell(row=r, column=WASFO0 + i, value=(
                f'=IF(Roster!$A{rr}="","",{was_prev})'))
            ca.cell(row=r, column=BLOCK0 + i, value=(
                f'=IF(Roster!$A{rr}="","",OR({get_column_letter(ABSENT0 + i)}{r},'
                f'{get_column_letter(WASFO0 + i)}{r}))'))
            # Mon-Thu-only variant for the off-Friday (Friday rows: r-4..r-1)
            mt = [f'INDEX(${col}$2:${col}${r - 1},{p - 1})={i}'
                  for col in fo_cols for p in (r - 4, r - 3, r - 2, r - 1) if p >= 2]
            mt = "OR(" + ",".join(mt) + ")" if mt else "FALSE"
            ca.cell(row=r, column=MT0 + i, value=(
                f'=IF(OR(Roster!$A{rr}="",{wd}<>5),"",{mt})'))
        for j in range(1, MAX_FO + 1):
            if j == 1:
                start = f'MOD($I{r},Config!$E$3)'
                guard = f'OR(Roster!$A{rr}="",Config!$E$4<{j})'
            else:
                prevcell = f'{fo_cols[j - 2]}{r}'
                start = f'MOD(IF({prevcell}="",0,{prevcell})+1,Config!$E$3)'
                guard = f'OR(Roster!$A{rr}="",Config!$E$4<{j},{prevcell}="")'

            def blocked(k, start=start, r=r):
                return (f'INDEX(${block_col}{r}:${block_last}{r},1,'
                        f'MOD({start}+{k},Config!$E$3)+1)')

            def assigned(k, start=start, r=r, j=j):
                terms = [f'${c}{r}=MOD({start}+{k},Config!$E$3)' for c in fo_cols[:j - 1]]
                return "OR(" + ",".join(terms) + ")" if terms else "FALSE"

            chain = '""'
            for k in range(FO_DEPTH - 1, -1, -1):
                chain = (f'IF(NOT(OR({blocked(k)},{assigned(k)})),'
                         f'MOD({start}+{k},Config!$E$3),{chain})')
            ca.cell(row=r, column=j,
                    value=f'=IF({guard},"",{chain})')
    ca.sheet_state = "hidden"
    protect(ca)

    # ---------------- Roster ----------------
    ro = wb.create_sheet("Roster")
    # row 1: employee names (outside the table, so table headers stay static)
    ro["A1"] = "Namen aus Config ->"
    ro["A1"].font = Font(italic=True)
    for i in range(MAX_EMP):
        c = ro.cell(row=1, column=4 + i,
                    value=f'=IF(Config!A{FIRST_EMP_ROW + i}="","",Config!A{FIRST_EMP_ROW + i})')
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")
    # row 2: table header
    headers = ["Datum", "Tag", "KW"] + [f"Mitarbeiter {i + 1}" for i in range(MAX_EMP)]
    for j, h in enumerate(headers, start=1):
        c = ro.cell(row=2, column=j, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")

    last_data_row = DATA_ROW0 + ROSTER_ROWS - 1
    for r in range(DATA_ROW0, last_data_row + 1):
        cr_calc = r - 1  # Calc row
        if r == DATA_ROW0:
            ro.cell(row=r, column=1,
                    value="=DATE(Config!$B$3,1,1)+IF(WEEKDAY(DATE(Config!$B$3,1,1),2)>5,8-WEEKDAY(DATE(Config!$B$3,1,1),2),0)")
        else:
            prev = f"$A{r - 1}"
            ro.cell(row=r, column=1,
                    value=(f'=IF({prev}="","",IF({prev}>=DATE(Config!$B$3,12,31),"",'
                           f'IF(WEEKDAY({prev},2)=5,{prev}+3,{prev}+1)))'))
        ro.cell(row=r, column=1).number_format = "DD.MM.YYYY"
        ro.cell(row=r, column=2,
                value=f'=IF($A{r}="","",CHOOSE(WEEKDAY($A{r},2),"Mo","Di","Mi","Do","Fr","Sa","So"))')
        ro.cell(row=r, column=3, value=f'=IF($A{r}="","",WEEKNUM($A{r},21))')

        for i in range(MAX_EMP):
            cr = FIRST_EMP_ROW + i
            h = (f'IFERROR(INDEX(Feiertage!$B$2:$Q$17,MATCH($A{r},Feiertage!$A$2:$A$17,0),'
                 f'MATCH(Config!$B${cr},Feiertage!$B$1:$Q$1,0)),"")')
            vac = (f'COUNTIFS(Urlaub!$A$2:$A${1 + URLAUB_ROWS},Config!$A${cr},'
                   f'Urlaub!$B$2:$B${1 + URLAUB_ROWS},"<="&$A{r},'
                   f'Urlaub!$C$2:$C${1 + URLAUB_ROWS},">="&$A{r})>0')
            fo_match = "OR(" + ",".join(
                f"Calc!${get_column_letter(1 + j)}{cr_calc}={i}" for j in range(MAX_FO)) + ")"
            formula = (
                f'=IF(OR($A{r}="",Config!$E$3=0,Config!$A${cr}=""),"",'
                f'IF({vac},"Urlaub",'
                f'IF({h}<>"",{h},'
                f'IF(AND(WEEKDAY($A{r},2)=5,Calc!${get_column_letter(MT0 + i)}{cr_calc}=TRUE),"frei",'
                f'IF({fo_match},"FO","BO")))))'
            )
            cell = ro.cell(row=r, column=4 + i, value=formula)
            cell.alignment = Alignment(horizontal="center")
            cell.border = THIN

    last_col = get_column_letter(3 + MAX_EMP)
    table_ref = f"A2:{last_col}{last_data_row}"
    cf_range = f"D{DATA_ROW0}:{last_col}{last_data_row}"
    ro.conditional_formatting.add(cf_range, CellIsRule(
        operator="equal", formula=['"FO"'], fill=FILL_FO, stopIfTrue=True))
    ro.conditional_formatting.add(cf_range, CellIsRule(
        operator="equal", formula=['"BO"'], fill=FILL_BO, stopIfTrue=True))
    ro.conditional_formatting.add(cf_range, CellIsRule(
        operator="equal", formula=['"frei"'], fill=FILL_FREI, stopIfTrue=True))
    ro.conditional_formatting.add(cf_range, CellIsRule(
        operator="equal", formula=['"Urlaub"'], fill=FILL_URLAUB, stopIfTrue=True))
    ro.conditional_formatting.add(cf_range, FormulaRule(
        formula=[f'AND(D{DATA_ROW0}<>"",D{DATA_ROW0}<>"FO",D{DATA_ROW0}<>"BO",D{DATA_ROW0}<>"frei",D{DATA_ROW0}<>"Urlaub")'],
        fill=FILL_HOLIDAY, stopIfTrue=True))
    # week banding across the whole row (lowest priority; status colours win)
    ro.conditional_formatting.add(
        f"A{DATA_ROW0}:{last_col}{last_data_row}",
        FormulaRule(formula=[f'MOD(INT(($A{DATA_ROW0}-DATE(2000,1,3))/7),2)=1'],
                    fill=FILL_WEEK))

    tab = Table(displayName="RosterTabelle", ref=table_ref,
                tableColumns=[TableColumn(id=j + 1, name=h)
                              for j, h in enumerate(headers)])
    tab.autoFilter = AutoFilter(ref=table_ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=False)
    ro.add_table(tab)

    ro.column_dimensions["A"].width = 12
    ro.column_dimensions["B"].width = 6
    ro.column_dimensions["C"].width = 6
    for i in range(MAX_EMP):
        ro.column_dimensions[get_column_letter(4 + i)].width = 12
    ro.freeze_panes = "D3"
    protect(ro)

    # ---------------- Auswertung ----------------
    aw = wb.create_sheet("Auswertung")
    headers_aw = ["Name", "Bundesland", "FO-Tage", "davon Freitag", "BO-Tage",
                  "Freitage", "Urlaub", "Feiertage", "Stunden"]
    for j, h in enumerate(headers_aw, start=1):
        c = aw.cell(row=1, column=j, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
    for i in range(MAX_EMP):
        r = 2 + i
        cr = FIRST_EMP_ROW + i
        col = get_column_letter(4 + i)
        rng = f"Roster!{col}${DATA_ROW0}:{col}${last_data_row}"
        aw.cell(row=r, column=1, value=f'=IF(Config!A{cr}="","",Config!A{cr})')
        aw.cell(row=r, column=2, value=f'=IF(Config!A{cr}="","",Config!B{cr})')
        aw.cell(row=r, column=3, value=f'=IF($A{r}="","",COUNTIF({rng},"FO"))')
        aw.cell(row=r, column=4,
                value=f'=IF($A{r}="","",COUNTIFS({rng},"FO",Roster!$B${DATA_ROW0}:$B${last_data_row},"Fr"))')
        aw.cell(row=r, column=5, value=f'=IF($A{r}="","",COUNTIF({rng},"BO"))')
        aw.cell(row=r, column=6, value=f'=IF($A{r}="","",COUNTIF({rng},"frei"))')
        aw.cell(row=r, column=7, value=f'=IF($A{r}="","",COUNTIF({rng},"Urlaub"))')
        aw.cell(row=r, column=8,
                value=(f'=IF($A{r}="","",COUNTIFS({rng},"<>FO",{rng},"<>BO",'
                       f'{rng},"<>frei",{rng},"<>Urlaub",{rng},"<>"))'))
        aw.cell(row=r, column=9, value=f'=IF($A{r}="","",10*(C{r}-D{r})+8*D{r}+8*E{r})')
    aw.column_dimensions["A"].width = 28
    aw.column_dimensions["B"].width = 12
    for j in range(3, 10):
        aw.column_dimensions[get_column_letter(j)].width = 13
    aw.freeze_panes = "A2"
    protect(aw)

    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    strip_empty_cached_values(path)
    check_formula_balance(path)
    print("saved", path)


def strip_empty_cached_values(path):
    """Remove empty <v/> elements that openpyxl writes after <f> formulas.

    A numeric cell with an empty cached <v/> is treated as corrupt by Excel,
    which then strips the formulas on open. With no cached values present,
    fullCalcOnLoad makes Excel recalculate everything when the file opens.
    """
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.match(r"xl/worksheets/sheet\d+\.xml$", item.filename):
                data = re.sub(rb"<v\s*/>|<v></v>", b"", data)
            zout.writestr(item, data)
    shutil.move(tmp, path)


def check_formula_balance(path):
    """Assert every stored formula has balanced parentheses (outside strings)."""
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if re.match(r"xl/worksheets/sheet\d+\.xml$", n):
                for f in re.findall(r"<f[^>]*>([^<]*)</f>", z.read(n).decode()):
                    f = f.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
                    depth = 0
                    instr = False
                    for ch in f:
                        if ch == '"':
                            instr = not instr
                        elif not instr:
                            depth += (ch == "(") - (ch == ")")
                    if depth != 0:
                        raise ValueError(f"unbalanced formula in {n}: {f[:80]}")


if __name__ == "__main__":
    build("/workspace/Dienstplan.xlsx")
