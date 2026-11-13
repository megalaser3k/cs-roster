"""Generates a formula-driven Excel roster planner for German teams.

Sheets:
  Konfiguration - year input, FO strength, employee list with Bundesland, legend
  Feiertage     - German public holidays for the configured year, per Bundesland
                  (incl. 24.12. Heiligabend and 31.12. Silvester for all states)
  Urlaub        - vacation periods per employee (start/end dates)
  Calc          - hidden helper: day engine (dates, FO assignees, absence/blocked)
  Dienstplan    - table: one row per (ISO week x employee), columns KW |
                  Mitarbeiter | Mo..Fr with FO/BO/frei/Urlaub/holiday
  Auswertung    - per-employee totals and hours
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
DAYS = 263              # max Mon-Fri days in a year is 262; one spare row
WEEKS = 54              # ISO weeks 0..53 (week 0 catches a Jan-1 Friday)
URLAUB_ROWS = 200       # max vacation entries
MAX_FO = 8              # max simultaneous front-office persons (Konfiguration!E4)
FO_DEPTH = 8            # absence fallback chain length

ALL = set(STATES)
HOLIDAYS = [
    # (label, date formula, short name, applicable states)
    ("Neujahr",                 "=DATE(JAHR,1,1)",    "Neujahr",        ALL),
    ("Heilige Drei Könige",     "=DATE(JAHR,1,6)",    "Dreikönige",     {"BW", "BY", "ST"}),
    ("Internationaler Frauentag", "=DATE(JAHR,3,8)",  "Frauentag",      {"BE", "MV"}),
    ("Karfreitag",              "=$U$13-2",          "Karfreitag",     ALL),
    ("Ostermontag",             "=$U$13+1",          "Ostermontag",    ALL),
    ("Tag der Arbeit",          "=DATE(JAHR,5,1)",    "1. Mai",         ALL),
    ("Christi Himmelfahrt",     "=$U$13+39",         "Himmelfahrt",    ALL),
    ("Pfingstmontag",           "=$U$13+50",         "Pfingstmontag",  ALL),
    ("Fronleichnam",            "=$U$13+60",         "Fronleichnam",   {"BW", "BY", "HE", "NW", "RP", "SL"}),
    ("Mariä Himmelfahrt",       "=DATE(JAHR,8,15)",   "Mariä Himmelfahrt", {"BY", "SL"}),
    ("Weltkindertag",           "=DATE(JAHR,9,20)",   "Weltkindertag",  {"TH"}),
    ("Tag der Deutschen Einheit", "=DATE(JAHR,10,3)", "Dt. Einheit",   ALL),
    ("Reformationstag",         "=DATE(JAHR,10,31)",  "Reformationstag", {"BB", "HB", "HH", "MV", "NI", "SN", "ST", "SH", "TH"}),
    ("Allerheiligen",           "=DATE(JAHR,11,1)",   "Allerheiligen",  {"BW", "BY", "NW", "RP", "SL"}),
    ("Buß- und Bettag",         "=DATE(JAHR,11,23)-(WEEKDAY(DATE(JAHR,11,23),1)-1)-4", "Buß-/Bettag", {"SN"}),
    ("Heiligabend",             "=DATE(JAHR,12,24)",  "Heiligabend",    ALL),  # betriebsfrei, kein gesetzl. Feiertag
    ("1. Weihnachtstag",        "=DATE(JAHR,12,25)",  "Weihnachten",    ALL),
    ("2. Weihnachtstag",        "=DATE(JAHR,12,26)",  "2. Weihnachtstag", ALL),
    ("Silvester",               "=DATE(JAHR,12,31)",  "Silvester",      ALL),  # betriebsfrei, kein gesetzl. Feiertag
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

    # ---------------- Konfiguration ----------------
    cfg = wb.active
    cfg.title = "Konfiguration"
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
        "MitarbeiterNamen", attr_text=f"Konfiguration!$A${FIRST_EMP_ROW}:$A${LAST_EMP_ROW}")
    for name, ref in [
        ("N_MA", "Konfiguration!$E$3"), ("K_FO", "Konfiguration!$E$4"),
        ("JAHR", "Konfiguration!$B$3"), ("ANKER", "Calc!$J$1"),
        ("URL_A", f"Urlaub!$A$2:$A${1 + URLAUB_ROWS}"),
        ("URL_B", f"Urlaub!$B$2:$B${1 + URLAUB_ROWS}"),
        ("URL_C", f"Urlaub!$C$2:$C${1 + URLAUB_ROWS}"),
        ("FT_A", "Feiertage!$A$2:$A$20"), ("FT_BQ", "Feiertage!$B$2:$Q$20"),
        ("FT_H", "Feiertage!$B$1:$Q$1"),
    ]:
        wb.defined_names[name] = DefinedName(name, attr_text=ref)

    legend = [
        "Legende:",
        "FO  = Front-Office (Mo-Do je 10 h, Fr 8 h) - blau",
        "BO  = Back-Office (8 h) - grün",
        "frei = Ausgleichs-Freitag nach kompletter FO-Woche - grau",
        "(nur bei FO an ALLEN vier Tagen Mo-Do; Teilwochen oder reine",
        "Freitags-Vertretungen geben keinen frei-Tag),",
        "Urlaub = Urlaubstag (Zeiträume auf Blatt 'Urlaub' eintragen) - gelb",
        "Feiertagsname = arbeitsfrei (Feiertag im eigenen Bundesland, - rot",
        "24.12. und 31.12. gelten betriebsfrei fuer alle)",
        "Jede zweite KW ist im Dienstplan hellgrau unterlegt.",
        "",
        "Blatt 'Dienstplan': eine Zeile pro (KW x Mitarbeiter),",
        "Spalten Mo-Fr. Tabelle ab Zeile 2 mit Filtern in der Kopfzeile.",
        "KW = ISO-Kalenderwoche (Woche 0 = evtl. 1.1., falls Freitag).",
        "",
        "Jedes/r Mitarbeiter/in ist alle N/K Wochen im Front-Office",
        "(N = Anzahl Mitarbeiter, K = FO-Staerke in E4): pro Woche eine",
        "Gruppe von K Personen - Freitag der Start, dann Mo-Do der",
        "Folgewoche, der folgende Freitag ist frei (10-h-Tage Mo-Do).",
        "",
        "Niemand arbeitet zwei FO-Wochen hintereinander: die Vertretung",
        "(Urlaub/Feiertag) geht an die naechste freie Person, aber nie an",
        "jemanden, der in der Vorwoche schon FO hatte. Ist niemand frei,",
        "bleibt FO an dem Tag unbesetzt.",
        "",
        "Zukunftssicher: Zum Jahreswechsel nur das Jahr in B3 ändern.",
        "",
        "Alle Blätter sind schreibgeschützt; nur die gelb markierten",
        "Eingabefelder (Jahr, FO-Staerke, Namen/Bundesländer, Urlaub)",
        "sind bearbeitbar. Schutz ohne Passwort (Überprüfen > Blattschutz",
        "aufheben, falls nötig).",
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
        ("Jahr", "=JAHR"),
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

    # ------------- Calc (hidden day engine; size-optimized) -------------
    # Week anchor: named cell ANKER = Calc!$J$1 = DATE(2000,1,3), a Monday in
    # both Excel's and LibreOffice's serial systems.
    # Calc row r (2..264) = the r-th working day (Mon-Fri) of the year.
    # Columns: DATECOL = date, A..H = FO assignees, I = FO base, J = FOSet,
    # K..BR = absence flags, BS.. = blocked flags, EA = MonRow helper,
    # EB.. = wasFOprev, GL.. = wasFO Mon-Thu (off-Friday)
    ca = wb.create_sheet("Calc")
    ABSENT0 = 11
    BLOCK0 = ABSENT0 + MAX_EMP
    MONROW = BLOCK0 + MAX_EMP
    WASFO0 = MONROW + 1
    MT0 = WASFO0 + MAX_EMP
    DATECOL = MT0 + MAX_EMP
    dl = get_column_letter(DATECOL)
    absent_last = get_column_letter(ABSENT0 + MAX_EMP - 1)
    block_col = get_column_letter(BLOCK0)
    block_last = get_column_letter(BLOCK0 + MAX_EMP - 1)
    mt_last = get_column_letter(MT0 + MAX_EMP - 1)
    fo_cols = [get_column_letter(1 + j) for j in range(MAX_FO)]
    ca["I1"] = "anker (Montag)"
    ca["J1"] = "=DATE(2000,1,3)"
    ca["J1"].number_format = "DD.MM.YYYY"
    wb.defined_names["CALCD"] = DefinedName("CALCD", attr_text=f"Calc!${dl}$2:${dl}${1 + DAYS}")
    for r in range(2, 2 + DAYS):
        dc = f'${dl}{r}'
        # date of this working day
        if r == 2:
            ca.cell(row=r, column=DATECOL,
                    value=f'=DATE(JAHR,1,1)+IF(WEEKDAY(DATE(JAHR,1,1),2)>5,8-WEEKDAY(DATE(JAHR,1,1),2),0)')
        else:
            prev = f'${dl}{r - 1}'
            ca.cell(row=r, column=DATECOL,
                    value=(f'=IF({prev}="","",IF({prev}>=DATE(JAHR,12,31),"",'
                           f'IF(WEEKDAY({prev},2)=5,{prev}+3,{prev}+1)))'))
        ca.cell(row=r, column=DATECOL).number_format = "DD.MM.YYYY"
        fr = f'$J$2:$J${r - 1}'
        w = f'INT(({dc}-ANKER)/7)'
        blockw = f'({w}-IF(WEEKDAY({dc},2)=5,0,1))'
        ca.cell(row=r, column=9,
                value=f'=IF({dc}="","",MOD({blockw}*K_FO,N_MA))')
        mc = f'${get_column_letter(MONROW)}{r}'
        ca.cell(row=r, column=MONROW,
                value=f'=IF({dc}="","",ROW()-WEEKDAY({dc},2)+1)')
        for j in range(1, MAX_FO + 1):
            if j == 1:
                start = f'MOD($I{r},N_MA)'
                guard = f'OR({dc}="",K_FO<{j})'
            else:
                prevcell = f'{fo_cols[j - 2]}{r}'
                start = f'MOD(IF({prevcell}="",0,{prevcell})+1,N_MA)'
                guard = f'OR({dc}="",K_FO<{j},{prevcell}="")'

            def blocked(k, start=start, r=r):
                return (f'INDEX(${block_col}{r}:${block_last}{r},1,'
                        f'MOD({start}+{k},N_MA)+1)')

            def assigned(k, start=start, r=r, j=j):
                terms = [f'${c}{r}=MOD({start}+{k},N_MA)' for c in fo_cols[:j - 1]]
                return "OR(" + ",".join(terms) + ")" if terms else "FALSE"

            chain = '""'
            for k in range(FO_DEPTH - 1, -1, -1):
                chain = (f'IF(NOT(OR({blocked(k)},{assigned(k)})),'
                         f'MOD({start}+{k},N_MA),{chain})')
            ca.cell(row=r, column=j, value=f'=IF({guard},"",{chain})')
        foset_sep = '&","&'
        ca.cell(row=r, column=10, value=(
            f'=IF({dc}="","",","&{foset_sep.join(f"{c}{r}" for c in fo_cols)}&",")'))
        wd = f'WEEKDAY({dc},2)'
        for i in range(MAX_EMP):
            cr = FIRST_EMP_ROW + i
            ca.cell(row=r, column=ABSENT0 + i, value=(
                f'=IF(OR({dc}="",Konfiguration!$A${cr}=""),"",OR('
                f'COUNTIFS(URL_A,Konfiguration!$A${cr},URL_B,"<="&{dc},'
                f'URL_C,">="&{dc})>0,'
                f'IFERROR(INDEX(FT_BQ,MATCH({dc},FT_A,0),'
                f'MATCH(Konfiguration!$B${cr},FT_H,0)),"")<>""))'))
            fri = [f'ISNUMBER(SEARCH(","&{i}&",",INDEX({fr},{p - 1})))'
                   for p in (r - 5, r - 4, r - 3, r - 2, r - 1) if p >= 2]
            mon = [f'IF(AND({mc}-{d}-1>=1,{mc}-{d}-1<={r - 2}),'
                   f'ISNUMBER(SEARCH(","&{i}&",",INDEX({fr},{mc}-{d}-1))),FALSE)'
                   for d in (6, 5, 4, 3, 2)] if r >= 3 else []
            if fri and mon:
                was_prev = f'IF({wd}=5,OR({",".join(fri)}),OR({",".join(mon)}))'
            elif fri:
                was_prev = f'IF({wd}=5,OR({",".join(fri)}),FALSE)'
            elif mon:
                was_prev = f'IF({wd}<5,OR({",".join(mon)}),FALSE)'
            else:
                was_prev = "FALSE"
            ca.cell(row=r, column=WASFO0 + i, value=(
                f'=IF(OR({dc}="",Konfiguration!$A${cr}=""),"",{was_prev})'))
            ca.cell(row=r, column=BLOCK0 + i, value=(
                f'=IF(OR({dc}="",Konfiguration!$A${cr}=""),"",'
                f'OR({get_column_letter(ABSENT0 + i)}{r},'
                f'{get_column_letter(WASFO0 + i)}{r}))'))
            days = [p for p in (r - 4, r - 3, r - 2, r - 1) if p >= 2]
            if len(days) == 4:
                mt = "AND(" + ",".join(
                    f'ISNUMBER(SEARCH(","&{i}&",",INDEX({fr},{p - 1})))'
                    for p in days) + ")"
            else:
                mt = "FALSE"
            ca.cell(row=r, column=MT0 + i, value=(
                f'=IF(OR({dc}="",Konfiguration!$A${cr}="",{wd}<>5),"",{mt})'))
    ca.sheet_state = "hidden"
    protect(ca)

    # ---------------- Dienstplan (week x employee table) ----------------
    ro = wb.create_sheet("Dienstplan")
    ro["A1"] = "KW = ISO-Kalenderwoche (Jahr in Konfiguration!B3); Mo-Fr = Status"
    ro["A1"].font = Font(italic=True)
    headers = ["KW", "Mitarbeiter", "Mo", "Di", "Mi", "Do", "Fr"]
    for j, h in enumerate(headers, start=1):
        c = ro.cell(row=2, column=j, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")
    # hidden helper column I: Monday of the row's week
    ro["I2"] = "MoDatum"
    last_row = 2 + WEEKS * MAX_EMP
    for w in range(0, WEEKS):
        for i in range(MAX_EMP):
            r = 3 + w * MAX_EMP + i
            cr = FIRST_EMP_ROW + i
            ro.cell(row=r, column=9,
                    value=f'=DATE(JAHR,1,4)-WEEKDAY(DATE(JAHR,1,4),3)+7*{w}')
            ro.cell(row=r, column=9).number_format = "DD.MM.YYYY"
            ro.cell(row=r, column=1, value=f'=IF($I{r}="","",WEEKNUM($I{r},21))')
            ro.cell(row=r, column=2,
                    value=f'=IF(Konfiguration!$A${cr}="","",Konfiguration!$A${cr})')
            mtcol = get_column_letter(MT0 + i)
            for d in range(1, 6):
                dt = f'($I{r}+{d - 1})'
                dr = f'MATCH({dt},CALCD,0)'
                vac = (f'COUNTIFS(URL_A,$B{r},URL_B,"<="&{dt},URL_C,">="&{dt})>0')
                hol = (f'IFERROR(INDEX(FT_BQ,MATCH({dt},FT_A,0),'
                       f'MATCH(Konfiguration!$B${cr},FT_H,0)),"")')
                frei = f'AND({d}=5,INDEX(Calc!${mtcol}$2:${mtcol}${1 + DAYS},{dr}-1)=TRUE)'
                fo = "OR(" + ",".join(
                    f'INDEX(Calc!$A$2:$H${1 + DAYS},{dr}-1,{j})={i}'
                    for j in range(1, MAX_FO + 1)) + ")"
                formula = (
                    f'=IF($B{r}="","",IF(ISNA({dr}),"",'
                    f'IF({vac},"Urlaub",'
                    f'IF({hol}<>"",{hol},'
                    f'IF({frei},"frei",'
                    f'IF({fo},"FO","BO"))))))'
                )
                cell = ro.cell(row=r, column=2 + d, value=formula)
                cell.alignment = Alignment(horizontal="center")
                cell.border = THIN
    ro.column_dimensions["A"].width = 6
    ro.column_dimensions["B"].width = 24
    for d in range(3, 8):
        ro.column_dimensions[get_column_letter(d)].width = 13
    ro.column_dimensions["I"].hidden = True
    cf_range = f"C3:G{last_row}"
    ro.conditional_formatting.add(cf_range, CellIsRule(
        operator="equal", formula=['"FO"'], fill=FILL_FO, stopIfTrue=True))
    ro.conditional_formatting.add(cf_range, CellIsRule(
        operator="equal", formula=['"BO"'], fill=FILL_BO, stopIfTrue=True))
    ro.conditional_formatting.add(cf_range, CellIsRule(
        operator="equal", formula=['"frei"'], fill=FILL_FREI, stopIfTrue=True))
    ro.conditional_formatting.add(cf_range, CellIsRule(
        operator="equal", formula=['"Urlaub"'], fill=FILL_URLAUB, stopIfTrue=True))
    ro.conditional_formatting.add(cf_range, FormulaRule(
        formula=['AND(C3<>"",C3<>"FO",C3<>"BO",C3<>"frei",C3<>"Urlaub")'],
        fill=FILL_HOLIDAY, stopIfTrue=True))
    ro.conditional_formatting.add(
        f"A3:G{last_row}",
        FormulaRule(formula=['MOD($A3,2)=1'], fill=FILL_WEEK))
    tab = Table(displayName="DienstplanTabelle", ref=f"A2:G{last_row}",
                tableColumns=[TableColumn(id=j + 1, name=h)
                              for j, h in enumerate(headers)])
    tab.autoFilter = AutoFilter(ref=f"A2:G{last_row}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=False)
    ro.add_table(tab)
    ro.freeze_panes = "C3"
    protect(ro)

    # ---------------- Auswertung ----------------
    aw = wb.create_sheet("Auswertung")
    headers_aw = ["Name", "Bundesland", "FO-Tage", "davon Freitag", "BO-Tage",
                  "Freitage", "Urlaub", "Feiertage", "Stunden"]
    for j, h in enumerate(headers_aw, start=1):
        c = aw.cell(row=1, column=j, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
    br = f"Dienstplan!$B$3:$B${last_row}"
    daycols = [f"Dienstplan!${get_column_letter(3 + d)}$3:${get_column_letter(3 + d)}${last_row}"
               for d in range(5)]
    for i in range(MAX_EMP):
        r = 2 + i
        cr = FIRST_EMP_ROW + i

        def countif_status(status):
            return "+".join(f'COUNTIFS({br},$A{r},{dc},"{status}")' for dc in daycols)

        def countif_not_statuses():
            terms = []
            for dc in daycols:
                terms.append(f'COUNTIFS({br},$A{r},{dc},"<>FO",{dc},"<>BO",'
                             f'{dc},"<>frei",{dc},"<>Urlaub",{dc},"<>")')
            return "+".join(terms)

        aw.cell(row=r, column=1, value=f'=IF(Konfiguration!A{cr}="","",Konfiguration!A{cr})')
        aw.cell(row=r, column=2, value=f'=IF(Konfiguration!A{cr}="","",Konfiguration!B{cr})')
        aw.cell(row=r, column=3, value=f'=IF($A{r}="","",{countif_status("FO")})')
        aw.cell(row=r, column=4, value=f'=IF($A{r}="","",COUNTIFS({br},$A{r},{daycols[4]},"FO"))')
        aw.cell(row=r, column=5, value=f'=IF($A{r}="","",{countif_status("BO")})')
        aw.cell(row=r, column=6, value=f'=IF($A{r}="","",{countif_status("frei")})')
        aw.cell(row=r, column=7, value=f'=IF($A{r}="","",{countif_status("Urlaub")})')
        aw.cell(row=r, column=8, value=f'=IF($A{r}="","",{countif_not_statuses()})')
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
