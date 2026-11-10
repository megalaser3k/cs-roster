"""Generates a formula-driven Excel roster planner for German teams.

Sheets:
  Config     - year input, employee list with Bundesland, legend
  Feiertage  - German public holidays for the configured year, per Bundesland
  Urlaub     - vacation periods per employee (start/end dates)
  Roster     - Excel table: one row per working day (Mon-Fri), one column per
               employee; FO/BO/frei/Urlaub/holiday resolved by formulas
  Auswertung - per-employee totals and hours
"""
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo, TableColumn
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.workbook.defined_name import DefinedName

STATES = ["BW", "BY", "BE", "BB", "HB", "HH", "HE", "MV",
          "NI", "NW", "RP", "SL", "SN", "ST", "SH", "TH"]

MAX_EMP = 60
FIRST_EMP_ROW = 6
LAST_EMP_ROW = FIRST_EMP_ROW + MAX_EMP - 1
ROSTER_ROWS = 263  # max Mon-Fri days in a year is 262; one spare row
URLAUB_ROWS = 200  # max vacation entries

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

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=14)
THIN = Border(*[Side(style="thin", color="BFBFBF")] * 4)

FILL_FO = PatternFill("solid", fgColor="BDD7EE")       # blue
FILL_BO = PatternFill("solid", fgColor="C6EFCE")       # green
FILL_FREI = PatternFill("solid", fgColor="D9D9D9")     # grey
FILL_URLAUB = PatternFill("solid", fgColor="FFEB9C")   # yellow
FILL_HOLIDAY = PatternFill("solid", fgColor="FFC7CE")  # red


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

    cfg["A5"] = "Name"
    cfg["B5"] = "Bundesland"
    for c in ("A5", "B5"):
        cfg[c].fill = HEADER_FILL
        cfg[c].font = HEADER_FONT
    dv = DataValidation(type="list", formula1='"' + ",".join(STATES) + '"',
                        allow_blank=True, showDropDown=False)
    cfg.add_data_validation(dv)
    dv.add(f"B{FIRST_EMP_ROW}:B{LAST_EMP_ROW}")

    wb.defined_names["MitarbeiterNamen"] = DefinedName(
        "MitarbeiterNamen", attr_text=f"Config!$A${FIRST_EMP_ROW}:$A${LAST_EMP_ROW}")

    legend = [
        "Legende:",
        "FO  = Front-Office (Mo-Do je 10 h, Fr 8 h) - blau",
        "BO  = Back-Office (8 h) - grün",
        "frei = Ausgleichs-Freitag nach der Front-Office-Woche - grau",
        "Urlaub = Urlaubstag (Zeiträume auf Blatt 'Urlaub' eintragen, - gelb",
        "  mehrere Urlaube pro Person = mehrere Zeilen mit gleichem Namen)",
        "Feiertagsname = arbeitsfrei (Feiertag im eigenen Bundesland) - rot",
        "",
        "Jedes/r Mitarbeiter/in ist alle N Wochen (N = Anzahl Mitarbeiter)",
        "im Front-Office: Freitag der Start, dann Mo-Do der Folgewoche,",
        "der folgende Freitag ist frei.",
        "",
        "Das Roster-Blatt ist eine Tabelle: Filtern über die Kopfzeile.",
        "Spalte D = 1. Person aus Config, Spalte E = 2. Person usw.",
        "",
        "Hinweis: Fällt ein Front-Office-Tag auf Urlaub oder einen Feiertag",
        "im eigenen Bundesland, ist die Abdeckung manuell zu planen.",
        "",
        "Zukunftssicher: Zum Jahreswechsel nur das Jahr in B3 ändern.",
    ]
    for i, line in enumerate(legend, start=5):
        cfg[f"D{i}"] = line
    cfg.column_dimensions["A"].width = 28
    cfg.column_dimensions["B"].width = 14
    cfg.column_dimensions["D"].width = 62

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

    # ---------------- Urlaub ----------------
    ur = wb.create_sheet("Urlaub")
    for j, h in enumerate(["Mitarbeiter", "Von", "Bis"], start=1):
        c = ur.cell(row=1, column=j, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
    for r in range(2, 2 + URLAUB_ROWS):
        ur.cell(row=r, column=2).number_format = "DD.MM.YYYY"
        ur.cell(row=r, column=3).number_format = "DD.MM.YYYY"
    dv2 = DataValidation(type="list", formula1="MitarbeiterNamen",
                         allow_blank=True)
    ur.add_data_validation(dv2)
    dv2.add(f"A2:A{1 + URLAUB_ROWS}")
    ur["E1"] = "Mehrere Urlaube pro Person: einfach mehrere Zeilen mit demselben Namen eintragen."
    ur["E1"].font = Font(italic=True)
    ur.column_dimensions["A"].width = 28
    ur.column_dimensions["B"].width = 12
    ur.column_dimensions["C"].width = 12
    ur.freeze_panes = "A2"

    # ---------------- Roster ----------------
    ro = wb.create_sheet("Roster")
    headers = ["Datum", "Tag", "KW"] + [f"Mitarbeiter {i + 1}" for i in range(MAX_EMP)]
    for j, h in enumerate(headers, start=1):
        c = ro.cell(row=1, column=j, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")

    for r in range(2, 2 + ROSTER_ROWS):
        if r == 2:
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
            vac = (f'SUMPRODUCT((Urlaub!$A$2:$A${1 + URLAUB_ROWS}=Config!$A${cr})'
                   f'*(Urlaub!$B$2:$B${1 + URLAUB_ROWS}<=$A{r})'
                   f'*(Urlaub!$C$2:$C${1 + URLAUB_ROWS}>=$A{r}))>0')
            formula = (
                f'=IF(OR($A{r}="",Config!$E$3=0,Config!$A${cr}=""),"",'
                f'IF({vac},"Urlaub",'
                f'IF({h}<>"",{h},'
                f'IF(AND(WEEKDAY($A{r},2)=5,MOD(INT(($A{r}-DATE(1900,1,2))/7)-2,Config!$E$3)={i}),"frei",'
                f'IF(MOD(INT(($A{r}-DATE(1900,1,2))/7)-1-IF(WEEKDAY($A{r},2)=5,0,1),Config!$E$3)={i},"FO","BO"))))))'
            )
            cell = ro.cell(row=r, column=4 + i, value=formula)
            cell.alignment = Alignment(horizontal="center")
            cell.border = THIN

    last_col = get_column_letter(3 + MAX_EMP)
    last_row = 1 + ROSTER_ROWS
    cf_range = f"D2:{last_col}{last_row}"
    ro.conditional_formatting.add(cf_range, CellIsRule(
        operator="equal", formula=['"FO"'], fill=FILL_FO, stopIfTrue=True))
    ro.conditional_formatting.add(cf_range, CellIsRule(
        operator="equal", formula=['"BO"'], fill=FILL_BO, stopIfTrue=True))
    ro.conditional_formatting.add(cf_range, CellIsRule(
        operator="equal", formula=['"frei"'], fill=FILL_FREI, stopIfTrue=True))
    ro.conditional_formatting.add(cf_range, CellIsRule(
        operator="equal", formula=['"Urlaub"'], fill=FILL_URLAUB, stopIfTrue=True))
    ro.conditional_formatting.add(cf_range, FormulaRule(
        formula=[f'AND(D2<>"",D2<>"FO",D2<>"BO",D2<>"frei",D2<>"Urlaub")'],
        fill=FILL_HOLIDAY, stopIfTrue=True))

    tab = Table(displayName="RosterTabelle",
                ref=f"A1:{last_col}{last_row}",
                tableColumns=[TableColumn(id=j + 1, name=h)
                              for j, h in enumerate(headers)])
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2",
                                         showRowStripes=True)
    ro.add_table(tab)

    ro.column_dimensions["A"].width = 12
    ro.column_dimensions["B"].width = 6
    ro.column_dimensions["C"].width = 6
    for i in range(MAX_EMP):
        ro.column_dimensions[get_column_letter(4 + i)].width = 12
    ro.freeze_panes = "D2"

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
        rng = f"Roster!{col}$2:{col}${last_row}"
        aw.cell(row=r, column=1, value=f'=IF(Config!A{cr}="","",Config!A{cr})')
        aw.cell(row=r, column=2, value=f'=IF(Config!A{cr}="","",Config!B{cr})')
        aw.cell(row=r, column=3, value=f'=IF($A{r}="","",COUNTIF({rng},"FO"))')
        aw.cell(row=r, column=4,
                value=f'=IF($A{r}="","",COUNTIFS({rng},"FO",Roster!$B$2:$B${last_row},"Fr"))')
        aw.cell(row=r, column=5, value=f'=IF($A{r}="","",COUNTIF({rng},"BO"))')
        aw.cell(row=r, column=6, value=f'=IF($A{r}="","",COUNTIF({rng},"frei"))')
        aw.cell(row=r, column=7, value=f'=IF($A{r}="","",COUNTIF({rng},"Urlaub"))')
        aw.cell(row=r, column=8,
                value=(f'=IF($A{r}="","",SUMPRODUCT(({rng}<>"FO")*({rng}<>"BO")*'
                       f'({rng}<>"frei")*({rng}<>"Urlaub")*({rng}<>"")))'))
        aw.cell(row=r, column=9, value=f'=IF($A{r}="","",10*(C{r}-D{r})+8*D{r}+8*E{r})')
    aw.column_dimensions["A"].width = 28
    aw.column_dimensions["B"].width = 12
    for j in range(3, 10):
        aw.column_dimensions[get_column_letter(j)].width = 13
    aw.freeze_panes = "A2"

    wb.save(path)
    print("saved", path)


if __name__ == "__main__":
    build("/workspace/Dienstplan.xlsx")
