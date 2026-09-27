"""Builds Credit_Note_Report.xlsx: yellow input cells feed live formulas and native Excel charts."""
import re
import zipfile
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col

from order_type_data import MONTHS, ORDER_TYPES


def labels(fmt):
    return DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False,
                         showPercent=False, showLeaderLines=False, numFmt=fmt)

HERE = Path(__file__).parent
OUT = HERE / "Credit_Note_Report.xlsx"

INVOICE_CNS = [  # month, invoices count, CNS count
    (date(2025, 9, 1), 5416, 1028),
    (date(2025, 10, 1), 10360, 2212),
    (date(2025, 11, 1), 5496, 1338),
    (date(2025, 12, 1), 5527, 598),
    (date(2026, 1, 1), 10836, 703),
    (date(2026, 2, 1), 5555, 1127),
    (date(2026, 3, 1), 5586, 986),
    (date(2026, 4, 1), 10638, 590),
    (date(2026, 5, 1), 5623, 801),
    (date(2026, 6, 1), 5668, 1038),
    (date(2026, 7, 1), 9048, 658),
    (date(2026, 8, 1), 5809, 1219),
    (date(2026, 9, 1), 6254, 775),
]
N = len(INVOICE_CNS)

TYPES = [t for t, v in ORDER_TYPES.items() if t != "Unknown" and sum(v) > 0]
NT = len(TYPES)
FIRST_T, LAST_T = 2, 1 + NT            # order-type columns B..T
TC0, TC1 = col(FIRST_T), col(LAST_T)
TOTAL_C = col(LAST_T + 1)

month_idx = {m: i for i, m in enumerate(MONTHS)}
def ot_count(t, d):
    return ORDER_TYPES[t][month_idx[d.strftime("%m-%Y")]]

# Inputs layout
INV_H, INV0 = 4, 5                     # invoice table header row / first data row
INV1 = INV0 + N - 1
OT_H, OT0 = INV1 + 4, INV1 + 5         # order-type table header row / first data row
OT1 = OT0 + N - 1

NAVY = "1F3B73"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
TOTAL_FILL = PatternFill("solid", fgColor="D9E1F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
PCT1, PCT2 = "0.0%", "0.00%"
MON = "mmm-yy"
INT = "#,##0"
MOM = '+0.0%;-0.0%;0.0%'

wb = Workbook()


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=15, color=NAVY)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = Font(italic=True, color="595959")


def header(ws, row, c0, values):
    for i, v in enumerate(values):
        c = ws.cell(row, c0 + i, v)
        c.fill, c.font, c.border = HDR_FILL, HDR_FONT, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def cell(ws, ref, value, fmt=None, fill=None, bold=False):
    c = ws[ref]
    c.value = value
    c.border = BOX
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if bold:
        c.font = Font(bold=True)
    return c


def section(ws, ref, text):
    ws[ref] = text
    ws[ref].font = Font(bold=True, size=12, color=NAVY)


def brief(ws, ref, formula):
    c = ws[ref]
    c.value = formula
    c.alignment = Alignment(wrap_text=True, vertical="top")


# ============================================================ Inputs
inp = wb.active
inp.title = "Inputs"
title(inp, "Inputs – edit the yellow cells, every report and chart recalculates",
      "Months must stay in order from Sep-2025 onwards. Order type 'Unknown' is excluded from all reports.")
section(inp, f"A{INV_H - 1}", "Invoice and credit note counts")
header(inp, INV_H, 1, ["Month", "Invoices Count", "CNS Count"])
for i, (d, iv, cn) in enumerate(INVOICE_CNS):
    r = INV0 + i
    cell(inp, f"A{r}", d, MON, INPUT_FILL)
    cell(inp, f"B{r}", iv, INT, INPUT_FILL)
    cell(inp, f"C{r}", cn, INT, INPUT_FILL)

section(inp, f"A{OT_H - 1}", "Credit notes per order type (counts used for the percentages)")
header(inp, OT_H, 1, ["Month"] + TYPES + ["Total"])
for i, (d, _, _) in enumerate(INVOICE_CNS):
    r = OT0 + i
    cell(inp, f"A{r}", f"=A{INV0 + i}", MON)
    for j, t in enumerate(TYPES):
        cell(inp, f"{col(FIRST_T + j)}{r}", ot_count(t, d), INT, INPUT_FILL)
    cell(inp, f"{TOTAL_C}{r}", f"=SUM({TC0}{r}:{TC1}{r})", INT, TOTAL_FILL, True)
inp.column_dimensions["A"].width = 11
for j in range(FIRST_T, LAST_T + 2):
    inp.column_dimensions[col(j)].width = 12.5
inp.row_dimensions[OT_H].height = 45
inp.freeze_panes = "B5"

# ============================================================ Report 1
r1 = wb.create_sheet("Report 1")
title(r1, "Report 1 – Month · Invoices Count · CNS Count", "MoM = change versus the previous month.")
header(r1, 4, 1, ["Month", "Invoices Count", "CNS Count", "Invoices MoM", "CNS MoM"])
R1_0, R1_1 = 5, 5 + N - 1
for i in range(N):
    r, src = R1_0 + i, INV0 + i
    cell(r1, f"A{r}", f"=Inputs!A{src}", MON)
    cell(r1, f"B{r}", f"=Inputs!B{src}", INT)
    cell(r1, f"C{r}", f"=Inputs!C{src}", INT)
    if i == 0:
        cell(r1, f"D{r}", "–")
        cell(r1, f"E{r}", "–")
    else:
        cell(r1, f"D{r}", f"=IFERROR(B{r}/B{r - 1}-1,0)", MOM)
        cell(r1, f"E{r}", f"=IFERROR(C{r}/C{r - 1}-1,0)", MOM)
TOT, AVG = R1_1 + 1, R1_1 + 2
cell(r1, f"A{TOT}", "Total", fill=TOTAL_FILL, bold=True)
cell(r1, f"B{TOT}", f"=SUM(B{R1_0}:B{R1_1})", INT, TOTAL_FILL, True)
cell(r1, f"C{TOT}", f"=SUM(C{R1_0}:C{R1_1})", INT, TOTAL_FILL, True)
cell(r1, f"A{AVG}", "Monthly avg", fill=TOTAL_FILL, bold=True)
cell(r1, f"B{AVG}", f"=AVERAGE(B{R1_0}:B{R1_1})", INT, TOTAL_FILL, True)
cell(r1, f"C{AVG}", f"=AVERAGE(C{R1_0}:C{R1_1})", INT, TOTAL_FILL, True)
for c_, w in zip("ABCDE", (12, 15, 12, 13, 12)):
    r1.column_dimensions[c_].width = w

ch = BarChart()
ch.type = "col"
ch.title = "Invoices Count vs CNS Count by month"
ch.y_axis.title = "Count"
ch.y_axis.numFmt = INT
ch.y_axis.majorGridlines = None
ch.add_data(Reference(r1, min_col=2, max_col=3, min_row=4, max_row=R1_1), titles_from_data=True)
ch.set_categories(Reference(r1, min_col=1, min_row=R1_0, max_row=R1_1))
ch.dataLabels = labels(INT)
ch.dataLabels.position = "outEnd"
ch.series[0].graphicalProperties.solidFill = NAVY
ch.series[1].graphicalProperties.solidFill = "C0392B"
ch.gapWidth = 60
ch.legend.position = "b"
ch.x_axis.delete = False
ch.y_axis.delete = False
ch.height, ch.width = 9, 26
r1.add_chart(ch, "G4")

B1 = AVG + 3
section(r1, f"A{B1}", "Brief on each number")
for i in range(N):
    r, out = R1_0 + i, B1 + 1 + i
    base = (f'TEXT(A{r},"mmm yyyy")&": "&TEXT(B{r},"#,##0")&" invoices and "&TEXT(C{r},"#,##0")'
            f'&" credit notes. "')
    if i == 0:
        mom = '"This is the first month of the range, so there is no previous month to compare with. "'
    else:
        mom = (f'"Versus "&TEXT(A{r - 1},"mmm yyyy")&", invoices "&IF(D{r}>=0,"rose ","fell ")'
               f'&TEXT(ABS(D{r}),"0.0%")&" and credit notes "&IF(E{r}>=0,"rose ","fell ")'
               f'&TEXT(ABS(E{r}),"0.0%")&". "')
    avg = (f'"CNS count is "&IF(C{r}>$C${AVG}*1.25,"well above",IF(C{r}<$C${AVG}*0.75,"well below","close to"))'
           f'&" the "&TEXT($C${AVG},"#,##0")&" monthly average."')
    brief(r1, f"A{out}", f"={base}&{mom}&{avg}")
    r1.merge_cells(f"A{out}:Q{out}")
    r1.row_dimensions[out].height = 30
out = B1 + 1 + N
brief(r1, f"A{out}",
      f'="Total: "&TEXT(B{TOT},"#,##0")&" invoices (avg "&TEXT(B{AVG},"#,##0")&" per month) and "'
      f'&TEXT(C{TOT},"#,##0")&" credit notes (avg "&TEXT(C{AVG},"#,##0")&" per month) from "'
      f'&TEXT(A{R1_0},"mmm yyyy")&" to "&TEXT(A{R1_1},"mmm yyyy")&"."')
r1.merge_cells(f"A{out}:Q{out}")
r1.row_dimensions[out].height = 30
r1.freeze_panes = "A5"

# ============================================================ Calc (helper sheet)
calc = wb.create_sheet("Calc")
title(calc, "Helper calculations (do not edit – feeds Reports 2 and 3)")

# ============================================================ Report 2
r2 = wb.create_sheet("Report 2")
title(r2, "Report 2 – Percentage of credit notes per order type, by month",
      "Each month adds up to 100%. Darker colour = bigger share. 'Unknown' excluded.")
header(r2, 4, 1, ["Month"] + [f"=Inputs!{col(FIRST_T + j)}{OT_H}" for j in range(NT)])
R2_0, R2_1 = 5, 5 + N - 1
for i in range(N):
    r, src = R2_0 + i, OT0 + i
    cell(r2, f"A{r}", f"=Inputs!A{src}", MON)
    for j in range(NT):
        c_ = col(FIRST_T + j)
        cell(r2, f"{c_}{r}", f"=IFERROR(Inputs!{c_}{src}/Inputs!${TOTAL_C}{src},0)", PCT1)
R2_AVG = R2_1 + 1
cell(r2, f"A{R2_AVG}", "Average", fill=TOTAL_FILL, bold=True)
for j in range(NT):
    c_ = col(FIRST_T + j)
    cell(r2, f"{c_}{R2_AVG}", f"=AVERAGE({c_}{R2_0}:{c_}{R2_1})", PCT1, TOTAL_FILL, True)
r2.conditional_formatting.add(
    f"{TC0}{R2_0}:{TC1}{R2_1}",
    ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                   mid_type="num", mid_value=0.15, mid_color="F8C471",
                   end_type="num", end_value=0.6, end_color="C0392B"))
r2.column_dimensions["A"].width = 11
for j in range(FIRST_T, LAST_T + 1):
    r2.column_dimensions[col(j)].width = 11
r2.row_dimensions[4].height = 45
r2.freeze_panes = "B5"

# Calc: tie-broken keys for ranking inside each month (rows mirror Report 2)
header(calc, 4, 1, ["Month key"] + [f"='Report 2'!{col(FIRST_T + j)}4" for j in range(NT)])
for i in range(N + 1):
    r = R2_0 + i
    calc[f"A{r}"] = f"='Report 2'!A{r}"
    for j in range(NT):
        c_ = col(FIRST_T + j)
        calc[f"{c_}{r}"] = f"='Report 2'!{c_}{r}+({LAST_T + 1}-COLUMN())/1000000000"
KEYS = lambda r: f"Calc!${TC0}${r}:${TC1}${r}"
NAMES = f"'Report 2'!${TC0}$4:${TC1}$4"
AVG_KEYS = KEYS(R2_AVG)

# Chart data: top 7 types by average share + all other types
TOPK = 7
CD_C0 = LAST_T + 3                     # chart-data block to the right of the table
section(r2, f"{col(CD_C0)}3", "Chart data (auto: top 7 order types by average share)")
header(r2, 4, CD_C0, ["Month"] + [None] * TOPK + ["All other types"])
for k in range(TOPK):
    c = r2.cell(4, CD_C0 + 1 + k)
    c.value = f"=INDEX({NAMES},MATCH(LARGE({AVG_KEYS},{k + 1}),{AVG_KEYS},0))"
for i in range(N):
    r = R2_0 + i
    cell(r2, f"{col(CD_C0)}{r}", f"=A{r}", MON)
    for k in range(TOPK):
        hc = f"{col(CD_C0 + 1 + k)}$4"
        cell(r2, f"{col(CD_C0 + 1 + k)}{r}", f"=INDEX(${TC0}{r}:${TC1}{r},MATCH({hc},${TC0}$4:${TC1}$4,0))", PCT1)
    cell(r2, f"{col(CD_C0 + TOPK + 1)}{r}",
         f"=MAX(0,1-SUM({col(CD_C0 + 1)}{r}:{col(CD_C0 + TOPK)}{r}))", PCT1)
for k in range(TOPK + 2):
    r2.column_dimensions[col(CD_C0 + k)].width = 12

ch2 = BarChart()
ch2.type = "bar"
ch2.grouping = "percentStacked"
ch2.overlap = 100
ch2.gapWidth = 40
ch2.title = "Credit-note mix per month (share of each order type)"
ch2.add_data(Reference(r2, min_col=CD_C0 + 1, max_col=CD_C0 + TOPK + 1, min_row=4, max_row=R2_1),
             titles_from_data=True)
ch2.set_categories(Reference(r2, min_col=CD_C0, min_row=R2_0, max_row=R2_1))
ch2.dataLabels = labels('[<0.035]"";0%')
ch2.y_axis.numFmt = "0%"
ch2.y_axis.majorGridlines = None
ch2.x_axis.scaling.orientation = "maxMin"
ch2.x_axis.delete = False
ch2.y_axis.delete = False
ch2.legend.position = "b"
palette = ["1F3B73", "2E86C1", "C0392B", "E67E22", "27AE60", "8E44AD", "16A085", "BDC3C7"]
for s, c_ in zip(ch2.series, palette):
    s.graphicalProperties.solidFill = c_
    s.graphicalProperties.line.solidFill = "FFFFFF"
ch2.height, ch2.width = 12, 26
CH2_ROW = R2_AVG + 3
r2.add_chart(ch2, f"A{CH2_ROW}")

# Calc: per-month top 3 and remainder for the briefs
TOP_C0 = LAST_T + 3
header(calc, 4, TOP_C0, ["Top 1", "Share 1", "Top 2", "Share 2", "Top 3", "Share 3",
                         "Other types >=1%", "Types <1%", "Top 3 total"])
for i in range(N):
    r = R2_0 + i
    for k in range(3):
        calc[f"{col(TOP_C0 + 2 * k)}{r}"] = f"=INDEX({NAMES},MATCH(LARGE({KEYS(r)},{k + 1}),{KEYS(r)},0))"
        calc[f"{col(TOP_C0 + 2 * k + 1)}{r}"] = f"=LARGE('Report 2'!${TC0}{r}:${TC1}{r},{k + 1})"
    rng = f"'Report 2'!${TC0}{r}:${TC1}{r}"
    calc[f"{col(TOP_C0 + 6)}{r}"] = f'=COUNTIF({rng},">=0.01")-COUNTIF({col(TOP_C0 + 1)}{r},">=0.01")' \
        f'-COUNTIF({col(TOP_C0 + 3)}{r},">=0.01")-COUNTIF({col(TOP_C0 + 5)}{r},">=0.01")'
    calc[f"{col(TOP_C0 + 7)}{r}"] = f'=COUNTIFS({rng},">0",{rng},"<0.01")'
    calc[f"{col(TOP_C0 + 8)}{r}"] = f"={col(TOP_C0 + 1)}{r}+{col(TOP_C0 + 3)}{r}+{col(TOP_C0 + 5)}{r}"

B2 = CH2_ROW + 25
section(r2, f"A{B2}", "Brief on each month")
T = lambda k, r: f"Calc!{col(TOP_C0 + k)}{r}"
for i in range(N):
    r, out = R2_0 + i, B2 + 1 + i
    f = (f'=TEXT(A{r},"mmm yyyy")&": "&{T(0, r)}&" "&TEXT({T(1, r)},"0.0%")&", "&{T(2, r)}&" "'
         f'&TEXT({T(3, r)},"0.0%")&", "&{T(4, r)}&" "&TEXT({T(5, r)},"0.0%")&". Top 3 = "'
         f'&TEXT({T(8, r)},"0.0%")&" of the month. The remaining "&TEXT(1-{T(8, r)},"0.0%")'
         f'&" is spread over "&{T(6, r)}&" other order types at 1% or more and "&{T(7, r)}&" under 1%."')
    brief(r2, f"A{out}", f)
    r2.merge_cells(f"A{out}:P{out}")
    r2.row_dimensions[out].height = 30

S2 = B2 + N + 3
section(r2, f"A{S2}", "Notable month-to-month shifts (top 5 order types by average share)")
for k in range(5):
    out = S2 + 1 + k
    hc = f"{col(CD_C0 + 1 + k)}$4"
    colrng = f"INDEX(${TC0}${R2_0}:${TC1}${R2_1},0,MATCH({hc},${TC0}$4:${TC1}$4,0))"
    mcol = f"$A${R2_0}:$A${R2_1}"
    f = (f'={hc}&": from "&TEXT(MIN({colrng}),"0.0%")&" ("&TEXT(INDEX({mcol},MATCH(MIN({colrng}),{colrng},0)),"mmm yyyy")'
         f'&") to "&TEXT(MAX({colrng}),"0.0%")&" ("&TEXT(INDEX({mcol},MATCH(MAX({colrng}),{colrng},0)),"mmm yyyy")'
         f'&"), average monthly share "&TEXT(AVERAGE({colrng}),"0.0%")&"."')
    brief(r2, f"A{out}", f)
    r2.merge_cells(f"A{out}:P{out}")
    r2.row_dimensions[out].height = 18

# ============================================================ Report 3
r3 = wb.create_sheet("Report 3")
title(r3, "Report 3 – Percentage of credit notes per order type over the period",
      "Change the yellow start / end months to re-cut the period. 'Unknown' excluded.")
r3["A3"], r3["A4"], r3["A5"] = "Start month", "End month", "Months in range"
for a in ("A3", "A4", "A5"):
    r3[a].font = Font(bold=True)
cell(r3, "B3", f"=Inputs!A{INV0}", MON, INPUT_FILL)
cell(r3, "B4", f"=Inputs!A{INV1}", MON, INPUT_FILL)
MRANGE = f"Inputs!$A${OT0}:$A${OT1}"
cell(r3, "B5", f'=COUNTIFS({MRANGE},">="&$B$3,{MRANGE},"<="&$B$4)')

# Calc: counts / shares / peaks per order type within the chosen range
C3H = R2_AVG + 4
C3_0, C3_1 = C3H + 1, C3H + NT
MASK_H = C3_1 + 3
MASK_0, MASK_1 = MASK_H + 1, MASK_H + N
header(calc, C3H, 1, ["Order type", "Credit notes in range", "Share", "Rank key",
                      "Months present", "Peak monthly share", "Peak month"])
header(calc, MASK_H, 1, ["Month (in range)"] + [f"='Report 2'!{col(FIRST_T + j)}4" for j in range(NT)])
for i in range(N):
    r = MASK_0 + i
    calc[f"A{r}"] = f"=Inputs!A{OT0 + i}"
    calc[f"A{r}"].number_format = MON
    for j in range(NT):
        c_ = col(FIRST_T + j)
        calc[f"{c_}{r}"] = (f"=IF(AND(Inputs!$A{OT0 + i}>='Report 3'!$B$3,Inputs!$A{OT0 + i}<='Report 3'!$B$4),"
                            f"'Report 2'!{c_}{R2_0 + i},-1)")
MASK = f"${TC0}${MASK_0}:${TC1}${MASK_1}"
for j in range(NT):
    r = C3_0 + j
    calc[f"A{r}"] = f"=Inputs!{col(FIRST_T + j)}{OT_H}"
    calc[f"B{r}"] = (f"=SUMIFS(Inputs!{col(FIRST_T + j)}${OT0}:{col(FIRST_T + j)}${OT1},"
                     f"{MRANGE},\">=\"&'Report 3'!$B$3,{MRANGE},\"<=\"&'Report 3'!$B$4)")
    calc[f"C{r}"] = f"=IFERROR(B{r}/SUM($B${C3_0}:$B${C3_1}),0)"
    calc[f"D{r}"] = f"=C{r}+({C3_1 + 1}-ROW())/1000000000"
    calc[f"E{r}"] = f'=COUNTIF(INDEX({MASK},0,{j + 1}),">0")'
    calc[f"F{r}"] = f"=MAX(0,MAX(INDEX({MASK},0,{j + 1})))"
    calc[f"G{r}"] = f"=INDEX($A${MASK_0}:$A${MASK_1},MATCH(MAX(INDEX({MASK},0,{j + 1})),INDEX({MASK},0,{j + 1}),0))"
    calc[f"G{r}"].number_format = MON
for c_ in range(1, LAST_T + 12):
    calc.column_dimensions[col(c_)].width = 13

KEY3 = f"Calc!$D${C3_0}:$D${C3_1}"
header(r3, 7, 1, ["Rank", "Order type", "Share", "Running total"])
R3_0, R3_1 = 8, 8 + NT - 1
for k in range(NT):
    r = R3_0 + k
    m = f"MATCH(LARGE({KEY3},A{r}),{KEY3},0)"
    cell(r3, f"A{r}", k + 1)
    cell(r3, f"B{r}", f"=INDEX(Calc!$A${C3_0}:$A${C3_1},{m})")
    cell(r3, f"C{r}", f"=INDEX(Calc!$C${C3_0}:$C${C3_1},{m})", PCT2)
    cell(r3, f"D{r}", f"=SUM($C${R3_0}:C{r})", PCT1)
r3.conditional_formatting.add(
    f"C{R3_0}:C{R3_1}",
    ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                   mid_type="num", mid_value=0.1, mid_color="F8C471",
                   end_type="num", end_value=0.6, end_color="C0392B"))

DN_H = R3_1 + 3
DN_0 = DN_H + 1
section(r3, f"A{DN_H - 1}", "Doughnut data (top 4 + all other types)")
header(r3, DN_H, 2, ["Order type", "Share"])
for k in range(4):
    cell(r3, f"B{DN_0 + k}", f"=B{R3_0 + k}")
    cell(r3, f"C{DN_0 + k}", f"=C{R3_0 + k}", PCT1)
cell(r3, f"B{DN_0 + 4}", "All other types")
cell(r3, f"C{DN_0 + 4}", f"=MAX(0,1-SUM(C{DN_0}:C{DN_0 + 3}))", PCT1)
for c_, w in zip("ABCD", (16, 26, 11, 14)):
    r3.column_dimensions[c_].width = w

ch3 = BarChart()
ch3.type = "bar"
ch3.title = "Share of credit notes by order type"
ch3.add_data(Reference(r3, min_col=3, min_row=7, max_row=R3_1), titles_from_data=True)
ch3.set_categories(Reference(r3, min_col=2, min_row=R3_0, max_row=R3_1))
ch3.dataLabels = labels('[<0.001]"<0.1%";0.0%')
ch3.dataLabels.position = "outEnd"
ch3.x_axis.scaling.orientation = "maxMin"
ch3.y_axis.numFmt = "0%"
ch3.y_axis.majorGridlines = None
ch3.x_axis.delete = False
ch3.y_axis.delete = False
ch3.legend = None
ch3.x_axis.tickLblSkip = 1
ch3.gapWidth = 40
ch3.series[0].graphicalProperties.solidFill = "C0392B"
ch3.height, ch3.width = 15, 18
r3.add_chart(ch3, "F3")

ch4 = DoughnutChart()
ch4.title = "Credit-note mix: top 4 order types"
ch4.add_data(Reference(r3, min_col=3, min_row=DN_H, max_row=DN_0 + 4), titles_from_data=True)
ch4.set_categories(Reference(r3, min_col=2, min_row=DN_0, max_row=DN_0 + 4))
ch4.dataLabels = labels("0.0%")
ch4.holeSize = 50
ch4.legend.position = "r"
ch4.height, ch4.width = 9, 18
r3.add_chart(ch4, "Q3")

B3R = DN_0 + 7
section(r3, f"A{B3R}", "Brief on each order type")
brief(r3, f"A{B3R + 1}",
      '="Period: "&TEXT($B$3,"mmm yyyy")&" to "&TEXT($B$4,"mmm yyyy")&" ("&$B$5&" months). The top 3 order types make up "'
      f'&TEXT(D{R3_0 + 2},"0.0%")&" of all credit notes."')
r3.merge_cells(f"A{B3R + 1}:P{B3R + 1}")
for k in range(NT):
    r, out = R3_0 + k, B3R + 2 + k
    m = f"MATCH(B{r},Calc!$A${C3_0}:$A${C3_1},0)"
    f = (f'=B{r}&": "&IF(C{r}<0.0001,"<0.01%",TEXT(C{r},IF(C{r}<0.01,"0.00%","0.0%")))&" of credit notes in the period (running total "'
         f'&TEXT(D{r},"0.0%")&"). Present in "&INDEX(Calc!$E${C3_0}:$E${C3_1},{m})&" of "&$B$5'
         f'&" months; highest monthly share "&TEXT(INDEX(Calc!$F${C3_0}:$F${C3_1},{m}),"0.0%")&" in "'
         f'&TEXT(INDEX(Calc!$G${C3_0}:$G${C3_1},{m}),"mmm yyyy")&"."')
    brief(r3, f"A{out}", f)
    r3.merge_cells(f"A{out}:P{out}")
    r3.row_dimensions[out].height = 18

# ============================================================ Summary
sm = wb.create_sheet("Summary", 0)
title(sm, "Invoices & Credit Notes – Summary", "All figures below are formulas and update when Inputs change.")
INVR = f"Inputs!$B${INV0}:$B${INV1}"
CNR = f"Inputs!$C${INV0}:$C${INV1}"
MR = f"Inputs!$A${INV0}:$A${INV1}"
lines = [
    ("Period", f'=TEXT(Inputs!A{INV0},"mmm yyyy")&" to "&TEXT(Inputs!A{INV1},"mmm yyyy")&" ("&COUNT({MR})&" months)."'),
    ("Totals", f'=TEXT(SUM({INVR}),"#,##0")&" invoices and "&TEXT(SUM({CNR}),"#,##0")&" credit notes; on average "'
               f'&TEXT(AVERAGE({INVR}),"#,##0")&" invoices and "&TEXT(AVERAGE({CNR}),"#,##0")&" credit notes per month."'),
    ("Highest CNS", f'=TEXT(INDEX({MR},MATCH(MAX({CNR}),{CNR},0)),"mmm yyyy")&" with "&TEXT(MAX({CNR}),"#,##0")&" credit notes."'),
    ("Lowest CNS", f'=TEXT(INDEX({MR},MATCH(MIN({CNR}),{CNR},0)),"mmm yyyy")&" with "&TEXT(MIN({CNR}),"#,##0")&" credit notes."'),
    ("Highest invoices", f'=TEXT(INDEX({MR},MATCH(MAX({INVR}),{INVR},0)),"mmm yyyy")&" with "&TEXT(MAX({INVR}),"#,##0")&" invoices."'),
    ("Top order types", f"=\"In \"&TEXT('Report 3'!B3,\"mmm yyyy\")&\" – \"&TEXT('Report 3'!B4,\"mmm yyyy\")&\": \""
                        f"&'Report 3'!B{R3_0}&\" \"&TEXT('Report 3'!C{R3_0},\"0.0%\")&\", \"&'Report 3'!B{R3_0 + 1}&\" \""
                        f"&TEXT('Report 3'!C{R3_0 + 1},\"0.0%\")&\", \"&'Report 3'!B{R3_0 + 2}&\" \"&TEXT('Report 3'!C{R3_0 + 2},\"0.0%\")"
                        f"&\" (together \"&TEXT('Report 3'!D{R3_0 + 2},\"0.0%\")&\" of credit notes).\""),
    ("Invoice peaks", f'="Invoices exceed 1.5x the median month in "&COUNTIF({INVR},">"&1.5*MEDIAN({INVR}))'
                      f'&" months, while CNS count does not rise with them."'),
]
for i, (k, f) in enumerate(lines):
    r = 4 + i
    sm[f"A{r}"] = k
    sm[f"A{r}"].font = Font(bold=True)
    brief(sm, f"B{r}", f)
    sm.row_dimensions[r].height = 32
sm.column_dimensions["A"].width = 18
sm.column_dimensions["B"].width = 120
section(sm, "A13", "How to use")
for i, t in enumerate([
    "Edit only the yellow cells on the Inputs sheet (and the start / end month on Report 3).",
    "Report 1: invoice and credit-note counts by month, with a chart and a written brief for every number.",
    "Report 2: percentage of credit notes per order type, by month (heatmap table and 100% stacked chart).",
    "Report 3: percentage of credit notes per order type over the chosen period (default Sep-2025 to Sep-2026).",
    "Order type 'Unknown' is excluded. Percentages use the sum of the remaining order types as 100%.",
]):
    sm[f"A{14 + i}"] = "• " + t

calc.sheet_state = "hidden"
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)


def fix_label_formats(path):
    """openpyxl omits the required sourceLinked attribute on data-label number formats."""
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename.startswith("xl/charts/chart"):
                data = re.sub(rb'(<(?:c:)?numFmt formatCode="[^"]*")\s*/>', rb'\1 sourceLinked="0"/>', data)
            dst.writestr(item, data)
    tmp.replace(path)


fix_label_formats(OUT)
print("wrote", OUT)
