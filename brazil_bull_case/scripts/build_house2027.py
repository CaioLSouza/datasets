"""Builds Ibov_DCF_house_2027_XP.xlsx: a replica of the DCF block of the house Ibovespa model (DCF Ibov 2026_Out.xlsx,
Target sheet, rows 17-65 and 120-151), rolled forward one year to a fair value at the end of 2027.

Same inputs (history 2015-2025 and IBOV Index consensus 2026-2028 as in the house file), same rules and the same mechanics:
FCFF = D&A + NOPAT + capex + NWC (level), discount rate = 5y real rate (NTN-B) + ERP, Gordon terminal value at g, Excel NPV
with the terminal value as one more period. The projection is extended to 2031 with the house's rules for 2029-2030, so the
end-2027 value uses four explicit years (2028-2031), as the house's end-2026 value uses 2027-2030. Four 0/1 switches correct
the mechanics one at a time (all off = the house). Checks tie the replica to the house file's own numbers.

Run: python build_house2027.py -> python excel_recalc.py ..\\Ibov_DCF_house_2027_XP.xlsx
"""
from __future__ import annotations

import warnings
from pathlib import Path

import openpyxl
import pandas as pd

from xp_template import Book, cn, NAVY, YEL, GREY, GREEN, INK, KEY, GBAND, PTS, PTS1, MULT, PCT, PCT2, UPS, PX, DATED, HEAT

warnings.filterwarnings("ignore")
HERE = Path(__file__).parent
HOUSE = Path(r"C:\Users\Caio\Documents\Documentos\Equity Strategy Dashboard\_pedidos\2026-10-05_ibov_valuation_groups\DCF Ibov 2026_Out.xlsx")
OUT = HERE.parent / "Ibov_DCF_house_2027_XP.xlsx"

# --------------------------------------------------------------------------- house file values
T = openpyxl.load_workbook(HOUSE, data_only=True)["Target"]
HY = {y: cn(1 + y - 2015) for y in range(2015, 2031)}          # house columns: 2015 -> B ... 2030 -> Q
hv = lambda row, y: T[f"{HY[y]}{row}"].value
H = {"ibov": T["B61"].value, "rr": T["C56"].value, "erp": T["B57"].value, "g": T["B54"].value, "infl": T["B63"].value,
     "fcff": {y: hv(44, y) for y in range(2025, 2031)}, "tv": T["R44"].value, "fv25": T["D59"].value, "fv26": T["F59"].value,
     "fv27": T["G59"].value, "irr": T["B62"].value, "erp_impl": T["B65"].value,
     "sens": [(T[f"B{r}"].value, T[f"C{r}"].value) for r in range(122, 130)], "nd_mult": T["F70"].value}
assert abs(H["fv26"] - 202259.1686) < 0.01, H["fv26"]

YEARS = list(range(2015, 2032))
YC = {y: cn(2 + y - 2015) for y in YEARS}                      # ours: 2015 -> C ... 2031 -> S
LASTA, LASTC = 2025, 2028                                       # last actual year; last year with consensus EBIT/EBITDA/EPS
NOTE_C = "U"

ORDER = ["Cover", "Summary", "Assumptions", "DCF", "Sensitivity", "Checks", "Read Me"]
bk = Book(OUT, ORDER, {"Assumptions": YEL})
F, W = bk.F, bk.W
IN, LV = bk.IN, bk.LV


def status(y):
    return "Actual" if y <= LASTA else ("Consensus" if y <= LASTC else "Projection")


# =========================================================================== Assumptions
ws = W["Assumptions"]
bk.title(ws, "DCF assumptions (house model, end-2027)", "Change the pink cells; the DCF sheet recalculates. Projection rules for 2028-2031 are pink cells in the DCF sheet.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 62)
ws.set_column("C:F", 15)
ws.set_column("G:G", 110)
A = {}
bk.section(ws, 5, "1. Market data", "B", "G")
bk.header_row(ws, 6, [("B", "Item"), ("C", "House file"), ("D", "Oct-5-2026"), ("E", "Used"), ("F", ""), ("G", "Source")])
for r, (lab, nm, hvv, nowv, fmt, src) in enumerate([
        ("Ibovespa (index pts)", "IbovUsed", H["ibov"], 208431.92, PX, "House file: Target!B61 (BDP IBOV Index, as saved). Oct-5: Bloomberg IBOV Index PX_LAST."),
        ("5-year real rate (NTN-B)", "RRUsed", H["rr"], 0.0676, PCT2, "House file: Target!C56 (BZRFB5PY Index, 7.42%). Oct-5: ANBIMA 5-year NTN-B (Bloomberg was down; check BZRFB5PY before publishing).")], start=7):
    ws.write(f"B{r}", lab, bk.TXT)
    ws.write_number(f"C{r}", hvv, F(**IN, num_format=fmt))
    ws.write_number(f"D{r}", nowv, F(**IN, num_format=fmt))
    ws.write_formula(f"E{r}", f"=CHOOSE(Vintage,C{r},D{r})", F(bold=True, num_format=fmt))
    ws.write(f"G{r}", src, bk.NOTE)
    bk.name(nm, "Assumptions", f"$E${r}")
    A[nm] = r
ws.write("B9", "Market data used: 1 = house file, 2 = Oct-5-2026", bk.TXT)
ws.write_number("C9", 2, F(**LV, num_format="0"))
ws.data_validation("C9", {"validate": "list", "source": [1, 2]})
ws.write("G9", "Consensus and history are the house file's in both cases (not refreshed: the consensus moved little).", bk.NOTE)
bk.name("Vintage", "Assumptions", "$C$9")

bk.section(ws, 11, "2. Valuation parameters (house Target sheet)", "B", "G")
bk.header_row(ws, 12, [("B", "Parameter"), ("C", "Value"), ("D", ""), ("E", ""), ("F", ""), ("G", "Source")])
for r, (lab, nm, v, fmt, src) in enumerate([
        ("Long-term growth of the FCFF (g)", "gLT", H["g"], PCT2, "Target!B54 (4.7%)."),
        ("Long-term inflation (Fisher switch and the implied-return block)", "InflLT", H["infl"], PCT2, "Target!B63 (4.0%)."),
        ("Working capital, % of sales", "NWCpct", 0.04, PCT, "Target row 43: NWC = -4% of sales."),
        ("Net debt / EBITDA of two years before (net debt switch)", "NDmult", H["nd_mult"], MULT,
         "Target rows 70-71 (EV/EBITDA method): net debt = 1.6x the EBITDA of two years before the value date.")], start=13):
    ws.write(f"B{r}", lab, bk.TXT)
    ws.write_number(f"C{r}", v, F(**LV, num_format=fmt))
    ws.write(f"G{r}", src, bk.NOTE)
    bk.name(nm, "Assumptions", f"$C${r}")
    A[nm] = r

bk.section(ws, 18, "3. Scenarios (fair value at the end of 2027)", "B", "G")
bk.header_row(ws, 19, [("B", "Lever"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), ("F", ""), ("G", "Source")])
SC = {"Bear": "C", "Base": "D", "Bull": "E"}
ws.write("B20", "5-year real rate (NTN-B)", bk.TXT)
ws.write_number("C20", 0.085, F(**LV, num_format=PCT2))
ws.write_formula("D20", "=RRUsed", F(**LV, num_format=PCT2))
ws.write_number("E20", 0.06, F(**LV, num_format=PCT2))
ws.write("G20", "As the house (Target rows 131-151): bear 8.5%, base = the market rate, bull 6.0%.", bk.NOTE)
ws.write("B21", "ERP", bk.TXT)
for c in "CDE":
    ws.write_number(f"{c}21", H["erp"], F(**LV, num_format=PCT2))
ws.write("G21", "Target!B57 (6.0%), the same in every scenario.", bk.NOTE)
ws.write("B22", "Discount rate (house: real rate + ERP)", bk.BOLD)
for c in "CDE":
    ws.write_formula(f"{c}22", f"=IF(FixInfl=1,(1+{c}20+{c}21)*(1+InflLT)-1,{c}20+{c}21)", F(bold=True, num_format=PCT2))
ws.write("G22", "Target!B55 = BZRFB5PY + ERP. With the inflation switch on: (1 + real rate + ERP) x (1 + inflation) - 1.", bk.NOTE)

bk.section(ws, 24, "4. Corrections to the house mechanics (0 = as the house, 1 = corrected)", "B", "G")
bk.header_row(ws, 25, [("B", "Switch"), ("C", "Value"), ("D", ""), ("E", ""), ("F", ""), ("G", "What it changes")])
for r, (lab, nm, note) in enumerate([
        ("Working capital: 0 = level (house), 1 = change", "FixNWC", "House row 44 adds the whole NWC level (-4% of sales) every year; the cash flow is the change in NWC."),
        ("Net debt: 0 = not subtracted (house), 1 = subtracted", "FixND", "FCFF values the firm; equity = firm value - net debt (1.6x EBITDA two years before, as the house's EV/EBITDA method)."),
        ("Terminal value: 0 = one period after the last flow (house), 1 = with the last flow", "FixTV", "NPV(rate, flows, TV) puts the TV one year after the last explicit year; it is a value at the end of that year."),
        ("Inflation in the rate: 0 = real rate + ERP (house), 1 = Fisher", "FixInfl", "The flows are nominal (sales grow with nominal GDP); real rate + ERP is a real rate.")], start=26):
    ws.write(f"B{r}", lab, bk.TXT)
    ws.write_number(f"C{r}", 0, F(**LV, num_format="0"))
    ws.data_validation(f"C{r}", {"validate": "list", "source": [0, 1]})
    ws.write(f"G{r}", note, bk.NOTE)
    bk.name(nm, "Assumptions", f"$C${r}")
    A[nm] = r
bk.section(ws, 31, "5. Projection rules (pink cells in the DCF sheet, 2028-2031 columns)", "B", "G")
for i, t in enumerate([
        "Real GDP and inflation 2028-2031: 2.0% and 4.0% (house); sales grow with nominal GDP x ratio (1.0).",
        "EBITDA margin 2029-2031: 30% (house). EBIT margin, interest and gross margin: held at the 2028 level (house: same % as the year before).",
        "Capex % of sales: 11% in 2028-2029, 9% from 2030 (house); 2031, the year added here, keeps 2030's 9%.",
        "Consensus: sales and capex to 2027, EBIT, EBITDA and EPS to 2028 (IBOV Index, as in the house file)."]):
    ws.write(f"B{32 + i}", t, bk.NOTE)
ws.freeze_panes(4, 0)

# =========================================================================== DCF
ws = W["DCF"]
bk.title(ws, "DCF (house model rolled to end-2027)", "Replica of Target rows 17-65 of DCF Ibov 2026_Out.xlsx, in Ibovespa points, extended to 2031. Blue = house file input; pink = projection rule.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 44)
ws.set_column(f"C:{YC[2031]}", 9.5)
ws.set_column("T:T", 2)
ws.set_column(f"{NOTE_C}:{NOTE_C}", 90)
bk.section(ws, 5, "Projections (index points)", "B", NOTE_C)
bk.header_row(ws, 6, [("B", "")] + [(YC[y], str(y)) for y in YEARS] + [(NOTE_C, "How it is calculated (house Target row)")])
for y in YEARS:
    ws.write(f"{YC[y]}7", status(y), F(italic=True, font_size=8, font_color=GREY if y <= LASTA else (INK if y <= LASTC else NAVY), align="center"))
KEYS = ["rgdp", "infl", "ngdp", "sg", "ratio", "sales", "cogs", "gm", "gp", "sga", "sgap", "ebit", "ebitm", "ebitda", "ebitdam", "da",
        "dap", "tax", "nopat", "int", "intp", "ni", "nim", "nwc", "dnwc", "capex", "capexp", "fcff_house", "fcff_fix", "fcff_used",
        "fcff_simple", "ibov", "fcfy"]
R = {k: 8 + i for i, k in enumerate(KEYS)}
LAB = {"rgdp": "Real GDP", "infl": "Inflation", "ngdp": "Nominal GDP", "sg": "Sales growth", "ratio": "Ratio (sales growth / nominal GDP)",
       "sales": "Sales", "cogs": "COGS", "gm": "Gross margin %", "gp": "Gross profit", "sga": "SG&A", "sgap": "SG&A % of sales", "ebit": "EBIT",
       "ebitm": "EBIT margin %", "ebitda": "EBITDA", "ebitdam": "EBITDA margin %", "da": "D&A", "dap": "D&A % of sales", "tax": "Tax rate",
       "nopat": "NOPAT", "int": "Net interest expense", "intp": "Interest % of sales", "ni": "Net income (EPS, index pts)", "nim": "Net margin %",
       "nwc": "Working capital (NWC)", "dnwc": "Change in NWC", "capex": "Capex", "capexp": "Capex % of sales",
       "fcff_house": "FCFF, house = D&A + NOPAT + capex + NWC (level)", "fcff_fix": "FCFF with the change in NWC",
       "fcff_used": "FCFF used in the valuation (switch FixNWC)", "fcff_simple": "memo: FCFF = EBITDA x (1 - tax) + capex",
       "ibov": "memo: Ibovespa (average to 2020, then the price used)", "fcfy": "memo: FCF yield"}
NOTES = {"ngdp": "(1 + real GDP) x (1 + inflation) - 1 (row 20).", "sg": "Actual/consensus: sales / previous - 1. Projection: nominal GDP x ratio (row 21).",
         "sales": "To 2027: house file (consensus 2026-27). 2028+: previous x (1 + growth) (row 25).", "cogs": "Sales x (gross margin - 1) (row 26).",
         "gm": "To 2028: house file. 2029+: previous year (row 27).", "sga": "To 2028: EBIT - gross profit. 2029+: sales x SG&A % (row 29).",
         "sgap": "2029+: previous year's % (row 30).", "ebit": "To 2028: house file (consensus). 2029+: gross profit + SG&A (row 31).",
         "ebitda": "To 2028: house file (consensus). 2029+: sales x EBITDA margin (row 33).", "ebitdam": "2029+: 30% (house, row 34).",
         "da": "EBITDA - EBIT (row 35).", "tax": "-25% to 2023, -24% from 2024 (row 37).", "nopat": "EBIT x (1 + tax rate) (row 38).",
         "int": "To 2028: net income - NOPAT. 2029+: sales x interest % (row 39).", "intp": "2029+: previous year's % (row 40).",
         "ni": "To 2028: house file (consensus EPS). 2029+: NOPAT + net interest (row 41).", "nwc": "-NWC % x sales (Assumptions; row 43).",
         "capex": "To 2027: house file. 2028+: sales x capex % (row 46).", "capexp": "2028-29: -11%; 2030: -9% (house); 2031: as 2030 (row 47).",
         "fcff_house": "D&A + NOPAT + capex + NWC level (row 44): what the house discounts.", "fcff_fix": "D&A + NOPAT + capex + change in NWC.",
         "fcff_used": "FixNWC = 0: the house row; 1: the corrected row.", "fcff_simple": "To 2025: house file. 2026+: EBITDA x (1 + tax) + capex (row 49).",
         "ibov": "Row 50: average to 2020 (house file), then the Ibovespa price used.", "fcfy": "Simple FCFF / Ibovespa (row 51)."}
PCTROWS = {"rgdp", "infl", "ngdp", "sg", "gm", "sgap", "ebitm", "ebitdam", "dap", "tax", "intp", "nim", "capexp", "fcfy"}
for k in KEYS:
    bold = k in ("sales", "ebitda", "ebit", "ni", "fcff_used")
    ws.write(f"B{R[k]}", LAB[k], F(bold=bold, font_color=GREY if k.startswith(("fcff_simple", "ibov", "fcfy")) else INK))
    if k in NOTES:
        ws.write(f"{NOTE_C}{R[k]}", NOTES[k], bk.NOTE)


def put(k, y, val, kind):
    """kind: 'in' (blue house input), 'lv' (pink rule), 'f' (formula)."""
    fmt = PCT if k in PCTROWS else ("0.00" if k == "ratio" else PTS)
    c = f"{YC[y]}{R[k]}"
    bold = k in ("sales", "ebitda", "ebit", "ni", "fcff_used")
    if kind == "f":
        ws.write_formula(c, val, F(num_format=fmt, bold=bold, bg_color=KEY if k == "fcff_used" and y >= 2026 else "#FFFFFF"))
    elif val is not None:
        ws.write_number(c, val, F(**(LV if kind == "lv" else IN), num_format=fmt))


C_ = lambda k, y: f"{YC[y]}{R[k]}"
for y in YEARS:
    p = y - 1
    hy = y if y <= 2030 else None
    # macro
    if 2022 <= y <= 2027:
        put("rgdp", y, hv(18, y), "in")
    elif y >= 2028:
        put("rgdp", y, hv(18, y) if y <= 2030 else 0.02, "lv")
    put("infl", y, hv(19, y) if y <= 2030 else 0.04, "in" if y <= 2027 else "lv")
    put("ngdp", y, f"=(1+{C_('rgdp', y)})*(1+{C_('infl', y)})-1", "f")
    # sales
    if y <= 2027:
        put("sales", y, hv(25, y), "in")
        if y > 2015:
            put("sg", y, f"={C_('sales', y)}/{C_('sales', p)}-1", "f")
            put("ratio", y, f"={C_('sg', y)}/{C_('ngdp', y)}", "f")
    else:
        put("ratio", y, 1.0, "lv")
        put("sg", y, f"={C_('ngdp', y)}*{C_('ratio', y)}", "f")
        put("sales", y, f"={C_('sales', p)}*(1+{C_('sg', y)})", "f")
    put("cogs", y, f"={C_('sales', y)}*({C_('gm', y)}-1)", "f")
    put("gm", y, hv(27, y) if y <= LASTC else f"={C_('gm', p)}", "in" if y <= LASTC else "f")
    put("gp", y, f"={C_('sales', y)}+{C_('cogs', y)}", "f")
    if y <= LASTC:
        put("ebit", y, hv(31, y), "in")
        put("sga", y, f"={C_('ebit', y)}-{C_('gp', y)}", "f")
        put("sgap", y, f"={C_('sga', y)}/{C_('sales', y)}", "f")
        put("ebitda", y, hv(33, y), "in")
        put("ebitdam", y, f"={C_('ebitda', y)}/{C_('sales', y)}", "f")
        put("ni", y, hv(41, y), "in")
        put("int", y, f"={C_('ni', y)}-{C_('nopat', y)}", "f")
        put("intp", y, f"={C_('int', y)}/{C_('sales', y)}", "f")
    else:
        put("sgap", y, f"={C_('sgap', p)}", "f")
        put("sga", y, f"={C_('sales', y)}*{C_('sgap', y)}", "f")
        put("ebit", y, f"={C_('gp', y)}+{C_('sga', y)}", "f")
        put("ebitdam", y, hv(34, y) if y <= 2030 else 0.30, "lv")
        put("ebitda", y, f"={C_('sales', y)}*{C_('ebitdam', y)}", "f")
        put("intp", y, f"={C_('intp', p)}", "f")
        put("int", y, f"={C_('sales', y)}*{C_('intp', y)}", "f")
        put("ni", y, f"={C_('nopat', y)}+{C_('int', y)}", "f")
    put("ebitm", y, f"={C_('ebit', y)}/{C_('sales', y)}", "f")
    put("da", y, f"={C_('ebitda', y)}-{C_('ebit', y)}", "f")
    put("dap", y, f"={C_('da', y)}/{C_('sales', y)}", "f")
    put("tax", y, hv(37, y) if y in (2015, 2024) else f"={C_('tax', p)}", "in" if y in (2015, 2024) else "f")
    put("nopat", y, f"={C_('ebit', y)}*(1+{C_('tax', y)})", "f")
    put("nim", y, f"={C_('ni', y)}/{C_('sales', y)}", "f")
    put("nwc", y, f"=-NWCpct*{C_('sales', y)}", "f")
    if y > 2015:
        put("dnwc", y, f"={C_('nwc', y)}-{C_('nwc', p)}", "f")
    if y <= 2027:
        put("capex", y, hv(46, y), "in")
        put("capexp", y, f"={C_('capex', y)}/{C_('sales', y)}", "f")
    else:
        if y in (2028, 2030):
            put("capexp", y, hv(47, y), "lv")
        else:
            put("capexp", y, f"={C_('capexp', p)}", "f")
        put("capex", y, f"={C_('sales', y)}*{C_('capexp', y)}", "f")
    put("fcff_house", y, f"={C_('da', y)}+{C_('nopat', y)}+{C_('capex', y)}+{C_('nwc', y)}", "f")
    if y > 2015:
        put("fcff_fix", y, f"={C_('da', y)}+{C_('nopat', y)}+{C_('capex', y)}+{C_('dnwc', y)}", "f")
        put("fcff_used", y, f"=IF(FixNWC=1,{C_('fcff_fix', y)},{C_('fcff_house', y)})", "f")
    else:
        put("fcff_used", y, f"={C_('fcff_house', y)}", "f")
    if y <= LASTA:
        put("fcff_simple", y, hv(49, y), "in")
    else:
        put("fcff_simple", y, f"={C_('ebitda', y)}*(1+{C_('tax', y)})+{C_('capex', y)}", "f")
    if y <= 2020:
        put("ibov", y, hv(50, y), "in")
    else:
        put("ibov", y, "=IbovUsed", "f")
    put("fcfy", y, f"={C_('fcff_simple', y)}/{C_('ibov', y)}", "f")

# --------------------------------------------------------------------------- valuation (rows 53-59 of the house)
V0 = R["fcfy"] + 2
bk.section(ws, V0, "Valuation: fair value at the end of the year (house Target rows 53-59)", "B", NOTE_C)
bk.header_row(ws, V0 + 1, [("B", ""), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), ("F", ""), ("G", "Base 2026")])
ws.merge_range(f"H{V0 + 1}:{NOTE_C}{V0 + 1}", "How it is calculated", F(bold=True, font_color="#FFFFFF", bg_color=NAVY))
VK = ["date", "years", "rr", "erp", "rate", "g", "last", "tv", "per", "pvf", "pvtv", "firm", "nd", "fv", "ibov", "up", "tvsh", "pe"]
V = {k: V0 + 2 + i for i, k in enumerate(VK)}
VL = {"date": "Value date (end of year)", "years": "Explicit FCFF years", "rr": "5-year real rate (NTN-B)", "erp": "ERP",
      "rate": "Discount rate (house 'WACC')", "g": "Long-term growth (g)", "last": "FCFF of the last explicit year",
      "tv": "Terminal value = last FCFF x (1 + g) / (rate - g)", "per": "Periods the terminal value is discounted",
      "pvf": "Present value of the explicit FCFF", "pvtv": "Present value of the terminal value", "firm": "Value of the cash flows",
      "nd": "Net debt subtracted (switch FixND)", "fv": "Fair value (index pts)", "ibov": "Ibovespa used", "up": "Upside",
      "tvsh": "Terminal value / value of the cash flows", "pe": "Implied P/E on next year's EPS"}
VN = {"rate": "Real rate + ERP (Target!B55); Fisher with FixInfl = 1.", "tv": "Target!R44.", "per": "4 explicit years + 1 (house NPV puts the TV one period later); 4 with FixTV = 1.",
      "pvf": "Excel NPV of the explicit FCFF (Target row 59).", "nd": "1.6x the EBITDA two years before the value date (Target rows 70-71), only with FixND = 1.",
      "fv": "Value of the cash flows - net debt subtracted.", "pe": "Fair value / net income of the year after the value date."}
for k in VK:
    ws.write(f"B{V[k]}", VL[k], F(bold=k in ("fv", "up", "rate")))
    if k in VN:
        ws.write(f"H{V[k]}", VN[k], bk.NOTE)
COLS = {"Bear": ("C", 2027), "Base": ("D", 2027), "Bull": ("E", 2027), "Base26": ("G", 2026)}
for s, (c, vy) in COLS.items():
    sc = SC["Base" if s == "Base26" else s]
    y0, y1 = vy + 1, vy + 4
    ws.write_number(f"{c}{V['date']}", vy, F(num_format="0", align="right"))
    ws.write_string(f"{c}{V['years']}", f"{y0}-{y1}", F(align="right"))
    ws.write_formula(f"{c}{V['rr']}", f"=Assumptions!{sc}20", F(num_format=PCT2))
    ws.write_formula(f"{c}{V['erp']}", f"=Assumptions!{sc}21", F(num_format=PCT2))
    ws.write_formula(f"{c}{V['rate']}", f"=IF(FixInfl=1,(1+{c}{V['rr']}+{c}{V['erp']})*(1+InflLT)-1,{c}{V['rr']}+{c}{V['erp']})", F(num_format=PCT2, bold=True))
    ws.write_formula(f"{c}{V['g']}", "=gLT", F(num_format=PCT2))
    ws.write_formula(f"{c}{V['last']}", f"={YC[y1]}{R['fcff_used']}", F(num_format=PTS))
    ws.write_formula(f"{c}{V['tv']}", f"={c}{V['last']}*(1+{c}{V['g']})/({c}{V['rate']}-{c}{V['g']})", F(num_format=PTS))
    ws.write_formula(f"{c}{V['per']}", "=5-FixTV", F(num_format="0"))
    ws.write_formula(f"{c}{V['pvf']}", f"=NPV({c}{V['rate']},{YC[y0]}{R['fcff_used']}:{YC[y1]}{R['fcff_used']})", F(num_format=PTS))
    ws.write_formula(f"{c}{V['pvtv']}", f"={c}{V['tv']}/(1+{c}{V['rate']})^{c}{V['per']}", F(num_format=PTS))
    ws.write_formula(f"{c}{V['firm']}", f"={c}{V['pvf']}+{c}{V['pvtv']}", F(num_format=PTS))
    ws.write_formula(f"{c}{V['nd']}", f"=IF(FixND=1,NDmult*{YC[vy - 2]}{R['ebitda']},0)", F(num_format=PTS))
    ws.write_formula(f"{c}{V['fv']}", f"={c}{V['firm']}-{c}{V['nd']}", F(num_format=PTS, bold=True, bg_color=KEY))
    ws.write_formula(f"{c}{V['ibov']}", "=IbovUsed", F(num_format=PTS))
    ws.write_formula(f"{c}{V['up']}", f"={c}{V['fv']}/{c}{V['ibov']}-1", F(num_format=UPS, bold=True, bg_color=KEY))
    ws.write_formula(f"{c}{V['tvsh']}", f"={c}{V['pvtv']}/{c}{V['firm']}", F(num_format=PCT))
    ws.write_formula(f"{c}{V['pe']}", f"={c}{V['fv']}/{YC[vy + 1]}{R['ni']}", F(num_format=MULT))
bk.name("FV27Bear", "DCF", f"$C${V['fv']}")
bk.name("FV27Base", "DCF", f"$D${V['fv']}")
bk.name("FV27Bull", "DCF", f"$E${V['fv']}")

# --------------------------------------------------------------------------- implied return (rows 61-65 of the house)
I0 = V["pe"] + 2
bk.section(ws, I0, "Implied return at the Ibovespa price used (house Target rows 61-65, rolled one year)", "B", NOTE_C)
bk.header_row(ws, I0 + 1, [("B", ""), ("C", "Today"), ("D", "2027"), ("E", "2028"), ("F", "2029"), ("G", "Terminal")])
ws.write(f"B{I0 + 2}", "Cash flows", bk.TXT)
ws.write_formula(f"C{I0 + 2}", "=-IbovUsed", F(num_format=PTS))
for c, y in zip("DEF", (2027, 2028, 2029)):
    ws.write_formula(f"{c}{I0 + 2}", f"={YC[y]}{R['fcff_used']}", F(num_format=PTS))
ws.write_formula(f"G{I0 + 2}", f"=D{V['tv']}", F(num_format=PTS))
for i, (lab, f_, fmt) in enumerate([("IRR", f"=IRR(C{I0 + 2}:G{I0 + 2})", PCT2), ("Inflation", "=InflLT", PCT2),
                                    ("Real IRR = IRR - inflation", f"=C{I0 + 3}-C{I0 + 4}", PCT2),
                                    ("Implied ERP = real IRR - 5-year real rate (base)", f"=C{I0 + 5}-D{V['rr']}", PCT2)]):
    ws.write(f"B{I0 + 3 + i}", lab, F(bold=i == 3))
    ws.write_formula(f"C{I0 + 3 + i}", f_, F(num_format=fmt, bold=i == 3))
ws.write(f"H{I0 + 2}", "As the house: price today, three yearly FCFF and the base terminal value as the fourth flow. An indicator, not a full DCF return.", bk.NOTE)
ws.freeze_panes(7, 2)
IRR_ROW = I0 + 6

# --------------------------------------------------------------------------- formula helper (fixed flags or the switches)


def fv_expr(y0, y1, rr, erp, g, nwc, tvfix, nd, infl, vy):
    """Fair value at the end of vy from the FCFF of y0..y1; flags 0/1 or 'cur' (the switches in Assumptions)."""
    row = R["fcff_used"] if nwc == "cur" else (R["fcff_fix"] if nwc else R["fcff_house"])
    k = {0: f"({rr}+{erp})", 1: f"((1+{rr}+{erp})*(1+InflLT)-1)", "cur": f"IF(FixInfl=1,(1+{rr}+{erp})*(1+InflLT)-1,{rr}+{erp})"}[infl]
    n = y1 - y0 + 1
    per = {0: f"{n + 1}", 1: f"{n}", "cur": f"({n}+1-FixTV)"}[tvfix]
    ndv = f"NDmult*DCF!${YC[vy - 2]}${R['ebitda']}"
    ndx = {0: "0", 1: ndv, "cur": f"IF(FixND=1,{ndv},0)"}[nd]
    return (f"NPV({k},DCF!${YC[y0]}${row}:${YC[y1]}${row})+DCF!${YC[y1]}${row}*(1+{g})/({k}-{g})/(1+{k})^{per}-{ndx}")


CUR = dict(nwc="cur", tvfix="cur", nd="cur", infl="cur")
HOUSEF = dict(nwc=0, tvfix=0, nd=0, infl=0)

# =========================================================================== Sensitivity
ws = W["Sensitivity"]
bk.title(ws, "Sensitivity: real rate and ERP", "The house's real-rate table (tie-out) and a real rate x ERP grid for the end-2027 fair value with the current settings.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 30)
ws.set_column("C:H", 15)
bk.section(ws, 5, "House table: 5-year real rate vs DCF value (Target rows 120-129)", "B", "G")
bk.header_row(ws, 6, [("B", "5-year real rate"), ("C", "House file, end-2026"), ("D", "Replica, end-2026 (house settings)"),
                      ("E", "Difference"), ("F", ""), ("G", "End-2027, current settings")])
ws.set_row(5, 30)
for i, (rr_, val) in enumerate(H["sens"]):
    r = 7 + i
    ws.write_number(f"B{r}", rr_, F(**IN, num_format=PCT2, align="center"))
    ws.write_number(f"C{r}", val, F(**IN, num_format=PTS))
    ws.write_formula(f"D{r}", "=" + fv_expr(2027, 2030, f"$B{r}", "0.06", "0.047", vy=2026, **HOUSEF), F(num_format=PTS))
    ws.write_formula(f"E{r}", f"=D{r}-C{r}", F(num_format=PX))
    ws.write_formula(f"G{r}", "=" + fv_expr(2028, 2031, f"$B{r}", "Assumptions!$D$21", "gLT", vy=2027, **CUR), F(num_format=PTS, bold=True))
SENS_LAST = 7 + len(H["sens"]) - 1
ws.write(f"B{SENS_LAST + 1}", "House settings: ERP 6.0%, g 4.7%, no corrections, house FCFF. The house table is an Excel data table on Target!C56.", bk.NOTE)
RRS = [0.05, 0.055, 0.06, 0.065, 0.0676, 0.07, 0.075, 0.08, 0.085, 0.09, 0.095]
ERPG = [0.04, 0.05, 0.06, 0.07, 0.08]
G1 = SENS_LAST + 4
for blk_i, (lab, kind) in enumerate([("End-2027 fair value (index pts), current settings", "v"), ("Upside vs the Ibovespa used", "u")]):
    top = G1 + blk_i * 16
    bk.section(ws, top, lab, "B", cn(1 + len(ERPG)))
    hr = top + 1
    ws.write(f"B{hr}", "Real rate \\ ERP", bk.HDRL)
    for j, e in enumerate(ERPG):
        if blk_i == 0:
            ws.write_number(hr - 1, 2 + j, e, F(**LV, num_format=PCT, align="center"))
        else:
            ws.write_formula(hr - 1, 2 + j, f"={cn(2 + j)}${G1 + 1}", F(bold=True, font_color="#FFFFFF", bg_color=NAVY, num_format=PCT, align="center"))
    for i, rr_ in enumerate(RRS):
        r = hr + 1 + i
        if blk_i == 0:
            ws.write_number(f"B{r}", rr_, F(**LV, num_format=PCT2, align="center"))
        else:
            ws.write_formula(f"B{r}", f"=$B${G1 + 2 + i}", F(bold=True, num_format=PCT2, align="center"))
        for j in range(len(ERPG)):
            if blk_i == 0:
                ws.write_formula(r - 1, 2 + j, "=" + fv_expr(2028, 2031, f"$B{r}", f"{cn(2 + j)}${hr}", "gLT", vy=2027, **CUR), F(num_format=PTS))
            else:
                ws.write_formula(r - 1, 2 + j, f"={cn(2 + j)}{G1 + 2 + i}/IbovUsed-1", F(num_format=UPS))
    last = hr + len(RRS)
    if kind == "u":
        ws.conditional_format(f"C{hr + 1}:{cn(1 + len(ERPG))}{last}", HEAT)
    ws.conditional_format(f"C{hr + 1}:{cn(1 + len(ERPG))}{last}", {"type": "formula",
                          "criteria": f"=AND(ROUND($B{hr + 1}-Assumptions!$D$20,4)=0,ROUND(C${hr}-Assumptions!$D$21,3)=0)",
                          "format": F(bold=True, border=2, border_color=NAVY)})
ws.write(f"B{G1 + 33}", "Framed cell = base real rate and base ERP. The grid uses the switches in Assumptions section 4.", bk.NOTE)

# =========================================================================== Summary
ws = W["Summary"]
bk.title(ws, "DCF (house model): fair value at the end of 2027", "Replica of the house DCF (DCF Ibov 2026_Out.xlsx), rolled forward one year. Construction: DCF sheet.", legend=False)
ws.set_column("A:A", 2)
ws.set_column("B:B", 52)
ws.set_column("C:H", 14)
ws.write("B3", "Ibovespa used", bk.NOTE)
ws.write_formula("C3", "=IbovUsed", F(num_format=PTS, bold=True))
ws.write("D3", "Market data", bk.NOTE)
ws.write_formula("E3", '=CHOOSE(Vintage,"house file","Oct-5-2026")', F(bold=True))
ws.write("F3", "Checks", bk.NOTE)
ws.write_formula("G3", "=Checks!C5", F(bold=True))
bk.section(ws, 5, "Fair value at the end of 2027 (index pts)", "B", "G")
bk.header_row(ws, 6, [("B", ""), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), ("F", ""), ("G", "Base, end-2026")])
for i, (lab, key, fmt) in enumerate([("5-year real rate (NTN-B)", "rr", PCT2), ("ERP", "erp", PCT2), ("Discount rate", "rate", PCT2),
                                     ("Fair value (index pts)", "fv", PTS), ("Upside vs the Ibovespa used", "up", UPS),
                                     ("Terminal value / value", "tvsh", PCT), ("Implied P/E on next year's EPS", "pe", MULT)]):
    r = 7 + i
    b = key in ("fv", "up")
    ws.write(f"B{r}", lab, F(bold=b))
    for c in "CDEG":
        ws.write_formula(f"{c}{r}", f"=DCF!{c}{V[key]}", F(num_format=fmt, bold=b, bg_color=KEY if b else "#FFFFFF"))
ws.conditional_format("C11:E11", HEAT)
ws.write("B14", "Corrections in use (Assumptions section 4)", bk.NOTE)
ws.write_formula("C14", '=IF(FixNWC+FixND+FixTV+FixInfl=0,"none: as the house",'
                        '"on: "&IF(FixNWC=1,"NWC change ","")&IF(FixND=1,"net debt ","")&IF(FixTV=1,"TV timing ","")&IF(FixInfl=1,"inflation",""))', F(italic=True, font_color=GREY))

bk.section(ws, 16, "Correcting the house mechanics one step at a time (base, end-2027)", "B", "G")
bk.header_row(ws, 17, [("B", "Step"), ("C", "Fair value"), ("D", "Change"), ("E", "Upside"), ("F", ""), ("G", "")])
steps = [("House mechanics", dict(nwc=0, tvfix=0, nd=0, infl=0)), ("+ terminal value discounted with the last flow", dict(nwc=0, tvfix=1, nd=0, infl=0)),
         ("+ working capital by change, not level", dict(nwc=1, tvfix=1, nd=0, infl=0)), ("+ net debt subtracted", dict(nwc=1, tvfix=1, nd=1, infl=0)),
         ("+ inflation in the rate (Fisher) = fully corrected", dict(nwc=1, tvfix=1, nd=1, infl=1))]
for i, (lab, fl) in enumerate(steps):
    r = 18 + i
    ws.write(f"B{r}", lab, F(bold=i in (0, len(steps) - 1)))
    ws.write_formula(f"C{r}", "=" + fv_expr(2028, 2031, "Assumptions!$D$20", "Assumptions!$D$21", "gLT", vy=2027, **fl), F(num_format=PTS, bold=i in (0, len(steps) - 1)))
    if i:
        ws.write_formula(f"D{r}", f"=C{r}-C{r - 1}", F(num_format='+#,##0;-#,##0;0'))
    ws.write_formula(f"E{r}", f"=C{r}/IbovUsed-1", F(num_format=UPS))
ws.write("B24", "Each step keeps the previous ones. Base real rate and ERP; projection rules as in the DCF sheet.", bk.NOTE)

bk.section(ws, 26, "Tie-out with the house file and reference values (not linked)", "B", "G")
bk.header_row(ws, 27, [("B", "Ibovespa fair value (index pts)"), ("C", "House file"), ("D", "Replica"), ("E", "Difference"), ("F", ""), ("G", "")])
ties = [("End-2026, house data and settings (Target!F59)", H["fv26"], fv_expr(2027, 2030, "Checks!$C$31", "Checks!$C$32", "Checks!$C$33", vy=2026, **HOUSEF)),
        ("End-2027 column of the house file, to 2030 (Target!G59)", H["fv27"], fv_expr(2028, 2030, "Checks!$C$31", "Checks!$C$32", "Checks!$C$33", vy=2027, **HOUSEF))]
for i, (lab, hval, f_) in enumerate(ties):
    r = 28 + i
    ws.write(f"B{r}", lab, bk.TXT)
    ws.write_number(f"C{r}", hval, F(**IN, num_format=PTS))
    ws.write_formula(f"D{r}", "=" + f_, F(num_format=PTS))
    ws.write_formula(f"E{r}", f"=D{r}-C{r}", F(num_format=PX))
for i, (lab, vals) in enumerate([("DCF by box (Ibov_DCF_by_box_XP.xlsx, 12 months)", (169580, 223718, 276936)),
                                  ("Multiples model, average of methods (Oct-6 table)", (174414, 229232, 294832))]):
    r = 31 + i
    ws.write(f"B{r}", lab, bk.TXT)
    for c, v in zip("CDE", vals):
        ws.write_number(f"{c}{r}", v, F(**IN, num_format=PTS))
ws.write("B30", "Other models, bear / base / bull:", bk.NOTE)
ws.freeze_panes(4, 0)

# =========================================================================== Checks
ws = W["Checks"]
bk.title(ws, "Checks", "'OK' / 'CHECK' drive the status in C5; 'Info' does not. Tie-outs hold while the projection rules are the house's.", legend=False)
ws.set_column("A:A", 2)
ws.set_column("B:B", 70)
ws.set_column("C:F", 15)
ws.set_column("G:G", 80)
ws.write("B5", "Overall status", bk.BOLD)
bk.header_row(ws, 7, [("B", "Check"), ("C", "Value"), ("D", "Target"), ("E", "Difference"), ("F", "Status"), ("G", "Note")])
chk = []
for y in range(2025, 2031):
    chk.append((f"FCFF {y}, house row = Target row 44", f"=DCF!{YC[y]}{R['fcff_house']}", f"={H['fcff'][y]}", 0.01, "House file value.", PX))
chk.append(("Terminal value on 2030, house rate (Target!R44)",
            f"=DCF!{YC[2030]}{R['fcff_house']}*(1+Checks!$C$33)/(Checks!$C$31+Checks!$C$32-Checks!$C$33)", f"={H['tv']}", 0.01, "", PX))
chk.append(("Fair value end-2025, house settings (Target!D59)", "=" + fv_expr(2025, 2030, "Checks!$C$31", "Checks!$C$32", "Checks!$C$33", vy=2025, **HOUSEF), f"={H['fv25']}", 0.01, "Flows 2025-2030 + TV.", PX))
chk.append(("Fair value end-2026, house settings (Target!F59)", "=Summary!D28", f"={H['fv26']}", 0.01, "", PX))
chk.append(("Fair value end-2027, house file column (Target!G59)", "=Summary!D29", f"={H['fv27']}", 0.01, "", PX))
chk.append(("Real-rate table, largest gap to the house (Sensitivity)", f"=MAX(MAX(Sensitivity!E7:E{SENS_LAST}),-MIN(Sensitivity!E7:E{SENS_LAST}))", "=0", 0.01, "", PX))
chk.append(("House implied return at the house price (Target!B62)", "=IRR(Checks!C35:G35)", f"={H['irr']}", 1e-6, "Helper row 35.", PCT2))
chk.append(("Corrections switched on (Assumptions section 4)", "=FixNWC+FixND+FixTV+FixInfl", "=0", None, "Info: 0 = the house mechanics.", "0"))
chk.append(("Market data used (1 = house file, 2 = Oct-5)", "=Vintage", "=2", None, "Info.", "0"))
chk.append(("Formula errors (DCF, Summary, Sensitivity)",
            "=SUMPRODUCT(--ISERROR(DCF!C8:S90))+SUMPRODUCT(--ISERROR(Summary!C5:G40))+SUMPRODUCT(--ISERROR(Sensitivity!C7:G60))", "=0", 0, "", "0"))
r = 8
for lab, val, tgt, tol, note, fmt in chk:
    ws.write(f"B{r}", lab, bk.TXT)
    ws.write_formula(f"C{r}", val, F(num_format=fmt))
    ws.write_formula(f"D{r}", tgt, F(num_format=fmt))
    ws.write_formula(f"E{r}", f'=IFERROR(C{r}-D{r},"")', F(num_format=fmt))
    if tol is None:
        ws.write_formula(f"F{r}", '="Info"', F(font_color=GREY, align="center"))
    else:
        ws.write_formula(f"F{r}", f'=IF(ABS(C{r}-D{r})<={tol:.10f},"OK","CHECK")', F(bold=True, align="center"))
    ws.write(f"G{r}", note, bk.NOTE)
    r += 1
ws.write_formula("C5", f'=IF(COUNTIF(F8:F{r},"CHECK")=0,"All checks OK","Some checks need attention")', F(bold=True, font_color=GREEN))
ws.conditional_format(f"F8:F{r}", {"type": "cell", "criteria": "==", "value": '"CHECK"', "format": F(font_color="#C00000", bold=True)})
assert r < 29
bk.section(ws, 30, "House file settings used by the tie-outs (Target sheet)", "B", "G")
for i, (lab, v, fmt) in enumerate([("5-year real rate (Target!C56)", H["rr"], PCT2), ("ERP (Target!B57)", H["erp"], PCT2),
                                    ("g (Target!B54)", H["g"], PCT2), ("Ibovespa (Target!B61)", H["ibov"], PX)]):
    ws.write(f"B{31 + i}", lab, bk.TXT)
    ws.write_number(f"C{31 + i}", v, F(**IN, num_format=fmt))
ws.write("B35", "Helper: house implied-return flows (Target row 61)", bk.NOTE)
ws.write_formula("C35", "=-C34", F(num_format=PTS))
for c, y in zip("DEF", (2026, 2027, 2028)):
    ws.write_formula(f"{c}35", f"=DCF!{YC[y]}{R['fcff_house']}", F(num_format=PTS))
ws.write_formula("G35", f"=DCF!{YC[2030]}{R['fcff_house']}*(1+C33)/(C31+C32-C33)", F(num_format=PTS))
ws.freeze_panes(7, 0)

# =========================================================================== Read Me
ws = W["Read Me"]
ws.set_column("A:A", 2)
ws.set_column("B:B", 26)
ws.set_column("C:C", 130)
ws.set_row(0, 26)
ws.write("B1", "Ibovespa DCF, house model rolled to end-2027", F(bold=True, font_size=16, font_color=NAVY))
ws.write("B2", "XP Research · Equity Strategy · built Oct-7-2026 · inputs from DCF Ibov 2026_Out.xlsx (Sep-26); market data Oct-5-2026", bk.NOTE)
rr_ = 4
for head, lines in [
    ("What it is", [("Replica", "The DCF block of the house Ibovespa model (DCF Ibov 2026_Out.xlsx, Target sheet rows 17-65 and 120-151), with the same inputs, rules and mechanics."),
                    ("2027", "Rolled forward one year: fair value at the end of 2027 from the FCFF of 2028-2031 + terminal value, as the house's end-2026 value uses 2027-2030. "
                             "2031 is projected with the house's 2029-2030 rules (nominal GDP growth, 30% EBITDA margin, 9% capex).")]),
    ("How to use it", [("1. Change", "Assumptions: market data (house file or Oct-5), g, NWC %, scenarios (real rate, ERP), the four correction switches. DCF sheet: pink projection rules."),
                       ("2. Read", "Summary: bear / base / bull, the correction steps and the tie-out with the house file. Sensitivity: real rate x ERP."),
                       ("3. Trace", "Summary -> DCF (valuation block -> FCFF rows -> house inputs). Checks ties every key number to the house file.")]),
    ("Mechanics (house)", [("Cash flow", "FCFF = D&A + NOPAT + capex + NWC, where NWC is the level (-4% of sales), not the change."),
                           ("Discount rate", "5-year real rate (BZRFB5PY) + ERP 6.0%, applied to nominal cash flows."),
                           ("Terminal value", "Last FCFF x (1 + g) / (rate - g), g = 4.7%, put one period after the last flow by Excel NPV."),
                           ("Equity", "The value of the FCFF is read as the index value; net debt is not subtracted.")]),
    ("Corrections", [("Switches", "Assumptions section 4, all off by default (= the house): NWC by change, net debt subtracted (1.6x EBITDA two years before), "
                                  "terminal value discounted with the last flow, inflation in the rate (Fisher). Summary shows the effect step by step.")]),
    ("Sources", [("Inputs", "History 2015-2025 and IBOV Index consensus 2026-2028 (sales and capex to 2027; EBIT, EBITDA and EPS to 2028) as saved in the house file; not refreshed."),
                 ("Market", "Ibovespa 208,431.92 (Oct-5 close); 5-year NTN-B 6.76% (ANBIMA, Oct-5). House file: 182,419.73 and 7.42%."),
                 ("Code", "Brazil Bull Case/scripts/build_house2027.py.")])]:
    ws.merge_range(f"B{rr_}:C{rr_}", head, F(bold=True, font_color="#FFFFFF", bg_color=NAVY))
    rr_ += 1
    for a_, b_ in lines:
        ws.write(f"B{rr_}", a_, F(bold=True, valign="top"))
        ws.write(f"C{rr_}", b_, bk.TXTW)
        n = -(-len(b_) // 170)
        if n > 1:
            ws.set_row(rr_ - 1, 13 * n + 2)
        rr_ += 1
    rr_ += 1

# =========================================================================== Cover
ws = bk.cover("Ibovespa", " | DCF (house model), end-2027", "VALUATION MODEL  ·  EQUITY STRATEGY",
              [("Summary", "Summary"), ("Assumptions", "Assumptions"), ("DCF", "DCF"), ("Sensitivity", "Sensitivity")],
              [[("Checks\n(tie-out with the house)", "Checks"), ("Read Me", "Read Me")]], readme="Read Me")
for i, (l_, f_, nf) in enumerate([("Ibovespa used", "=IbovUsed", PTS), ("Fair value end-2027, base", "=FV27Base", PTS),
                                   ("Upside, base", "=FV27Base/IbovUsed-1", UPS), ("Fair value end-2027, bear", "=FV27Bear", PTS),
                                   ("Fair value end-2027, bull", "=FV27Bull", PTS), ("Model checks", "=Checks!C5", None)]):
    r = 8 + i
    bb = i == 1
    ws.merge_range(f"K{r}:L{r}", l_, F(font_name="Roboto Light", bold=bb, bottom=4))
    fmt_ = F(align="right", bottom=4, bold=bb, num_format=nf) if nf else F(align="right", bottom=4, bold=True, font_color=GREEN)
    ws.merge_range(f"M{r}:N{r}", "", fmt_)
    ws.write_formula(f"M{r}", f_, fmt_)
ws.write("K16", "Equity Strategy | XP Research", F(bold=True))
ws.write("C19", "Replica of the house DCF (DCF Ibov 2026_Out.xlsx) rolled to end-2027: FCFF 2028-2031 + terminal value at 5y NTN-B + 6.0% ERP, g 4.7%.", bk.NOTE)
bk.close()
print("written", OUT)
