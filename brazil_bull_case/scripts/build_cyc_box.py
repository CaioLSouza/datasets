"""Cyclicals box as a sum of the parts: WEG & Embraer (a global part: risk free weighted by revenue geography, as their
analysts) and Cyclicals ex-WEG & Embraer (unified rules), consensus only, 8 explicit years (Caio, Oct-8-2026).

Cash to shareholders with the analysts' bridge in every year (Caio, Oct-8-2026: not the Financials method):
net income + D&A + capex (negative) +/- change in working capital, net debt held constant (the analysts' FCFF bridges,
NOPAT + D&A - capex +/- NWC, taken to the equity holder and discounted at the unified Ke). 2026-28: Bloomberg consensus of
the members, summed in each sub-box. From 2029, projected at the sub-box level only: sales grow from the 2028 consensus
sales growth to g by the last explicit year; net margin and D&A as % of sales held at FY28 (editable), capex
% of sales converging linearly to D&A by the last explicit year (Caio, Oct-8-2026); working
capital 4% of the change in sales. g = IPCA + real GDP = 6% (nominal GDP, as Financials). Ke: (NTN-B + IPCA) x 0.85
+ sub-box beta x 5.5%. Scenarios: real rate 8.5% / today / 5.5%; earnings shock = normalized earnings beta x +/-20%, on
net income (capex does not move with the cycle). Value at end-2027 vs. today's points.
Usage: python build_cyc_box.py [output file name]
"""
import sys
from pathlib import Path

import openpyxl
import pandas as pd
from xlsxwriter.utility import xl_col_to_name as cn

from xp_template import Book, PCT, PCT2, PTS1, UPS

HERE = Path(__file__).parent
OUT = HERE.parent / (sys.argv[1] if len(sys.argv) > 1 else "Cyclicals_DCF_box_XP.xlsx")
BOXF = HERE.parent / "Ibov_house_2027_by_box_XP.xlsx"
YRS = list(range(2026, 2038))
Y = [cn(2 + i) for i in range(12)]          # C .. N
YC = dict(zip(YRS, Y))
E28, N37 = YC[2028], YC[2037]
NOTE_COL = "P"
SUBS = ["WEG & Embraer", "Cyclicals ex-WEG & Embraer"]
SHORT = {"WEG & Embraer": "W&E", "Cyclicals ex-WEG & Embraer": "Ex"}

_mem = [r for r in openpyxl.load_workbook(BOXF, read_only=True, data_only=True)["Members"].iter_rows(min_row=6, values_only=True)
        if r[0] and r[2] == "Cyclicals"]
_roe = pd.read_parquet(HERE / "dcf_members.parquet")
IDX = {"sales": 11, "ebit": 14, "ebitda": 17, "ni": 20, "capex": 23}
MET = ["sales", "ebitda", "ebit", "ni", "capex", "fleet", "roe"]
FLEET_RATIO = 27715 / 38099          # Localiza model 2028: book value of cars sold / gross capex (DCF & Sensitivity r9 / r8)
CONS = []
for r in sorted(_mem, key=lambda x: (x[3] != "WEG & Embraer", -x[9])):
    t = r[0]
    d = {m: [r[IDX[m] + i] if r[IDX[m] + i] is not None else 0.0 for i in range(3)] for m in IDX}
    d["fleet"] = [f"=-{cn(17 + i)}{{r}}*FleetRatio" for i in range(3)] if t in ("RENT3", "VAMO3") else [0.0, 0.0, 0.0]
    roe = [None if pd.isna(_roe.loc[t, f"roe{y}"]) else float(_roe.loc[t, f"roe{y}"]) / 100 for y in (26, 27, 28)]
    d["roe"] = roe
    CONS.append((f"{r[1]} ({t})", t, r[3], r[9], r[10], d))

bk = Book(OUT, ["DCF Cyclicals", "Box Consensus", "Analyst Models"], prefix="Cyclicals - ")
F, LV, IN, NOTE = bk.F, bk.LV, bk.IN, bk.NOTE


def put(ws, r, label, cells, fmt, note="", bold=False, style=None):
    ws.write(f"B{r}", label, F(bold=bold))
    for c, v in cells.items():
        if isinstance(v, str):
            ws.write_formula(f"{c}{r}", v, F(num_format=fmt, bold=bold))
        elif v is not None:
            st = dict(style or IN)
            st["bold"] = bold or st.get("bold", False)
            ws.write_number(f"{c}{r}", v, F(**st, num_format=fmt))
    if note:
        ws.write(f"{NOTE_COL}{r}", note, NOTE)


def years_header(ws, r, first="Line"):
    bk.header_row(ws, r, [("B", first)] + [(c, str(y)) for c, y in zip(Y, YRS)] + [(NOTE_COL, "How it is calculated / source")])


# ============================================================================= Box Consensus
ws = bk.W["Box Consensus"]
bk.title(ws, "consensus 2026-28 of the members, summed into the two sub-boxes",
         "Bloomberg BEST consensus (Oct-7-2026, company totals in R$ mn; Members sheet of the main model; ROE from dcf_members.parquet). "
         "Index points = R$ mn x points / market cap. Only 2026-28 are built from the members.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 30)
ws.set_column("C:C", 22)
ws.set_column("D:Y", 9.5)
ws.write("B4", "Fleet sales / gross capex (Localiza model, 2028)", F())
ws.write_number("D4", FLEET_RATIO, F(**LV, num_format=PCT))
ws.write("E4", "Capex net of the used cars sold, as the analysts' bridge (Caio, Oct-8-2026): Localiza model 2028, book value of cars sold / gross capex = 73%; applied to Localiza and Vamos (rental fleets). 0 = gross capex.", NOTE)
bk.name("FleetRatio", "Box Consensus", "$D$4")
LAB = {"sales": "Sales", "ebitda": "EBITDA", "ebit": "EBIT", "ni": "Net income", "capex": "Capex", "fleet": "Fleet sales", "roe": "ROE"}
MC = {m: [cn(5 + 3 * k + i) for i in range(3)] for k, m in enumerate(MET)}      # F .. Z


def header(r):
    hd = [("B", "Member"), ("C", "Sub-box"), ("D", "Index points"), ("E", "Points / market cap")]
    for m in MET:
        for i, c in enumerate(MC[m]):
            hd.append((c, f"{LAB[m]} {2026 + i}"))
    bk.header_row(ws, r, hd)
    ws.set_row(r - 1, 44)


R1 = 8
bk.section(ws, R1 - 1, "1. Consensus by member (R$ mn; ROE in %)", "B", "Z")
header(R1)
ROWS1 = {}
r = R1 + 1
for name, t, sub, pts, mc, d in CONS:
    ws.write(f"B{r}", name, F())
    ws.write(f"C{r}", sub, F())
    ws.write_number(f"D{r}", pts, F(**IN, num_format=PTS1))
    ws.write_formula(f"E{r}", f"=D{r}/{mc}", F(num_format="0.000000"))
    for m in MET:
        for i, c in enumerate(MC[m]):
            v = d[m][i]
            if isinstance(v, str):
                ws.write_formula(f"{c}{r}", v.format(r=r), F(num_format="#,##0"))
            elif v is None:
                ws.write_blank(f"{c}{r}", None, F())
            else:
                ws.write_number(f"{c}{r}", v, F(**IN, num_format=PCT if m == "roe" else "#,##0"))
    ROWS1[t] = r
    r += 1
L1 = r - 1
R2 = r + 3
bk.section(ws, R2 - 1, "2. In index points (= R$ mn x points / market cap); book value = net income / ROE", "B", "Z")
header(R2)
ws.write(f"{MC['roe'][0]}{R2}", "Book value 2026", bk.HDR)
ws.write(f"{MC['roe'][1]}{R2}", "Book value 2027", bk.HDR)
ws.write(f"{MC['roe'][2]}{R2}", "Book value 2028", bk.HDR)
r = R2 + 1
for name, t, sub, pts, mc, d in CONS:
    r1 = ROWS1[t]
    ws.write(f"B{r}", name, F())
    ws.write_formula(f"C{r}", f"=C{r1}", F())
    ws.write_formula(f"D{r}", f"=D{r1}", F(num_format=PTS1))
    for m in MET[:-1]:
        for c in MC[m]:
            ws.write_formula(f"{c}{r}", f"={c}{r1}*$E${r1}", F(num_format=PTS1))
    for i, c in enumerate(MC["roe"]):
        ni = MC["ni"][i]
        ws.write_formula(f"{c}{r}", f'=IF(ISNUMBER({c}{r1}),{ni}{r1}*$E${r1}/{c}{r1},0)', F(num_format=PTS1))
    r += 1
L2 = r - 1
TOT = {}
for k, s in enumerate(SUBS + ["Cyclicals"]):
    rr = L2 + 1 + k
    TOT[s] = rr
    ws.write(f"B{rr}", s, F(bold=True, top=1 if k == 0 else 0))
    crit = f"$C${R2 + 1}:$C${L2},\"{s}\""
    if s == "Cyclicals":
        ws.write_formula(f"D{rr}", f"=SUM(D{R2 + 1}:D{L2})", F(num_format=PTS1, bold=True))
    else:
        ws.write_formula(f"D{rr}", f"=SUMIFS(D{R2 + 1}:D{L2},{crit})", F(num_format=PTS1, bold=True, top=1 if k == 0 else 0))
    for m in MET:
        for c in MC[m]:
            rng = f"{c}{R2 + 1}:{c}{L2}"
            f_ = f"=SUM({rng})" if s == "Cyclicals" else f"=SUMIFS({rng},{crit})"
            ws.write_formula(f"{c}{rr}", f_, F(num_format=PTS1, bold=True, top=1 if k == 0 else 0))
rr = L2 + 4
ws.write(f"B{rr}", "Check: net income vs. Box Shares!M9:O9 (Cyclicals) of the main model", F())
for c, v in zip(MC["ni"], [1607.0455, 1944.9107, 2304.3006]):
    ws.write_formula(f"{c}{rr}", f"={c}{TOT['Cyclicals']}-{v}", F(num_format=PTS1))
ws.write(f"B{rr + 2}", "Members without a BEST_ROE get book value 0: the sub-box ROE uses the members with an ROE (box ROE = sum of net income "
         "with ROE / sum of book value; see DCF Cyclicals). Leases (IFRS 16) and concession fees are not in the consensus: pending the Bloomberg pull.", NOTE)
ws.freeze_panes(R1, 3)
BC = "'Box Consensus'!"

# ============================================================================= Analyst Models (memo)
ws = bk.W["Analyst Models"]
bk.title(ws, "how the XP Cyclicals analysts value their companies (memo)",
         "Files in inputs/analyst_models/Cyclicals (DCF & Sensitivity / Valuation sheets), read on Oct-8-2026.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 20)
ws.set_column("C:I", 18)
hd = ["Model (weight in the box)", "Currency / method", "Explicit to", "Perpetuity g", "Ke (beta)", "WACC", "Cash flow", "Note"]
for i, h in enumerate(hd):
    ws.write(4, 1 + i, h, bk.HDRL if i == 0 else bk.HDR)
ws.set_row(4, 30)
rows = [("WEG (20.5%)", "BRL; FCFF - net debt", "2035", "7.6% (about 3.4% real)", "11.4% (0.8)", "11.2%", "NOPAT + D&A - capex +/- NWC", "Permanent real growth"),
        ("Embraer (18.3%)", "USD; FCFF - net debt", "~2035", "2.4% (real 0%)", "10.6% (1.0)", "9.8%", "NOPAT + D&A - capex +/- NWC", "Exporter"),
        ("Localiza (13.4%)", "BRL; FCFF / FCFE", "2035", "6.1% (about 2% real, GDP-linked)", "15.2% (0.95)", "12.3%",
         "NOPAT + D&A - capex + BV of cars sold +/- NWC", "Fleet net of used-car sales (resale at 79% of purchase price)"),
        ("Rumo (7.2%)", "BRL; FCFE to concession end", "~2070", "4.0% (real 0%)", "15.3% (1.0)", "11.7%", "FCFF - concession fees + net funding",
         "Lucas do Rio Verde project apart"),
        ("Motiva (5.4%)", "BRL; sum of concessions", "Each contract", "0% (finite lives)", "14.7% (0.92)", "10.8%", "FCFF by concession", "No perpetuity"),
        ("Totvs (5.2%)", "BRL; FCFE", "2035", "Exit multiple on 2035 FCFE", "14.4% (0.8)", "-", "Operating cash flow - capex - lease principal", "Ke without the 0.85"),
        ("Lojas Renner (3.6%)", "BRL; FCFF", "~2035", "4.0% (real 0%)", "14.1% (0.95)", "13.4%", "EBIT - taxes + D&A - capex +/- NWC", "USD + country-risk Ke"),
        ("Cyrela (2.3%)", "BRL; FCFF", "-", "4.5% (real 0%)", "20.0% (1.46)", "17.5%", "NOPAT + D&A - capex +/- NWC", "ERP 7.7%")]
for k, row in enumerate(rows):
    for i, v in enumerate(row):
        ws.write(5 + k, 1 + i, v, F(text_wrap=True, valign="top", **({} if i else {"bold": True})))
    ws.set_row(5 + k, 32)

# ============================================================================= DCF Cyclicals
ws = bk.W["DCF Cyclicals"]
ws.set_column("A:A", 2)
ws.set_column("B:B", 66)
ws.set_column("C:N", 10.5)
ws.set_column("O:O", 2)
ws.set_column("P:P", 90)
bk.title(ws, "DCF by sub-box with the analysts' cash-flow method (end-2027)",
         "Cash to shareholders = net income + D&A - capex +/- NWC (the analysts' bridge, net debt constant): consensus 2026-28, then sales, margin, "
         "D&A and capex projected for each sub-box. "
         "Two sub-boxes, summed. Ke unified with the other boxes.")
bk.section(ws, 5, "1. Inputs", "B", NOTE_COL)
bk.header_row(ws, 6, [("B", "Item"), ("C", "W&E"), ("D", "Ex"), (NOTE_COL, "Source / note")])
put(ws, 7, "Index points today", {"C": f"={BC}D{TOT['WEG & Embraer']}", "D": f"={BC}D{TOT['Cyclicals ex-WEG & Embraer']}"}, PTS1,
    "W&E = WEG & Embraer; Ex = Cyclicals ex-WEG & Embraer (Box Consensus).")
put(ws, 8, "Long-term IPCA", {"C": 0.04}, PCT, "House and the analysts' Macro sheets.", style=LV)
put(ws, 9, "Share of the risk free kept after tax", {"C": 0.85}, "0.00", "As the analysts: risk free = (NTN-B + IPCA) x 0.85.", style=LV)
put(ws, 10, "Equity risk premium, the same for every box", {"C": 0.055}, PCT2, "Unified (Caio, Oct-8-2026).", style=LV)
put(ws, 11, "Earnings shock per 1% of the Ibovespa ex-Commodities shock (beta x normalization)", {"C": 0.127107 * 1.1868588, "D": 1.510297 * 1.1868588},
    "0.000", "Box Earnings!I13:I14 x I16 of the main model (WEG & Embraer 0.13, ex 1.51, normalized by 1.19).")
put(ws, 12, "Working capital, % of sales", {"C": 0.04}, PCT, "House assumption.", style=LV)
put(ws, 13, "Net margin from 2029 (blank = FY28 consensus of the sub-box)", {}, PCT, "Typed value replaces the sub-box margin.", style=LV)
put(ws, 14, "Capex / D&A in the last explicit year (capex % of sales goes linearly from FY28 to it)", {"C": 1.0, "D": 1.0}, '0.00"x"',
    "Caio, Oct-8-2026: capex converges to D&A (maintenance) by the last explicit year, so the perpetuity does not carry expansion capex forever.", style=LV)
put(ws, 15, "D&A, % of sales, from 2029 (blank = FY28 consensus)", {}, PCT, "Typed value replaces the sub-box ratio.", style=LV)
put(ws, 16, "Growth after 2028 starts from the 2028 consensus sales growth and goes linearly to g in the last explicit year", {}, PCT,
    "As the analysts: the top line drives the model; margins and capital intensity at FY28.")
put(ws, 17, "WEG & Embraer (global part): Brazil share of the real risk free", {"C": (5386.4 * 0.35 + 4818.2 * 0.20) / 10204.6}, "0%",
    "Analysts' weights by revenue geography: WEG 35% Brazil (DCF & Sensitivity C55), Embraer 20% (C66); weighted by index points.", style=LV)
put(ws, 18, "US real risk free", {"C": 0.021}, PCT2, "Analysts: WEG 2.08% (F56), Embraer 2.51% (F69).", style=LV)
ws.write_blank("C13", None, F(**LV, num_format=PCT))
ws.write_blank("D13", None, F(**LV, num_format=PCT))
ws.write_blank("C15", None, F(**LV, num_format=PCT))
ws.write_blank("D15", None, F(**LV, num_format=PCT))

bk.section(ws, 19, "2. House vs. analysts, variable by variable (choice: 1 = house, 2 = analysts)", "B", NOTE_COL)
bk.header_row(ws, 20, [("B", "Variable"), ("C", "House W&E"), ("D", "House Ex"), ("E", "Analysts / GDP W&E"), ("F", "Analysts / GDP Ex"), ("G", "Choice"),
                       ("H", "Used W&E"), ("I", "Used Ex"), (NOTE_COL, "House / analysts")])
ws.set_row(19, 30)
ws.set_column("C:I", 11)
HV = dict(beta=21, rg=22, n=23, tv=24)
CHOICE = dict(beta=1, rg=2, n=2, tv=2)     # Caio, Oct-8-2026: 8 explicit years
vals = {"beta": (0.7214, 1.3240, 0.89, 1.02, "0.000",
                 "House: sub-box betas, weekly 5 years from the local indices (Box Betas). Analysts: WEG 0.8 / Embraer 1.0; ex: Localiza 0.95, Rumo 1.0, Motiva 0.92, Renner 0.95, Totvs 0.8, Cyrela 1.46."),
        "rg": (0.007, 0.007, 0.02, 0.02, PCT,
               "Long-term real growth. House: g 4.7% = about 0.7% real over a 4% IPCA. Nominal GDP (Caio, Oct-8-2026): real GDP 2%, as Financials - "
               "Cyclicals sell to the domestic economy and should grow with nominal GDP in the perpetuity."),
        "n": (4, 4, 8, 8, "0", "Explicit years after end-2027. House: 2028-31. Analysts: to 2035."),
        "tv": (1, 1, 0, 0, "0", "Extra periods in the discount of the terminal value. House: one more than the last flow. Analysts: none.")}
LABS = {"beta": "Beta", "rg": "Long-term real growth (g = IPCA + real growth)", "n": "Explicit years after the value date",
        "tv": "Extra periods in the discount of the terminal value"}
for k, r in HV.items():
    a, b, c, d, fm, note = vals[k]
    ws.write(f"B{r}", LABS[k], F())
    for col, v in zip("CDEF", (a, b, c, d)):
        ws.write_number(f"{col}{r}", v, F(**LV, num_format=fm))
    ws.write_number(f"G{r}", CHOICE[k], F(**LV, num_format="0", align="center"))
    ws.write_formula(f"H{r}", f"=IF($G{r}=1,C{r},E{r})", F(num_format=fm, bold=True))
    ws.write_formula(f"I{r}", f"=IF($G{r}=1,D{r},F{r})", F(num_format=fm, bold=True))
    ws.write(f"{NOTE_COL}{r}", note, NOTE)
put(ws, 25, "g used = IPCA + real growth (choice 2 = nominal GDP 6.0%, as Financials)", {"H": f"=$C$8+H{HV['rg']}", "I": f"=$C$8+I{HV['rg']}"}, PCT, bold=True)
ws.write_formula(f"H{HV['n']}", f"=IF($G{HV['n']}=1,C{HV['n']},E{HV['n']})", F(num_format="0", bold=True))
bk.name("NExp", "DCF Cyclicals", f"$H${HV['n']}")
bk.name("TVx", "DCF Cyclicals", f"$H${HV['tv']}")

bk.section(ws, 27, "3. Cost of equity by scenario", "B", NOTE_COL)
bk.header_row(ws, 28, [("B", "Item"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), (NOTE_COL, "How it is calculated")])
V = "CDE"
put(ws, 29, "5-year real rate (NTN-B)", {"C": 0.085, "D": 0.0676, "E": 0.055}, PCT2, "Same scenarios as the other boxes.", style=LV)
put(ws, 30, "Risk free, Brazil = (NTN-B + IPCA) x share kept after tax", {c: f"=({c}29+$C$8)*$C$9" for c in V}, PCT2, "The rest of the index (unified).")
put(ws, 31, "Risk free, global (WEG & Embraer) = (Brazil share x NTN-B + US share x US real + IPCA) x share kept",
    {c: f"=($C$17*{c}29+(1-$C$17)*$C$18+$C$8)*$C$9" for c in V}, PCT2,
    "As the WEG and Embraer analysts: real risk free weighted by revenue geography; nominal in R$ with the IPCA.")
put(ws, 32, "Ke, WEG & Embraer = global risk free + beta x ERP", {c: f"={c}31+$H${HV['beta']}*$C$10" for c in V}, PCT2, bold=True)
put(ws, 33, "Ke, Cyclicals ex-WEG & Embraer = Brazil risk free + beta x ERP", {c: f"={c}30+$I${HV['beta']}*$C$10" for c in V}, PCT2, bold=True)

# ---- 4. Projection, one block per sub-box
P = {}
r0 = 35
for j, s in enumerate(SUBS):
    tot = TOT[s]
    gcol = "H" if j == 0 else "I"
    g_ = f"${gcol}$25"
    io = "C" if j == 0 else "D"
    bk.section(ws, r0, f"4{'ab'[j]}. {s}: projection (index points)", "B", NOTE_COL)
    years_header(ws, r0 + 1)
    R = dict(t=r0 + 2, gr=r0 + 3, sales=r0 + 4, mg=r0 + 5, ni=r0 + 6, dap=r0 + 7, da=r0 + 8, cxp=r0 + 9, capex=r0 + 10,
             nwc=r0 + 11, fcfe=r0 + 12, po=r0 + 13, roe=r0 + 14, rr=r0 + 15)
    P[s] = R
    put(ws, R["t"], "Years after end-2027 (grey = after the explicit years)", {c: f"={YRS[i]}-2027" for i, c in enumerate(Y) if YRS[i] >= 2028}, "0")
    ws.conditional_format(f"{E28}{R['t']}:{N37}{R['t']}", {"type": "formula", "criteria": f"={E28}${R['t']}>NExp",
                                                             "format": bk.wb.add_format({"font_color": "#A6A6A6"})})
    gr = {c: f"={c}{R['sales']}/{p}{R['sales']}-1" for p, c in zip(Y[:2], Y[1:3])}
    g0 = f"${E28}${R['gr']}"
    gr.update({c: f"=IF({c}{R['t']}>=NExp,{g_},{g0}+({g_}-{g0})*({c}{R['t']}-1)/(NExp-1))" for c in Y[3:]})
    put(ws, R["gr"], "Sales growth: 2027-28 consensus; then from the 2028 growth linearly to g in the last explicit year", gr, PCT)
    sales = {c: f"={BC}{MC['sales'][i]}{tot}" for i, c in enumerate(Y[:3])}
    sales.update({c: f"={p}{R['sales']}*(1+{c}{R['gr']})" for p, c in zip(Y[2:], Y[3:])})
    put(ws, R["sales"], "Sales", sales, PTS1, "2026-28: Box Consensus (sum of the members).")
    mg = {c: f"={BC}{MC['ni'][i]}{tot}/{BC}{MC['sales'][i]}{tot}" for i, c in enumerate(Y[:3])}
    mg.update({c: f"=IF(ISNUMBER(${io}$13),${io}$13,${E28}${R['mg']})" for c in Y[3:]})
    put(ws, R["mg"], "Net margin (consensus; held at FY28 unless typed in the inputs)", mg, PCT)
    put(ws, R["ni"], "Net income = sales x net margin", {c: f"={c}{R['sales']}*{c}{R['mg']}" for c in Y}, PTS1, bold=True)
    dap = {c: f"=({BC}{MC['ebitda'][i]}{tot}-{BC}{MC['ebit'][i]}{tot})/{BC}{MC['sales'][i]}{tot}" for i, c in enumerate(Y[:3])}
    dap.update({c: f"=IF(ISNUMBER(${io}$15),${io}$15,${E28}${R['dap']})" for c in Y[3:]})
    put(ws, R["dap"], "D&A, % of sales (EBITDA - EBIT, consensus; held at FY28)", dap, PCT)
    put(ws, R["da"], "(+) D&A", {c: f"={c}{R['sales']}*{c}{R['dap']}" for c in Y}, PTS1)
    cxp = {c: f"=-({BC}{MC['capex'][i]}{tot}+{BC}{MC['fleet'][i]}{tot})/{BC}{MC['sales'][i]}{tot}" for i, c in enumerate(Y[:3])}
    cxp.update({c: f"=IF({c}{R['t']}>=NExp,{c}{R['dap']}*${io}$14,${E28}${R['cxp']}+({c}{R['dap']}*${io}$14-${E28}${R['cxp']})*({c}{R['t']}-1)/(NExp-1))" for c in Y[3:]})
    put(ws, R["cxp"], "Capex, % of sales (consensus; then linearly to D&A x C14 in the last explicit year)", cxp, PCT)
    put(ws, R["capex"], "(-) Capex", {c: f"=-{c}{R['sales']}*{c}{R['cxp']}" for c in Y}, PTS1)
    put(ws, R["nwc"], "(+/-) Change in working capital = -NWC % x change in sales",
        {c: f"=-$C$12*({c}{R['sales']}-{p}{R['sales']})" for p, c in zip(Y[:-1], Y[1:])}, PTS1)
    put(ws, R["fcfe"], "Cash to shareholders = net income + D&A - capex +/- NWC", {c: f"={c}{R['ni']}+{c}{R['da']}+{c}{R['capex']}+{c}{R['nwc']}" for c in Y[1:]},
        PTS1, "The analysts' bridge, with net debt held constant (no new debt funds the capex); leases pending the Bloomberg pull.", bold=True)
    put(ws, R["po"], "Cash to shareholders / net income", {c: f"={c}{R['fcfe']}/{c}{R['ni']}" for c in Y[1:]}, PCT)
    roe = {c: f"={BC}{MC['ni'][i]}{tot}/{BC}{MC['roe'][i]}{tot}" for i, c in enumerate(Y[:3])}
    put(ws, R["roe"], "Memo: ROE (consensus)", roe, PCT, "Sum of net income / sum of book value (net income / BEST_ROE).")
    put(ws, R["rr"], "Memo: implied return on reinvestment = sales growth / (1 - cash / net income)",
        {c: f"=IFERROR({c}{R['gr']}/(1-{c}{R['po']}),\"\")" for c in Y[3:]}, PCT,
        "Sanity check of the perpetuity: the return the reinvested cash must earn to deliver the growth; compare with the ROE.")
    r0 += 17

# ---- 5. Valuation
rv = r0
bk.section(ws, rv, "5. Valuation at the end of 2027", "B", NOTE_COL)
bk.header_row(ws, rv + 1, [("B", "Item"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), (NOTE_COL, "How it is calculated")])
put(ws, rv + 2, "Earnings shock of the Ibovespa ex-Commodities", {"C": -0.2, "D": 0, "E": 0.2}, UPS, "Symmetric +/-20%, as the other boxes.", style=LV)
VAL = {}
r = rv + 3
for j, s in enumerate(SUBS):
    R = P[s]
    kerow = 32 + j
    g_ = f"${'H' if j == 0 else 'I'}$25"
    FL = f"${E28}${R['fcfe']}:${N37}${R['fcfe']}"
    NL = f"${E28}${R['ni']}:${N37}${R['ni']}"
    TT = f"${E28}${R['t']}:${N37}${R['t']}"
    sh = f"{'$C$11' if j == 0 else '$D$11'}"
    put(ws, r, f"{s}: earnings shock = beta x normalization x shock", {c: f"={sh}*{c}${rv + 2}" for c in V}, UPS)
    pvx = {c: f"=SUMPRODUCT({FL},--({TT}<=NExp),1/(1+{c}{kerow})^{TT})+{c}{r}*SUMPRODUCT({NL},--({TT}<=NExp),1/(1+{c}{kerow})^{TT})" for c in V}
    put(ws, r + 1, f"{s}: PV of the explicit flows (+ shock x PV of net income)", pvx, PTS1)
    tv = {c: f"=(INDEX({FL},NExp)+{c}{r}*INDEX({NL},NExp))*(1+{g_})/({c}{kerow}-{g_})/(1+{c}{kerow})^(NExp+TVx)" for c in V}
    put(ws, r + 2, f"{s}: PV of the terminal value", tv, PTS1, "Gordon on the last explicit flow.")
    put(ws, r + 3, f"{s}: fair value (index pts)", {c: f"={c}{r + 1}+{c}{r + 2}" for c in V}, PTS1, bold=True)
    put(ws, r + 4, f"{s}: upside", {c: f"={c}{r + 3}/${'C' if j == 0 else 'D'}$7-1" for c in V}, UPS, bold=True)
    VAL[s] = r + 3
    r += 6
put(ws, r, "Cyclicals: fair value (index pts) = sum of the parts (WEG & Embraer, global + the rest, unified)", {c: f"={c}{VAL[SUBS[0]]}+{c}{VAL[SUBS[1]]}" for c in V}, PTS1, bold=True)
put(ws, r + 1, "Cyclicals: upside vs. today", {c: f"={c}{r}/($C$7+$D$7)-1" for c in V}, UPS, bold=True)
put(ws, r + 2, "Memo: upside in the unified DCF by box (main model, v7b)", {"C": -0.5811, "D": -0.2218, "E": 0.2317}, UPS, "Cyclicals_DCF_XP.xlsx.")
put(ws, r + 3, "Memo: upside in the multiples model (Oct-6)", {"C": 0.02, "D": 0.40, "E": 0.85}, UPS, "Ibov_valuation_by_box_model_XP.xlsx.")
ROWS_OUT = dict(total=r + 1, we=VAL[SUBS[0]] + 1, ex=VAL[SUBS[1]] + 1)

r += 5
bk.section(ws, r, "6. What is fixed (not switchable) and where it comes from", "B", NOTE_COL)
fixed = [("Risk free nominal: (NTN-B + IPCA) x 0.85; ERP 5.5% for every box", "Unified (Caio)"),
         ("Sum of the parts: WEG & Embraer valued as a global part (risk free by revenue geography, as their analysts); the rest unified", "Caio, Oct-8-2026"),
         ("Consensus only (2026-28 sales, net income, EBITDA, EBIT, capex); no company-specific adjustments", "Caio"),
         ("Cash to shareholders = net income + D&A - capex +/- NWC, net debt constant (the analysts' bridge); perpetuity on the last explicit flow", "Analysts (Caio, Oct-8-2026)"),
         ("From 2029: sales from the 2028 growth to g; net margin and D&A % of sales at FY28; capex converges to D&A by the last explicit year", "Box level only (Caio)"),
         ("Scenario real rates 8.5% / today / 5.5%; earnings shock +/-20% x earnings beta, on net income", "House / main model"),
         ("Pending the Bloomberg pull: leases (IFRS 16) and concession fees", "Not in the consensus")]
for i, (a, b) in enumerate(fixed):
    ws.write(f"B{r + 1 + i}", a, F())
    ws.write(f"{NOTE_COL}{r + 1 + i}", b, NOTE)
ws.freeze_panes(4, 2)
ws.activate()
bk.close()
print("saved", OUT, ROWS_OUT)
