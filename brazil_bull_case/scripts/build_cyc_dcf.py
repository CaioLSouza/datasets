"""DCF of the Cyclicals box alone, as in the main model (Ibov_house_2027_by_box_XP.xlsx, DCF by Box, v7b defaults):
house lines x the box's consensus shares, FCFF with the box rules (BoxCF = 1), CAPM rate (BoxRate = 2), Gordon
perpetuity (BoxGrowth = 0), no net debt, no rescale (ScaleBox = 0), scenarios = real rate + earnings level shock.
Output: Cyclicals_DCF_XP.xlsx in the project root.
"""
import datetime as dt
from pathlib import Path

from xp_template import Book, PCT, PCT2, PTS1, UPS, PP, NAVY

HERE = Path(__file__).parent
OUT = HERE.parent / "Cyclicals_DCF_XP.xlsx"
Y = "CDEFGH"                                  # 2026 .. 2031
YRS = list(range(2026, 2032))
HOUSE = {   # DCF Ibov 2026_Out.xlsx (DCF sheet of the main model), index points, columns N..S
    "sales": [158921, 163769, 173726.1552, 184288.70543616, 195493.45872667854, 207379.4610172606],
    "ebit": [33080, 34883, 37322.01, 39591.188208, 41998.3324510464, 44551.83106407003],
    "ebitda": [49324, 50717, 51996.99, 55286.611630848, 58648.03761800356, 62213.83830517817],
    "ni": [20958, 21666, 23177, 24586.1616, 26081.000225279997, 27666.725038977027],
    "capex": [-19159, -19333, -19109.877072, -20271.7575979776, -17594.411285401067, -18664.151491553454],
}
SH = {   # Box Shares section 2 (members' Bloomberg consensus, Oct-7-2026); None = carried from the previous year
    "sales": [0.10558338915495054, 0.11218492273504395] + [None] * 4,
    "ebit": [0.0935153053518142, 0.10053885129004453, 0.10599900182268422] + [None] * 3,
    "ebitda": [0.08548133145043356, 0.09186583026851028, 0.09744481626063714] + [None] * 3,
    "ni": [0.07665992334547189, 0.08891948976656959, 0.09854214837266893] + [None] * 3,
    "capex": [0.1369081089086292, 0.14132059013869358] + [None] * 4,
}

bk = Book(OUT, ["DCF Cyclicals"], prefix="Cyclicals - ")
F, LV, IN, NOTE = bk.F, bk.LV, bk.IN, bk.NOTE
ws = bk.W["DCF Cyclicals"]
bk.title(ws, "DCF (end-2027)",
         "The Cyclicals block of the main model (DCF by Box), on its own: the box's share of each house line, FCFF with "
         "the box rules, CAPM rate, Gordon perpetuity. Same result as Ibov_house_2027_by_box_XP.xlsx.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 66)
ws.set_column("C:H", 12)
ws.set_column("I:I", 2)
ws.set_column("J:J", 100)


def put(r, label, cells, fmt, note="", bold=False, style=None):
    ws.write(f"B{r}", label, F(bold=bold))
    for c, v in cells.items():
        if isinstance(v, str):
            ws.write_formula(f"{c}{r}", v, F(num_format=fmt, bold=bold))
        else:
            ws.write_number(f"{c}{r}", v, F(**{**(style or IN), "bold": bold or (style or IN).get("bold", False)}, num_format=fmt))
    if note:
        ws.write(f"J{r}", note, NOTE)


def years_header(r, first="Item"):
    bk.header_row(ws, r, [("B", first)] + [(c, str(y)) for c, y in zip(Y, YRS)] + [("J", "How it is calculated")])


# ------------------------------------------------------------------ 1. Assumptions
bk.section(ws, 5, "1. Assumptions", "B", "J")
bk.header_row(ws, 6, [("B", "Item"), ("C", "Value"), ("J", "Source / note")])
put(7, "Cyclicals, index points today", {"C": 26271.074646272005}, PTS1,
    "Members' weight x Ibovespa close of Oct-5-2026 (208,432); Box Shares!C9 of the main model.")
put(8, "Long-term growth of the FCFF (g)", {"C": 0.047}, PCT, "House perpetuity growth (DCF Ibov 2026_Out.xlsx, Target!B54).", style=LV)
put(9, "Equity risk premium (ERP)", {"C": 0.06}, PCT, "House ERP (Target!B57).", style=LV)
put(10, "Beta of Cyclicals (weekly, 5 years, vs. the Ibovespa)", {"C": 1.150451096712058}, "0.00",
    "Computed from the local box index (Box Betas!I9 of the main model).")
put(11, "Working capital, % of sales", {"C": 0.04}, PCT, "House assumption (Assumptions!C15).", style=LV)
put(12, "ROE of Cyclicals, FY28 consensus", {"C": 0.17906566532369428}, PCT,
    "Members' BEST_ROE FY28, aggregated as net income / book value (Box Shares!E46). Sets the reinvestment that funds growth.")
put(13, "Earnings beta of Cyclicals to the Ibovespa ex-Commodities", {"C": 1.490738235843181}, "0.00",
    "12-month log change of the 12m fwd EPS index, Oct-07 to Sep-26 (Box Earnings!I9).")
put(14, "Normalization factor of the domestic earnings betas", {"C": 1.1868587775150152}, "0.000",
    "Weighted average beta of Financials, Defensives and Cyclicals = 1 (Box Earnings!I16).")
bk.name("gLT", "DCF Cyclicals", "$C$8")
bk.name("ERP", "DCF Cyclicals", "$C$9")

# ------------------------------------------------------------------ 2. House lines
bk.section(ws, 16, "2. House lines for the Ibovespa (index points; DCF Ibov 2026_Out.xlsx, Target sheet)", "B", "J")
years_header(17, "Line")
R = {"sales": 18, "ebit": 20, "ebitda": 21, "ni": 22, "capex": 23}
lab = {"sales": "Sales", "ebit": "EBIT", "ebitda": "EBITDA", "ni": "Net income", "capex": "Capex"}
for k, r in R.items():
    put(r, lab[k], {c: v for c, v in zip(Y, HOUSE[k])}, PTS1,
        "House model; 2028+ are the house projections." if k == "sales" else "")
put(19, "Sales growth", {c: f"={c}18/{p}18-1" for p, c in zip(Y, Y[1:])}, PCT,
    "2028+: the house nominal GDP growth (6.08%). Drives the reinvestment in section 4.")
put(24, "Tax rate", {c: -0.24 for c in Y}, PCT, "House (-24%).")

# ------------------------------------------------------------------ 3. Cyclicals' share and lines
bk.section(ws, 26, "3. Cyclicals' share of each line (members' consensus, Bloomberg, Oct-7-2026) and the box's lines", "B", "J")
years_header(27, "Line")
S0 = {"sales": 28, "ebit": 29, "ebitda": 30, "ni": 31, "capex": 32}
notes = {"sales": "Consensus 2026-27; 2028+ as 2027 (the house projects sales from 2027).",
         "ebit": "Consensus 2026-28; 2029+ as 2028.", "ebitda": "Consensus 2026-28; 2029+ as 2028.",
         "ni": "Consensus 2026-28; 2029+ as 2028.", "capex": "Consensus 2026-27; 2028+ as 2027."}
for k, r in S0.items():
    cells = {}
    for i, c in enumerate(Y):
        v = SH[k][i]
        cells[c] = v if v is not None else f"={Y[i - 1]}{r}"
    put(r, f"Share of {lab[k].lower()}", cells, PCT2, notes[k])
L = {"sales": 34, "ebit": 35, "ebitda": 36, "ni": 37, "capex": 38}
for k, r in L.items():
    put(r, f"{lab[k]} of Cyclicals", {c: f"={c}{R[k]}*{c}{S0[k]}" for c in Y}, PTS1,
        "House line x share." if k == "sales" else "", bold=k == "ni")

# ------------------------------------------------------------------ 4. FCFF
bk.section(ws, 40, "4. Free cash flow to the firm (box rules)", "B", "J")
years_header(41, "Line")
put(42, "D&A = EBITDA - EBIT", {c: f"={c}36-{c}35" for c in Y}, PTS1)
put(43, "NOPAT = EBIT x (1 + tax rate)", {c: f"={c}35*(1+{c}24)" for c in Y}, PTS1)
put(44, "Working capital = -NWC % x sales", {c: f"=-$C$11*{c}34" for c in Y}, PTS1)
put(45, "Change in working capital", {c: f"={c}44-{p}44" for p, c in zip(Y, Y[1:])}, PTS1)
put(46, "Capex: consensus (2026-27); 2028+ = -(D&A + sales growth x NOPAT / ROE)",
    {"C": "=C38", "D": "=D38", **{c: f"=-({c}42+{c}19*{c}43/$C$12)" for c in Y[2:]}}, PTS1,
    "2028+: replaces the D&A and pays for the growth: reinvestment = growth / ROE of the NOPAT.")
put(47, "FCFF = D&A + NOPAT + capex + change in working capital",
    {"C": "=C42+C43+C46+C44", **{c: f"={c}42+{c}43+{c}46+{c}45" for c in Y[1:]}}, PTS1,
    "2026 uses the working-capital level (no 2025 for the change); only 2028-31 enter the value.", bold=True)

# ------------------------------------------------------------------ 5. Valuation
bk.section(ws, 49, "5. Valuation at the end of 2027", "B", "J")
bk.header_row(ws, 50, [("B", "Item"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), ("J", "How it is calculated")])
V = "CDE"
put(51, "5-year real rate (NTN-B)", {"C": 0.085, "D": 0.0676, "E": 0.055}, PCT2,
    "Bear 8.5% as the house; base = today's 5y real rate; bull 5.5%.", style=LV)
put(52, "Discount rate = real rate + beta x ERP (CAPM)", {c: f"={c}51+$C$10*ERP" for c in V}, PCT2,
    "The house convention: real rate + premium discounts the nominal flows.")
put(53, "Earnings level shock of the Ibovespa ex-Commodities", {"C": -0.2, "D": 0, "E": 0.2}, UPS,
    "Symmetric +/-20% (the house bottom-up uses the same size).", style=LV)
put(54, "Earnings shock of Cyclicals = beta x normalization x shock", {c: f"=$C$13*$C$14*{c}53" for c in V}, UPS,
    "Moves the level of every FCFF, 2028 on.")
put(55, "PV of FCFF 2028-31 = (1 + shock) x NPV(rate, FCFF 2028-31)", {c: f"=(1+{c}54)*NPV({c}52,$E$47:$H$47)" for c in V}, PTS1,
    "Excel NPV: 2028 discounted one year (value at end-2027).")
put(56, "FCFF 2031 in the scenario = FCFF 2031 x (1 + shock)", {c: f"=$H$47*(1+{c}54)" for c in V}, PTS1)
put(57, "Terminal value at end-2031 = FCFF 2031 x (1 + g) / (rate - g)", {c: f"={c}56*(1+gLT)/({c}52-gLT)" for c in V}, PTS1,
    "Gordon perpetuity, as the house.")
put(58, "Periods the terminal value is discounted", {c: 5 for c in V}, "0", "As the house (FixTV = 0): one period more than the last flow.", style=LV)
put(59, "PV of the terminal value", {c: f"={c}57/(1+{c}52)^{c}58" for c in V}, PTS1)
put(60, "Fair value of Cyclicals at the end of 2027 (index pts)", {c: f"={c}55+{c}59" for c in V}, PTS1,
    "No net debt subtracted and no rescale to the house value (defaults of the main model).", bold=True)
put(61, "Upside vs. today", {c: f"={c}60/$C$7-1" for c in V}, UPS, bold=True)
put(62, "Terminal value share of the fair value", {c: f"={c}59/{c}60" for c in V}, PCT)
put(63, "FCFF 2028 yield on today's points", {"D": "=E47/C7"}, PCT, "Compare with the base rate: the gap is the growth the price assumes.")

# ------------------------------------------------------------------ 6. Sensitivity
bk.section(ws, 65, "6. Base upside (no earnings shock): 5-year real rate (rows) x g (columns)", "B", "J")
gs = [0.037, 0.042, 0.047, 0.052, 0.057, 0.062]
ws.write("B66", "Real rate \\ g", bk.HDRL)
for c, g in zip(Y, gs):
    ws.write_number(f"{c}66", g, F(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", num_format=PCT))
for i, rr in enumerate((0.055, 0.06, 0.0676, 0.075, 0.085)):
    r = 67 + i
    ws.write_number(f"B{r}", rr, F(bold=True, num_format=PCT2, align="left"))
    for c in Y:
        k = f"($B{r}+$C$10*ERP)"
        ws.write_formula(f"{c}{r}", f"=(NPV({k},$E$47:$H$47)+$H$47*(1+{c}$66)/({k}-{c}$66)/(1+{k})^$D$58)/$C$7-1",
                         F(num_format=UPS))
ws.conditional_format("C67:H71", {"type": "3_color_scale", "min_color": "#F8696B", "mid_type": "num", "mid_value": 0,
                                  "mid_color": "#FFEB84", "max_color": "#63BE7B"})
ws.freeze_panes(4, 0)
bk.close()
print("saved", OUT)
