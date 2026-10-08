"""Simplified DCF of the Cyclicals box (FCFE, value-driver form, market-calibrated cost of equity).

Value = earnings x (1 - g / ROE) / (Ke - g). Ke = house Ke (5y real rate + ERP) + the box's spread to the Ibovespa
implied Ke, moved part of the way from today's spread to its 10-year median. Scenarios: real rate (through the
analysts' Ke sensitivity) and the earnings level (earnings beta x Ibovespa ex-Commodities shock).
Output: Cyclicals_DCF_simple_XP.xlsx in the project root.
"""
import datetime as dt
from pathlib import Path

import pandas as pd

from xp_template import Book, PCT, PCT2, PTS1, UPS, PP, MULT, DATE, DATED, NAVY

HERE = Path(__file__).parent
OUT = HERE.parent / "Cyclicals_DCF_simple_XP.xlsx"

# ---------------------------------------------------------------- history: 12m fwd P/E and ROE, Cyclicals and Ibovespa
h = pd.read_parquet(HERE / "dcf_roe_hist.parquet").dropna(subset=["pe", "roe"])
h = h[(h.pe > 0) & (h.roe > 0)].copy()
h["e"] = h.w / h.pe
h["b"] = h.e / (h.roe / 100)


def agg(d):
    return pd.Series({"pe": d.w.sum() / d.e.sum(), "roe": d.e.sum() / d.b.sum()})


cyc = h[h.box == "Cyclicals"].groupby("date").apply(agg)
ibo = h.groupby("date").apply(agg)
hist = cyc.join(ibo, lsuffix="_c", rsuffix="_i").sort_index()

bk = Book(OUT, ["DCF Cyclicals", "History"], prefix="Cyclicals - ")
F = bk.F
LV, IN = bk.LV, bk.IN

# ---------------------------------------------------------------- History sheet
ws = bk.W["History"]
bk.title(ws, "implied cost of equity, 10 years",
         "Monthly 12m fwd consensus of the Ibovespa members (Bloomberg BEST_PE_RATIO / BEST_ROE), aggregated by box: "
         "P/E = sum of weights / sum of earnings; ROE = sum of earnings / sum of book value.")
for c, wdt in zip("ABCDEFGHI", (2, 11, 12, 12, 12, 12, 14, 14, 14)):
    ws.set_column(f"{c}:{c}", wdt)
bk.section(ws, 5, "Implied Ke = g + (1 - g / ROE) / (P/E)   (Gordon in value-driver form, g from DCF Cyclicals)", "B", "I")
bk.header_row(ws, 6, [("B", "Month-end"), ("C", "Cyclicals P/E"), ("D", "Cyclicals ROE"), ("E", "Ibovespa P/E"),
                      ("F", "Ibovespa ROE"), ("G", "Cyclicals implied Ke"), ("H", "Ibovespa implied Ke"),
                      ("I", "Spread, Cyclicals - Ibovespa")])
ws.set_row(5, 30)
H0 = 7
for i, (d, r) in enumerate(hist.iterrows()):
    x = H0 + i
    ws.write_datetime(f"B{x}", d.to_pydatetime(), F(num_format=DATE))
    ws.write_number(f"C{x}", r.pe_c, F(**IN, num_format=MULT))
    ws.write_number(f"D{x}", r.roe_c, F(**IN, num_format=PCT))
    ws.write_number(f"E{x}", r.pe_i, F(**IN, num_format=MULT))
    ws.write_number(f"F{x}", r.roe_i, F(**IN, num_format=PCT))
    ws.write_formula(f"G{x}", f"=gLT+(1-gLT/D{x})/C{x}", F(num_format=PCT2))
    ws.write_formula(f"H{x}", f"=gLT+(1-gLT/F{x})/E{x}", F(num_format=PCT2))
    ws.write_formula(f"I{x}", f"=(G{x}-H{x})*100", F(num_format='+0.00" pp";-0.00" pp"'))
HL = H0 + len(hist) - 1
ws.freeze_panes(H0 - 1, 2)

# ---------------------------------------------------------------- DCF Cyclicals sheet
ws = bk.W["DCF Cyclicals"]
bk.title(ws, "simplified DCF (FCFE)",
         "Value = earnings x (1 - g / ROE) / (Ke - g): the cash left to shareholders after the reinvestment that funds "
         "growth g, discounted at the box's cost of equity. No net debt (equity flows).")
ws.set_column("A:A", 2)
ws.set_column("B:B", 72)
ws.set_column("C:E", 13)
ws.set_column("F:F", 95)
NOTE = bk.NOTE


def row(r, label, cells, fmt, note="", bold=False):
    ws.write(f"B{r}", label, F(bold=bold))
    for c, v in cells.items():
        f = fmt if not isinstance(fmt, dict) else fmt[c]
        if isinstance(v, str) and v.startswith("="):
            ws.write_formula(f"{c}{r}", v, F(num_format=f, bold=bold))
        elif isinstance(v, tuple):          # (value, style) for inputs / levers
            val, sty = v
            if isinstance(val, dt.datetime):
                ws.write_datetime(f"{c}{r}", val, F(**sty, num_format=f))
            else:
                ws.write_number(f"{c}{r}", val, F(**sty, num_format=f))
    if note:
        ws.write(f"F{r}", note, NOTE)


# 1. Assumptions
bk.section(ws, 5, "1. Assumptions", "B", "F")
bk.header_row(ws, 6, [("B", "Item"), ("C", "Value"), ("D", ""), ("E", ""), ("F", "Source / note")])
row(7, "Pricing date", {"C": (dt.datetime(2026, 10, 5), IN)}, DATED, "Close of Oct-5-2026, as in the multiples model.")
row(8, "Target date", {"C": (dt.datetime(2027, 12, 31), LV)}, DATED, "End-2027, as the house target model rolled to 2027.")
row(9, "Years from pricing to target", {"C": "=(C8-C7)/365"}, "0.00")
row(10, "Cyclicals, index points today", {"C": (26271.0746, IN)}, PTS1,
    "Members' weight x Ibovespa close 208,432 (Box Shares sheet of Ibov_house_2027_by_box_XP.xlsx).")
row(11, "Long-term nominal growth (g)", {"C": (0.047, LV)}, PCT, "House perpetuity growth (DCF Ibov 2026_Out.xlsx, Target!B54).")
row(12, "Equity risk premium (ERP)", {"C": (0.06, LV)}, PCT, "House ERP (Target!B57).")
row(13, "Share of the gap to the 10-year median spread closed by the target", {"C": (0.5, LV)}, "0%",
    "0% = Cyclicals keep today's spread to the Ibovespa; 100% = back to the 10-year median.")
row(14, "Earnings level shock, Ibovespa ex-Commodities (bull +, bear -)", {"C": (0.20, LV)}, PCT,
    "Symmetric, as agreed (house bottom-up uses +/-20%).")
row(15, "Earnings beta of Cyclicals to the Ibovespa ex-Commodities", {"C": (1.4907, IN)}, "0.00",
    "12-month log change of the 12m fwd EPS index, Oct-07 to Sep-26 (Box Earnings sheet).")
row(16, "Normalization factor of the domestic earnings betas", {"C": (1.1869, IN)}, "0.000",
    "Makes the weighted average beta of Financials, Defensives and Cyclicals = 1 (Box Earnings!I16).")
row(17, "Change in target price per 100 bp of Ke (analysts' models)", {"C": (0.174, IN)}, PCT,
    "XP analysts' models, Ke -100 bp, Cyclicals box (Ke_sensitivity_by_box_2026-10-06.xlsx).")
for n, ref in (("gLT", "$C$11"), ("ERP", "$C$12")):
    bk.name(n, "DCF Cyclicals", ref)

# 2. Implied Ke
bk.section(ws, 19, "2. Cost of equity the market implies today (12m fwd consensus, last month of History)", "B", "F")
bk.header_row(ws, 20, [("B", "Item"), ("C", "Cyclicals"), ("D", "Ibovespa"), ("E", ""), ("F", "Note")])
row(21, "P/E, 12m fwd", {"C": f"=History!C{HL}", "D": f"=History!E{HL}"}, MULT, f"History, {hist.index[-1]:%b-%y}.")
row(22, "ROE, 12m fwd", {"C": f"=History!D{HL}", "D": f"=History!F{HL}"}, PCT)
row(23, "Implied Ke = g + (1 - g / ROE) / (P/E)", {"C": "=gLT+(1-gLT/C22)/C21", "D": "=gLT+(1-gLT/D22)/D21"}, PCT2,
    "The discount rate at which the value-driver formula gives today's price.")
row(24, "Spread to the Ibovespa today (pp)", {"C": "=(C23-D23)*100"}, PP)
row(25, "Spread to the Ibovespa, 10-year median (pp)", {"C": f"=MEDIAN(History!I{H0}:I{HL})"}, PP,
    "Median, not average: the Cyclicals series is noisy when ROE is close to g.")
row(26, "Spread used = today + share closed x (median - today) (pp)", {"C": "=C24+C13*(C25-C24)"}, PP, bold=True)

# 3. Scenarios
bk.section(ws, 28, "3. Scenarios", "B", "F")
bk.header_row(ws, 29, [("B", "Item"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), ("F", "Note")])
S = "CDE"
row(30, "5-year real rate (NTN-B)", {"C": (0.085, LV), "D": (0.0676, LV), "E": (0.055, LV)}, PCT2,
    "Bear 8.5% as the house; base = today's 5y real rate; bull 5.5%.")
row(31, "Fair Ke of Cyclicals = real rate + ERP + spread used", {c: f"={c}30+ERP+$C$26/100" for c in S}, PCT2)
row(32, "Change in the real rate vs the base (pp)", {c: f"=({c}30-$D$30)*100" for c in S}, PP)
row(33, "Base valuation factor = (implied Ke - g) / (fair Ke base - g)", {c: "=($C$23-gLT)/($D$31-gLT)" for c in S}, "0.000",
    "Same in every scenario: how far today's price is from the base fair value.")
row(34, "Effective Ke - g from the analysts' sensitivity = 1% / sensitivity (pp)", {c: "=1/$C$17" for c in S}, "0.00",
    "Gordon form that reproduces the analysts' change per 100 bp of Ke.")
row(35, "Rate factor = (Ke - g) / (Ke - g + change in the real rate)", {c: f"={c}34/({c}34+{c}32)" for c in S}, "0.000")
row(36, "Earnings shock, Ibovespa ex-Commodities", {"C": "=-$C$14", "D": "=0", "E": "=$C$14"}, UPS)
row(37, "Earnings shock of Cyclicals = beta x normalization x shock", {c: f"=$C$15*$C$16*{c}36" for c in S}, UPS,
    "Moves the level of earnings and of the value one to one.")
row(38, "Fair value today (index pts)", {c: f"=$C$10*{c}33*{c}35*(1+{c}37)" for c in S}, PTS1,
    "Points x base factor x rate factor x (1 + earnings shock).")
row(39, "Fair value at the target date (index pts) = today x (1 + g) ^ years", {c: f"={c}38*(1+gLT)^$C$9" for c in S}, PTS1,
    "In the Gordon model the value grows at g once the dividends are paid.", bold=True)
row(40, "Upside vs today", {c: f"={c}39/$C$10-1" for c in S}, UPS, bold=True)

# 4. Sensitivity of the base upside
bk.section(ws, 42, "4. Base upside: share of the gap closed (rows) x 5-year real rate (columns)", "B", "F")
rates = [0.085, 0.0676, 0.055]
ws.write("B43", "Share closed \\ real rate", bk.HDRL)
cols = "CDE"
for c, rr in zip(cols, rates):
    ws.write_number(f"{c}43", rr, F(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", num_format=PCT2))
for i, sh in enumerate((0, 0.25, 0.5, 0.75, 1.0)):
    r = 44 + i
    ws.write_number(f"B{r}", sh, F(bold=True, num_format="0%", align="left"))
    for c in cols:
        spread = f"($C$24+$B{r}*($C$25-$C$24))/100"
        base_f = f"($C$23-gLT)/($D$30+ERP+{spread}-gLT)"
        rate_f = f"(1/$C$17)/(1/$C$17+({c}$43-$D$30)*100)"
        ws.write_formula(f"{c}{r}", f"={base_f}*{rate_f}*(1+gLT)^$C$9-1", F(num_format=UPS))
ws.conditional_format("C44:E48", {"type": "3_color_scale", "min_color": "#F8696B", "mid_type": "num", "mid_value": 0,
                                  "mid_color": "#FFEB84", "max_color": "#63BE7B"})
ws.write("B50", "Columns move the real rate in the base (no earnings shock); the base valuation factor uses the base real "
         "rate (D30) and the rate effect uses the analysts' sensitivity, as in section 3.", NOTE)
ws.activate()
bk.close()
print("saved", OUT)
