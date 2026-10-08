"""Builds Ibov_house_2027_by_box_XP.xlsx: the house Ibovespa target model (DCF Ibov 2026_Out.xlsx, Target sheet) rolled forward
to the end of 2027 and split into the four boxes (Financials, Defensives, Cyclicals, Commodities): the DCF by box, the house
multiples (P/E, EV/EBITDA) and the bottom-up by box, and the average of the methods (the house's target).
DCF by box: CAPM rates (box betas from the boxes' return indices), box cash-flow rules (capex funding growth, NWC by change)
and the consensus growth of each box fading to g after 2031 (year by year, sheet Box Fade; or the H-model formula); the base is
scaled to the house DCF value. Scenarios: the house real-rate shock passes through each box's beta, and earnings growth moves
with each box's earnings beta to the domestic cycle (Commodities: a commodity shock), so the Ibovespa's bear / bull open up
(sheet Box Earnings; Assumptions section 8).

The aggregate is the replica of build_house2027.py (same inputs, rules and mechanics; checks tie it to the house file).
Each house line (sales, EBIT, EBITDA, net income, capex) is split among the boxes by their share of the IBOV Index line,
measured from the members' Bloomberg consensus (company totals x index points / market cap, FY26-FY28). Shares of the
non-financial boxes = their members / the IBOV Index line; Financials = the rest, which includes the EBITDA and capex the
index imputes to the banks (Bloomberg scales the members with data up to the whole index). Defensives, Cyclicals and
Commodities are valued with the house's FCFF mechanics; Financials with FCFE = net income x (1 - g / ROE), same g.
Box discount rates (switch BoxRate): box rate - g = (house rate - g) x the box's 10-year cash yield relative to the Ibovespa
(average 12m fwd earnings yield x the box's cash conversion), so each box keeps its historical relative valuation, as the
multiples do; BoxRate = 0 gives every box the house rate; BoxRate = 2 is the CAPM, real rate + box beta x ERP, with the betas
computed from the boxes' weekly total-return indices (sheet Box Betas). The four box values are scaled pro rata to the
house's Ibovespa value (switch ScaleBox).

Run: python box_lines_data.py (Bloomberg) -> python build_house2027_box.py -> python excel_recalc.py ..\\Ibov_house_2027_by_box_XP.xlsx
"""
from __future__ import annotations

import json
import warnings
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd

from xp_template import Book, cn, NAVY, YEL, GREY, GREEN, INK, KEY, GBAND, PTS, PTS1, MULT, PCT, PCT2, UPS, PX, DATE, DATED, HEAT

warnings.filterwarnings("ignore")
HERE = Path(__file__).parent
HOUSE = Path(r"C:\Users\Caio\Documents\Documentos\Equity Strategy Dashboard\_pedidos\2026-10-05_ibov_valuation_groups\DCF Ibov 2026_Out.xlsx")
OUT = HERE.parent / "Ibov_house_2027_by_box_XP.xlsx"
SRCM = Path(r"C:\Users\Caio\Documents\Documentos\Equity Strategy Dashboard\_pedidos\2026-10-05_ibov_valuation_groups\scripts")

# --------------------------------------------------------------------------- house file values
T = openpyxl.load_workbook(HOUSE, data_only=True)["Target"]
HY = {y: cn(1 + y - 2015) for y in range(2015, 2031)}          # house columns: 2015 -> B ... 2030 -> Q
hv = lambda row, y: T[f"{HY[y]}{row}"].value
H = {"ibov": T["B61"].value, "rr": T["C56"].value, "erp": T["B57"].value, "g": T["B54"].value, "infl": T["B63"].value,
     "fcff": {y: hv(44, y) for y in range(2025, 2031)}, "tv": T["R44"].value, "fv25": T["D59"].value, "fv26": T["F59"].value,
     "fv27": T["G59"].value, "irr": T["B62"].value, "erp_impl": T["B65"].value,
     "sens": [(T[f"B{r}"].value, T[f"C{r}"].value) for r in range(122, 130)], "nd_mult": T["F70"].value}
assert abs(H["fv26"] - 202259.1686) < 0.01, H["fv26"]

# --------------------------------------------------------------------------- members and their consensus lines
BOXES = ["Financials", "Defensives", "Cyclicals", "Commodities"]
NONFIN = BOXES[1:]
MEMO = ["Cyclicals ex-WEG & Embraer", "WEG & Embraer"]
GROUPS = BOXES + MEMO
LINES = [("sales", "Sales", "BEST_SALES"), ("ebit", "EBIT", "BEST_EBIT"), ("ebitda", "EBITDA", "BEST_EBITDA"),
         ("ni", "Net income", "BEST_NET_INCOME"), ("capex", "Capex", "BEST_CAPEX")]
CY = ["26", "27", "28"]
LASTSH = {"sales": 2027, "capex": 2027, "ebit": 2028, "ebitda": 2028, "ni": 2028}   # last year of consensus in the house lines
mem = pd.read_parquet(SRCM / "bx_members.parquet")
mem["bo"] = mem["box"].map({b: i for i, b in enumerate(BOXES)})
mem = mem.sort_values(["bo", "glob", "w30"], ascending=[True, True, False])
bl = pd.read_parquet(HERE / "box_lines.parquet").reindex(mem.index)
dmm = pd.read_parquet(HERE / "dcf_members.parquet").reindex(mem.index)
BLI = json.loads((HERE / "box_lines.json").read_text(encoding="utf-8"))
IDX = BLI["ibov_index_lines"]
# 10-year history of the 12m fwd earnings yield by box: 1 / harmonic P/E of the members of each month (bx_panel20, the
# same Bloomberg panel as the multiples model), Oct-16..Sep-26
pan = pd.read_parquet(SRCM / "bx_panel20.parquet")
pan = pan[(pan["date"] >= "2016-10-01") & (pan["date"] <= "2026-09-30")].copy()
pan["key2"] = np.where(pan["box"] != "Cyclicals", pan["box"], np.where(pan["glob"] == 1, "WEG & Embraer", "Cyclicals ex-WEG & Embraer"))


def _ey(x):
    ok = x["pe"].notna() & (x["pe"] >= 1) & (x["pe"] <= 100)
    return float((x["w"][ok] / x["pe"][ok]).sum() / x["w"][ok].sum())


EYG = BOXES + ["Ibovespa"] + MEMO
EYH = pd.DataFrame({g: (pan if g == "Ibovespa" else (pan[pan["key2"] == g] if g in MEMO else pan[pan["box"] == g])).groupby("date").apply(_ey)
                    for g in EYG})
assert len(EYH) == 120, len(EYH)
REFV = json.loads((HERE / "ref_v6.json").read_text(encoding="utf-8"))
MULTREF = {(g, sc): REFV["fv"]["Average"][f"{g}|{sc}"] for g in BOXES + ["Ibovespa"] + MEMO for sc in ("Bear", "Base", "Bull")}
# box total-return indices, daily (Economatica adjusted close; within each month a buy-and-hold portfolio of the members at the
# previous month-end index weights), from the Oct-5 performance study; weekly (Friday) for the CAPM betas, Sep-16..Sep-26
BXD = pd.read_parquet(SRCM.parent / "studies" / "performance_vs_rates" / "box_daily.parquet")
BXD.index = pd.to_datetime(BXD.index)
BXW = BXD[BXD.index <= "2026-09-30"].resample("W-FRI").last()
BXW = BXW[(BXW.index >= "2016-09-23") & (BXW.index <= "2026-09-25")]
# 12m fwd EPS index by group, matched sample chain-linked (Oct-6 EPS study), monthly Oct-06..Sep-26: earnings betas to the
# domestic cycle (Ibovespa ex-Commodities) for the DCF scenarios
EPSI = pd.read_parquet(SRCM.parent / "studies" / "eps_vs_rates" / "eps_index_matched.parquet")
EPSI.index = pd.to_datetime(EPSI.index)
# commodity prices: iPath Bloomberg Commodity ETN (DJP), month-end, from the Oct-5 performance study (local file)
DJP = pd.read_csv(SRCM.parent / "studies" / "performance_vs_rates" / "djp_daily.csv", index_col=0, parse_dates=True)["DJP"].resample("ME").last().reindex(EPSI.index)
assert DJP.notna().all()
# analysts' 12m target prices (house COMP SHEET TP, consensus when the house has none), as the multiples model's bottom-up
mem["tp_used"] = mem["tp_house"].fillna(mem["tp_cons"])
mem["tp_src2"] = np.where(mem["tp_house"].notna(), "XP", np.where(mem["tp_cons"].notna(), "consensus", ""))


def _hm(x, col, lo, hi):
    ok = x[col].notna() & (x[col] >= lo) & (x[col] <= hi)
    return float(x["w"][ok].sum() / (x["w"][ok] / x[col][ok]).sum()) if ok.any() else np.nan


# monthly 12m fwd P/E (all members) and EV/EBITDA (non-financial members) by group, harmonic means, Oct-16..Sep-26
MPE = pd.DataFrame({g: (pan if g == "Ibovespa" else (pan[pan["key2"] == g] if g in MEMO else pan[pan["box"] == g])).groupby("date").apply(lambda x: _hm(x, "pe", 1, 100))
                    for g in EYG})
EVG = ["Defensives", "Cyclicals", "Commodities", "Ibovespa ex-Financials", "Cyclicals ex-WEG & Embraer", "WEG & Embraer"]
_nf = pan[pan["box"] != "Financials"]
MEV = pd.DataFrame({g: (_nf if g == "Ibovespa ex-Financials" else (_nf[_nf["key2"] == g] if g in MEMO else _nf[_nf["box"] == g])).groupby("date").apply(lambda x: _hm(x, "evx", 1, 50))
                    for g in EVG})

YEARS = list(range(2015, 2032))
YC = {y: cn(2 + y - 2015) for y in YEARS}                      # ours: 2015 -> C ... 2031 -> S
LASTA, LASTC = 2025, 2028                                       # last actual year; last year with consensus EBIT/EBITDA/EPS
NOTE_C = "U"

ORDER = ["Cover", "Summary", "Assumptions", "DCF by Box", "Multiples by Box", "DCF", "Sensitivity", "Support >", "Box Rates", "Box Betas", "Box Earnings",
         "Box Fade", "Box Shares", "Members", "Checks", "Read Me"]
bk = Book(OUT, ORDER, {"Assumptions": YEL, "Support >": YEL})
F, W = bk.F, bk.W
IN, LV = bk.IN, bk.LV


def status(y):
    return "Actual" if y <= LASTA else ("Consensus" if y <= LASTC else "Projection")


# =========================================================================== Assumptions
ws = W["Assumptions"]
bk.title(ws, "DCF assumptions (house model by box, end-2027)", "Change the pink cells; the DCF sheet recalculates. Projection rules for 2028-2031 are pink cells in the DCF sheet.")
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
ws.write_number("E20", 0.055, F(**LV, num_format=PCT2))
ws.write("G20", "House (Target rows 131-151): bear 8.5%, base = the market rate, bull 6.0%. Bull here at 5.5% (a row of the house's real-rate table).", bk.NOTE)
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
bk.section(ws, 37, "6. By box", "B", "G")
bk.header_row(ws, 38, [("B", "Setting"), ("C", "Value"), ("D", ""), ("E", ""), ("F", ""), ("G", "What it does")])
ws.write("B39", "Scale the boxes to the house's Ibovespa value: 1 = yes, 0 = no", bk.TXT)
ws.write_number("C39", 0, F(**LV, num_format="0"))
ws.data_validation("C39", {"validate": "list", "source": [0, 1]})
ws.write("G39", "1: each box x (house Ibovespa value / sum of the boxes), so the boxes add up to the house number. 0: each box with its own value.", bk.NOTE)
bk.name("ScaleBox", "Assumptions", "$C$39")
ws.write("B40", "Financials ROE for the FCFE (blank = consensus FY28, Box Shares)", bk.TXT)
ws.write_blank("C40", None, F(**LV, num_format=PCT))
ws.write("G40", "FCFE = net income x (1 - g / ROE): the earnings left after keeping the bank's equity growing at g.", bk.NOTE)
bk.name("FinROEovr", "Assumptions", "$C$40")
ws.write("B41", "Index points of the members: Ibovespa on Oct-5-2026 (Assumptions D7) x price-drifted Sep-30 weights.", bk.NOTE)
ws.write("B42", "Box discount rates: 0 = house rate, 1 = relative to history, 2 = CAPM (box beta)", bk.TXT)
ws.write_number("C42", 2, F(**LV, num_format="0"))
ws.data_validation("C42", {"validate": "list", "source": [0, 1, 2]})
ws.write("G42", "0: every box at the house rate. 1: box rate - g = (house rate - g) x the box's 10y cash yield relative to the Ibovespa (Box Rates), "
                "as the multiples. 2: CAPM, real rate + box beta x ERP (Box Betas).", bk.NOTE)
bk.name("BoxRate", "Assumptions", "$C$42")
ws.write("B43", "Lowest risk premium of a box (box rate >= real rate + this)", bk.TXT)
ws.write_number("C43", 0.0, F(**LV, num_format=PCT))
ws.write("G43", "Keeps a box's risk premium from going below this; at 0% it binds only for Cyclicals in the bear.", bk.NOTE)
bk.name("ERPfloor", "Assumptions", "$C$43")
ws.write("B44", "CAPM beta window (years of weekly returns: 2, 5 or 10)", bk.TXT)
ws.write_number("C44", 5, F(**LV, num_format="0"))
ws.data_validation("C44", {"validate": "list", "source": [2, 5, 10]})
ws.write("G44", "Weekly total returns of the box indices up to Sep-25-2026 (Box Betas).", bk.NOTE)
bk.name("BetaYears", "Assumptions", "$C$44")
ws.write("B45", "CAPM beta: 1 = raw, 2 = Blume adjusted (0.67 x raw + 0.33)", bk.TXT)
ws.write_number("C45", 1, F(**LV, num_format="0"))
ws.data_validation("C45", {"validate": "list", "source": [1, 2]})
ws.write("G45", "The adjustment pulls the betas toward 1 (Bloomberg's default adjusted beta).", bk.NOTE)
bk.name("BetaAdj", "Assumptions", "$C$45")
for r_, lab_, nm_, v_, fmt_, note_, src_ in [
        (46, "Box cash-flow rules: 1 = capex funding growth and NWC by change, 0 = the house rules", "BoxCF", 1, "0",
         "1: capex = D&A + growth x NOPAT / ROE (Box Shares section 5) and working capital by its change, in the box values only; the house aggregate is untouched.", [0, 1]),
        (47, "Growth in the terminal value: 1 = consensus growth fading to g, 0 = g (house)", "BoxGrowth", 0, "0",
         "1: the box's consensus net income growth 2026-28 (capped) fades to g over FadeYears after 2031; 0: the fade starts at g. Valued year by year or by the "
         "H-model formula (FadeMode, section 8).", [0, 1]),
        (48, "Years over which the growth fades to g (after 2031)", "FadeYears", 10, "0", "At most 10 with the year-by-year fade (Box Fade).", None),
        (49, "Cap on the starting growth (+/-)", "GrowthCap", 0.15, PCT, "Defensives (+23%) and Cyclicals (+20%) are capped.", None)]:
    ws.write(f"B{r_}", lab_, bk.TXT)
    ws.write_number(f"C{r_}", v_, F(**LV, num_format=fmt_))
    if src_:
        ws.data_validation(f"C{r_}", {"validate": "list", "source": src_})
    ws.write(f"G{r_}", note_, bk.NOTE)
    bk.name(nm_, "Assumptions", f"$C${r_}")
bk.section(ws, 51, "7. Target by box: average of the methods (house Target rows 83-90)", "B", "G")
bk.header_row(ws, 52, [("B", "Lever"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), ("F", ""), ("G", "Source")])
for r_, lab_, vals_, fmt_, note_ in [
        (53, "Target P/E on 2027 EPS", [7.5, 10.0, 12.0], MULT, "House: 10.0x for 2027 (Target!F76); bear / bull 7.5x / 12x (Target rows 134 / 148)."),
        (54, "Target EV/EBITDA on 2027 EBITDA", [4.5, 5.5, 6.0], MULT, "House: 5.5x for 2027 (Target!F67); bear / bull 4.5x / 6.0x (rows 135 / 149). Net debt = 1.6x EBITDA 2025."),
        (55, "Bottom-up: change vs. the analysts' target prices", [-0.2, 0.0, 0.2], UPS, "House: -20% / +20% in the bear / bull (rows 136 / 150); 2027 without the / 1.15 of 2026.")]:
    ws.write(f"B{r_}", lab_, bk.TXT)
    for c_, v_ in zip("CDE", vals_):
        ws.write_number(f"{c_}{r_}", v_, F(**LV, num_format=fmt_))
    ws.write(f"G{r_}", note_, bk.NOTE)
ws.write("B56", "Box multiples: 1 = relative to history, 0 = house multiple", bk.TXT)
ws.write_number("C56", 1, F(**LV, num_format="0"))
ws.data_validation("C56", {"validate": "list", "source": [0, 1]})
ws.write("G56", "1: the boxes keep their historical premium or discount to the Ibovespa, calibrated so they add up to the house value (Multiples by Box).", bk.NOTE)
bk.name("RelMult", "Assumptions", "$C$56")
ws.write("B57", "Weights of the methods: DCF, P/E, EV/EBITDA, bottom-up", bk.TXT)
for c_, nm_ in zip("CDEF", ("WDCF", "WPE", "WEV", "WBU")):
    ws.write_number(f"{c_}57", 1, F(**LV, num_format="0.0"))
    bk.name(nm_, "Assumptions", f"${c_}$57")
ws.write("G57", "House: simple average of the four methods (Target row 88).", bk.NOTE)
bk.section(ws, 59, "8. Scenarios of the DCF by box: the house's two levers (real rate, section 3; earnings level, below)", "B", "G")
bk.header_row(ws, 60, [("B", "Lever"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), ("F", ""), ("G", "What it does")])
for r_, lab_, vals_, fmt_, note_ in [
        (61, "Earnings level of the Ibovespa ex-Commodities vs. consensus", [-0.20, 0.0, 0.20], UPS,
         "As the house's EPS / EBITDA lower / higher, symmetric: every year's earnings and flows x (1 + shock). Split by each domestic box's earnings beta "
         "(Box Earnings), normalized so the three domestic boxes average this shock."),
        (62, "Commodity prices vs. base", [-0.20, 0.0, 0.20], UPS,
         "Commodities' own method: its earnings level moves by its earnings elasticity to commodity prices (Box Earnings, estimated 2007-26) x this change."),
        (63, "Probability of the scenario", [0.25, 0.5, 0.25], PCT, "Expected value by box (Summary section 10); normalized by their sum.")]:
    ws.write(f"B{r_}", lab_, bk.TXT)
    for c_, v_ in zip("CDE", vals_):
        ws.write_number(f"{c_}{r_}", v_, F(**LV, num_format=fmt_))
    ws.write(f"G{r_}", note_, bk.NOTE)
for r_, lab_, nm_, v_, src_, note_ in [
        (64, "Rate shock: 1 = x box beta, 0 = the same for every box (house)", "RateShock", 0, [0, 1],
         "1: box rate = base box rate + beta x (scenario house rate - base house rate), a shock to the market's required return; 0: the real-rate change moves "
         "every box equally. CAPM rates only (BoxRate = 2)."),
        (65, "Normalize the domestic earnings betas: 1 = yes, 0 = raw", "EBNorm", 1, [0, 1],
         "1: the betas are scaled so the index-point-weighted average of Financials, Defensives and Cyclicals is 1 (the domestic boxes get the shock on average)."),
        (66, "Ibovespa bear / bull: 1 = sum of the boxes, 0 = house DCF", "ScaleScen", 1, [0, 1],
         "1: every scenario keeps the base-case scale factor (DCF by Box), so the shocks open the Ibovespa's bear / bull; 0: each scenario is scaled to the house DCF."),
        (67, "Fade after 2031: 0 = H-model, 1 = year by year, 2 = growth paid", "FadeMode", 0, [0, 1, 2],
         "1: the flow grows year by year from the starting growth to g (Box Fade); the H-model formula approximates it and undervalues negative growth "
         "(Commodities). 2: flow = earnings x (1 - growth / ROE): growth costs reinvestment at the box's FY28 ROE.")]:
    ws.write(f"B{r_}", lab_, bk.TXT)
    ws.write_number(f"C{r_}", v_, F(**LV, num_format="0"))
    ws.data_validation(f"C{r_}", {"validate": "list", "source": src_})
    ws.write(f"G{r_}", note_, bk.NOTE)
    bk.name(nm_, "Assumptions", f"$C${r_}")
ws.write("B68", "Earnings betas: first month of the window", bk.TXT)
ws.write_datetime("C68", pd.Timestamp("2007-10-31").to_pydatetime(), F(**LV, num_format=DATED))
ws.write("G68", "Window from this month to Sep-26; Oct-07 is the first 12-month change (the whole history).", bk.NOTE)
bk.name("EBStart", "Assumptions", "$C$68")
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

# =========================================================================== Members
ws = W["Members"]
bk.title(ws, "Members: consensus lines by stock", "Bloomberg BEST consensus, company totals in BRL mn (Oct-7-2026), and market cap; in index points = points x line / market cap.")
MF = 6
ML = MF + len(mem) - 1
MT = ML + 2
LC = {}
ci = 11
for k, lab, f in LINES:
    for y in CY:
        LC[(k, y)] = [cn(ci), None]
        ci += 1
for k, lab, f in LINES:
    for y in CY:
        LC[(k, y)][1] = cn(ci)
        ci += 1
CROE, CNIV, CBV, CFLAG = cn(ci), cn(ci + 1), cn(ci + 2), cn(ci + 3)
CTP, CTPS, CBP, CBU = cn(ci + 4), cn(ci + 5), cn(ci + 6), cn(ci + 7)
mcols = [("A", "Ticker", 8), ("B", "Company", 18), ("C", "Box", 12), ("D", "Sub-box", 24), ("E", "Ibovespa weight Sep-30", 9),
         ("F", "Price Sep-30 (R$)", 9), ("G", "Price Oct-5 (R$)", 9), ("H", "Price-drifted weight (unscaled)", 10), ("I", "Weight Oct-5", 9),
         ("J", "Index points", 10), ("K", "Market cap (BRL mn)", 12)]
for (k, y), (cb, cp) in LC.items():
    lab = dict((a, b) for a, b, _ in LINES)[k]
    mcols += [(cb, f"{lab} FY{y} (BRL mn)", 10), (cp, f"{lab} FY{y} (index pts)", 10)]
mcols += [(CROE, "ROE FY28", 8), (CNIV, "Net income FY28 with ROE (pts)", 10), (CBV, "Book value FY28 (pts)", 10), (CFLAG, "Missing lines", 30),
          (CTP, "Target price 12m (R$)", 10), (CTPS, "Source", 10), (CBP, "Index points with a target price", 10), (CBU, "Index points at the target price", 10)]
for a_, b_, t_ in [("A", "D", "Identification"), ("E", "K", "Weight, index points, market cap"), (LC[("sales", "26")][0], LC[("capex", "28")][0], "Consensus, BRL mn (Bloomberg BEST, EQY_FUND_CRNCY = BRL)"),
                   (LC[("sales", "26")][1], LC[("capex", "28")][1], "In index points = points x line / market cap"), (CROE, CBV, "ROE FY28"),
                   (CTP, CBU, "Bottom-up (analysts' 12m target prices)")]:
    ws.merge_range(f"{a_}4:{b_}4", t_, bk.GROUPH)
ws.write(f"{CFLAG}4", "", bk.GROUPH)
ws.set_row(4, 54)
for c, h, wdt in mcols:
    ws.write(f"{c}5", h, bk.HDR)
    ws.set_column(f"{c}:{c}", wdt)
for i, (t, x) in enumerate(mem.iterrows()):
    r = MF + i
    y_ = bl.loc[t]
    ws.write_string(f"A{r}", t, F(**IN, bold=True))
    ws.write_string(f"B{r}", x["name"], F(**IN))
    ws.write_string(f"C{r}", x["box"], F(**IN))
    sub = ("WEG & Embraer" if x["glob"] == 1 else "Cyclicals ex-WEG & Embraer") if x["box"] == "Cyclicals" else x["box"]
    ws.write_string(f"D{r}", sub, F(**IN))
    ws.write_number(f"E{r}", x["w30"], F(**IN, num_format="0.00%"))
    ws.write_number(f"F{r}", x["px30"], F(**IN, num_format=PX))
    ws.write_number(f"G{r}", x["px_now"], F(**IN, num_format=PX))
    ws.write_formula(f"H{r}", f"=E{r}*G{r}/F{r}", F(num_format="0.000%"))
    ws.write_formula(f"I{r}", f"=H{r}/SUM($H${MF}:$H${ML})", F(num_format="0.00%"))
    ws.write_formula(f"J{r}", f"=I{r}*Assumptions!$D$7", F(num_format=PTS))
    ws.write_number(f"K{r}", y_["mcap"] / 1e6, F(**IN, num_format=PTS))
    miss = []
    for (k, yy), (cb, cp) in LC.items():
        v = y_[f"{k}{yy}"]
        if pd.notna(v):
            ws.write_number(f"{cb}{r}", v, F(**IN, num_format=PTS))
        else:
            miss.append(f"{k}{yy}")
        ws.write_formula(f"{cp}{r}", f"=IF(ISNUMBER({cb}{r}),$J{r}*{cb}{r}/$K{r},0)", F(num_format=PTS1))
    roe = dmm.loc[t, "roe28"]
    if pd.notna(roe):
        ws.write_number(f"{CROE}{r}", roe / 100, F(**IN, num_format=PCT))
    ni28 = LC[("ni", "28")][1]
    ws.write_formula(f"{CNIV}{r}", f"=IF(AND(ISNUMBER({CROE}{r}),N({CROE}{r})>0),{ni28}{r},0)", F(num_format=PTS1))
    ws.write_formula(f"{CBV}{r}", f"=IF({CNIV}{r}<>0,{CNIV}{r}/{CROE}{r},0)", F(num_format=PTS1))
    if miss:
        ws.write_string(f"{CFLAG}{r}", ", ".join(miss), F(font_color=GREY))
    if pd.notna(x["tp_used"]):
        ws.write_number(f"{CTP}{r}", float(x["tp_used"]), F(**IN, num_format=PX))
        ws.write_string(f"{CTPS}{r}", x["tp_src2"], F(**IN))
    ws.write_formula(f"{CBP}{r}", f"=IF(ISNUMBER({CTP}{r}),J{r},0)", F(num_format=PTS1))
    ws.write_formula(f"{CBU}{r}", f"=IF(ISNUMBER({CTP}{r}),J{r}*{CTP}{r}/G{r},0)", F(num_format=PTS1))
ws.write(f"A{MT}", "Total", bk.BOLD)
for c in ["E", "I", "J"] + [cp for _, cp in LC.values()] + [CNIV, CBV, CBP, CBU]:
    fmt = "0.00%" if c in ("E", "I") else PTS
    ws.write_formula(f"{c}{MT}", f"=SUM({c}{MF}:{c}{ML})", F(bold=True, num_format=fmt, top=1))
for i, t in enumerate(["Index points of a line = index points x line / market cap: the stock's share of the line, in points. Company totals and market cap in BRL, "
                       "so the currency of reporting does not matter; both share classes of a company use the company's market cap.",
                       "Banks have no EBITDA and no capex in Bloomberg; the IBOV Index line scales the members with data up to the whole index, "
                       "so part of the index EBITDA and capex is imputed to the banks (Box Shares section 3)."]):
    ws.write(f"A{MT + 2 + i}", t, bk.NOTE)
ws.freeze_panes(5, 2)
ws.autofilter(f"A5:{CBU}{ML}")
MR = lambda c: f"Members!${c}${MF}:${c}${ML}"

# =========================================================================== Box Shares
ws = W["Box Shares"]
bk.title(ws, "Box shares of the house lines", "Members' consensus by box (index pts) as a share of the IBOV Index line (Bloomberg, Oct-7-2026); Financials = the rest. Net income: share of the members' total.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 58)
ws.set_column("C:R", 10.5)
SB = {}
for i, (k, lab, f) in enumerate(LINES):
    for j, y in enumerate(CY):
        SB[(k, y)] = cn(3 + 3 * i + j)
LABL = dict((a, b) for a, b, _ in LINES)
SR1 = {"Financials": 7, "Defensives": 8, "Cyclicals": 9, "Commodities": 10, "Members": 11, "IBOV Index": 12,
       "Cyclicals ex-WEG & Embraer": 14, "WEG & Embraer": 15}
SR2 = {"Financials": 19, "Defensives": 20, "Cyclicals": 21, "Commodities": 22, "Total": 23, "Cyclicals ex-WEG & Embraer": 25, "WEG & Embraer": 26}


def share_header(row):
    ws.write(f"B{row}", "Group", bk.HDRL)
    ws.write(f"C{row}", "Index points", bk.HDR)
    for (k, y), c in SB.items():
        ws.write(f"{c}{row}", f"{LABL[k]} {2000 + int(y)}", bk.HDR)


bk.section(ws, 5, "1. Members' consensus by box (index pts, Members sheet)", "B", "R")
share_header(6)
ws.set_row(5, 30)
for g, r in SR1.items():
    tot = g in ("Members", "IBOV Index")
    ws.write(f"B{r}", {"Members": "Sum of the members", "IBOV Index": "IBOV Index line (Bloomberg BEST, Oct-7)"}.get(g, ("   " if g in MEMO else "") + g),
             F(bold=tot, top=1 if g == "Members" else 0))
    if g == "Members":
        for c in ["C"] + list(SB.values()):
            ws.write_formula(f"{c}{r}", f"=SUM({c}7:{c}10)", F(bold=True, num_format=PTS, top=1))
    elif g == "IBOV Index":
        ws.write_formula(f"C{r}", "=Assumptions!$D$7", F(num_format=PTS))
        for (k, y), c in SB.items():
            f = dict((a, x) for a, _, x in LINES)[k]
            v = IDX[y].get(f) if k != "ni" else None
            if v is None:
                ws.write_string(f"{c}{r}", "n/a", bk.NA)
            else:
                ws.write_number(f"{c}{r}", v, F(**IN, num_format=PTS))
    else:
        crit = "C" if g in BOXES else "D"
        ws.write_formula(f"C{r}", f"=SUMIFS({MR('J')},{MR(crit)},\"{g}\")", F(num_format=PTS))
        for (k, y), c in SB.items():
            ws.write_formula(f"{c}{r}", f"=SUMIFS({MR(LC[(k, y)][1])},{MR(crit)},\"{g}\")", F(num_format=PTS))
ws.write("B13", "Memo (inside Cyclicals)", bk.NOTE)

bk.section(ws, 17, "2. Shares used to split the house lines (DCF by Box)", "B", "R")
share_header(18)
ws.set_row(17, 30)
for g, r in SR2.items():
    tot = g == "Total"
    ws.write(f"B{r}", ("   " if g in MEMO else "") + g, F(bold=tot, top=1 if tot else 0))
    if tot:
        for c in SB.values():
            ws.write_formula(f"{c}{r}", f"=SUM({c}19:{c}22)", F(bold=True, num_format=PCT, top=1))
        continue
    ws.write_formula(f"C{r}", f"=C{SR1[g]}/$C$11", F(num_format=PCT))
    for (k, y), c in SB.items():
        if k == "ni":
            f_ = f"={c}{SR1[g]}/{c}$11"
        elif g == "Financials":
            f_ = f"=1-SUM({c}20:{c}22)"
        else:
            f_ = f"={c}{SR1[g]}/{c}$12"
        ws.write_formula(f"{c}{r}", f_, F(num_format=PCT, bold=g == "Financials"))
ws.write("B24", "Memo (inside Cyclicals)", bk.NOTE)
ws.write("B27", "Non-financial boxes: members' line / IBOV Index line. Financials: 1 - the other three (net income: members' share of the members' total).", bk.NOTE)

bk.section(ws, 29, "3. Financials: reported by the members vs. taken by Financials from the index line (index pts)", "B", "R")
share_header(30)
ws.set_row(29, 30)
for i, (lab, kind) in enumerate([("Financials, members' reported consensus", "rep"), ("Financials share x IBOV Index line", "idx"),
                                  ("Imputed by the index (banks' EBITDA and capex) + index gap", "imp")]):
    r = 31 + i
    ws.write(f"B{r}", lab, F(bold=kind == "imp"))
    for (k, y), c in SB.items():
        if k == "ni":
            ws.write_string(f"{c}{r}", "n/a", bk.NA)
            continue
        f_ = {"rep": f"={c}7", "idx": f"={c}19*{c}12", "imp": f"={c}32-{c}31"}[kind]
        ws.write_formula(f"{c}{r}", f_, F(num_format=PTS, bold=kind == "imp"))

bk.section(ws, 35, "4. Financials ROE for the FCFE (FY28 consensus)", "B", "R")
for i, (lab, f_, fmt) in enumerate([("Net income FY28 of the Financials with an ROE (index pts)", f"=SUMIFS({MR(CNIV)},{MR('C')},\"Financials\")", PTS),
                                    ("Book value FY28 = net income / ROE (index pts)", f"=SUMIFS({MR(CBV)},{MR('C')},\"Financials\")", PTS),
                                    ("ROE FY28, consensus = net income / book value", "=C36/C37", PCT),
                                    ("ROE used (Assumptions override if typed)", "=IF(ISNUMBER(FinROEovr),FinROEovr,C38)", PCT)]):
    r = 36 + i
    ws.write(f"B{r}", lab, F(bold=i == 3))
    ws.write_formula(f"C{r}", f_, F(num_format=fmt, bold=i == 3))
bk.name("FinROE", "Box Shares", "$C$39")
bk.section(ws, 42, "5. Box ROE FY28 and consensus growth (DCF by Box: capex funding growth and the growth fade)", "B", "R")
bk.header_row(ws, 43, [("B", "Group"), ("C", "Net income FY28 with ROE (pts)"), ("D", "Book value FY28 (pts)"), ("E", "ROE FY28"),
                       ("F", "Net income 2026 (pts)"), ("G", "Net income 2028 (pts)"), ("H", "Growth 2026-28 (a year)"), ("I", "Starting growth (capped)")])
ws.set_row(42, 44)
SR5 = {"Financials": 44, "Defensives": 45, "Cyclicals": 46, "Commodities": 47, "Cyclicals ex-WEG & Embraer": 49, "WEG & Embraer": 50}
for g, r in SR5.items():
    crit = "C" if g in BOXES else "D"
    ws.write(f"B{r}", ("   " if g in MEMO else "") + g, bk.TXT)
    ws.write_formula(f"C{r}", f"=SUMIFS({MR(CNIV)},{MR(crit)},\"{g}\")", F(num_format=PTS))
    ws.write_formula(f"D{r}", f"=SUMIFS({MR(CBV)},{MR(crit)},\"{g}\")", F(num_format=PTS))
    ws.write_formula(f"E{r}", f"=C{r}/D{r}", F(num_format=PCT))
    ws.write_formula(f"F{r}", f"={SB[('ni', '26')]}{SR1[g]}", F(num_format=PTS))
    ws.write_formula(f"G{r}", f"={SB[('ni', '28')]}{SR1[g]}", F(num_format=PTS))
    ws.write_formula(f"H{r}", f"=(G{r}/F{r})^(1/2)-1", F(num_format=PCT))
    ws.write_formula(f"I{r}", f"=MAX(MIN(H{r},GrowthCap),-GrowthCap)", F(num_format=PCT, bold=True))
ws.write("B48", "Memo (inside Cyclicals)", bk.NOTE)
ws.write("B51", "ROE: members' consensus FY28 (BEST_ROE), aggregated as net income / book value. Growth: the box's consensus net income 2026 to 2028 (section 1), capped (Assumptions).", bk.NOTE)
ws.freeze_panes(6, 2)

# Box Earnings rows (section 1: earnings betas; section 2: shocks) and Box Fade blocks, referenced from DCF by Box
DOMC = "Ibov ex-Comm"
EGR = ["Financials", "Defensives", "Cyclicals", "Commodities", "Cyclicals ex-WEG & Embraer", "WEG & Embraer", "Ibovespa", DOMC]
assert list(EPSI.columns) == EGR and len(EPSI) == 240, EPSI.columns
assert (EPSI > 0).all().all()
EBR = {"Financials": 7, "Defensives": 8, "Cyclicals": 9, "Commodities": 10, DOMC: 11, "Cyclicals ex-WEG & Embraer": 13, "WEG & Embraer": 14}
EKR = {"Financials": 21, "Defensives": 22, "Cyclicals": 23, "Commodities": 24, "Cyclicals ex-WEG & Embraer": 27, "WEG & Embraer": 28}
E0 = 32                                                              # Box Earnings: first data row (Oct-06)
EL = E0 + len(EPSI) - 1                                              # last data row (Sep-26)
ELV = {g_: cn(4 + i_) for i_, g_ in enumerate(EGR)}                  # EPS index levels E..L
EG12 = {g_: cn(4 + len(EGR) + i_) for i_, g_ in enumerate(EGR)}     # 12-month log growth M..T
EG48 = {g_: cn(4 + 2 * len(EGR) + i_) for i_, g_ in enumerate(EGR)}  # 4-year log growth a year U..AB
FADE0, FADEH = 18, 10                                                # Box Fade: first block row, rows per block
FADE = {}
for i_, (g_, c_) in enumerate([(g__, c__) for g__ in GROUPS for c__ in "CDE"]):
    FADE[(g_, c_)] = FADE0 + FADEH * i_
FADE_LAST = max(FADE.values()) + FADEH
FADEV = lambda g_, c_: f"$C${FADE[(g_, c_)] + 8}"                    # value at end-2031 of the flows after 2031
FSR = {"Financials": 7, "Defensives": 8, "Cyclicals": 9, "Commodities": 10, "Cyclicals ex-WEG & Embraer": 12, "WEG & Embraer": 13}

# =========================================================================== DCF by Box
ws = W["DCF by Box"]
bk.title(ws, "DCF by box (end-2027)", "Each box gets its share of every house line (Box Shares); box rules for capex and NWC, consensus growth fading in the TV, CAPM rates (Box Rates). FCFE for Financials.")
BY = list(range(2026, 2032))
YB = {y: cn(2 + y - 2026) for y in BY}
NB = "J"
ws.set_column("A:A", 2)
ws.set_column("B:B", 60)
ws.set_column("C:H", 11)
ws.set_column("I:I", 2)
ws.set_column(f"{NB}:{NB}", 100)
SCN = [("C", "Bear"), ("D", "Base"), ("E", "Bull")]
bk.section(ws, 5, "Ibovespa = sum of the four boxes (end-2027)", "B", NB)
bk.header_row(ws, 6, [("B", ""), ("C", "Bear"), ("D", "Base"), ("E", "Bull")])
KRS = {"Financials": 7, "Defensives": 8, "Cyclicals": 9, "Commodities": 10, "Ibovespa": 11, "Cyclicals ex-WEG & Embraer": 13, "WEG & Embraer": 14}
KR = {"Financials": 21, "Defensives": 22, "Cyclicals": 23, "Commodities": 24, "Cyclicals ex-WEG & Embraer": 26, "WEG & Embraer": 27}
TOPL = {7: "Sum of the box values, own mechanics and box rates (Financials by FCFE)", 8: "House DCF fair value for the Ibovespa (DCF sheet)",
        9: "Scale factor = house value / sum of the boxes", 10: "Factor applied (ScaleBox; ScaleScen = 1: the base factor in every scenario)",
        11: "Sum of the box fair values = the Ibovespa by box", 15: "Ibovespa by box vs. the house DCF (sum / house - 1)",
        12: "Memo: sum of the FCFF pieces at the house rate (house mechanics) = house value",
        13: "Memo: sum of the box values at the house rate (Financials by FCFE)", 14: "Memo: scale factor at the house rate"}
for r, lab in TOPL.items():
    ws.write(f"B{r}", lab, F(bold=r in (11,)))
TOPN = {7: "Defensives + Cyclicals + Commodities (FCFF) + Financials (FCFE), each at its box rate (Box Rates).",
        9: "Close to 1: the boxes' own values add up to about the house number.",
        10: "ScaleScen = 1: the bear / bull keep the base factor, so the earnings shocks (Box Earnings) open the Ibovespa's bear / bull; 0: every scenario = house DCF.",
        12: "The split itself is exact: with every box on the house mechanics and the house rate, the boxes add up to the house value.",
        13: "The boxes as in the house (one rate), for comparison."}
for r, n in TOPN.items():
    ws.write(f"{NB}{r}", n, bk.NOTE)
BKEYS2 = ["sh_sales", "sales", "sh_ebit", "ebit", "sh_ebitda", "ebitda", "sh_capex", "capex", "sh_ni", "ni", "da", "tax", "nopat", "nwc", "dnwc",
          "fcff_house", "fcff_fix", "fcff_used", "capex_g", "fcff_box", "fcff_v"]
BLAB = {"sh_sales": "Share of sales", "sales": "Sales", "sh_ebit": "Share of EBIT", "ebit": "EBIT", "sh_ebitda": "Share of EBITDA", "ebitda": "EBITDA",
        "sh_capex": "Share of capex", "capex": "Capex", "sh_ni": "Share of net income", "ni": "Net income", "da": "D&A = EBITDA - EBIT", "tax": "Tax rate (house)",
        "nopat": "NOPAT = EBIT x (1 + tax rate)", "nwc": "Working capital (NWC) = -NWC % x sales", "dnwc": "Change in NWC",
        "fcff_house": "FCFF, house = D&A + NOPAT + capex + NWC (level)", "fcff_fix": "FCFF with the change in NWC", "fcff_used": "FCFF, house mechanics (switch FixNWC)",
        "capex_g": "Capex funding growth = -(D&A + growth x NOPAT / ROE)", "fcff_box": "FCFF, box rules = D&A + NOPAT + capex funding growth + change in NWC",
        "fcff_v": "FCFF used for the box value (switch BoxCF)",
        "roe": "ROE (Financials, Box Shares section 4)", "fcfe": "FCFE = net income x (1 - g / ROE)"}
BNOTE = {"sh_sales": "Box Shares: 2026-27 consensus; 2028+ as 2027 (the house projects sales from 2027).", "sales": "House sales (DCF sheet) x share.",
         "sh_ebit": "Box Shares: 2026-28 consensus; 2029+ as 2028.", "sh_ebitda": "Box Shares: 2026-28; 2029+ as 2028. Financials includes the EBITDA the index imputes to banks.",
         "sh_capex": "Box Shares: 2026-27; 2028+ as 2027.", "sh_ni": "Box Shares: 2026-28; 2029+ as 2028.", "tax": "DCF sheet (house, -24%).",
         "fcff_used": "FixNWC = 0: house row; 1: corrected row. Sum of the boxes = house FCFF (Checks); used for the house-rate memo.",
         "capex_g": "2028+: D&A + the reinvestment that funds the year's growth (house sales growth) at the box's FY28 ROE (Box Shares section 5); 2026-27: consensus.",
         "fcff_v": "BoxCF = 1: box rules (row above); 0: the house mechanics.", "fcfe": "Financials only: the value used for the box (same g as the house)."}
VK2 = ["rate", "shock", "gstart", "last", "tv", "per", "pvf", "pvtv", "firm", "nd", "fcff_val"]
VFIN = ["e_last", "e_tv", "e_pvf", "e_pvtv", "fcfe_val"]
VEND = ["val_own", "pts", "up_own", "val", "up", "gimp", "rate_h", "fcff_h", "val_h"]
VLAB = {"rate": "Discount rate of the box (Box Rates)", "shock": "Earnings level shock of the scenario (Box Earnings)",
        "gstart": "Starting growth of the fade = consensus 2026-28 (capped; g if BoxGrowth = 0)",
        "last": "FCFF 2031 in the scenario = FCFF 2031 x (1 + shock)",
        "tv": "Value at end-2031 of the FCFF after 2031 (FadeMode 0: H-model formula; 1-2: Box Fade)", "per": "Periods the TV is discounted",
        "pvf": "Present value of FCFF 2028-2031 (each year x (1 + shock))", "pvtv": "Present value of the terminal value", "firm": "Value of the FCFF",
        "nd": "Net debt subtracted (switch FixND)", "fcff_val": "Value by the house mechanics (FCFF)",
        "e_last": "FCFE 2031 in the scenario = FCFE 2031 x (1 + shock)",
        "e_tv": "Value at end-2031 of the FCFE after 2031 (FadeMode 0: H-model formula; 1-2: Box Fade)",
        "e_pvf": "Present value of FCFE 2028-2031 (each year x (1 + shock))", "e_pvtv": "Present value of the FCFE terminal value", "fcfe_val": "Value by FCFE",
        "val_own": "Box value, own mechanics", "pts": "Index points today", "up_own": "Upside, own value", "val": "Fair value at the end of 2027 (scaled)", "up": "Upside",
        "gimp": "Starting growth priced in by today's index points (reverse DCF)",
        "rate_h": "Memo: house discount rate (DCF sheet)", "fcff_h": "Memo: value of the FCFF at the house rate (house mechanics)",
        "val_h": "Memo: box value at the house rate (Financials by FCFE)"}
VNOTE = {"rate": "BoxRate 0: the house rate; 1: g + (house rate - g) x the box's relative cash yield, at least real rate + floor; 2: CAPM, real rate + beta x ERP.",
         "nd": "1.6x the house EBITDA of 2025 x the box's 2026 EBITDA share, only with FixND = 1.",
         "gstart": "Growth fades from this to g over FadeYears after 2031 (H = FadeYears / 2); the same in every scenario.",
         "shock": "Domestic boxes: normalized earnings beta x the Ibovespa ex-Commodities shock; Commodities: price elasticity x commodity prices (Assumptions section 8). Zero in the base.",
         "tv": "FadeMode 0: flow 2031 x [(1 + g) + H x (starting growth - g)] / (rate - g). 1: year by year; 2: year by year with growth paid (Box Fade).",
         "gimp": "Starting growth + (points - own value) / (flow 2031 x H / ((rate - g) x (1 + rate)^periods)): exact with the H-model formula, first order with the year-by-year fade.",
         "fcff_h": "Same FCFF at the house rate: the four boxes add up to the house value exactly (Checks).",
         "val_h": "The box as in the house (one rate); Summary compares it with the box-rate value.",
         "fcff_val": "Financials: shown for the reconciliation with the house value; not used.", "val_own": "FCFF value (Financials: FCFE value).",
         "val": "Own value x factor applied (top of the sheet)."}
BB = {}


def box_block(g, r0):
    fin = g == "Financials"
    keys = BKEYS2 + (["roe", "fcfe"] if fin else [])
    vkeys = VK2 + (VFIN if fin else []) + VEND
    R2 = {k: r0 + 2 + i for i, k in enumerate(keys)}
    vh = r0 + 2 + len(keys) + 1
    R2.update({k: vh + 1 + i for i, k in enumerate(vkeys)})
    bk.section(ws, r0, g + (" (memo, inside Cyclicals; not added to the Ibovespa)" if g in MEMO else (" (FCFE)" if fin else " (house FCFF)")),
               "B", NB, fill=YEL, color=INK)
    bk.header_row(ws, r0 + 1, [("B", "")] + [(YB[y], str(y)) for y in BY] + [(NB, "How it is calculated")])
    for k in keys:
        ws.write(f"B{R2[k]}", BLAB[k], F(bold=k in ("fcff_v", "fcfe"), font_color=GREY if k.startswith("sh_") else INK))
        if k in BNOTE:
            ws.write(f"{NB}{R2[k]}", BNOTE[k], bk.NOTE)
    for k, lab, f in LINES:
        for y in BY:
            c = f"{YB[y]}{R2['sh_' + k]}"
            if y <= LASTSH[k]:
                ws.write_formula(c, f"='Box Shares'!{SB[(k, str(y - 2000))]}{SR2[g]}", F(num_format=PCT, font_color=GREY))
            else:
                ws.write_formula(c, f"={YB[y - 1]}{R2['sh_' + k]}", F(num_format=PCT, font_color=GREY))
            ws.write_formula(f"{YB[y]}{R2[k]}", f"=DCF!{YC[y]}{R[k]}*{c}", F(num_format=PTS))
    for y in BY:
        c = YB[y]
        q = lambda k: f"{c}{R2[k]}"
        ws.write_formula(q("da"), f"={q('ebitda')}-{q('ebit')}", F(num_format=PTS))
        ws.write_formula(q("tax"), f"=DCF!{YC[y]}{R['tax']}", F(num_format=PCT))
        ws.write_formula(q("nopat"), f"={q('ebit')}*(1+{q('tax')})", F(num_format=PTS))
        ws.write_formula(q("nwc"), f"=-NWCpct*{q('sales')}", F(num_format=PTS))
        ws.write_formula(q("fcff_house"), f"={q('da')}+{q('nopat')}+{q('capex')}+{q('nwc')}", F(num_format=PTS))
        if y > 2026:
            ws.write_formula(q("dnwc"), f"={q('nwc')}-{YB[y - 1]}{R2['nwc']}", F(num_format=PTS))
            ws.write_formula(q("fcff_fix"), f"={q('da')}+{q('nopat')}+{q('capex')}+{q('dnwc')}", F(num_format=PTS))
            ws.write_formula(q("fcff_used"), f"=IF(FixNWC=1,{q('fcff_fix')},{q('fcff_house')})", F(num_format=PTS))
        else:
            ws.write_formula(q("fcff_used"), f"={q('fcff_house')}", F(num_format=PTS))
        if y >= 2028:
            ws.write_formula(q("capex_g"), f"=-({q('da')}+DCF!{YC[y]}{R['sg']}*{q('nopat')}/'Box Shares'!$E${SR5[g]})", F(num_format=PTS))
        else:
            ws.write_formula(q("capex_g"), f"={q('capex')}", F(num_format=PTS))
        if y > 2026:
            ws.write_formula(q("fcff_box"), f"={q('da')}+{q('nopat')}+{q('capex_g')}+{q('dnwc')}", F(num_format=PTS))
        else:
            ws.write_formula(q("fcff_box"), f"={q('fcff_house')}", F(num_format=PTS))
        ws.write_formula(q("fcff_v"), f"=IF(BoxCF=1,{q('fcff_box')},{q('fcff_used')})", F(num_format=PTS, bold=True, bg_color=KEY if not fin else "#FFFFFF"))
        if fin:
            ws.write_formula(q("roe"), "=FinROE", F(num_format=PCT))
            ws.write_formula(q("fcfe"), f"={q('ni')}*(1-gLT/{q('roe')})", F(num_format=PTS, bold=True, bg_color=KEY))
    ws.write(f"B{vh}", "Valuation at the end of 2027", bk.HDRL)
    for c, s_ in SCN:
        ws.write(f"{c}{vh}", s_, bk.HDR)
    for k in vkeys:
        ws.write(f"B{R2[k]}", VLAB[k], F(bold=k in ("val", "up", "val_own"), font_color=GREY if ((fin and k in VK2) or k.endswith("_h")) else INK))
        if k in VNOTE:
            ws.write(f"{NB}{R2[k]}", VNOTE[k], bk.NOTE)
    fu, fe, fvr = R2["fcff_used"], R2.get("fcfe"), R2["fcff_v"]
    hfac = lambda c_: f"((1+gLT)+FadeYears/2*({c_}{R2['gstart']}-gLT))"
    for c, s_ in SCN:
        v = lambda k: f"{c}{R2[k]}"
        ws.write_formula(v("rate"), f"='Box Rates'!{c}{KR[g]}", F(num_format=PCT2, bold=True))
        ws.write_formula(v("shock"), f"='Box Earnings'!{c}{EKR[g]}", F(num_format=UPS))
        ws.write_formula(v("gstart"), f"=IF(BoxGrowth=1,'Box Shares'!$I${SR5[g]},gLT)", F(num_format=PCT))
        ws.write_formula(v("last"), f"=$H${fvr}*(1+{v('shock')})", F(num_format=PTS))
        h_tv = f"{v('last')}*{hfac(c)}/({v('rate')}-gLT)"
        ws.write_formula(v("tv"), f"={h_tv}" if fin else f"=IF(FadeMode=0,{h_tv},'Box Fade'!{FADEV(g, c)})", F(num_format=PTS))
        ws.write_formula(v("per"), "=5-FixTV", F(num_format="0"))
        ws.write_formula(v("pvf"), f"=(1+{v('shock')})*NPV({v('rate')},$E${fvr}:$H${fvr})", F(num_format=PTS))
        ws.write_formula(v("pvtv"), f"={v('tv')}/(1+{v('rate')})^{v('per')}", F(num_format=PTS))
        ws.write_formula(v("firm"), f"={v('pvf')}+{v('pvtv')}", F(num_format=PTS))
        ws.write_formula(v("nd"), f"=IF(FixND=1,NDmult*DCF!${YC[2025]}${R['ebitda']}*$C${R2['sh_ebitda']},0)", F(num_format=PTS))
        ws.write_formula(v("fcff_val"), f"={v('firm')}-{v('nd')}", F(num_format=PTS, bold=not fin))
        if fin:
            ws.write_formula(v("e_last"), f"=$H${fe}*(1+{v('shock')})", F(num_format=PTS))
            h_etv = f"{v('e_last')}*{hfac(c)}/({v('rate')}-gLT)"
            ws.write_formula(v("e_tv"), f"=IF(FadeMode=0,{h_etv},'Box Fade'!{FADEV(g, c)})", F(num_format=PTS))
            ws.write_formula(v("e_pvf"), f"=(1+{v('shock')})*NPV({v('rate')},$E${fe}:$H${fe})", F(num_format=PTS))
            ws.write_formula(v("e_pvtv"), f"={v('e_tv')}/(1+{v('rate')})^{v('per')}", F(num_format=PTS))
            ws.write_formula(v("fcfe_val"), f"={v('e_pvf')}+{v('e_pvtv')}", F(num_format=PTS, bold=True))
        ws.write_formula(v("val_own"), f"={v('fcfe_val' if fin else 'fcff_val')}", F(num_format=PTS, bold=True))
        ws.write_formula(v("pts"), f"='Box Shares'!$C${SR1[g]}", F(num_format=PTS))
        ws.write_formula(v("up_own"), f"={v('val_own')}/{v('pts')}-1", F(num_format=UPS))
        ws.write_formula(v("val"), f"={v('val_own')}*{c}$10", F(num_format=PTS, bold=True, bg_color=KEY))
        ws.write_formula(v("up"), f"={v('val')}/{v('pts')}-1", F(num_format=UPS, bold=True, bg_color=KEY))
        ll_ = v("e_last") if fin else v("last")
        ws.write_formula(v("gimp"), f"={v('gstart')}+({v('pts')}-{v('val_own')})/({ll_}*FadeYears/2/(({v('rate')}-gLT)*(1+{v('rate')})^{v('per')}))",
                         F(num_format=PCT, font_color=INK if c == "D" else GREY))
        rh = v("rate_h")
        ws.write_formula(rh, f"=DCF!{c}{V['rate']}", F(num_format=PCT2, font_color=GREY))
        ws.write_formula(v("fcff_h"), f"=NPV({rh},$E${fu}:$H${fu})+$H${fu}*(1+gLT)/({rh}-gLT)/(1+{rh})^{v('per')}-{v('nd')}", F(num_format=PTS, font_color=GREY))
        if fin:
            ws.write_formula(v("val_h"), f"=NPV({rh},$E${fe}:$H${fe})+$H${fe}*(1+gLT)/({rh}-gLT)/(1+{rh})^{v('per')}", F(num_format=PTS, font_color=GREY))
        else:
            ws.write_formula(v("val_h"), f"={v('fcff_h')}", F(num_format=PTS, font_color=GREY))
    BB[g] = R2
    return R2["val_h"] + 3


r0 = 16
for g in GROUPS:
    r0 = box_block(g, r0)
LASTBOX = r0
for c, s_ in SCN:
    ws.write_formula(f"{c}7", "=" + "+".join(f"{c}{BB[g]['val_own']}" for g in BOXES), F(num_format=PTS))
    ws.write_formula(f"{c}8", f"=DCF!{c}{V['fv']}", F(num_format=PTS))
    ws.write_formula(f"{c}9", f"={c}8/{c}7", F(num_format="0.000"))
    ws.write_formula(f"{c}10", f"=IF(ScaleBox=1,IF(ScaleScen=1,$D$9,{c}9),1)", F(num_format="0.000"))
    ws.write_formula(f"{c}11", "=" + "+".join(f"{c}{BB[g]['val']}" for g in BOXES), F(num_format=PTS, bold=True, bg_color=KEY))
    ws.write_formula(f"{c}12", "=" + "+".join(f"{c}{BB[g]['fcff_h']}" for g in BOXES), F(num_format=PTS, font_color=GREY))
    ws.write_formula(f"{c}13", "=" + "+".join(f"{c}{BB[g]['val_h']}" for g in BOXES), F(num_format=PTS, font_color=GREY))
    ws.write_formula(f"{c}14", f"={c}8/{c}13", F(num_format="0.000", font_color=GREY))
    ws.write_formula(f"{c}15", f"={c}11/{c}8-1", F(num_format=UPS))
ws.freeze_panes(4, 2)

# =========================================================================== Box Rates
ws = W["Box Rates"]
bk.title(ws, "Box discount rates: house, relative to history or CAPM",
         "Box rate - g = (house rate - g) x relative cash yield; relative cash yield = (10y average earnings yield x cash conversion) / the same for the Ibovespa.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 46)
ws.set_column("C:I", 14)
bk.section(ws, 5, "1. Historical cash yield relative to the Ibovespa (12m fwd, Oct-16..Sep-26)", "B", "I")
bk.header_row(ws, 6, [("B", "Group"), ("C", "Cash conversion 2031 (FCFF or FCFE / net income)"), ("D", "Average earnings yield (section 3)"),
                      ("E", "Average P/E = 1 / yield"), ("F", "Cash yield = conversion x earnings yield"), ("G", "Relative to the Ibovespa"),
                      ("H", "CAPM beta used (Box Betas)"), ("I", "")])
ws.set_row(5, 58)
EYC = {g: cn(2 + i) for i, g in enumerate(EYG)}
DBQ = "'DCF by Box'"
for g, r in KRS.items():
    tot = g == "Ibovespa"
    t1 = dict(top=1) if tot else {}
    ws.write(f"B{r}", ("   " if g in MEMO else "") + g, F(bold=tot, **t1))
    if tot:
        conv = f"=DCF!${YC[2031]}${R['fcff_used']}/DCF!${YC[2031]}${R['ni']}"
    elif g == "Financials":
        conv = f"={DBQ}!$H${BB[g]['fcfe']}/{DBQ}!$H${BB[g]['ni']}"
    else:
        conv = f"={DBQ}!$H${BB[g]['fcff_used']}/{DBQ}!$H${BB[g]['ni']}"
    ws.write_formula(f"C{r}", conv, F(num_format=PCT, **t1))
    ws.write_formula(f"D{r}", f"=AVERAGE({EYC[g]}$32:{EYC[g]}$151)", F(num_format=PCT2, **t1))
    ws.write_formula(f"E{r}", f"=1/D{r}", F(num_format=MULT, **t1))
    ws.write_formula(f"F{r}", f"=C{r}*D{r}", F(num_format=PCT2, **t1))
    ws.write_formula(f"G{r}", f"=F{r}/$F$11", F(num_format="0.000", bold=True, **t1))
    ws.write_formula(f"H{r}", f"='Box Betas'!I{r}", F(num_format="0.00", **t1))
ws.write("B12", "Memo (inside Cyclicals)", bk.NOTE)
ws.write("B15", "Cash conversion: the box's FCFF (Financials: FCFE) / net income in 2031, the year the terminal value grows from (DCF by Box; Ibovespa: DCF sheet).", bk.NOTE)
bk.section(ws, 17, "2. Discount rate by box (fair value at the end of 2027)", "B", "I")
bk.header_row(ws, 18, [("B", ""), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), ("F", ""), ("G", "Risk premium: Bear"), ("H", "Base"), ("I", "Bull")])
ws.set_row(17, 30)
ws.write("B19", "House discount rate (DCF sheet)", bk.BOLD)
ws.write("B20", "5-year real rate (DCF sheet)", bk.TXT)
for c, ec in zip("CDE", "GHI"):
    ws.write_formula(f"{c}19", f"=DCF!{c}{V['rate']}", F(num_format=PCT2, bold=True))
    ws.write_formula(f"{c}20", f"=DCF!{c}{V['rr']}", F(num_format=PCT2))
    ws.write_formula(f"{ec}19", f"={c}19-{c}20", F(num_format=PCT2, bold=True))
for g, r in KR.items():
    ws.write(f"B{r}", ("   " if g in MEMO else "") + g, bk.TXT)
    for c, ec in zip("CDE", "GHI"):
        x_ = f"{c}$20+$H${KRS[g]}*DCF!{c}${V['erp']}+RateShock*($H${KRS[g]}-1)*({c}$20-$D$20)"
        capm = f"IF(FixInfl=1,(1+{x_})*(1+InflLT)-1,{x_})"
        ws.write_formula(f"{c}{r}", f"=CHOOSE(BoxRate+1,{c}$19,MAX(gLT+({c}$19-gLT)*$G${KRS[g]},{c}$20+ERPfloor),{capm})", F(num_format=PCT2, bold=True))
        ws.write_formula(f"{ec}{r}", f"={c}{r}-{c}$20", F(num_format=PCT2))
ws.write("B25", "Memo (inside Cyclicals)", bk.NOTE)
ws.write("B28", "BoxRate 0: the house rate. 1: g + (house rate - g) x relative cash yield, at least real rate + floor. 2: CAPM, real rate + beta x house ERP (Fisher if FixInfl = 1); "
                "with RateShock = 1 the scenario's real-rate change passes through the beta: + (beta - 1) x change. Risk premium = rate - real rate.", bk.NOTE)
bk.section(ws, 30, "3. Monthly 12m fwd earnings yield by box (1 / harmonic P/E of the month's members; Bloomberg panel of the multiples model)", "B", "I")
ws.write("B31", "Month-end", bk.HDRL)
for g in EYG:
    ws.write(f"{EYC[g]}31", g, bk.HDR)
ws.set_row(30, 30)
for i_, (dt, rowv) in enumerate(EYH.iterrows()):
    r = 32 + i_
    ws.write_datetime(f"B{r}", pd.Timestamp(dt).to_pydatetime(), F(**IN, num_format=DATE, align="left"))
    for g in EYG:
        ws.write_number(f"{EYC[g]}{r}", float(rowv[g]), F(**IN, num_format=PCT2))
ws.freeze_panes(4, 2)

# =========================================================================== Box Betas
ws = W["Box Betas"]
bk.title(ws, "Box betas vs the Ibovespa (CAPM)", "Beta = covariance of the box's weekly total return with the Ibovespa's / variance of the Ibovespa's, over the window in Assumptions.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 46)
ws.set_column("C:Q", 11)
bk.section(ws, 5, "1. Beta of each box vs the Ibovespa (weekly total returns, window in Assumptions)", "B", "I")
bk.header_row(ws, 6, [("B", "Group"), ("C", "Weeks in the window"), ("D", "Average weekly return"), ("E", "Covariance with the Ibovespa"),
                      ("F", "Variance of the Ibovespa"), ("G", "Beta (raw)"), ("H", "Beta (Blume adjusted)"), ("I", "Beta used")])
ws.set_row(5, 44)
BW0 = 22                                   # first data row (the base week, levels only)
BWL = BW0 + len(BXW) - 1
LV_ = {g: cn(3 + i) for i, g in enumerate(EYG)}           # levels D..J
RT_ = {g: cn(3 + len(EYG) + i) for i, g in enumerate(EYG)}  # weekly returns K..Q
fl = f"$C${BW0 + 1}:$C${BWL}"
rng = lambda g: f"{RT_[g]}${BW0 + 1}:{RT_[g]}${BWL}"
for g, r in KRS.items():
    tot = g == "Ibovespa"
    t1 = dict(top=1) if tot else {}
    ws.write(f"B{r}", ("   " if g in MEMO else "") + g, F(bold=tot, **t1))
    ws.write_formula(f"C{r}", f"=SUM({fl})", F(num_format="0", **t1))
    ws.write_formula(f"D{r}", f"=SUMPRODUCT({fl},{rng(g)})/C{r}", F(num_format="0.000%", **t1))
    ws.write_formula(f"E{r}", f"=(SUMPRODUCT({fl},{rng(g)},{rng('Ibovespa')})-C{r}*D{r}*$D$11)/(C{r}-1)", F(num_format="0.000000", **t1))
    ws.write_formula(f"F{r}", f"=(SUMPRODUCT({fl},{rng('Ibovespa')},{rng('Ibovespa')})-C{r}*$D$11^2)/(C{r}-1)", F(num_format="0.000000", **t1))
    ws.write_formula(f"G{r}", f"=E{r}/F{r}", F(num_format="0.00", **t1))
    ws.write_formula(f"H{r}", f"=0.67*G{r}+0.33", F(num_format="0.00", **t1))
    ws.write_formula(f"I{r}", f"=IF(BetaAdj=2,H{r},G{r})", F(num_format="0.00", bold=True, **t1))
ws.write("B12", "Memo (inside Cyclicals)", bk.NOTE)
ws.write("B16", "Weighted average of the four box betas (~1)", bk.TXT)
ws.write_formula("C16", "=SUMPRODUCT(I7:I10,'Box Shares'!C7:C10)/SUM('Box Shares'!C7:C10)", F(num_format="0.00"))
ws.write("B17", "Window: weeks after", bk.TXT)
ws.write_formula("C17", f"=EDATE($B${BWL},-12*BetaYears)", F(num_format=DATED))
ws.write("B18", "Source: box total-return indices (Economatica adjusted close, members at the previous month-end index weights; Oct-5 performance study, "
                "box_daily.parquet). The Ibovespa here is the sum of the four boxes.", bk.NOTE)
bk.section(ws, 20, "2. Weekly data (Friday closes): index levels (inputs) and returns (formulas)", "B", "Q")
ws.write("B21", "Week", bk.HDRL)
ws.write("C21", "In window (1/0)", bk.HDR)
for g in EYG:
    ws.write(f"{LV_[g]}21", f"{g}, level", bk.HDR)
    ws.write(f"{RT_[g]}21", f"{g}, return", bk.HDR)
ws.set_row(20, 44)
for k_, (dt, rowv) in enumerate(BXW.iterrows()):
    r = BW0 + k_
    ws.write_datetime(f"B{r}", pd.Timestamp(dt).to_pydatetime(), F(**IN, num_format=DATED, align="left"))
    for g in EYG:
        ws.write_number(f"{LV_[g]}{r}", float(rowv[g]), F(**IN, num_format="0.0000"))
    if k_ == 0:
        continue
    ws.write_formula(f"C{r}", f"=IF(B{r}>$C$17,1,0)", F(num_format="0", align="center"))
    for g in EYG:
        ws.write_formula(f"{RT_[g]}{r}", f"={LV_[g]}{r}/{LV_[g]}{r - 1}-1", F(num_format="0.00%"))
ws.freeze_panes(22, 2)

# =========================================================================== Box Earnings
ws = W["Box Earnings"]
bk.title(ws, "Box earnings: sensitivity to the domestic and commodity cycles (DCF scenarios)",
         "Earnings beta = covariance of the box's 12-month EPS growth with the domestic cycle's (Ibovespa ex-Commodities) / variance of the domestic cycle's, over the window in Assumptions.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 46)
ws.set_column("C:AD", 11)
bk.section(ws, 5, "1. Earnings beta of each box to the domestic cycle (12-month change of the 12m fwd EPS, log; window in Assumptions)", "B", "L")
bk.header_row(ws, 6, [("B", "Group"), ("C", "Months in the window"), ("D", "Average 12-month growth"), ("E", "Std. dev. of 12-month growth"),
                      ("F", "Covariance with the domestic cycle"), ("G", "Earnings beta, estimated"), ("H", "Override (blank = estimated)"),
                      ("I", "Domestic beta used"), ("J", "Commodity price elasticity"), ("K", "Correlation with the domestic cycle"), ("L", "Std. dev. of 4-year growth, a year")])
ws.set_row(5, 58)
f12 = f"$C${E0 + 12}:$C${EL}"
f48 = f"$D${E0 + 48}:$D${EL}"
y12 = lambda g_: f"{EG12[g_]}${E0 + 12}:{EG12[g_]}${EL}"
y48 = lambda g_: f"{EG48[g_]}${E0 + 48}:{EG48[g_]}${EL}"
DJL, DJC = cn(4 + 3 * len(EGR)), cn(5 + 3 * len(EGR))           # commodity prices: level, 12-month log change
DJ12 = f"{DJC}${E0 + 12}:{DJC}${EL}"
for g, r in EBR.items():
    dom = g == DOMC
    t1 = dict(top=1) if dom else {}
    ws.write(f"B{r}", "Domestic cycle: Ibovespa ex-Commodities" if dom else ("   " if g in MEMO else "") + g, F(bold=dom, **t1))
    ws.write_formula(f"C{r}", f"=SUM({f12})", F(num_format="0", **t1))
    ws.write_formula(f"D{r}", f"=SUMPRODUCT({f12},{y12(g)})/C{r}", F(num_format=PCT, **t1))
    ws.write_formula(f"E{r}", f"=SQRT((SUMPRODUCT({f12},{y12(g)},{y12(g)})-C{r}*D{r}^2)/(C{r}-1))", F(num_format=PCT, **t1))
    ws.write_formula(f"F{r}", f"=(SUMPRODUCT({f12},{y12(g)},{y12(DOMC)})-C{r}*D{r}*$D${EBR[DOMC]})/(C{r}-1)", F(num_format="0.0000", **t1))
    ws.write_formula(f"G{r}", f"=F{r}/$E${EBR[DOMC]}^2", F(num_format="0.00", **t1))
    if dom:
        ws.write_formula(f"I{r}", f"=G{r}", F(num_format="0.00", bold=True, **t1))
        ws.write_number(f"J{r}", 0, F(num_format="0.00", **t1))
    else:
        if g == "Commodities":
            ws.write_number(f"H{r}", 0, F(**LV, num_format="0.00"))
        else:
            ws.write_blank(f"H{r}", None, F(**LV, num_format="0.00"))
        ws.write_formula(f"I{r}", f"=IF(ISNUMBER(H{r}),H{r},G{r})", F(num_format="0.00", bold=True))
        if g == "Commodities":
            mx = f"SUMPRODUCT({f12},{DJ12})/C{r}"
            ws.write_formula(f"J{r}", f"=(SUMPRODUCT({f12},{y12(g)},{DJ12})-C{r}*D{r}*{mx})/(SUMPRODUCT({f12},{DJ12},{DJ12})-C{r}*({mx})^2)", F(num_format="0.00", bold=True))
        else:
            ws.write_number(f"J{r}", 0, F(**LV, num_format="0.00"))
    ws.write_formula(f"K{r}", f"=F{r}/(E{r}*$E${EBR[DOMC]})", F(num_format="0.00", **t1))
    n48 = f"SUM({f48})"
    ws.write_formula(f"L{r}", f"=SQRT((SUMPRODUCT({f48},{y48(g)},{y48(g)})-{n48}*(SUMPRODUCT({f48},{y48(g)})/{n48})^2)/({n48}-1))", F(num_format=PCT, **t1))
ws.write("B12", "Memo (inside Cyclicals)", bk.NOTE)
ws.write("B15", "Weighted average of the domestic boxes (Financials, Defensives, Cyclicals; index points)", bk.TXT)
ws.write_formula("I15", "=SUMPRODUCT(I7:I9,'Box Shares'!$C$7:$C$9)/SUM('Box Shares'!$C$7:$C$9)", F(num_format="0.00"))
ws.write("B16", "Normalization factor of the domestic betas (EBNorm)", bk.TXT)
ws.write_formula("I16", "=IF(EBNorm=1,1/I15,1)", F(num_format="0.000", bold=True))
ws.write("B17", "Window: 12-month changes from", bk.TXT)
ws.write_formula("C17", "=EBStart", F(num_format=DATE))
ws.write("B18", "Commodities follow commodity prices (correlation ~0.5 with the domestic cycle): domestic beta 0; its earnings move by the elasticity to "
                "commodity prices (column J, 12-month changes vs. the DJP commodity ETN). WEG & Embraer barely move with the domestic cycle.", bk.NOTE)
bk.section(ws, 19, "2. Earnings level shock by box and scenario: normalized domestic beta x Ibovespa ex-Commodities shock + price elasticity x commodity prices", "B", "L")
bk.header_row(ws, 20, [("B", "Group"), ("C", "Bear"), ("D", "Base"), ("E", "Bull")])
ws.set_row(19, 44)
for g, r in EKR.items():
    ws.write(f"B{r}", ("   " if g in MEMO else "") + g, bk.TXT)
    for c in "CDE":
        ws.write_formula(f"{c}{r}", f"=$I${EBR[g]}*$I$16*Assumptions!{c}$61+$J${EBR[g]}*Assumptions!{c}$62", F(num_format=UPS, bold=True))
ws.write("B25", "Ibovespa (weighted by index points)", F(bold=True, top=1))
for c in "CDE":
    ws.write_formula(f"{c}25", f"=SUMPRODUCT({c}21:{c}24,'Box Shares'!$C$7:$C$10)/SUM('Box Shares'!$C$7:$C$10)", F(num_format=UPS, bold=True, top=1))
ws.write("B26", "Memo (inside Cyclicals)", bk.NOTE)
ws.write("B29", "The shock moves the level of the box's earnings and flows in every year, 2028 onward (DCF by Box, Box Fade); growth is unchanged. Zero in the base.", bk.NOTE)
bk.section(ws, 30, "3. Monthly 12m fwd EPS index by group (matched sample, chain-linked; inputs) and its growth (formulas); commodity prices (DJP)", "B", "AD")
ws.write("B31", "Month-end", bk.HDRL)
ws.write("C31", "In window, 12-month (1/0)", bk.HDR)
ws.write("D31", "In window, 4-year (1/0)", bk.HDR)
for g in EGR:
    nm_ = "Ibovespa ex-Commodities" if g == DOMC else g
    ws.write(f"{ELV[g]}31", f"{nm_}, EPS index", bk.HDR)
    ws.write(f"{EG12[g]}31", f"{nm_}, 12-month growth (log)", bk.HDR)
    ws.write(f"{EG48[g]}31", f"{nm_}, 4-year growth a year (log)", bk.HDR)
ws.write(f"{DJL}31", "Commodity prices (DJP ETN)", bk.HDR)
ws.write(f"{DJC}31", "Commodity prices, 12-month change (log)", bk.HDR)
ws.set_row(30, 58)
for k_, (dt, rowv) in enumerate(EPSI.iterrows()):
    r = E0 + k_
    ws.write_datetime(f"B{r}", pd.Timestamp(dt).to_pydatetime(), F(**IN, num_format=DATE, align="left"))
    for g in EGR:
        ws.write_number(f"{ELV[g]}{r}", float(rowv[g]), F(**IN, num_format="0.0000"))
    ws.write_number(f"{DJL}{r}", float(DJP.loc[dt]), F(**IN, num_format="0.00"))
    if k_ >= 12:
        ws.write_formula(f"{DJC}{r}", f"=LN({DJL}{r}/{DJL}{r - 12})", F(num_format=PCT))
    if k_ >= 12:
        ws.write_formula(f"C{r}", f"=IF(B{r}>=EBStart,1,0)", F(num_format="0", align="center"))
        for g in EGR:
            ws.write_formula(f"{EG12[g]}{r}", f"=LN({ELV[g]}{r}/{ELV[g]}{r - 12})", F(num_format=PCT))
    if k_ >= 48:
        ws.write_formula(f"D{r}", f"=IF(B{r}>=EBStart,1,0)", F(num_format="0", align="center"))
        for g in EGR:
            ws.write_formula(f"{EG48[g]}{r}", f"=LN({ELV[g]}{r}/{ELV[g]}{r - 48})/4", F(num_format=PCT))
ws.freeze_panes(32, 2)

# =========================================================================== Box Fade
ws = W["Box Fade"]
bk.title(ws, "Fade after 2031, year by year (FadeMode 1 and 2)",
         "Growth steps down from the starting growth to g over FadeYears (mid-year steps: the H-model's area), then the Gordon formula; values at end-2031.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 64)
ws.set_column("C:N", 10)
DBQ = "'DCF by Box'"
bk.section(ws, 5, "1. Value at end-2031 of the flows after 2031: year by year vs. the H-model formula", "B", "N")
bk.header_row(ws, 6, [("B", "Group"), ("C", "Year by year: Bear"), ("D", "Base"), ("E", "Bull"), ("F", "H-model formula: Bear"), ("G", "Base"), ("H", "Bull"),
                      ("I", "Base: year by year / H-model - 1")])
ws.set_row(5, 44)
for g, r in FSR.items():
    fin = g == "Financials"
    ws.write(f"B{r}", ("   " if g in MEMO else "") + g + (" (FCFE)" if fin else ""), bk.TXT)
    for c, hc in zip("CDE", "FGH"):
        ws.write_formula(f"{c}{r}", f"={FADEV(g, c)}", F(num_format=PTS))
        lst = f"{DBQ}!{c}{BB[g]['e_last' if fin else 'last']}"
        ws.write_formula(f"{hc}{r}", f"={lst}*((1+gLT)+FadeYears/2*({DBQ}!{c}{BB[g]['gstart']}-gLT))/({DBQ}!{c}{BB[g]['rate']}-gLT)", F(num_format=PTS))
    ws.write_formula(f"I{r}", f"=D{r}/G{r}-1", F(num_format=UPS, bold=True))
ws.write("B11", "Memo (inside Cyclicals)", bk.NOTE)
ws.write("B14", "Fade in use (Assumptions section 8)", bk.TXT)
ws.write_formula("C14", '=CHOOSE(FadeMode+1,"0: H-model formula (this sheet is not used)","1: year by year","2: year by year, growth paid by reinvestment")', F(bold=True))
ws.write("B15", "FadeMode 1: the 2031 flow grows year by year. FadeMode 2: flow = earnings x (1 - growth / ROE) - NWC % x change in sales (Financials: net income x "
                "(1 - growth / ROE)): growth costs reinvestment at the box's FY28 ROE (Box Shares section 5).", bk.NOTE)
for g in GROUPS:
    fin = g == "Financials"
    roe = "FinROE" if fin else f"'Box Shares'!$E${SR5[g]}"
    for c, s_ in SCN:
        r0 = FADE[(g, c)]
        yr, gr, ea, sa, fl_, df, pv, vl = (r0 + i for i in range(1, 9))
        ws.write(f"B{r0}", f"{g} · {s_}", bk.HDRL)
        ws.write(f"C{r0}", "2031", bk.HDR)
        for k in range(1, 11):
            ws.write(f"{cn(2 + k)}{r0}", str(2031 + k), bk.HDR)
        ws.write(f"N{r0}", "2042: g", bk.HDR)
        ws.write(f"B{yr}", "Year of the fade", bk.TXT)
        ws.write(f"B{gr}", "Growth: starting growth (DCF by Box) stepping down to g over FadeYears", bk.TXT)
        ws.write(f"B{ea}", ("Net income" if fin else "NOPAT") + " 2031 x (1 + shock), then x (1 + growth)", bk.TXT)
        ws.write(f"B{sa}", "Sales (not used: FCFE)" if fin else "Sales 2031 x (1 + shock), then x (1 + growth)", F(font_color=GREY) if fin else bk.TXT)
        ws.write(f"B{fl_}", "Flow: FadeMode 1 = 2031 flow x (1 + growth); 2 = earnings x (1 - growth / ROE)" + ("" if fin else " - NWC % x change in sales"), bk.BOLD)
        ws.write(f"B{df}", "Discount factor from end-2031 at the box rate", bk.TXT)
        ws.write(f"B{pv}", "Present value at end-2031", bk.TXT)
        ws.write(f"B{vl}", "Value at end-2031 = PV 2032-41 + flow 2042 / (rate - g), discounted", bk.BOLD)
        gs = f"{DBQ}!${c}${BB[g]['gstart']}"
        rt = f"{DBQ}!${c}${BB[g]['rate']}"
        sh = f"{DBQ}!${c}${BB[g]['shock']}"
        ws.write_formula(f"C{ea}", f"={DBQ}!$H${BB[g]['ni' if fin else 'nopat']}*(1+{sh})", F(num_format=PTS))
        if not fin:
            ws.write_formula(f"C{sa}", f"={DBQ}!$H${BB[g]['sales']}*(1+{sh})", F(num_format=PTS))
        ws.write_formula(f"C{fl_}", f"={DBQ}!${c}${BB[g]['e_last' if fin else 'last']}", F(num_format=PTS, bold=True))
        for k in range(1, 12):
            col, prv = cn(2 + k), cn(1 + k)
            ws.write_number(f"{col}{yr}", k, F(num_format="0", align="center"))
            ws.write_formula(f"{col}{gr}", f"=IF({col}{yr}<=FadeYears,{gs}-({gs}-gLT)*({col}{yr}-0.5)/FadeYears,gLT)", F(num_format=PCT))
            ws.write_formula(f"{col}{ea}", f"={prv}{ea}*(1+{col}{gr})", F(num_format=PTS))
            if fin:
                paid = f"{col}{ea}*(1-{col}{gr}/{roe})"
            else:
                ws.write_formula(f"{col}{sa}", f"={prv}{sa}*(1+{col}{gr})", F(num_format=PTS))
                paid = f"{col}{ea}*(1-{col}{gr}/{roe})-NWCpct*({col}{sa}-{prv}{sa})"
            ws.write_formula(f"{col}{fl_}", f"=IF(FadeMode=2,{paid},{prv}{fl_}*(1+{col}{gr}))", F(num_format=PTS, bold=True))
            if k <= 10:
                ws.write_formula(f"{col}{df}", f"=1/(1+{rt})^{col}{yr}", F(num_format="0.0000"))
                ws.write_formula(f"{col}{pv}", f"={col}{fl_}*{col}{df}", F(num_format=PTS))
        ws.write_formula(f"C{vl}", f"=SUM(D{pv}:M{pv})+N{fl_}/({rt}-gLT)*M{df}", F(num_format=PTS, bold=True, bg_color=KEY))
ws.freeze_panes(4, 2)

# =========================================================================== Multiples by Box
ws = W["Multiples by Box"]
bk.title(ws, "Target by box: house multiples, bottom-up and the average of the methods (end-2027)",
         "House Target methods for 2027 split by box: P/E and EV/EBITDA at each box's 10y relative multiple (calibrated to the house value), bottom-up from the analysts' TPs.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 60)
ws.set_column("C:Z", 11)
ASC = {"C": "C", "D": "D", "E": "E"}            # scenario columns here = Assumptions section 7 columns (Bear, Base, Bull)
MSC = [("C", "Bear"), ("D", "Base"), ("E", "Bull")]
NIc, EBc, EB25 = f"DCF!${YC[2027]}${R['ni']}", f"DCF!${YC[2027]}${R['ebitda']}", f"DCF!${YC[2025]}${R['ebitda']}"
# ---- section 6 layout first (the relative multiples are referenced above)
PEC = {g: cn(2 + i) for i, g in enumerate(EYG)}                       # C..I: P/E of Fin, Def, Cyc, Comm, Ibov, CycEx, WEG
EVC = {g: cn(9 + i) for i, g in enumerate(EVG)}                       # J..O: EV/EBITDA of Def, Cyc, Comm, Ibov ex-Fin, CycEx, WEG
RPG = [g for g in EYG if g != "Ibovespa"]
RPC = {g: cn(15 + i) for i, g in enumerate(RPG)}                       # P..U: relative P/E
REG = [g for g in EVG if g != "Ibovespa ex-Financials"]
REC = {g: cn(21 + i) for i, g in enumerate(REG)}                       # V..Z: relative EV/EBITDA
H0, H1 = 70, 70 + len(MPE) - 1
assert len(MPE) == 120 and len(MEV) == 120
RELPE = {g: f"${RPC[g]}$68" for g in RPG}
RELEV = {g: f"${REC[g]}$68" for g in REG}

bk.section(ws, 5, "1. House methods for the Ibovespa, end-2027 (Target rows 67-90)", "B", "J")
bk.header_row(ws, 6, [("B", ""), ("C", "Bear"), ("D", "Base"), ("E", "Bull")])
S1 = [(7, "Target P/E on 2027 EPS (Assumptions)", lambda c: f"=Assumptions!{c}53", MULT),
      (8, "EPS 2027, house (DCF sheet)", lambda c: f"={NIc}", PTS),
      (9, "P/E method = target P/E x EPS 2027", lambda c: f"={c}7*{c}8", PTS),
      (10, "Target EV/EBITDA on 2027 EBITDA (Assumptions)", lambda c: f"=Assumptions!{c}54", MULT),
      (11, "EBITDA 2027, house (DCF sheet)", lambda c: f"={EBc}", PTS),
      (12, "Net debt = 1.6x EBITDA 2025 (house)", lambda c: f"=NDmult*{EB25}", PTS),
      (13, "EV/EBITDA method = target x EBITDA 2027 - net debt", lambda c: f"={c}10*{c}11-{c}12", PTS),
      (14, "Bottom-up: members at the analysts' target prices (section 4)", lambda c: f"={c}50", PTS),
      (15, "DCF: Ibovespa by box (DCF by Box; = the house DCF in the base)", lambda c: f"='DCF by Box'!{c}11", PTS),
      (16, "Average of the methods (weights in Assumptions)", lambda c: f"=(WDCF*{c}15+WPE*{c}9+WEV*{c}13+WBU*{c}14)/(WDCF+WPE+WEV+WBU)", PTS),
      (17, "Ibovespa used", lambda c: "=IbovUsed", PTS),
      (18, "Upside of the average", lambda c: f"={c}16/{c}17-1", UPS)]
for r, lab, fn, fmt in S1:
    b = r in (16, 18)
    ws.write(f"B{r}", lab, F(bold=b))
    for c, _ in MSC:
        ws.write_formula(f"{c}{r}", fn(c), F(num_format=fmt, bold=b, bg_color=KEY if b else "#FFFFFF"))

GR7 = BOXES + ["Ibovespa", None] + MEMO
# ---- section 2: P/E by box
bk.section(ws, 20, "2. P/E method by box: house P/E x the box's 10y relative P/E, calibrated to the house value", "B", "J")
bk.header_row(ws, 21, [("B", "Group"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), ("F", ""), ("G", "Relative P/E (10y average)"),
                       ("H", "EPS 2027 (pts)"), ("I", "P/E today on 2027 EPS"), ("J", "Target P/E, base")])
ws.set_row(20, 42)
ws.write("B22", "Calibration: house P/E value / sum of (relative P/E x box EPS)", bk.TXT)
for c, _ in MSC:
    ws.write_formula(f"{c}22", f"={c}9/SUMPRODUCT($G$23:$G$26,$H$23:$H$26)", F(num_format="0.00"))
PER = {}
for i, g in enumerate(GR7):
    r = 23 + i
    if g is None:
        ws.write(f"B{r}", "Memo (inside Cyclicals; own relative multiple, not additive)", bk.NOTE)
        continue
    PER[g] = r
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", g, F(**b))
    if tot:
        for c in "CDEH":
            ws.write_formula(f"{c}{r}", f"=SUM({c}23:{c}26)", F(num_format=PTS, **b))
        ws.write_formula(f"I{r}", f"='Box Shares'!$C$11/H{r}", F(num_format=MULT, **b))
        continue
    ws.write_formula(f"G{r}", f"=IF(RelMult=1,{RELPE[g]},1)", F(num_format="0.00"))
    ws.write_formula(f"H{r}", f"={NIc}*'Box Shares'!{SB[('ni', '27')]}{SR2[g]}", F(num_format=PTS))
    ws.write_formula(f"I{r}", f"='Box Shares'!$C${SR1[g]}/H{r}", F(num_format=MULT))
    for c, _ in MSC:
        ws.write_formula(f"{c}{r}", f"={c}$22*$G{r}*$H{r}", F(num_format=PTS, bold=True))
    ws.write_formula(f"J{r}", f"=D$22*G{r}", F(num_format=MULT))

# ---- section 3: EV/EBITDA by box
bk.section(ws, 32, "3. EV/EBITDA method by box: house multiple x the box's 10y relative EV/EBITDA, calibrated (Financials: P/E value)", "B", "J")
bk.header_row(ws, 33, [("B", "Group"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), ("F", ""), ("G", "Relative EV/EBITDA (10y average)"),
                       ("H", "EBITDA 2027 (pts)"), ("I", "Net debt (pts)"), ("J", "Target EV/EBITDA, base")])
ws.set_row(32, 42)
ws.write("B34", "Calibration: (house value - Financials + others' net debt) / boxes' EV", bk.TXT)
for c, _ in MSC:
    ws.write_formula(f"{c}34", f"=({c}13-{c}35+SUM($I$36:$I$38))/({c}10*SUMPRODUCT($G$36:$G$38,$H$36:$H$38))", F(num_format="0.000"))
EVR = {}
for i, g in enumerate(GR7):
    r = 35 + i
    if g is None:
        ws.write(f"B{r}", "Memo (inside Cyclicals; own relative multiple, not additive)", bk.NOTE)
        continue
    EVR[g] = r
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", g + (" (P/E value: EV/EBITDA does not apply)" if g == "Financials" else ""), F(**b))
    if tot:
        for c in "CDEHI":
            ws.write_formula(f"{c}{r}", f"=SUM({c}35:{c}38)", F(num_format=PTS, **b))
        continue
    if g == "Financials":
        for c, _ in MSC:
            ws.write_formula(f"{c}{r}", f"={c}{PER['Financials']}", F(num_format=PTS, bold=True))
        continue
    ws.write_formula(f"G{r}", f"=IF(RelMult=1,{RELEV[g]},1)", F(num_format="0.00"))
    ws.write_formula(f"H{r}", f"={EBc}*'Box Shares'!{SB[('ebitda', '27')]}{SR2[g]}", F(num_format=PTS))
    ws.write_formula(f"I{r}", f"=NDmult*{EB25}*'Box Shares'!{SB[('ebitda', '26')]}{SR2[g]}", F(num_format=PTS))
    for c, _ in MSC:
        ws.write_formula(f"{c}{r}", f"={c}$34*{c}$10*$G{r}*$H{r}-$I{r}", F(num_format=PTS, bold=True))
    ws.write_formula(f"J{r}", f"=D$34*D$10*G{r}", F(num_format=MULT))

# ---- section 4: bottom-up by box
bk.section(ws, 44, "4. Bottom-up by box: the members at the analysts' 12m target prices (Members), x (1 + change) in the bear / bull", "B", "J")
bk.header_row(ws, 45, [("B", "Group"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), ("F", ""), ("G", "Index points with a target price"),
                       ("H", "Index points at the target price"), ("I", "Upside of the covered members"), ("J", "")])
ws.set_row(44, 42)
BUR = {}
for i, g in enumerate(GR7):
    r = 46 + i
    if g is None:
        ws.write(f"B{r}", "Memo (inside Cyclicals)", bk.NOTE)
        continue
    BUR[g] = r
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", g, F(**b))
    if tot:
        for c in "CDEGH":
            ws.write_formula(f"{c}{r}", f"=SUM({c}46:{c}49)", F(num_format=PTS, **b))
        ws.write_formula(f"I{r}", f"=H{r}/G{r}-1", F(num_format=UPS, **b))
        continue
    crit = "C" if g in BOXES else "D"
    ws.write_formula(f"G{r}", f"=SUMIFS({MR(CBP)},{MR(crit)},\"{g}\")", F(num_format=PTS))
    ws.write_formula(f"H{r}", f"=SUMIFS({MR(CBU)},{MR(crit)},\"{g}\")", F(num_format=PTS))
    ws.write_formula(f"I{r}", f"=H{r}/G{r}-1", F(num_format=UPS))
    for c, _ in MSC:
        ws.write_formula(f"{c}{r}", f"='Box Shares'!$C${SR1[g]}*(1+$I{r})*(1+Assumptions!{c}$55)", F(num_format=PTS, bold=True))

# ---- section 5: average of the methods by box
bk.section(ws, 55, "5. Average of the methods by box (DCF by Box + sections 2-4; weights in Assumptions)", "B", "J")
bk.header_row(ws, 56, [("B", "Group"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), ("F", ""), ("G", "Upside: Bear"), ("H", "Base"), ("I", "Bull"), ("J", "")])
AVR = {}
for i, g in enumerate(GR7):
    r = 57 + i
    if g is None:
        ws.write(f"B{r}", "Memo (inside Cyclicals; not additive)", bk.NOTE)
        continue
    AVR[g] = r
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", g, F(**b))
    for c, uc in zip("CDE", "GHI"):
        if tot:
            ws.write_formula(f"{c}{r}", f"=SUM({c}57:{c}60)", F(num_format=PTS, bg_color=KEY, **b))
            ws.write_formula(f"{uc}{r}", f"={c}{r}/'Box Shares'!$C$11-1", F(num_format=UPS, **b))
        else:
            ws.write_formula(f"{c}{r}", f"=(WDCF*'DCF by Box'!{c}{BB[g]['val']}+WPE*{c}{PER[g]}+WEV*{c}{EVR[g]}+WBU*{c}{BUR[g]})/(WDCF+WPE+WEV+WBU)",
                             F(num_format=PTS, bold=True, bg_color=KEY))
            ws.write_formula(f"{uc}{r}", f"={c}{r}/'Box Shares'!$C${SR1[g]}-1", F(num_format=UPS))
ws.conditional_format("G57:I61", HEAT)
ws.conditional_format("G63:I64", HEAT)

# ---- section 6: monthly multiples
bk.section(ws, 66, "6. Monthly 12m fwd P/E and EV/EBITDA (harmonic means of the month's members; Bloomberg panel of the multiples model) and the relative multiples", "B", "Z")
ws.write("B67", "Month-end", bk.HDRL)
for g in EYG:
    ws.write(f"{PEC[g]}67", f"P/E {g}", bk.HDR)
for g in EVG:
    ws.write(f"{EVC[g]}67", f"EV/EBITDA {g}", bk.HDR)
for g in RPG:
    ws.write(f"{RPC[g]}67", f"Relative P/E {g}", bk.HDR)
for g in REG:
    ws.write(f"{REC[g]}67", f"Relative EV/EBITDA {g}", bk.HDR)
ws.set_row(66, 54)
ws.write("B68", "10-year average (Oct-16..Sep-26)", bk.BOLD)
for g in RPG:
    ws.write_formula(f"{RPC[g]}68", f"=AVERAGE({RPC[g]}{H0}:{RPC[g]}{H1})", F(num_format="0.00", bold=True))
for g in REG:
    ws.write_formula(f"{REC[g]}68", f"=AVERAGE({REC[g]}{H0}:{REC[g]}{H1})", F(num_format="0.00", bold=True))
for g in EYG:
    ws.write_formula(f"{PEC[g]}68", f"=AVERAGE({PEC[g]}{H0}:{PEC[g]}{H1})", F(num_format=MULT, bold=True))
for g in EVG:
    ws.write_formula(f"{EVC[g]}68", f"=AVERAGE({EVC[g]}{H0}:{EVC[g]}{H1})", F(num_format=MULT, bold=True))
for k_, (dt, rowv) in enumerate(MPE.iterrows()):
    r = H0 + k_
    ws.write_datetime(f"B{r}", pd.Timestamp(dt).to_pydatetime(), F(**IN, num_format=DATE, align="left"))
    for g in EYG:
        ws.write_number(f"{PEC[g]}{r}", float(rowv[g]), F(**IN, num_format="0.00"))
    ev_row = MEV.loc[dt]
    for g in EVG:
        if pd.notna(ev_row[g]):
            ws.write_number(f"{EVC[g]}{r}", float(ev_row[g]), F(**IN, num_format="0.00"))
    for g in RPG:
        ws.write_formula(f"{RPC[g]}{r}", f"={PEC[g]}{r}/{PEC['Ibovespa']}{r}", F(num_format="0.00"))
    for g in REG:
        ws.write_formula(f"{REC[g]}{r}", f"=IFERROR({EVC[g]}{r}/{EVC['Ibovespa ex-Financials']}{r},\"\")", F(num_format="0.00"))
ws.freeze_panes(4, 2)

# =========================================================================== Summary
ws = W["Summary"]
bk.title(ws, "Target by box at the end of 2027 (house model)",
         "House Target methods rolled to end-2027 and split by box: DCF (CAPM), P/E, EV/EBITDA, bottom-up, and their average. Construction: Multiples by Box, DCF by Box.", legend=False)
ws.set_column("A:A", 2)
ws.set_column("B:B", 56)
ws.set_column("C:J", 13)
ws.write("B3", "Ibovespa used", bk.NOTE)
ws.write_formula("C3", "=IbovUsed", F(num_format=PTS, bold=True))
ws.write("D3", "Market data", bk.NOTE)
ws.write_formula("E3", '=CHOOSE(Vintage,"house file","Oct-5-2026")', F(bold=True))
ws.write("F3", "Checks", bk.NOTE)
ws.write_formula("G3", "=Checks!C5", F(bold=True))
ws.write("H3", "Boxes scaled", bk.NOTE)
ws.write_formula("I3", '=IF(ScaleBox=1,"yes","no")', F(bold=True))
ws.write("B4", "DCF box settings", bk.NOTE)
ws.write_formula("C4", '=CHOOSE(BoxRate+1,"house rate","rate relative to history","CAPM rates")&IF(BoxCF=1,", capex funding growth and NWC by change","")'
                       '&IF(BoxGrowth=1,", consensus growth fading after 2031","")&CHOOSE(FadeMode+1," (H-model)"," (year by year)"," (year by year, growth paid)")'
                       '&IF(RateShock=1,"; scenario rate shock x beta","")&IF(ScaleScen=1,"; earnings shocks open the bear / bull","")'
                       '&IF(RelMult=1,"; multiples relative to history","")', F(italic=True, font_color=GREY))
GR8 = BOXES + ["Ibovespa", None] + MEMO


def table(top, title, src, kind):
    """Bear / base / bull values and upsides by group; src(g, c) gives the value formula for a box or memo."""
    bk.section(ws, top, title, "B", "J")
    bk.header_row(ws, top + 1, [("B", "Group"), ("C", "Index points"), ("D", "Bear"), ("E", "Base"), ("F", "Bull"), ("G", ""),
                                ("H", "Upside: Bear"), ("I", "Base"), ("J", "Bull")])
    rows = {}
    for i, g in enumerate(GR8):
        r = top + 2 + i
        if g is None:
            ws.write(f"B{r}", "Memo (inside Cyclicals; not additive)", bk.NOTE)
            continue
        rows[g] = r
        tot = g == "Ibovespa"
        b = dict(bold=True, top=1) if tot else {}
        ws.write(f"B{r}", g, F(**b))
        if tot:
            ws.write_formula(f"C{r}", f"=SUM(C{top + 2}:C{top + 5})", F(num_format=PTS, **b))
            for c in "DEF":
                ws.write_formula(f"{c}{r}", f"=SUM({c}{top + 2}:{c}{top + 5})", F(num_format=PTS, **b))
        else:
            ws.write_formula(f"C{r}", f"='Box Shares'!C{SR1[g]}", F(num_format=PTS, **b))
            for c, sc in zip("DEF", "CDE"):
                ws.write_formula(f"{c}{r}", src(g, sc), F(num_format=PTS, **b))
        for c, uc in zip("DEF", "HIJ"):
            ws.write_formula(f"{uc}{r}", f"={c}{r}/$C{r}-1", F(num_format=UPS, **b))
    ws.conditional_format(f"H{top + 2}:J{top + 6}", HEAT)
    ws.conditional_format(f"H{top + 8}:J{top + 9}", HEAT)
    return rows


SUMR = table(5, "1. Target at the end of 2027 by box: average of the methods (index pts)", lambda g, sc: f"='Multiples by Box'!{sc}{AVR[g]}", "avg")

MREFROW = {"Bear": 79, "Base": 80, "Bull": 81}
MREFCOL = {"Financials": "D", "Defensives": "E", "Cyclicals": "F", "Commodities": "G", "Ibovespa": "H", "Cyclicals ex-WEG & Embraer": "I", "WEG & Embraer": "J"}
bk.section(ws, 16, "2. Base case by method (upside vs. index points) and the multiples model", "B", "J")
bk.header_row(ws, 17, [("B", "Group"), ("C", ""), ("D", "DCF (CAPM)"), ("E", "P/E"), ("F", "EV/EBITDA"), ("G", "Bottom-up"), ("H", "Average of the methods"),
                       ("I", "Multiples model (Oct-6)"), ("J", "")])
ws.set_row(16, 30)
for i, g in enumerate(GR8):
    r = 18 + i
    if g is None:
        ws.write(f"B{r}", "Memo (inside Cyclicals)", bk.NOTE)
        continue
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", g, F(**b))
    pts_ = f"$C${SUMR[g]}"
    if tot:
        srcs = [f"='DCF by Box'!$D$11/{pts_}-1", f"='Multiples by Box'!$D${PER[g]}/{pts_}-1", f"='Multiples by Box'!$D${EVR[g]}/{pts_}-1",
                f"='Multiples by Box'!$D${BUR[g]}/{pts_}-1", f"='Multiples by Box'!$D${AVR[g]}/{pts_}-1"]
    else:
        srcs = [f"='DCF by Box'!$D${BB[g]['up']}", f"='Multiples by Box'!$D${PER[g]}/{pts_}-1", f"='Multiples by Box'!$D${EVR[g]}/{pts_}-1",
                f"='Multiples by Box'!$D${BUR[g]}/{pts_}-1", f"='Multiples by Box'!$D${AVR[g]}/{pts_}-1"]
    for c, f_ in zip("DEFGH", srcs):
        ws.write_formula(f"{c}{r}", f_, F(num_format=UPS, bold=(c == "H") or tot, top=1 if tot else 0))
    ws.write_formula(f"I{r}", f"=${MREFCOL[g]}${MREFROW['Base']}/{pts_}-1", F(num_format=UPS, font_color=GREY, **b))
ws.write("B26", "Average absolute gap to the multiples model, four boxes", bk.BOLD)
for c in "DEFGH":
    ws.write_formula(f"{c}26", "=AVERAGE(" + ",".join(f"ABS({c}{r}-I{r})" for r in range(18, 22)) + ")", F(num_format=PCT, bold=True, top=1))
ws.conditional_format("D18:H22", HEAT)

DSUMR = table(28, "3. DCF by box (CAPM, box cash flows, Gordon after 2031 as the house; not scaled: the Ibovespa is the sum of the boxes)",
              lambda g, sc: f"='DCF by Box'!{sc}{BB[g]['val']}", "dcf")

bk.section(ws, 39, "4. Discount rate by box (Box Rates)", "B", "J")
bk.header_row(ws, 40, [("B", "Group"), ("C", "Relative cash yield"), ("D", "Rate: Bear"), ("E", "Base"), ("F", "Bull"), ("G", "CAPM beta"),
                       ("H", "Risk premium: Bear"), ("I", "Base"), ("J", "Bull")])
ws.set_row(39, 30)
for i, g in enumerate(GR8):
    r = 41 + i
    if g is None:
        ws.write(f"B{r}", "Memo (inside Cyclicals)", bk.NOTE)
        continue
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", "Ibovespa (house rate)" if tot else g, F(**b))
    ws.write_formula(f"C{r}", f"='Box Rates'!G{KRS[g]}", F(num_format="0.000", **b))
    ws.write_formula(f"G{r}", f"='Box Betas'!I{KRS[g]}", F(num_format="0.00", **b))
    src_r = 19 if tot else KR[g]
    for c, sc in zip("DEF", "CDE"):
        ws.write_formula(f"{c}{r}", f"='Box Rates'!{sc}{src_r}", F(num_format=PCT2, **b))
    for c, sc in zip("HIJ", "GHI"):
        ws.write_formula(f"{c}{r}", f"='Box Rates'!{sc}{src_r}", F(num_format=PCT2, **b))
ws.write("B49", "Rates in use (BoxRate in Assumptions; default 2 = CAPM, real rate + box beta x the house ERP). Betas: Box Betas.", bk.NOTE)

bk.section(ws, 51, "5. DCF: each box with its own value, before scaling (Financials by FCFE)", "B", "J")
bk.header_row(ws, 52, [("B", ""), ("C", ""), ("D", "Bear"), ("E", "Base"), ("F", "Bull"), ("G", ""), ("H", "Upside: Bear"), ("I", "Base"), ("J", "Bull")])
for i, g in enumerate(BOXES):
    r = 53 + i
    ws.write(f"B{r}", g, bk.TXT)
    for c, sc, uc in zip("DEF", "CDE", "HIJ"):
        ws.write_formula(f"{c}{r}", f"='DCF by Box'!{sc}{BB[g]['val_own']}", F(num_format=PTS))
        ws.write_formula(f"{uc}{r}", f"='DCF by Box'!{sc}{BB[g]['up_own']}", F(num_format=UPS))
for i, (lab, row_, fmt, b) in enumerate([("Sum of the boxes", 7, PTS, True), ("House DCF fair value, Ibovespa (DCF sheet)", 8, PTS, False),
                                          ("Scale factor = house / sum of the boxes", 9, "0.000", False),
                                          ("Factor applied (ScaleScen = 1: the base factor in every scenario)", 10, "0.000", False)]):
    r = 57 + i
    ws.write(f"B{r}", lab, F(bold=b, top=1 if i == 0 else 0))
    for c, sc in zip("DEF", "CDE"):
        ws.write_formula(f"{c}{r}", f"='DCF by Box'!{sc}{row_}", F(num_format=fmt, bold=b, top=1 if i == 0 else 0))

bk.section(ws, 61, "6. Scenario parameters and settings (Assumptions)", "B", "J")
bk.header_row(ws, 62, [("B", ""), ("C", ""), ("D", "Bear"), ("E", "Base"), ("F", "Bull")])
for i, (lab, f_, fmt) in enumerate([("5-year real rate (NTN-B)", lambda sc: f"=DCF!{sc}{V['rr']}", PCT2), ("ERP (house)", lambda sc: f"=DCF!{sc}{V['erp']}", PCT2),
                                     ("House discount rate", lambda sc: f"=DCF!{sc}{V['rate']}", PCT2), ("Target P/E on 2027 EPS", lambda sc: f"=Assumptions!{sc}53", MULT),
                                     ("Target EV/EBITDA on 2027 EBITDA", lambda sc: f"=Assumptions!{sc}54", MULT), ("Bottom-up change", lambda sc: f"=Assumptions!{sc}55", UPS)]):
    r = 63 + i
    ws.write(f"B{r}", lab, bk.TXT)
    for c, sc in zip("DEF", "CDE"):
        ws.write_formula(f"{c}{r}", f_(sc), F(num_format=fmt, align="right"))
ws.write("B69", "Corrections to the house aggregate (Assumptions section 4)", bk.NOTE)
ws.write_formula("D69", '=IF(FixNWC+FixND+FixTV+FixInfl=0,"none: as the house",'
                        '"on: "&IF(FixNWC=1,"NWC change ","")&IF(FixND=1,"net debt ","")&IF(FixTV=1,"TV timing ","")&IF(FixInfl=1,"inflation",""))', F(italic=True, font_color=GREY))

bk.section(ws, 70, "7. Box shares of the house lines (Box Shares; last consensus year)", "B", "J")
bk.header_row(ws, 71, [("B", "Line"), ("C", ""), ("D", "Financials"), ("E", "Defensives"), ("F", "Cyclicals"), ("G", "Commodities")])
for i, (k, y) in enumerate([("sales", "27"), ("ebit", "28"), ("ebitda", "28"), ("ni", "28"), ("capex", "27")]):
    r = 72 + i
    ws.write(f"B{r}", f"{LABL[k]} {2000 + int(y)}", bk.TXT)
    for c, g in zip("DEFG", BOXES):
        ws.write_formula(f"{c}{r}", f"='Box Shares'!{SB[(k, y)]}{SR2[g]}", F(num_format=PCT))

bk.section(ws, 77, "8. Reference (not linked): multiples model and our DCF by box (index pts)", "B", "J")
bk.header_row(ws, 78, [("B", "Model"), ("C", ""), ("D", "Financials"), ("E", "Defensives"), ("F", "Cyclicals"), ("G", "Commodities"), ("H", "Ibovespa"),
                       ("I", "Cyc. ex-WEG & EMB"), ("J", "WEG & Embraer")])
ws.set_row(77, 30)
for s_, r in MREFROW.items():
    ws.write(f"B{r}", f"Multiples model, average of methods (Oct-6 table): {s_}", bk.TXT)
    for g, c in MREFCOL.items():
        ws.write_number(f"{c}{r}", MULTREF[(g, s_)], F(**IN, num_format=PTS))
ws.write("B82", "DCF by box (Ibov_DCF_by_box_XP.xlsx, 12 months): Base", bk.TXT)
for c, v in zip("DEFGHIJ", (70757, 42152, 24906, 85903, 223718, 17780, 7140)):
    ws.write_number(f"{c}82", v, F(**IN, num_format=PTS))
ws.write("B83", "Multiples model: Ibov_valuation_by_box_model_XP.xlsx (prices of Oct-5; each box at its own 10y average multiples, plus bottom-up).", bk.NOTE)

ws.set_column("K:L", 13)
bk.section(ws, 85, "9. Scenarios of the DCF by box: shocks, rates and the change vs. the base (multiples model for comparison)", "B", "L")
bk.header_row(ws, 86, [("B", "Group"), ("C", "Earnings beta, domestic (normalized)"), ("D", "Commodity price elasticity"), ("E", "Earnings shock: Bear"), ("F", "Bull"),
                       ("G", "Rate: Bear"), ("H", "Bull"), ("I", "DCF vs. base: Bear"), ("J", "Bull"), ("K", "Multiples model vs. base: Bear"), ("L", "Bull")])
ws.set_row(85, 44)
for i, g in enumerate(GR8):
    r = 87 + i
    if g is None:
        ws.write(f"B{r}", "Memo (inside Cyclicals)", bk.NOTE)
        continue
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", g, F(**b))
    if tot:
        src9 = ["='Box Earnings'!I15*'Box Earnings'!I16", "=0", "='Box Earnings'!C25", "='Box Earnings'!E25", "='Box Rates'!C19", "='Box Rates'!E19",
                "='DCF by Box'!C11/'DCF by Box'!D11-1", "='DCF by Box'!E11/'DCF by Box'!D11-1"]
    else:
        vr = BB[g]["val"]
        src9 = [f"='Box Earnings'!I{EBR[g]}*'Box Earnings'!$I$16", f"='Box Earnings'!J{EBR[g]}", f"='Box Earnings'!C{EKR[g]}", f"='Box Earnings'!E{EKR[g]}",
                f"='Box Rates'!C{KR[g]}", f"='Box Rates'!E{KR[g]}", f"='DCF by Box'!C{vr}/'DCF by Box'!D{vr}-1", f"='DCF by Box'!E{vr}/'DCF by Box'!D{vr}-1"]
    for c, f_, nf in zip("CDEFGHIJ", src9, ["0.00", "0.00", UPS, UPS, PCT2, PCT2, UPS, UPS]):
        ws.write_formula(f"{c}{r}", f_, F(num_format=nf, **b) if tot else F(num_format=nf, bold=c in "IJ"))
    mc = MREFCOL[g]
    ws.write_formula(f"K{r}", f"=${mc}${MREFROW['Bear']}/${mc}${MREFROW['Base']}-1", F(num_format=UPS, font_color=GREY, **b))
    ws.write_formula(f"L{r}", f"=${mc}${MREFROW['Bull']}/${mc}${MREFROW['Base']}-1", F(num_format=UPS, font_color=GREY, **b))
ws.conditional_format("I87:L91", HEAT)
ws.conditional_format("I93:L94", HEAT)
ws.write("B95", "Bear / bull: the house real-rate scenarios through each box's beta (RateShock) and the earnings shocks of Assumptions section 8 x each box's earnings "
                "beta (Commodities: the commodity shock). Multiples model: the Oct-6 table (section 8).", bk.NOTE)

bk.section(ws, 97, "10. DCF by box: expected value (scenario probabilities) and the growth priced in today", "B", "L")
bk.header_row(ws, 98, [("B", "Group"), ("C", "Expected value (pts)"), ("D", "Expected upside"), ("E", "Upside: Bear"), ("F", "Base"), ("G", "Bull"),
                       ("H", "Bull gain / bear loss vs. base"), ("I", "Consensus growth 2026-28"), ("J", "Starting growth used (base)"),
                       ("K", "Growth priced in today (reverse DCF)"), ("L", "")])
ws.set_row(97, 44)
prs = "SUM(Assumptions!$C$63:$E$63)"
for i, g in enumerate(GR8):
    r = 99 + i
    if g is None:
        ws.write(f"B{r}", "Memo (inside Cyclicals)", bk.NOTE)
        continue
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", g, F(**b))
    vals10 = [f"'DCF by Box'!{c}11" for c in "CDE"] if tot else [f"'DCF by Box'!{c}{BB[g]['val']}" for c in "CDE"]
    pts_ = f"$C${DSUMR[g]}"
    fb = (lambda nf: F(num_format=nf, **b)) if tot else (lambda nf: F(num_format=nf, bold=True))
    ws.write_formula(f"C{r}", f"=(Assumptions!$C$63*{vals10[0]}+Assumptions!$D$63*{vals10[1]}+Assumptions!$E$63*{vals10[2]})/{prs}", fb(PTS))
    ws.write_formula(f"D{r}", f"=C{r}/{pts_}-1", fb(UPS))
    for c, v_ in zip("EFG", vals10):
        ws.write_formula(f"{c}{r}", f"={v_}/{pts_}-1", F(num_format=UPS, **b))
    ws.write_formula(f"H{r}", f"=({vals10[2]}-{vals10[1]})/({vals10[1]}-{vals10[0]})", F(num_format='0.00"x"', **b))
    if not tot:
        ws.write_formula(f"I{r}", f"='Box Shares'!H{SR5[g]}", F(num_format=PCT))
        ws.write_formula(f"J{r}", f"='DCF by Box'!D{BB[g]['gstart']}", F(num_format=PCT))
        ws.write_formula(f"K{r}", f"='DCF by Box'!D{BB[g]['gimp']}", F(num_format=PCT, bold=True))
ws.conditional_format("D99:G103", HEAT)
ws.conditional_format("D105:G106", HEAT)
ws.write("B107", "Expected value: probabilities in Assumptions section 8. Gain / loss = (bull - base) / (base - bear). Growth priced in: the starting growth of the "
                 "fade at which the box's own value equals today's index points (base rate and flows; DCF by Box).", bk.NOTE)
ws.freeze_panes(4, 0)

# =========================================================================== Checks
ws = W["Checks"]
bk.title(ws, "Checks", "'OK' / 'CHECK' drive the status in C5; 'Info' does not. House tie-outs hold while the projection rules are the house's.", legend=False)
ws.set_column("A:A", 2)
ws.set_column("B:B", 78)
ws.set_column("C:F", 15)
ws.set_column("G:G", 80)
ws.write("B5", "Overall status", bk.BOLD)
bk.header_row(ws, 7, [("B", "Check"), ("C", "Value"), ("D", "Target"), ("E", "Difference"), ("F", "Status"), ("G", "Note")])
HS = 90                                    # house file settings: rr, ERP, g, Ibovespa in C90:C93; helper flows in row 94
hs = (f"Checks!$C${HS}", f"Checks!$C${HS + 1}", f"Checks!$C${HS + 2}")
DB = "'DCF by Box'"
chk = [("House replica", None, None, None, None, None)]
for y in range(2025, 2031):
    chk.append((f"FCFF {y}, house row = Target row 44", f"=DCF!{YC[y]}{R['fcff_house']}", f"={H['fcff'][y]}", 0.01, "House file value.", PX))
chk += [("Terminal value on 2030, house rate (Target!R44)", f"=DCF!{YC[2030]}{R['fcff_house']}*(1+{hs[2]})/({hs[0]}+{hs[1]}-{hs[2]})", f"={H['tv']}", 0.01, "", PX),
        ("Fair value end-2025, house settings (Target!D59)", "=" + fv_expr(2025, 2030, *hs, vy=2025, **HOUSEF), f"={H['fv25']}", 0.01, "", PX),
        ("Fair value end-2026, house settings (Target!F59)", "=" + fv_expr(2027, 2030, *hs, vy=2026, **HOUSEF), f"={H['fv26']}", 0.01, "", PX),
        ("Fair value end-2027, house file column (Target!G59)", "=" + fv_expr(2028, 2030, *hs, vy=2027, **HOUSEF), f"={H['fv27']}", 0.01, "", PX),
        ("Real-rate table, largest gap to the house (Sensitivity)", f"=MAX(MAX(Sensitivity!E7:E{SENS_LAST}),-MIN(Sensitivity!E7:E{SENS_LAST}))", "=0", 0.01, "", PX),
        ("House implied return at the house price (Target!B62)", f"=IRR(Checks!C{HS + 4}:G{HS + 4})", f"={H['irr']}", 1e-6, f"Helper row {HS + 4}.", PCT2),
        ("Split by box", None, None, None, None, None),
        ("Member weights on Oct-5 add up to 100%", f"=Members!I{MT}", "=1", 1e-9, "", "0.0000%"),
        ("Members' index points = Ibovespa on Oct-5", f"=Members!J{MT}", "=Assumptions!$D$7", 0.01, "", PX),
        ("Box shares add up to 100% (all lines and years; sum of absolute gaps)", "=SUMPRODUCT(ABS('Box Shares'!D23:R23-1))", "=0", 1e-9, "", "0.000000"),
        ("Memo sub-boxes = Cyclicals (index points)", "='Box Shares'!C14+'Box Shares'!C15", "='Box Shares'!C9", 0.01, "", PX),
        ("FCFF used: sum of the boxes = house, 2028-2031 (sum of absolute gaps)",
         "=SUMPRODUCT(ABS(" + "+".join(f"{DB}!E{BB[g]['fcff_used']}:H{BB[g]['fcff_used']}" for g in BOXES) + f"-DCF!{YC[2028]}{R['fcff_used']}:{YC[2031]}{R['fcff_used']}))",
         "=0", 0.01, "", PX)]
for c, s_ in SCN:
    chk.append((f"{s_}: sum of the FCFF pieces = house value (exact split)", f"={DB}!{c}12", f"={DB}!{c}8", 0.01, "", PX))
chk += [("Base: memo sub-boxes = Cyclicals at the house rate", f"={DB}!D{BB['Cyclicals ex-WEG & Embraer']['val_h']}+{DB}!D{BB['WEG & Embraer']['val_h']}",
         f"={DB}!D{BB['Cyclicals']['val_h']}", 0.01, "At the box rates the memo rows have their own rates and are not additive.", PX),
        ("Base: sum of the box fair values = house value (or the own sum if not scaled)", f"={DB}!D11", f"=IF(ScaleBox=1,{DB}!D8,{DB}!D7)", 0.01, "", PX),
        ("Base: scale factor (house / sum of own values)", f"={DB}!D9", "=1", None, "Info: Financials by FCFE vs. its FCFF piece explains the gap.", "0.000"),
        ("Base: Financials FCFE value / its FCFF piece - 1", f"={DB}!D{BB['Financials']['fcfe_val']}/{DB}!D{BB['Financials']['fcff_val']}-1", "=0", None, "Info.", UPS),
        ("Bear: Ibovespa by box vs. the house DCF (sum / house - 1)", f"={DB}!C15", "=0", None, "Info: ScaleScen = 1 lets the earnings shocks open the bear / bull.", UPS),
        ("Bull: Ibovespa by box vs. the house DCF (sum / house - 1)", f"={DB}!E15", "=0", None, "Info.", UPS),
        ("Base: the earnings shocks are zero (the base is the consensus)", "=ABS(Assumptions!D61)+ABS(Assumptions!D62)", "=0", None, "Info.", PCT),
        ("Year-by-year fade: FadeYears is at most 10 (Box Fade has 10 years)", "=MAX(FadeYears-10,0)", "=0", 0, "", "0"),
        ("Fade after 2031 (0 = H-model formula, 1 = year by year, 2 = growth paid)", "=FadeMode", "=0", None, "Info.", "0"),
        ("Base: Commodities, year-by-year fade / H-model formula - 1 (Box Fade)", "='Box Fade'!I10", "=0", None,
         "Info: the H-model formula undervalues negative starting growth.", UPS),
        ("Earnings betas: months of 12-month EPS changes in the window (Box Earnings)", "='Box Earnings'!C7", "=228", None, "Info: Oct-07..Sep-26 by default.", "0"),
        ("Earnings betas: the domestic cycle's own beta = 1", "='Box Earnings'!G11", "=1", 1e-9, "", "0.00"),
        ("Earnings betas: domestic boxes, weighted average after normalization", "='Box Earnings'!I15*'Box Earnings'!I16", "=1", None, "Info: 1 with EBNorm = 1.", "0.00"),
        ("Commodities: earnings elasticity to commodity prices (Box Earnings)", "='Box Earnings'!J10", "=1", None, "Info.", "0.00"),
        ("Box rates: months of earnings-yield history (Box Rates)", "=COUNT('Box Rates'!C32:C151)", "=120", 0, "", "0"),
        ("Box rates in use (0 = house rate, 1 = relative to history, 2 = CAPM)", "=BoxRate", "=2", None, "Info.", "0"),
        ("CAPM betas: weighted average of the four boxes (index points)", "='Box Betas'!C16", "=1", None, "Info: close to 1, the Ibovespa's beta.", "0.00"),
        ("Base: average gap to the multiples model, four boxes, DCF (upside, pp)", "=Summary!D26*100", "=0", None, "Info: Summary section 2.", "0.0"),
        ("Base: average gap to the multiples model, four boxes, average of the methods (pp)", "=Summary!H26*100", "=0", None, "Info: Summary section 2.", "0.0"),
        ("Members with a missing consensus line", f"=COUNTIF(Members!{CFLAG}{MF}:{CFLAG}{ML},\"?*\")", "=0", None, "Info: missing lines count as 0 (banks: EBITDA, capex).", "0"),
        ("Target by box (Multiples by Box)", None, None, None, None, None),
        ("P/E 2027 with the house inputs (Target!F77)", f"={T['F76'].value}*DCF!${YC[2027]}${R['ni']}", f"={T['F77'].value}", 0.01, "House file value.", PX),
        ("EV/EBITDA 2027 with the house inputs (Target!F72)", f"={T['F67'].value}*DCF!${YC[2027]}${R['ebitda']}-{T['F70'].value}*DCF!${YC[2025]}${R['ebitda']}",
         f"={T['F72'].value}", 0.01, "House file value.", PX)]
MB = "'Multiples by Box'"
for c, s_ in SCN:
    chk.append((f"{s_}: P/E method, the boxes add up to the house value", f"={MB}!{c}27", f"={MB}!{c}9", 0.01, "", PX))
    chk.append((f"{s_}: EV/EBITDA method, the boxes add up to the house value", f"={MB}!{c}39", f"={MB}!{c}13", 0.01, "", PX))
for c, s_ in SCN:
    chk.append((f"{s_}: average of the methods, the boxes add up to the Ibovespa average", f"={MB}!{c}61", f"={MB}!{c}16", 0.01, "", PX))
chk += [("Months of multiples history (P/E)", f"=COUNT({MB}!C70:C189)", "=120", 0, "", "0"),
        ("Members with a target price (index points, share of the Ibovespa)", f"=Members!{CBP}{MT}/Members!J{MT}", "=1", None, "Info: bottom-up coverage.", PCT),
        ("Settings", None, None, None, None, None),
        ("Corrections switched on (Assumptions section 4)", "=FixNWC+FixND+FixTV+FixInfl", "=0", None, "Info: 0 = the house mechanics.", "0"),
        ("Market data used (1 = house file, 2 = Oct-5)", "=Vintage", "=2", None, "Info.", "0"),
        ("Formula errors (DCF, DCF by Box, Multiples by Box, Summary, Sensitivity, Box Shares, Box Rates, Box Betas, Box Earnings, Box Fade, Members)",
         f"=SUMPRODUCT(--ISERROR(DCF!C8:S90))+SUMPRODUCT(--ISERROR({DB}!C5:H{LASTBOX}))+SUMPRODUCT(--ISERROR(Summary!C4:L110))+SUMPRODUCT(--ISERROR({MB}!C5:Z190))"
         f"+SUMPRODUCT(--ISERROR('Box Earnings'!C7:AB{EL}))+SUMPRODUCT(--ISERROR('Box Fade'!C5:N{FADE_LAST}))"
         f"+SUMPRODUCT(--ISERROR(Sensitivity!C7:G60))+SUMPRODUCT(--ISERROR('Box Shares'!C7:R40))+SUMPRODUCT(--ISERROR('Box Rates'!C7:I151))+SUMPRODUCT(--ISERROR('Box Betas'!C7:Q600))+SUMPRODUCT(--ISERROR(Members!E{MF}:{CBU}{MT}))", "=0", 0, "", "0")]
r = 8
for lab, val, tgt, tol, note, fmt in chk:
    if val is None:
        ws.write(f"B{r}", lab, bk.SUB)
        for c in "CDEFG":
            ws.write_blank(f"{c}{r}", None, bk.SUB)
        r += 1
        continue
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
assert r < HS - 2, r
bk.section(ws, HS - 1, "House file settings used by the tie-outs (Target sheet)", "B", "G")
for i, (lab, v, fmt) in enumerate([("5-year real rate (Target!C56)", H["rr"], PCT2), ("ERP (Target!B57)", H["erp"], PCT2),
                                    ("g (Target!B54)", H["g"], PCT2), ("Ibovespa (Target!B61)", H["ibov"], PX)]):
    ws.write(f"B{HS + i}", lab, bk.TXT)
    ws.write_number(f"C{HS + i}", v, F(**IN, num_format=fmt))
ws.write(f"B{HS + 4}", "Helper: house implied-return flows (Target row 61)", bk.NOTE)
ws.write_formula(f"C{HS + 4}", f"=-C{HS + 3}", F(num_format=PTS))
for c, y in zip("DEF", (2026, 2027, 2028)):
    ws.write_formula(f"{c}{HS + 4}", f"=DCF!{YC[y]}{R['fcff_house']}", F(num_format=PTS))
ws.write_formula(f"G{HS + 4}", f"=DCF!{YC[2030]}{R['fcff_house']}*(1+C{HS + 2})/(C{HS}+C{HS + 1}-C{HS + 2})", F(num_format=PTS))
ws.freeze_panes(7, 0)

# =========================================================================== Read Me
ws = W["Read Me"]
ws.set_column("A:A", 2)
ws.set_column("B:B", 26)
ws.set_column("C:C", 130)
ws.set_row(0, 26)
ws.write("B1", "Ibovespa by box, house target model rolled to end-2027", F(bold=True, font_size=16, font_color=NAVY))
ws.write("B2", "XP Research · Equity Strategy · built Oct-7-2026 · house file inputs (Sep-26); members' consensus Oct-7-2026; market data Oct-5-2026", bk.NOTE)
rr_ = 4
for head, lines in [
    ("What it is", [("Target", "The house Ibovespa target (DCF Ibov 2026_Out.xlsx, Target sheet) rolled forward to the end of 2027 and split by box: the average of four "
                               "methods (DCF, target P/E, target EV/EBITDA, bottom-up), as the house's Target row 88."),
                    ("Aggregate", "The house DCF block is replicated with the same inputs, rules and mechanics (DCF sheet); Checks tie it, and the 2027 P/E and EV/EBITDA, "
                                  "to the house file.")]),
    ("DCF by box", [("Lines", "Every house line (sales, EBIT, EBITDA, net income, capex) is split by the box's share of the IBOV Index line, measured with the members' "
                              "consensus (Box Shares). Banks have no EBITDA or capex; the index imputes them, and the imputed part goes to Financials."),
                    ("Rates", "CAPM (BoxRate = 2, default): real rate + box beta x 6.0% ERP. Betas from the boxes' weekly total-return indices (Box Betas). "
                              "BoxRate = 0: the house rate; 1: rate relative to the box's historical cash yield."),
                    ("Cash flows", "BoxCF = 1 (default): capex = D&A + the reinvestment that funds the year's growth at the box's FY28 ROE, and working capital by its "
                                   "change; the house aggregate keeps the house rules. Financials: FCFE = net income x (1 - g / ROE)."),
                    ("Growth", "As the house: after 2031 the flows grow at g = 4.7% forever (Gordon; BoxGrowth = 0, FadeMode = 0), no fade and no cap. Optional: BoxGrowth = 1 lets the "
                               "box's consensus net income growth 2026-28 fade to g over 10 years (FadeMode 1 = year by year, Box Fade)."),
                    ("Scenarios", "Two levers, as the house: the real rate (8.5% bear / 5.5% bull, the same change for every box) and the earnings level (-20% / +20% for the "
                                  "Ibovespa ex-Commodities, split by each domestic box's normalized earnings beta; Commodities by its earnings elasticity to commodity prices "
                                  "x -20% / +20% commodity prices). Growth after 2031 is the same in every scenario. The base is the consensus."),
                    ("Sum", "Not scaled (ScaleBox = 0, default): the Ibovespa by box is the sum of the four boxes and may differ from the house DCF (DCF sheet). ScaleBox = 1 scales the base to the house value. Summary sections 9-10: sensitivity, expected value and the growth priced in today."),]),
    ("Multiples and bottom-up by box", [("P/E", "House target P/E (10.0x base, 7.5x / 12x bear / bull) x the box's 10y average P/E relative to the Ibovespa, calibrated so the "
                                                "boxes add up to the house value (P/E x 2027 EPS). RelMult = 0: the house P/E for every box."),
                                        ("EV/EBITDA", "House target EV/EBITDA (5.5x base, 4.5x / 6.0x) x the box's 10y relative EV/EBITDA, less net debt (1.6x EBITDA 2025 "
                                                      "split by EBITDA share), calibrated to the house value; Financials take their P/E value."),
                                        ("Bottom-up", "Members at the analysts' 12m target prices (house COMP SHEET, consensus when missing), -20% / +20% in the bear / bull."),
                                        ("Average", "Weighted average of the four methods (equal weights, as the house); the boxes add up to the house's Ibovespa average.")]),
    ("How to use it", [("1. Change", "Assumptions: market data, g, scenarios, corrections, box settings (section 6: ScaleBox, BoxRate, BoxCF, BoxGrowth, fade, cap, betas), "
                                     "the target methods (section 7: multiples, bottom-up change, RelMult, weights) and the DCF scenarios (section 8: earnings shocks, "
                                     "probabilities, RateShock, EBNorm, ScaleScen, FadeMode, earnings-beta window)."),
                       ("2. Read", "Summary: the target by box (average of the methods), each method in the base case, the DCF by box and the rates."),
                       ("3. Trace", "Summary -> Multiples by Box / DCF by Box -> Box Rates, Box Betas, Box Earnings, Box Fade, Box Shares -> Members; DCF (house aggregate).")]),
    ("Sources", [("House file", "History 2015-2025 and IBOV Index consensus 2026-2028 as saved in DCF Ibov 2026_Out.xlsx (Sep-26); not refreshed."),
                 ("Members", "Bloomberg BEST_SALES, BEST_EBIT, BEST_EBITDA, BEST_NET_INCOME, BEST_CAPEX (FY26-FY28, BRL), CUR_MKT_CAP, BEST_ROE (Oct-7-2026); "
                             "target prices from the house COMP SHEET (consensus when missing). Weights: Economatica Sep-30, drifted by price to Oct-5."),
                 ("Histories", "Monthly 12m fwd P/E and EV/EBITDA (Bloomberg panel of the multiples model); weekly box total returns (Economatica, Oct-5 performance study)."),
                 ("Earnings", "Monthly 12m fwd EPS index by box, matched sample chain-linked, Oct-06..Sep-26 (Oct-6 EPS study, eps_index_matched.parquet)."),
                 ("Market", "Ibovespa 208,431.92 (Oct-5 close); 5-year NTN-B 6.76% (ANBIMA, Oct-5)."),
                 ("Code", "Brazil Bull Case/scripts: box_lines_data.py (Bloomberg) and build_house2027_box.py (this workbook).")])]:
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
ws = bk.cover("Ibovespa", " | target by box (house model), end-2027", "VALUATION MODEL  ·  EQUITY STRATEGY",
              [("Summary", "Summary"), ("Assumptions", "Assumptions"), ("DCF by Box", "DCF by Box"), ("Multiples\nby Box", "Multiples by Box")],
              [[("DCF\n(Ibovespa)", "DCF"), ("Sensitivity", "Sensitivity"), ("Box Rates", "Box Rates"), ("Box Betas", "Box Betas")],
               [("Box\nEarnings", "Box Earnings"), ("Box Fade", "Box Fade"), ("Box Shares", "Box Shares"), ("Members\n(76 stocks)", "Members")],
               [("Checks", "Checks")]], readme="Read Me")
Ib = SUMR["Ibovespa"]
for i, (l_, f_, nf) in enumerate([("Ibovespa used", "=IbovUsed", PTS), ("Target end-2027, base (average)", f"=Summary!E{Ib}", PTS),
                                   ("Upside, base", f"=Summary!I{Ib}", UPS), ("Target end-2027, bear", f"=Summary!D{Ib}", PTS),
                                   ("Target end-2027, bull", f"=Summary!F{Ib}", PTS), ("Model checks", "=Checks!C5", None)]):
    r = 8 + i
    bb = i == 1
    ws.merge_range(f"K{r}:L{r}", l_, F(font_name="Roboto Light", bold=bb, bottom=4))
    fmt_ = F(align="right", bottom=4, bold=bb, num_format=nf) if nf else F(align="right", bottom=4, bold=True, font_color=GREEN)
    ws.merge_range(f"M{r}:N{r}", "", fmt_)
    ws.write_formula(f"M{r}", f_, fmt_)
ws.write("K16", "Equity Strategy | XP Research", F(bold=True))
gl = F(bold=True, italic=True, font_size=9, font_color=GREY)
hd = F(bold=True, bottom=2, bottom_color="#8EB3DF", align="center")
ws.write("C21", "Target by box, base (average of the methods) >>", gl)
ws.merge_range("C22:E22", "Box", F(bold=True, bottom=2, bottom_color="#8EB3DF"))
for j, h_ in enumerate(["Index points", "Target", "Upside"]):
    ws.write(21, 5 + j, h_, hd)
for i, g in enumerate(BOXES + ["Ibovespa"]):
    r = 23 + i
    b = g == "Ibovespa"
    ws.merge_range(f"C{r}:E{r}", g, F(bold=b, top=1 if b else 0))
    for j, (c, nf) in enumerate([("C", PTS), ("E", PTS), ("I", UPS)]):
        ws.write_formula(r - 1, 5 + j, f"=Summary!{c}{SUMR[g]}", F(bold=b, num_format=nf, align="center", top=1 if b else 0))
ws.write("C29", "House target methods rolled to end-2027 and split by box: DCF (CAPM), P/E, EV/EBITDA and bottom-up, and their average.", bk.NOTE)
bk.close()
print("written", OUT)
