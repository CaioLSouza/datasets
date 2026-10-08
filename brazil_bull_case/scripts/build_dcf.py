"""Builds Ibov_DCF_by_box_XP.xlsx: a DCF of the Ibovespa by box (FCFE, value-driver form), in the XP model template look.

Each box: earnings in Ibovespa points from consensus (Bloomberg, FY26-FY28, blended to forward years 1-2); growth converges
from the consensus year-2 growth to the long term by GrowthYear and ROE to its 10-year average by RoeYear; FCFE = earnings -
increase in book value (book = earnings / ROE); Gordon terminal value on the long-term ROE. Ibovespa = sum of the four boxes.
Scenarios Bear / Base / Bull / Custom, one sheet each.
Main premises in line with the house Ibovespa model (DCF Ibov 2026_Out.xlsx, Target sheet): 5y real rate (NTN-B) + ERP 6.0%,
long-term growth 4.7% nominal, IPCA 4.0%, bear / bull real rates 8.5% / 6.0%. Discount-rate convention (Assumptions RateConv):
1 = house, real rate + ERP used as the discount rate of the nominal cash flows; 2 = Fisher, (1 + real rate + ERP) x (1 + IPCA) - 1.
ERP basis: 1 typed, 2 Ibovespa 10-year average implied ERP, 3 each box's 10-year average implied ERP, where the implied ERP is
measured each month from the box's P/E and ROE in the same convention (History sheet).
Separate from the multiples model.

Run: python dcf_data.py (Bloomberg) -> python build_dcf.py -> python excel_recalc.py ..\\Ibov_DCF_by_box_XP.xlsx
"""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd

from xp_template import (Book, cn, NAVY, YEL, GREY, GREEN, INK, KEY, PTS, PTS1, MULT, PCT, PCT2, UPS, PP, PPC, PX, DATE,
                         DATED, HEAT)

HERE = Path(__file__).parent
SRCM = Path(r"C:\Users\Caio\Documents\Documentos\Equity Strategy Dashboard\_pedidos\2026-10-05_ibov_valuation_groups\scripts")
OUT = HERE.parent / "Ibov_DCF_by_box_XP.xlsx"

BOXES = ["Financials", "Defensives", "Cyclicals", "Commodities"]
MEMO = ["Cyclicals ex-WEG & Embraer", "WEG & Embraer"]
G6 = BOXES + MEMO
HG = ["Financials", "Defensives", "Cyclicals", "Commodities", "Ibovespa", "Cyclicals ex-WEG & Embraer", "WEG & Embraer"]
SCEN = ["Bear", "Base", "Bull", "Custom"]
SC = dict(zip(SCEN, "DEFG"))
GICS_MAP = [("Financials", "Financials"), ("Consumer Staples", "Defensives"), ("Health Care", "Defensives"),
            ("Utilities", "Defensives"), ("Communication Services", "Defensives"), ("Consumer Discretionary", "Cyclicals"),
            ("Industrials", "Cyclicals"), ("Information Technology", "Cyclicals"), ("Real Estate", "Cyclicals"),
            ("Energy", "Commodities"), ("Materials", "Commodities")]
CRIT = {"Financials": ("Financials", "*"), "Defensives": ("Defensives", "*"), "Cyclicals": ("Cyclicals", "*"),
        "Commodities": ("Commodities", "*"), "Ibovespa": ("*", "*"), "Cyclicals ex-WEG & Embraer": ("Cyclicals", "Cyclicals ex-WEG & Embraer"),
        "WEG & Embraer": ("Cyclicals", "WEG & Embraer")}
NY = 5                                       # explicit forecast years (house model: 4 years after its value date + terminal)
YEARS = list(range(1, NY + 1))
YC = {t: cn(2 + t) for t in YEARS}          # year t -> column (D..H)
TCOL = cn(3 + NY)                            # terminal column
NC = cn(4 + NY)                              # notes column
HORIZON, ROELT = 1, 2                        # defaults: 12-month fair value; long-term ROE = consensus (as the house model)
ERPK = [0, 0, 0, 0]                          # ERP move in std. dev. (bear, base, bull, custom)
GADD = [-2, 0, 2, 0]                         # earnings growth vs consensus path, pp per year (years 2..NY)
GLT = [0.042, 0.047, 0.052, 0.047]           # long-term nominal growth (house 4.7% in the base)
PPF = '+0.0" pp";-0.0" pp";0.0" pp"'

# --------------------------------------------------------------------------- data
m = pd.read_parquet(SRCM / "bx_members.parquet")
m["bo"] = m["box"].map({b: i for i, b in enumerate(BOXES)})
m = m.sort_values(["bo", "glob", "w30"], ascending=[True, True, False])
dm = pd.read_parquet(HERE / "dcf_members.parquet").reindex(m.index)
hist = pd.read_parquet(HERE / "dcf_roe_hist.parquet").sort_values(["cod_ativo", "date"]).reset_index(drop=True)
bxd = json.loads((SRCM / "bx_data.json").read_text(encoding="utf-8"))
macro = json.loads((SRCM / "bx_macro.json").read_text(encoding="utf-8"))
rr5 = pd.Series(macro["real_rate"]).rename(lambda x: pd.Timestamp(x)) / 100
REF = json.loads((HERE / "ref_v6.json").read_text(encoding="utf-8"))
months = sorted(hist["date"].unique())
assert len(months) == 120, len(months)

ORDER = ["Cover", "Summary", "Assumptions", "DCF Base", "DCF Bear", "DCF Bull", "DCF Custom", "Implied ERP", "Sensitivity",
         "Support >", "Boxes Now", "History", "Members", "ROE Panel", "Checks", "Read Me"]
bk = Book(OUT, ORDER, {"Assumptions": YEL, "Support >": YEL})
F, W = bk.F, bk.W
IN, LV = bk.IN, bk.LV


def q(sheet):
    return f"'{sheet}'" if any(ch in sheet for ch in " &->") else sheet


def nz(x):
    return None if (x is None or (isinstance(x, float) and np.isnan(x))) else x


# History layout (needed by Assumptions names): ROE D..J, P/E K..Q, implied ERP R..X
HCOL = {k: {g: cn(3 + 7 * j + i) for i, g in enumerate(HG)} for j, k in enumerate(["roe", "pe", "erp"])}
HST = {"avg": 8, "sd": 9, "n": 10, "now": 11, "z": 12}
HF = 17
HL = HF + len(months) - 1

# =========================================================================== Assumptions
ws = W["Assumptions"]
bk.title(ws, "DCF assumptions", "The only sheet to edit: change the pink cells (dotted border); the DCF sheets recalculate.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 64)
ws.set_column("C:G", 14)
ws.set_column("H:H", 115)
A = {}


def one(r, lab, nm, v, fmt, note, lever=False):
    ws.write(f"B{r}", lab, bk.TXT)
    style = F(**(LV if lever else IN), num_format=fmt)
    if isinstance(v, str) and v.startswith("="):
        ws.write_formula(f"C{r}", v, F(num_format=fmt))
    elif isinstance(v, pd.Timestamp):
        ws.write_datetime(f"C{r}", v.to_pydatetime(), style)
    else:
        ws.write_number(f"C{r}", v, style)
    ws.write(f"H{r}", note, bk.NOTE)
    if nm:
        bk.name(nm, "Assumptions", f"$C${r}")
        A[nm] = r


r = 5
bk.section(ws, r, "1. Market snapshot")
for item in [("Pricing date (prices and index weights)", "PxDate", pd.Timestamp("2026-10-05"), DATED, "Close of Oct-5-2026, the same snapshot as the multiples model."),
             ("Ibovespa close (index pts)", "IbovNow", bxd["ibov_now"], PX, "Bloomberg IBOV Index PX_LAST."),
             ("Selic target today", "SelicNow", 0.1375, PCT2, "BCB (BZSTSETA Index)."),
             ("5-year real rate today (NTN-B)", "RR5Now", 0.0676, PCT2, "ANBIMA, Oct-5 (as in the multiples model). The house model uses the 5y (BZRFB5PY Index, Target!B56). "
                                                                            "For reference, the 10y (IPCA+ 2035/2037) was 6.86%."),
             ("Ibovespa membership date", "MemDate", pd.Timestamp("2026-09-30"), DATED, "Members and weights: Economatica Ibovespa composition, drifted by price to the pricing date."),
             ("Ibovespa close on the membership date (index pts)", "IbovMem", bxd["ibov_sep30"], PX, "Used only in Checks."),
             ("Consensus snapshot date (Bloomberg)", "ConsDate", pd.Timestamp("2026-10-07"), DATED, "BEST_PE_RATIO and BEST_ROE, FY26/FY27/FY28; earnings = price on this date / P/E.")]:
    r += 1
    one(r, *item)
r += 2
bk.section(ws, r, "2. Macro")
for lab, nm, v, fmt, note, lev in [("IPCA, forward year 1", "IPCA1", 0.040, PCT2, "House Ibovespa model (DCF Ibov 2026_Out.xlsx, Target sheet): 2027 4.0%. Used only with the Fisher convention (RateConv = 2).", True),
                                    ("IPCA, forward year 2", "IPCA2", 0.040, PCT2, "House model: 2028 4.0%.", True),
                                    ("IPCA, long term (year 3 onwards, terminal and history)", "IPCALT", 0.040, PCT2, "House model long-term inflation 4.0% (Target!B63).", True),
                                    ("Share of FY2026 still ahead on the pricing date", "FYleft", "=(DATE(2026,12,31)-PxDate)/365", "0.00",
                                     "Forward year 1 = this share of FY26 + the rest of FY27; year 2 = the same split of FY27 and FY28 (12m-forward convention).", False)]:
    r += 1
    one(r, lab, nm, v, fmt, note, lever=lev)
r += 2
bk.section(ws, r, "3. Scenario levers")
r += 1
bk.header_row(ws, r, [("B", "Lever"), ("C", "Today"), ("D", "Bear"), ("E", "Base"), ("F", "Bull"), ("G", "Custom"), ("H", "How it is used")])
keys = ["narr", "selic", "cut", "rr", "erp", "erps", "erpk", "erpu", "ke", "gadd", "g", "roe"]
for i, k in enumerate(keys):
    A[k] = r + 1 + i
r = A["roe"]
lev = [("narr", "Narrative", ["Moderate populist", "Base", "Reformist", "User-defined"], None, "Labels only."),
       ("selic", "Selic at the end of the horizon", [0.132, 0.115, 0.095, 0.115], PCT2, "Same scenarios as the multiples model; drives the rate-driven change in earnings (section 5)."),
       ("rr", "5-year real rate (NTN-B)", [0.085, "=RR5Now", 0.06, "=RR5Now"], PCT2,
        "As the house model (Target sheet): bear 8.5%, base = today's 5y real rate, bull 6.0%."),
       ("erp", "ERP, typed (used when the ERP basis = 1)", [0.06, 0.06, 0.06, 0.06], PCT2,
        "House ERP 6.0% (Target!B57), the same in every scenario."),
       ("erps", "ERP shift (pp, added to the ERP of any basis)", [0, 0, 0, 0], PPF, "Added to the ERP of the basis in use (typed, Ibovespa or box 10-year average)."),
       ("erpk", "ERP move in std. dev. of its history (k; +1 = ERP 1 sd lower)", ERPK, '+0.0;-0.0;0.0',
        "Mirrors the multiples model (target multiple = 10y average + k x sd): ERP moves by -k x the sd of the 10-year implied ERP (Ibovespa; each box with basis 3)."),
       ("gadd", f"Earnings growth vs consensus path, years 2-{NY} (pp per year)", GADD, PPF,
        f"Added to the earnings growth of every box in years 2-{NY} (consensus in year 2, then the path to the long term); bear -2 / bull +2 pp (proposal). The terminal grows at g."),
       ("g", "Long-term nominal growth (g)", GLT, PCT2,
        "Base = house perpetuity growth 4.7% (Target!B54), nominal. Bear / bull -/+ 0.5 pp (proposal). Earnings growth reaches it by GrowthYear."),
       ("roe", "Long-term ROE vs the box's 10-year average (pp)", [0, 0, 0, 0], PPF,
        "Shifts every box's long-term ROE (10-year average of its 12m fwd ROE, History).")]
for k, lab, vals, fmt, note in lev:
    rr_ = A[k]
    ws.write(f"B{rr_}", lab, bk.TXT)
    for c, v in zip("DEFG", vals):
        if isinstance(v, str) and v.startswith("="):
            ws.write_formula(f"{c}{rr_}", v, F(**LV, num_format=fmt))
        elif isinstance(v, str):
            ws.write_string(f"{c}{rr_}", v, F(**LV, align="center"))
        else:
            ws.write_number(f"{c}{rr_}", v, F(**LV, num_format=fmt))
    ws.write(f"H{rr_}", note, bk.NOTE)
ws.write_formula(f"C{A['selic']}", "=SelicNow", F(num_format=PCT2))
ws.write_formula(f"C{A['rr']}", "=RR5Now", F(num_format=PCT2))
ws.write(f"B{A['cut']}", "Selic cut vs today (pp; positive = cut)", bk.TXT)
ws.write(f"B{A['erpu']}", "ERP used, Ibovespa level", bk.BOLD)
ws.write(f"B{A['ke']}", "Cost of equity, Ibovespa level = real rate + ERP", bk.BOLD)
for c in "DEFG":
    ws.write_formula(f"{c}{A['cut']}", f"=(SelicNow-{c}{A['selic']})*100", F(num_format=PPC))
    ws.write_formula(f"{c}{A['erpu']}", f"=CHOOSE(ERPBasis,{c}{A['erp']},ErpHistIbov,ErpHistIbov)+{c}{A['erps']}/100-{c}{A['erpk']}*ErpSdIbov", F(num_format=PCT2, bold=True))
    ws.write_formula(f"{c}{A['ke']}", f"={c}{A['rr']}+{c}{A['erpu']}", F(num_format=PCT2, bold=True))
ws.write(f"H{A['cut']}", "Formula: (Selic today - Selic in the scenario) x 100.", bk.NOTE)
ws.write(f"H{A['erpu']}", "Basis ERP (typed, or the Ibovespa's 10-year average) + shift - k x sd. With basis 3 each box uses its own average and sd (DCF sheets).", bk.NOTE)
ws.write(f"H{A['ke']}", "RateConv 1 (house): the discount rate of the nominal cash flows, as Target!B55. RateConv 2 (Fisher): turned nominal year by year with the IPCA.", bk.NOTE)
r += 2
bk.section(ws, r, "4. Settings")
r += 1
bk.header_row(ws, r, [("B", "Setting"), ("C", "Value"), ("D", ""), ("E", ""), ("F", ""), ("G", ""), ("H", "Notes")])
for lab, nm, v, fmt, note in [
        ("Discount-rate convention: 1 = house (real rate + ERP), 2 = Fisher (with IPCA)", "RateConv", 1, "0",
         "1: real rate + ERP discounts the nominal cash flows, as the house model (Target!B55 = BZRFB5PY + ERP). "
         "2: (1 + real rate + ERP) x (1 + IPCA) - 1. Same ERP, about 4 pp more discount rate."),
        ("ERP basis: 1 = typed, 2 = Ibovespa 10y average, 3 = each box's 10y average", "ERPBasis", 1, "0",
         "1 = the house's 6.0%. 2 and 3: implied ERP measured each month in the same convention (History): Ke = g + (1 - g / ROE) / P/E, minus the 5y real rate."),
        ("Valuation horizon: 0 = today, 1 = 12 months (end of forward year 1)", "Horizon", HORIZON, "0",
         "1: fair value in 12 months = value today x (1 + year-1 discount rate) - year-1 FCFE (paid out during the year), as a 12-month target price. 0: value today."),
        ("Long-term ROE: 1 = box's 10-year average, 2 = consensus year 2 (no mean reversion)", "RoeLT", ROELT, "0",
         "1: ROE converges to the box's 10-year average 12m fwd ROE (History). 2: ROE stays at the consensus year-2 level, as the house model keeps consensus profitability."),
        ("Long-term nominal growth behind the historical implied ERP", "HistG", 0.047, PCT2, "g used to read the ERP out of each month's P/E and ROE (History); house 4.7%."),
        ("ROE and ERP history window: first month", "WinStart", pd.Timestamp("2016-10-31"), DATE, "Averages over the window (any day counts the whole month)."),
        ("ROE and ERP history window: last month", "WinEnd", pd.Timestamp("2026-09-30"), DATE, ""),
        ("Year by which growth reaches the long term", "GrowthYear", NY, "0", f"Growth converges linearly from the consensus year-2 growth (years 3 to GrowthYear). 3-{NY}."),
        ("Year by which ROE reaches the long term", "RoeYear", NY, "0", f"ROE converges linearly from the consensus year-2 ROE (years 3 to RoeYear). 3-{NY}."),
        ("Apply the rate-driven change in earnings (1 = yes, 0 = no)", "UseEPS", 1, "0", "1: years 1-2 earnings move with the Selic as in the multiples model (section 5). 0: consensus as is."),
        ("P/E: lowest valid value", "PEmin", 1.0, MULT, "Members with any FY26-FY28 P/E or ROE outside the ranges are left out and the box is scaled to 100% of its points."),
        ("P/E: highest valid value", "PEmax", 100.0, MULT, ""),
        ("ROE: lowest valid value", "ROEmin", 0.01, PCT, ""),
        ("ROE: highest valid value", "ROEmax", 1.0, PCT, "")]:
    r += 1
    one(r, lab, nm, v, fmt, note, lever=True)
ws.data_validation(f"C{A['ERPBasis']}", {"validate": "list", "source": [1, 2, 3]})
ws.data_validation(f"C{A['RateConv']}", {"validate": "list", "source": [1, 2]})
ws.data_validation(f"C{A['UseEPS']}", {"validate": "list", "source": [0, 1]})
ws.data_validation(f"C{A['Horizon']}", {"validate": "list", "source": [0, 1]})
ws.data_validation(f"C{A['RoeLT']}", {"validate": "list", "source": [1, 2]})
for nm in ("GrowthYear", "RoeYear"):
    ws.data_validation(f"C{A[nm]}", {"validate": "integer", "criteria": "between", "minimum": 3, "maximum": NY})
r += 1
ws.write(f"B{r}", "Explicit forecast years", bk.TXT)
ws.write_number(f"C{r}", NY, F(num_format="0", align="right"))
ws.write(f"H{r}", f"Fixed by the layout of the DCF sheets (years 1-{NY} + terminal), like the house model (4 years after its value date + terminal). NY in build_dcf.py.", bk.NOTE)
r += 2
bk.section(ws, r, "5. Inputs from the multiples model (Ibov_valuation_by_box_model_XP.xlsx)")
r += 1
bk.header_row(ws, r, [("B", "Box"), ("C", "EPS change per -100bp of Selic"), ("D", ""), ("E", ""), ("F", ""), ("G", ""), ("H", "Source")])
SENS = {}
for i, g in enumerate(G6):
    r += 1
    SENS[g] = r
    ws.write(f"B{r}", ("   " if g in MEMO else "") + g, bk.TXT)
    ws.write_number(f"C{r}", REF["sens"][g] / 100, F(**IN, num_format=PCT2))
    if i == 0:
        ws.write(f"H{r}", "Rates & EPS sheet of the multiples model: Ibovespa +2.07% per -100bp, split by box return beta to the 2y pre (Oct-6).", bk.NOTE)
r += 2
bk.section(ws, r, "6. Optional overrides (leave blank to use the model rule)")
r += 1
ws.write(f"B{r}", "A number typed here replaces the model value for that box and scenario. Checks counts the overrides in use.", bk.NOTE)
r += 1
bk.header_row(ws, r, [("B", "Box"), ("C", ""), ("D", "Bear"), ("E", "Base"), ("F", "Bull"), ("G", "Custom"), ("H", "")])
OVR = {}
ovr_first = r + 2
for lab, key, fmt in [("Long-term ROE (%)", "roe", PCT), ("Long-term nominal growth (%)", "g", PCT2), ("ERP (%), replaces the ERP basis", "erp", PCT2)]:
    r += 1
    ws.write(f"B{r}", lab, bk.SUB)
    for g in G6:
        r += 1
        OVR[(key, g)] = r
        ws.write(f"B{r}", ("   " if g in MEMO else "") + g, bk.TXT)
        for c in "DEFG":
            ws.write_blank(f"{c}{r}", None, F(**LV, num_format=fmt))
A["ovr"] = (ovr_first, r)
r += 2
bk.section(ws, r, "7. Classification")
r += 1
ws.write(f"B{r}", "XP-adjusted GICS sector", bk.HDRL)
ws.write(f"C{r}", "Box", bk.HDR)
m0 = r + 1
for gs, bx in GICS_MAP:
    r += 1
    ws.write(f"B{r}", gs, F(**IN))
    ws.write(f"C{r}", bx, F(**IN))
bk.name("GicsMap", "Assumptions", f"$B${m0}:$C${r}")
r += 2
ws.write(f"B{r}", "Global industrials (shown apart as a memo)", bk.HDRL)
g0 = r + 1
for t in ["WEGE3", "EMBJ3"]:
    r += 1
    ws.write(f"B{r}", t, F(**IN))
for _ in range(2):
    r += 1
    ws.write_blank(f"B{r}", None, F(**LV))
bk.name("GlobalList", "Assumptions", f"$B${g0}:$B${r}")
bk.name("ErpHistIbov", "History", f"${HCOL['erp']['Ibovespa']}${HST['avg']}")
bk.name("ErpSdIbov", "History", f"${HCOL['erp']['Ibovespa']}${HST['sd']}")
ws.freeze_panes(4, 0)

# =========================================================================== Members
ws = W["Members"]
bk.title(ws, "Members: consensus earnings and book value by stock", "Inputs in blue (Bloomberg, Economatica); earnings and book value in Ibovespa points are formulas.")
MF, ML = 6, 6 + len(m) - 1
MT = ML + 2
mcols = [("A", "Ticker", 8), ("B", "Company", 18), ("C", "XP-adjusted GICS sector", 21), ("D", "Box", 12), ("E", "Global industrial (1/0)", 9),
         ("F", "Sub-box", 24), ("G", "Ibovespa weight on the membership date", 11), ("H", "Price on the membership date (R$)", 11),
         ("I", "Price on the pricing date (R$)", 11), ("J", "Price-drifted weight (unscaled)", 11), ("K", "Weight on the pricing date", 10),
         ("L", "Index points", 10), ("M", "Price on the consensus date (R$)", 11), ("N", "P/E FY26 (x)", 8), ("O", "P/E FY27 (x)", 8),
         ("P", "P/E FY28 (x)", 8), ("Q", "ROE FY26", 8), ("R", "ROE FY27", 8), ("S", "ROE FY28", 8), ("T", "Valid data (1/0)", 8),
         ("U", "Index points with valid data", 10), ("V", "Earnings FY26 (index pts)", 10), ("W", "Earnings FY27 (index pts)", 10),
         ("X", "Earnings FY28 (index pts)", 10), ("Y", "Book value FY26 (index pts)", 10), ("Z", "Book value FY27 (index pts)", 10),
         ("AA", "Book value FY28 (index pts)", 10), ("AB", "Data flags", 34)]
for a_, b_, t_ in [("A", "C", "Identification"), ("D", "F", "Classification"), ("G", "L", "Weight and index points"),
                   ("M", "T", "Consensus (Bloomberg, bst)"), ("U", "AA", "In Ibovespa points"), ("AB", "AB", "")]:
    if a_ == b_:
        ws.write(f"{a_}4", t_, bk.GROUPH)
    else:
        ws.merge_range(f"{a_}4:{b_}4", t_, bk.GROUPH)
ws.set_row(4, 54)
for c, h, wdt in mcols:
    ws.write(f"{c}5", h, bk.HDR)
    ws.set_column(f"{c}:{c}", wdt)
for i, (t, x) in enumerate(m.iterrows()):
    r = MF + i
    y = dm.loc[t]
    ws.write_string(f"A{r}", t, F(**IN, bold=True))
    ws.write_string(f"B{r}", x["name"], F(**IN))
    ws.write_string(f"C{r}", x["gics"], F(**IN))
    ws.write_formula(f"D{r}", f"=VLOOKUP(C{r},GicsMap,2,FALSE)", bk.TXT)
    ws.write_formula(f"E{r}", f"=IF(COUNTIF(GlobalList,A{r})>0,1,0)", F(align="center"))
    ws.write_formula(f"F{r}", f'=IF(D{r}="Cyclicals",IF(E{r}=1,"WEG & Embraer","Cyclicals ex-WEG & Embraer"),D{r})', bk.TXT)
    ws.write_number(f"G{r}", x["w30"], F(**IN, num_format="0.00%"))
    ws.write_number(f"H{r}", x["px30"], F(**IN, num_format=PX))
    ws.write_number(f"I{r}", x["px_now"], F(**IN, num_format=PX))
    ws.write_formula(f"J{r}", f"=G{r}*I{r}/H{r}", F(num_format="0.000%"))
    ws.write_formula(f"K{r}", f"=J{r}/SUM($J${MF}:$J${ML})", F(num_format="0.00%"))
    ws.write_formula(f"L{r}", f"=K{r}*IbovNow", F(num_format=PTS))
    if nz(y["px_today"]) is not None:
        ws.write_number(f"M{r}", y["px_today"], F(**IN, num_format=PX))
    for c, k in zip("NOP", ["pe26", "pe27", "pe28"]):
        if nz(y[k]) is not None:
            ws.write_number(f"{c}{r}", y[k], F(**IN, num_format=MULT))
    for c, k in zip("QRS", ["roe26", "roe27", "roe28"]):
        if nz(y[k]) is not None:
            ws.write_number(f"{c}{r}", y[k] / 100, F(**IN, num_format=PCT))
    ws.write_formula(f"T{r}", f"=IF(AND(COUNT(M{r}:S{r})=7,MIN(N{r}:P{r})>=PEmin,MAX(N{r}:P{r})<=PEmax,MIN(Q{r}:S{r})>=ROEmin,MAX(Q{r}:S{r})<=ROEmax),1,0)", F(align="center"))
    ws.write_formula(f"U{r}", f"=T{r}*L{r}", F(num_format=PTS))
    for c, pc in zip("VWX", "NOP"):
        ws.write_formula(f"{c}{r}", f"=IF(T{r}=1,L{r}*(M{r}/{pc}{r})/I{r},0)", F(num_format=PTS1))
    for c, ec, rc in zip(["Y", "Z", "AA"], "VWX", "QRS"):
        ws.write_formula(f"{c}{r}", f"=IF(T{r}=1,{ec}{r}/{rc}{r},0)", F(num_format=PTS1))
    ws.write_formula(f"AB{r}", f'=IF(T{r}=1,"","Left out: P/E or ROE missing or out of range")', F(font_color=GREY))
ws.write(f"A{MT}", "Total", bk.BOLD)
for c, f in (("G", "0.00%"), ("K", "0.00%"), ("L", PTS), ("U", PTS), ("V", PTS1), ("W", PTS1), ("X", PTS1), ("Y", PTS1), ("Z", PTS1), ("AA", PTS1)):
    ws.write_formula(f"{c}{MT}", f"=SUM({c}{MF}:{c}{ML})", F(bold=True, num_format=f, top=1))
for i, t in enumerate([
        "Earnings (index pts) = index points x (price on the consensus date / P/E) / price on the pricing date: the stock's earnings yield on the pricing-date price, in points. P/E-based, so the currency of reporting does not matter.",
        "Book value (index pts) = earnings / ROE of the same fiscal year (Bloomberg BEST_ROE, end-of-year equity)."]):
    ws.write(f"A{MT + 2 + i}", t, bk.NOTE)
ws.freeze_panes(5, 2)
ws.autofilter(f"A5:AB{ML}")
mr = lambda c: f"Members!${c}${MF}:${c}${ML}"

# =========================================================================== Boxes Now
ws = W["Boxes Now"]
bk.title(ws, "Boxes now: consensus earnings, book value and ROE by box", "SUMIFS on Members; boxes scaled from the members with valid data to 100% of their index points.")
BN = {"Financials": 6, "Defensives": 7, "Cyclicals": 8, "Commodities": 9, "Ibovespa": 10, "Cyclicals ex-WEG & Embraer": 12, "WEG & Embraer": 13}
bcols = [("B", "Group", 26), ("C", "Box criterion", 12), ("D", "Sub-box criterion", 24), ("E", "Members", 8), ("F", "Index points", 10),
         ("G", "Points with valid data", 10), ("H", "Coverage", 9), ("I", "Earnings FY26 (pts)", 10), ("J", "Earnings FY27 (pts)", 10),
         ("K", "Earnings FY28 (pts)", 10), ("L", "Book value FY26 (pts)", 10), ("M", "Book value FY27 (pts)", 10), ("N", "Book value FY28 (pts)", 10),
         ("O", "ROE FY26", 8), ("P", "ROE FY27", 8), ("Q", "ROE FY28", 8), ("R", "Earnings, forward year 1 (pts)", 11),
         ("S", "Earnings, forward year 2 (pts)", 11), ("T", "Book value, year 1 (pts)", 10), ("U", "Book value, year 2 (pts)", 10),
         ("V", "ROE, year 1", 8), ("W", "ROE, year 2", 8), ("X", "Earnings growth, year 2", 9), ("Y", "P/E on year-1 earnings (x)", 9),
         ("Z", "Long-term ROE (10y average)", 10), ("AA", "Implied ERP, 10y average", 10)]
ws.set_column("A:A", 2)
ws.set_row(4, 54)
for c, h, wdt in bcols:
    ws.write(f"{c}5", h, bk.HDRL if c == "B" else bk.HDR)
    ws.set_column(f"{c}:{c}", wdt)
ws.write("B11", "Memo (not added to the Ibovespa)", bk.NOTE)


def sif(col, r):
    return f"SUMIFS({mr(col)},{mr('D')},$C{r},{mr('F')},$D{r})"


for g, r in BN.items():
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", g, F(**b))
    ws.write_string(f"C{r}", CRIT[g][0], F(**IN, **b))
    ws.write_string(f"D{r}", CRIT[g][1], F(**IN, **b))
    if tot:
        for c, f in (("E", "0"), ("F", PTS), ("G", PTS), ("I", PTS1), ("J", PTS1), ("K", PTS1), ("L", PTS1), ("M", PTS1), ("N", PTS1)):
            ws.write_formula(f"{c}{r}", f"=SUM({c}6:{c}9)", F(num_format=f, **b))
    else:
        ws.write_formula(f"E{r}", f"=COUNTIFS({mr('D')},$C{r},{mr('F')},$D{r})", F(num_format="0", **b))
        ws.write_formula(f"F{r}", "=" + sif("L", r), F(num_format=PTS, **b))
        ws.write_formula(f"G{r}", "=" + sif("U", r), F(num_format=PTS, **b))
        for c, mc in zip("IJKLMN", ["V", "W", "X", "Y", "Z", "AA"]):
            ws.write_formula(f"{c}{r}", f"={sif(mc, r)}/$G{r}*$F{r}", F(num_format=PTS1, **b))
    ws.write_formula(f"H{r}", f"=G{r}/F{r}", F(num_format=PCT, **b))
    for c, e, bv in zip("OPQ", "IJK", "LMN"):
        ws.write_formula(f"{c}{r}", f"={e}{r}/{bv}{r}", F(num_format=PCT, **b))
    ws.write_formula(f"R{r}", f"=FYleft*I{r}+(1-FYleft)*J{r}", F(num_format=PTS1, **b))
    ws.write_formula(f"S{r}", f"=FYleft*J{r}+(1-FYleft)*K{r}", F(num_format=PTS1, **b))
    ws.write_formula(f"T{r}", f"=FYleft*L{r}+(1-FYleft)*M{r}", F(num_format=PTS1, **b))
    ws.write_formula(f"U{r}", f"=FYleft*M{r}+(1-FYleft)*N{r}", F(num_format=PTS1, **b))
    ws.write_formula(f"V{r}", f"=R{r}/T{r}", F(num_format=PCT, **b))
    ws.write_formula(f"W{r}", f"=S{r}/U{r}", F(num_format=PCT, **b))
    ws.write_formula(f"X{r}", f"=S{r}/R{r}-1", F(num_format=UPS, **b))
    ws.write_formula(f"Y{r}", f"=F{r}/R{r}", F(num_format=MULT, **b))
    ws.write_formula(f"Z{r}", f"=History!{HCOL['roe'][g]}{HST['avg']}", F(num_format=PCT, **b))
    ws.write_formula(f"AA{r}", f"=History!{HCOL['erp'][g]}{HST['avg']}", F(num_format=PCT2, **b))
for i, t in enumerate([
        "Forward year 1 = share of FY26 still ahead x FY26 + the rest x FY27 (Assumptions, FYleft); year 2 = the same blend of FY27 and FY28.",
        "ROE of a box = its earnings / its book value (both in index points), i.e. the aggregate ROE, not an average of the members' ROEs.",
        "The Ibovespa row adds up the four boxes (points, earnings, book value) and derives its ratios from the sums."]):
    ws.write(f"B{16 + i}", t, bk.NOTE)
ws.freeze_panes(5, 2)

# =========================================================================== ROE Panel
ws = W["ROE Panel"]
ws.hide_gridlines(0)
bk.title(ws, "ROE panel: monthly 12m fwd P/E and ROE of the members", "One row per member and month (Oct-16..Sep-26); Bloomberg BEST_PE_RATIO and BEST_ROE (BF).", legend=False)
pcols = [("A", "Month-end", 9), ("B", "Ticker", 8), ("C", "XP-adjusted GICS sector", 22), ("D", "Box", 12), ("E", "Sub-box", 24),
         ("F", "Ibovespa weight", 9), ("G", "12m fwd P/E (x)", 8), ("H", "12m fwd ROE", 8), ("I", "Valid (1/0)", 7),
         ("J", "Earnings per unit of weight (w / P/E)", 11), ("K", "Book value per unit of weight (earnings / ROE)", 12), ("L", "Weight if valid", 9)]
ws.set_row(3, 54)
for c, h, wdt in pcols:
    ws.write(f"{c}4", h, bk.HDR)
    ws.set_column(f"{c}:{c}", wdt)
PF = 5
PL = PF + len(hist) - 1
fd, fb, fw = F(**IN, num_format=DATE), F(**IN, num_format="0.00"), F(**IN, num_format="0.000%")
fpc, f0, f5, fwt = F(**IN, num_format=PCT), F(align="center"), F(num_format="0.00000"), F(num_format="0.000%")
for i, x in enumerate(hist.itertuples(index=False)):
    r = PF + i
    ws.write_datetime(r - 1, 0, pd.Timestamp(x.date).to_pydatetime(), fd)
    ws.write_string(r - 1, 1, x.cod_ativo, F(**IN))
    ws.write_string(r - 1, 2, x.gics, F(**IN))
    ws.write_formula(r - 1, 3, f"=VLOOKUP(C{r},GicsMap,2,FALSE)", bk.TXT)
    ws.write_formula(r - 1, 4, f'=IF(D{r}="Cyclicals",IF(COUNTIF(GlobalList,B{r})>0,"WEG & Embraer","Cyclicals ex-WEG & Embraer"),D{r})', bk.TXT)
    ws.write_number(r - 1, 5, x.w, fw)
    if not np.isnan(x.pe):
        ws.write_number(r - 1, 6, x.pe, fb)
    if not np.isnan(x.roe):
        ws.write_number(r - 1, 7, x.roe / 100, fpc)
    ws.write_formula(r - 1, 8, f"=IF(AND(ISNUMBER(G{r}),ISNUMBER(H{r}),G{r}>=PEmin,G{r}<=PEmax,H{r}>=ROEmin,H{r}<=ROEmax),1,0)", f0)
    ws.write_formula(r - 1, 9, f"=IF(I{r}=1,F{r}/G{r},0)", f5)
    ws.write_formula(r - 1, 10, f"=IF(I{r}=1,J{r}/H{r},0)", f5)
    ws.write_formula(r - 1, 11, f"=I{r}*F{r}", fwt)
ws.freeze_panes(4, 2)
ws.autofilter(f"A4:L{PL}")
PR = lambda c: f"'ROE Panel'!${c}${PF}:${c}${PL}"

def erp_ss(roe, pe, rr):
    """Steady-state implied ERP (Excel expression): Ke = g + (1 - g / ROE) / P/E, net of the real rate, in the RateConv convention."""
    ke = f"(HistGn+(1-HistGn/{roe})/{pe})"
    return f"IF(RateConv=1,{ke},(1+{ke})/(1+IPCALT)-1)-{rr}"


# =========================================================================== History
ws = W["History"]
bk.title(ws, "History: ROE, P/E and implied ERP of each box (monthly, 12m fwd)",
         "Box ROE = sum of earnings / sum of book value; P/E = sum of weights / sum of earnings (ROE Panel). Implied ERP: the ERP that makes a steady-state DCF match the month's P/E.")
ws.set_column("A:A", 12)
ws.set_column("B:B", 9)
ws.set_column("C:C", 10)
ws.set_column("D:X", 11)
for k, a_, b_, lab in [("roe", "D", "J", "12m fwd ROE"), ("pe", "K", "Q", "12m fwd P/E (x), members with ROE"),
                       ("erp", "R", "X", "Implied ERP = Ke - 5y real rate, Ke = g + (1 - g / ROE) / P/E (made real with the IPCA if RateConv = 2)")]:
    ws.merge_range(f"{a_}4:{b_}4", lab, bk.GROUPH)
ws.set_row(3, 30)
ws.write("A5", "Group", bk.HDRL)
ws.write("A6", "Box criterion", bk.NOTE)
ws.write("A7", "Sub-box criterion", bk.NOTE)
for k, lab in (("avg", "Average (window)"), ("sd", "Std. dev. (window)"), ("n", "Months in window"), ("now", "Now (pricing date)"), ("z", "Now vs average (std. dev.)")):
    ws.write(f"A{HST[k]}", lab, bk.BOLD if k in ("avg", "now") else bk.TXT)
ws.write("A14", "Long-term nominal growth used for the implied ERP (Assumptions, HistG)", bk.NOTE)
ws.write_formula("C14", "=HistG", F(num_format=PCT2, bold=True))
bk.name("HistGn", "History", "$C$14")
ws.write(f"A{HF - 1}", "Month-end", bk.HDR)
ws.write(f"B{HF - 1}", "In window (1/0)", bk.HDR)
ws.write(f"C{HF - 1}", "5y real rate (NTN-B)", bk.HDR)
ws.set_row(HF - 2, 30)
for kind in ("roe", "pe", "erp"):
    for g in HG:
        c = HCOL[kind][g]
        ws.write(f"{c}5", g, bk.HDR)
        ws.write(f"{c}{HF - 1}", g, bk.HDR)
        if kind != "erp":
            ws.write_string(f"{c}6", CRIT[g][0], F(**IN, align="center"))
            ws.write_string(f"{c}7", CRIT[g][1], F(**IN, align="center", font_size=8))
        fmt = {"roe": PCT, "pe": MULT, "erp": PCT2}[kind]
        rng_ = f"{c}${HF}:{c}${HL}"
        ws.write_formula(f"{c}{HST['avg']}", f"=AVERAGEIFS({rng_},$B${HF}:$B${HL},1)", F(num_format=fmt, bold=True))
        ws.write_formula(f"{c}{HST['n']}", f"=SUMPRODUCT($B${HF}:$B${HL},--ISNUMBER({rng_}))", F(num_format="0"))
        ws.write_formula(f"{c}{HST['sd']}", f"=SQRT((SUMPRODUCT($B${HF}:$B${HL},{rng_},{rng_})-{c}{HST['n']}*{c}{HST['avg']}^2)/({c}{HST['n']}-1))",
                         F(num_format=PCT2 if kind != "pe" else "0.00"))
        if kind == "roe":
            now = f"='Boxes Now'!V{BN[g]}"
        elif kind == "pe":
            now = f"='Boxes Now'!Y{BN[g]}"
        else:
            now = "=" + erp_ss(f"{HCOL['roe'][g]}{HST['now']}", f"{HCOL['pe'][g]}{HST['now']}", "RR5Now")
        ws.write_formula(f"{c}{HST['now']}", now, F(num_format=fmt, bold=True))
        ws.write_formula(f"{c}{HST['z']}", f"=({c}{HST['now']}-{c}{HST['avg']})/{c}{HST['sd']}", F(num_format="+0.00;-0.00"))
for i, dt in enumerate(months):
    r = HF + i
    ws.write_datetime(f"A{r}", pd.Timestamp(dt).to_pydatetime(), F(**IN, num_format=DATE))
    ws.write_formula(f"B{r}", f"=IF(AND(A{r}>EOMONTH(WinStart,-1),A{r}<=EOMONTH(WinEnd,0)),1,0)", F(align="center"))
    ws.write_number(f"C{r}", float(rr5[pd.Timestamp(dt)]), F(**IN, num_format=PCT2))
    for g in HG:
        cr, cp, ce = HCOL["roe"][g], HCOL["pe"][g], HCOL["erp"][g]
        crit = f"{PR('A')},$A{r},{PR('D')},{cr}$6,{PR('E')},{cr}$7"
        ws.write_formula(f"{cr}{r}", f'=IFERROR(SUMIFS({PR("J")},{crit})/SUMIFS({PR("K")},{crit}),"")', F(num_format=PCT))
        ws.write_formula(f"{cp}{r}", f'=IFERROR(SUMIFS({PR("L")},{crit})/SUMIFS({PR("J")},{crit}),"")', F(num_format=MULT))
        ws.write_formula(f"{ce}{r}", f'=IF(AND(ISNUMBER({cr}{r}),ISNUMBER({cp}{r})),{erp_ss(f"{cr}{r}", f"{cp}{r}", f"$C{r}")},"")', F(num_format=PCT2))
ws.write(f"A{HL + 2}", "Steady-state reading: P/E = (1 - g / ROE) / (Ke - g), solved for Ke, net of the 5y real rate (with RateConv = 2, Ke is first made real with the IPCA). "
                       "It is the ERP a DCF with the month's ROE and growth g would need to match the month's P/E, in the convention of the DCF sheets.", bk.NOTE)
ws.write(f"A{HL + 3}", "5y real rate: NTN-B constant maturity (Bloomberg BZRFB5PY to Aug-21, Tesouro Direto after), as in the multiples model.", bk.NOTE)
ws.freeze_panes(HF - 1, 1)

# =========================================================================== DCF scenario sheets
TL = {"t": 19, "date": 20, "pi": 21, "cpi": 22, "defl": 23}
P_ = {"rr": 6, "erp": 7, "erps": 8, "erpk": 9, "basis": 10, "gadd": 11, "g": 12, "roe": 13, "cut": 14, "conv": 15}
IBOV0, BOX0 = 25, 47                         # first row of the Ibovespa block and of the first box block
BLK = {}
BKEYS = ["pts", "epsadj", "erp", "ke_r", "ke_n", "df", "roelt", "glt", "g", "e", "roe", "bv", "dbv", "fcfe", "payout", "pv", "tv", "pvtv",
         "val0", "val", "up", "pe_impl", "pe_now", "tvshare", "rfcfe"]


def years_header(ws, r):
    bk.header_row(ws, r, [("B", "")] + [("C", "Today")] + [(YC[t], f"Year {t}") for t in YEARS] + [(TCOL, "Terminal"), (NC, "How it is calculated")])


def dcf_sheet(s):
    ws = W[f"DCF {s}"]
    col = SC[s]
    bk.title(ws, f"DCF, {s} case", f"FCFE by box in Ibovespa points. Scenario levers: Assumptions column {col}. Years 1-{NY} + terminal; Ibovespa = sum of the four boxes.")
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 64)
    ws.set_column("C:C", 12)
    ws.set_column(f"D:{YC[NY]}", 11)
    ws.set_column(f"{TCOL}:{TCOL}", 12)
    ws.set_column(f"{NC}:{NC}", 110)
    bk.section(ws, 5, f"Scenario parameters ({s}; Assumptions column {col})", "B", NC)
    par = [("rr", "5-year real rate (NTN-B)", f"=Assumptions!{col}{A['rr']}", PCT2),
           ("erp", "ERP, typed (ERP basis 1)", f"=Assumptions!{col}{A['erp']}", PCT2),
           ("erps", "ERP shift (pp)", f"=Assumptions!{col}{A['erps']}", PPF),
           ("erpk", "ERP move in std. dev. of its history (k)", f"=Assumptions!{col}{A['erpk']}", '+0.0;-0.0;0.0'),
           ("basis", "ERP basis (1 typed, 2 Ibovespa 10y average, 3 each box's 10y average)", "=ERPBasis", "0"),
           ("gadd", f"Earnings growth vs consensus path, years 2-{NY} (pp per year)", f"=Assumptions!{col}{A['gadd']}", PPF),
           ("g", "Long-term nominal growth (default)", f"=Assumptions!{col}{A['g']}", PCT2),
           ("roe", "Long-term ROE vs 10y average (pp)", f"=Assumptions!{col}{A['roe']}", PPF),
           ("cut", "Selic cut vs today (pp)", f"=Assumptions!{col}{A['cut']}", PPC),
           ("conv", "Discount-rate convention (1 house: real rate + ERP; 2 Fisher: with IPCA)", "=RateConv", "0")]
    for k, lab, f_, fmt in par:
        ws.write(f"B{P_[k]}", lab, bk.TXT)
        ws.write_formula(f"C{P_[k]}", f_, F(num_format=fmt))
    bk.section(ws, TL["t"] - 2, "Timeline (shared by every box)", "B", NC)
    years_header(ws, TL["t"] - 1)
    ws.write(f"B{TL['t']}", "Year (t)", bk.TXT)
    ws.write_number(f"C{TL['t']}", 0, F(num_format="0", align="right"))
    for t in YEARS:
        ws.write_number(f"{YC[t]}{TL['t']}", t, F(num_format="0", align="right"))
    ws.write(f"B{TL['date']}", "End of the forward year", bk.TXT)
    ws.write_formula(f"C{TL['date']}", "=PxDate", F(num_format=DATE, align="right"))
    for t in YEARS:
        ws.write_formula(f"{YC[t]}{TL['date']}", f"=EDATE(PxDate,12*{YC[t]}{TL['t']})", F(num_format=DATE, align="right"))
    ws.write(f"B{TL['pi']}", "Inflation (IPCA)", bk.TXT)
    for t in YEARS:
        ws.write_formula(f"{YC[t]}{TL['pi']}", "=IPCA1" if t == 1 else ("=IPCA2" if t == 2 else "=IPCALT"), F(num_format=PCT))
    ws.write_formula(f"{TCOL}{TL['pi']}", "=IPCALT", F(num_format=PCT))
    ws.write(f"B{TL['cpi']}", "Price index (today = 1)", bk.TXT)
    ws.write_number(f"C{TL['cpi']}", 1, F(num_format="0.000"))
    for t in YEARS:
        p = "C" if t == 1 else YC[t - 1]
        ws.write_formula(f"{YC[t]}{TL['cpi']}", f"={p}{TL['cpi']}*(1+{YC[t]}{TL['pi']})", F(num_format="0.000"))
    ws.write(f"B{TL['defl']}", "Deflator for discounting (1 under the house convention)", bk.TXT)
    ws.write_number(f"C{TL['defl']}", 1, F(num_format="0.000"))
    for t in YEARS:
        ws.write_formula(f"{YC[t]}{TL['defl']}", f"=IF(RateConv=1,1,{YC[t]}{TL['cpi']})", F(num_format="0.000"))
    ws.write(f"{NC}{TL['cpi']}", "Cash flows are nominal (consensus). House convention: discounted at real rate + ERP. Fisher: at (1 + real rate + ERP) x (1 + IPCA) - 1.", bk.NOTE)
    ws.write(f"{NC}{TL['defl']}", "Used by the closed-form checks: FCFE / deflator discounted at (1 + real rate + ERP) gives the same value under either convention.", bk.NOTE)
    r0 = BOX0
    for g in G6:
        r0 = box_block(ws, s, g, r0)
    ibov_block(ws, s, IBOV0)
    ws.freeze_panes(4, 2)


def box_block(ws, s, g, r0):
    R = {k: r0 + 2 + i for i, k in enumerate(BKEYS)}
    memo = g in MEMO
    bk.section(ws, r0, g + (" (memo, not added to the Ibovespa)" if memo else ""), "B", NC, fill=YEL, color=INK)
    years_header(ws, r0 + 1)
    bn, col = BN[g], SC[s]
    lab = {"pts": "Index points", "epsadj": "Earnings change from the Selic scenario (years 1-2)", "erp": "Equity risk premium (ERP)",
           "ke_r": "Cost of equity = real rate + ERP", "ke_n": "Discount rate of the nominal cash flows", "df": "Discount factor",
           "roelt": "Long-term ROE", "glt": "Long-term nominal growth", "g": "Earnings growth", "e": "Earnings (index pts)", "roe": "ROE",
           "bv": "Book value at the start of the year (index pts)", "dbv": "Increase in book value during the year",
           "fcfe": "Free cash flow to equity (FCFE)", "payout": "   FCFE / earnings", "pv": "Present value of FCFE",
           "tv": f"Terminal value at the end of year {NY}", "pvtv": "Present value of the terminal value", "val0": "Equity value today (index pts)",
           "val": "Fair value at the horizon (index pts)", "up": "Upside (fair value / index points today - 1)",
           "pe_impl": "Implied 12m fwd P/E at the horizon (fair value / next 12m earnings)", "pe_now": "12m fwd P/E today (consensus year-1 earnings)",
           "tvshare": "Terminal value / equity value", "rfcfe": "   FCFE / deflator (for Implied ERP)"}
    for k in BKEYS:
        bold = k in ("fcfe", "val", "up", "ke_r")
        ws.write(f"B{R[k]}", lab[k], F(bold=bold, font_color=GREY) if k in ("payout", "rfcfe") else F(bold=bold))
    fp, fp1 = F(num_format=PTS), F(num_format=PTS1)
    ws.write_formula(f"C{R['pts']}", f"='Boxes Now'!F{bn}", fp)
    ws.write_formula(f"C{R['epsadj']}", f"=IF(UseEPS=1,Assumptions!$C${SENS[g]}*$C${P_['cut']},0)", F(num_format=UPS))
    o_erp = f"Assumptions!{col}{OVR[('erp', g)]}"
    ws.write_formula(f"C{R['erp']}", f"=IF(ISNUMBER({o_erp}),{o_erp},CHOOSE($C${P_['basis']},$C${P_['erp']},ErpHistIbov,History!${HCOL['erp'][g]}${HST['avg']})"
                                     f"+$C${P_['erps']}/100-$C${P_['erpk']}*CHOOSE($C${P_['basis']},ErpSdIbov,ErpSdIbov,History!${HCOL['erp'][g]}${HST['sd']}))", F(num_format=PCT2))
    ws.write_formula(f"C{R['ke_r']}", f"=$C${P_['rr']}+C{R['erp']}", F(num_format=PCT2, bold=True))
    for t in YEARS:
        ws.write_formula(f"{YC[t]}{R['ke_n']}", f"=IF(RateConv=1,$C{R['ke_r']},(1+$C{R['ke_r']})*(1+{YC[t]}${TL['pi']})-1)", F(num_format=PCT2))
    ws.write_formula(f"{TCOL}{R['ke_n']}", f"=IF(RateConv=1,$C{R['ke_r']},(1+$C{R['ke_r']})*(1+IPCALT)-1)", F(num_format=PCT2))
    ws.write_number(f"C{R['df']}", 1, F(num_format="0.0000"))
    for t in YEARS:
        p = "C" if t == 1 else YC[t - 1]
        ws.write_formula(f"{YC[t]}{R['df']}", f"={p}{R['df']}/(1+{YC[t]}{R['ke_n']})", F(num_format="0.0000"))
    o_roe, o_g = f"Assumptions!{col}{OVR[('roe', g)]}", f"Assumptions!{col}{OVR[('g', g)]}"
    ws.write_formula(f"C{R['roelt']}", f"=IF(ISNUMBER({o_roe}),{o_roe},CHOOSE(RoeLT,'Boxes Now'!Z{bn},'Boxes Now'!W{bn})+$C${P_['roe']}/100)", F(num_format=PCT, bold=True))
    ws.write_formula(f"C{R['glt']}", f"=IF(ISNUMBER({o_g}),{o_g},$C${P_['g']})", F(num_format=PCT2))
    gx, ga = f"'Boxes Now'!$X${bn}", f"$C${P_['gadd']}/100"
    ws.write_formula(f"{YC[2]}{R['g']}", f"={gx}+{ga}", F(num_format=PCT))
    for t in range(3, NY + 1):
        ws.write_formula(f"{YC[t]}{R['g']}", f"={gx}+($C{R['glt']}-{gx})*MIN(1,({YC[t]}${TL['t']}-2)/(GrowthYear-2))+{ga}", F(num_format=PCT))
    ws.write_formula(f"{TCOL}{R['g']}", f"=$C{R['glt']}", F(num_format=PCT))
    ws.write_formula(f"{YC[1]}{R['e']}", f"='Boxes Now'!R{bn}*(1+$C{R['epsadj']})", fp1)
    for t in range(2, NY + 1):
        ws.write_formula(f"{YC[t]}{R['e']}", f"={YC[t - 1]}{R['e']}*(1+{YC[t]}{R['g']})", fp1)
    ws.write_formula(f"{TCOL}{R['e']}", f"={YC[NY]}{R['e']}*(1+{TCOL}{R['g']})", fp1)
    ws.write_formula(f"{YC[1]}{R['roe']}", f"='Boxes Now'!V{bn}", F(num_format=PCT))
    ws.write_formula(f"{YC[2]}{R['roe']}", f"='Boxes Now'!W{bn}", F(num_format=PCT))
    for t in range(3, NY + 1):
        ws.write_formula(f"{YC[t]}{R['roe']}", f"=${YC[2]}{R['roe']}+($C{R['roelt']}-${YC[2]}{R['roe']})*MIN(1,({YC[t]}${TL['t']}-2)/(RoeYear-2))", F(num_format=PCT))
    ws.write_formula(f"{TCOL}{R['roe']}", f"=$C{R['roelt']}", F(num_format=PCT))
    for c in [YC[t] for t in YEARS] + [TCOL]:
        ws.write_formula(f"{c}{R['bv']}", f"={c}{R['e']}/{c}{R['roe']}", fp1)
    for t in YEARS:
        c, nxt = YC[t], (YC[t + 1] if t < NY else TCOL)
        ws.write_formula(f"{c}{R['dbv']}", f"={nxt}{R['bv']}-{c}{R['bv']}", fp1)
        ws.write_formula(f"{c}{R['fcfe']}", f"={c}{R['e']}-{c}{R['dbv']}", F(num_format=PTS1, bold=True))
        ws.write_formula(f"{c}{R['payout']}", f"={c}{R['fcfe']}/{c}{R['e']}", F(num_format=PCT, font_color=GREY))
        ws.write_formula(f"{c}{R['pv']}", f"={c}{R['fcfe']}*{c}{R['df']}", fp1)
        ws.write_formula(f"{c}{R['rfcfe']}", f"={c}{R['fcfe']}/{c}${TL['defl']}", F(num_format=PTS1, font_color=GREY))
    ws.write_formula(f"C{R['pv']}", f"=SUM({YC[1]}{R['pv']}:{YC[NY]}{R['pv']})", fp)
    ws.write_formula(f"{TCOL}{R['tv']}", f"={TCOL}{R['e']}*(1-{TCOL}{R['g']}/{TCOL}{R['roe']})/({TCOL}{R['ke_n']}-{TCOL}{R['g']})", fp)
    ws.write_formula(f"{TCOL}{R['pvtv']}", f"={TCOL}{R['tv']}*{YC[NY]}{R['df']}", fp)
    ws.write_formula(f"C{R['pvtv']}", f"={TCOL}{R['pvtv']}", fp)
    ws.write_formula(f"C{R['val0']}", f"=C{R['pv']}+C{R['pvtv']}", F(num_format=PTS, bold=True))
    ws.write_formula(f"C{R['val']}", f"=IF(Horizon=1,C{R['val0']}*(1+{YC[1]}{R['ke_n']})-{YC[1]}{R['fcfe']},C{R['val0']})", F(num_format=PTS, bold=True, bg_color=KEY))
    ws.write_formula(f"C{R['up']}", f"=C{R['val']}/C{R['pts']}-1", F(num_format=UPS, bold=True, bg_color=KEY))
    ws.write_formula(f"C{R['pe_impl']}", f"=C{R['val']}/IF(Horizon=1,{YC[2]}{R['e']},{YC[1]}{R['e']})", F(num_format=MULT))
    ws.write_formula(f"C{R['pe_now']}", f"='Boxes Now'!Y{bn}", F(num_format=MULT))
    ws.write_formula(f"C{R['tvshare']}", f"=C{R['pvtv']}/C{R['val0']}", F(num_format=PCT))
    notes = {"pts": "Boxes Now (members' weight x Ibovespa close).",
             "epsadj": "Box EPS sensitivity to the Selic (Assumptions section 5) x Selic cut; applied to years 1 and 2 (0 if UseEPS = 0).",
             "erp": "Override if typed; else the basis ERP (1 typed, 2 Ibovespa 10y average, 3 this box's 10y average) + shift - k x sd of its history (History).",
             "ke_r": "5y real rate (scenario) + ERP. House convention: this is the discount rate (Target!B55).",
             "ke_n": "House convention: = cost of equity. Fisher: (1 + cost of equity) x (1 + IPCA of the year) - 1; terminal uses the long-term IPCA.",
             "df": "Product of 1 / (1 + nominal cost of equity) up to the year.",
             "roelt": "RoeLT 1: 10-year average of the box's 12m fwd ROE (History); 2: consensus year-2 ROE. Plus the shift in Assumptions, or the override.",
             "glt": "Long-term nominal growth of the scenario (house 4.7%), or the override.",
             "g": "Year 2 = consensus; converges linearly to the long-term growth by GrowthYear; + the scenario's growth add-on in years 2-NY; terminal = long term.".replace("NY", str(NY)),
             "e": "Year 1 = consensus blended to the forward year (Boxes Now) x (1 + change from rates); then grows at the row above.",
             "roe": "Years 1-2 = consensus; converges linearly to the long-term ROE by RoeYear (Assumptions).",
             "bv": "Equity at the start of the year needed to earn the year's earnings at the year's ROE = earnings / ROE.",
             "dbv": "Book value at the start of next year - at the start of this year = earnings retained to fund growth.",
             "fcfe": "Earnings - increase in book value (what can be paid out). In the long term = earnings x (1 - g / ROE).",
             "pv": f"FCFE x discount factor. 'Today' = sum of years 1-{NY}.",
             "tv": "Year-11 earnings x (1 - g / ROE) / (Ke - g), nominal long-term values (Gordon growth).",
             "val0": f"Present value of FCFE (years 1-{NY}) + present value of the terminal value.",
             "val": "Horizon 1: value today x (1 + year-1 discount rate) - year-1 FCFE (paid out during the year) = value in 12 months. Horizon 0: value today.",
             "pe_now": "Index points / consensus year-1 earnings (without the change from rates).",
             "rfcfe": "FCFE / deflator (timeline): used by Implied ERP and Sensitivity to re-discount the same cash flows at other rates."}
    for k, n in notes.items():
        ws.write(f"{NC}{R[k]}", n, bk.NOTE)
    BLK[(s, g)] = R
    return r0 + len(BKEYS) + 4


def ibov_block(ws, s, r0):
    R = {}
    bk.section(ws, r0, f"Ibovespa = sum of the four boxes ({s})", "B", NC)
    years_header(ws, r0 + 1)
    rows = [("pts", "Index points", "C", PTS), ("e", "Earnings (index pts)", "Y", PTS1), ("bv", "Book value at the start of the year", "YT", PTS1),
            ("fcfe", "Free cash flow to equity (FCFE)", "Y", PTS1), ("pv", "Present value of FCFE", "CY", PTS1),
            ("pvtv", "Present value of the terminal value", "C", PTS), ("val0", "Equity value today (index pts)", "C", PTS),
            ("val", "Fair value at the horizon (index pts)", "C", PTS)]
    for i, (k, lab, where, fmt) in enumerate(rows):
        r = r0 + 2 + i
        R[k] = r
        bold = k in ("fcfe", "val")
        ws.write(f"B{r}", lab, F(bold=bold))
        cols = (["C"] if "C" in where else []) + ([YC[t] for t in YEARS] if "Y" in where else []) + ([TCOL] if "T" in where else [])
        for c in cols:
            terms = "+".join(f"{c}{BLK[(s, g)][k]}" for g in BOXES)
            ws.write_formula(f"{c}{r}", "=" + terms, F(num_format=fmt, bold=bold, bg_color=KEY if k == "val" else "#FFFFFF"))
    r = r0 + 10
    R.update(up=r, roe=r + 1, pe_impl=r + 2, pe_now=r + 3, tvshare=r + 4)
    ws.write(f"B{r}", "Upside (fair value / index points today - 1)", bk.BOLD)
    ws.write_formula(f"C{r}", f"=C{R['val']}/C{R['pts']}-1", F(num_format=UPS, bold=True, bg_color=KEY))
    ws.write(f"B{r + 1}", "ROE (sum of earnings / sum of book value)", bk.TXT)
    for t in YEARS:
        ws.write_formula(f"{YC[t]}{r + 1}", f"={YC[t]}{R['e']}/{YC[t]}{R['bv']}", F(num_format=PCT))
    ws.write(f"B{r + 2}", "Implied 12m fwd P/E at the horizon (fair value / next 12m earnings)", bk.TXT)
    ws.write_formula(f"C{r + 2}", f"=C{R['val']}/IF(Horizon=1,{YC[2]}{R['e']},{YC[1]}{R['e']})", F(num_format=MULT))
    ws.write(f"B{r + 3}", "12m fwd P/E today (consensus year-1 earnings)", bk.TXT)
    ws.write_formula(f"C{r + 3}", f"='Boxes Now'!Y{BN['Ibovespa']}", F(num_format=MULT))
    ws.write(f"B{r + 4}", "Terminal value / equity value", bk.TXT)
    ws.write_formula(f"C{r + 4}", f"=C{R['pvtv']}/C{R['val0']}", F(num_format=PCT))
    ws.write(f"B{r + 5}", "Contribution to the Ibovespa upside (pp)", bk.SUB)
    for i, g in enumerate(BOXES):
        rr = r + 6 + i
        R["ctb_" + g] = rr
        ws.write(f"B{rr}", "   " + g, bk.TXT)
        ws.write_formula(f"C{rr}", f"=(C{BLK[(s, g)]['val']}-C{BLK[(s, g)]['pts']})/C{R['pts']}*100", F(num_format=PP))
    BLK[(s, "Ibovespa")] = R


for s in SCEN:
    dcf_sheet(s)


def closed_form(sh, R, kr):
    """Value of a box's cash flows at cost of equity kr = 1 + real rate + ERP (Excel expression), as in the DCF sheet, in either convention."""
    return (f"SUMPRODUCT({sh}!${YC[1]}${R['rfcfe']}:${YC[NY]}${R['rfcfe']},{kr}^(-{sh}!${YC[1]}${TL['t']}:${YC[NY]}${TL['t']}))"
            f"+{sh}!${TCOL}${R['e']}*(1-{sh}!${TCOL}${R['g']}/{sh}!${TCOL}${R['roe']})"
            f"/({kr}*IF(RateConv=1,1,1+IPCALT)-1-{sh}!${TCOL}${R['g']})/{sh}!${YC[NY]}${TL['defl']}*{kr}^(-{NY})")


def target_form(sh, R, kr):
    """Fair value at the horizon from the closed-form value today (same roll-forward as the DCF sheet)."""
    cf = closed_form(sh, R, kr)
    return f"IF(Horizon=1,({cf})*IF(RateConv=1,{kr},{kr}*(1+IPCA1))-{sh}!${YC[1]}${R['fcfe']},{cf})"


# =========================================================================== Implied ERP
ws = W["Implied ERP"]
bk.title(ws, "Implied ERP: the risk premium that makes the DCF equal today's price",
         "Each column re-discounts the scenario's cash flows at real rate + ERP (convention in use) for ERP = 0.0% to 15.0%; the implied ERP is interpolated where the value crosses the index points.")
GR = BOXES + ["Ibovespa"]
ERPS = [round(0.001 * i, 3) for i in range(151)]
G0 = 21
GLAST = G0 + len(ERPS) - 1
ws.set_column("A:A", 2)
ws.set_column("B:B", 34)
ws.set_column("C:V", 11)
bk.section(ws, 5, "Implied ERP at today's prices, by box and scenario (each scenario's real rate and cash flows)", "B", "G")
bk.header_row(ws, 6, [("B", "Group"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), ("F", "Custom"), ("G", "10y average (History)")])
IMP, gcol = {}, {}
for j, s in enumerate(SCEN):
    for i, g in enumerate(GR):
        gcol[(s, g)] = cn(2 + 5 * j + i)
for i, g in enumerate(GR):
    r = 7 + i
    tot = g == "Ibovespa"
    ws.write(f"B{r}", g, F(bold=tot, top=1 if tot else 0))
    for j, s in enumerate(SCEN):
        vc = gcol[(s, g)]
        price = f"{vc}${G0 - 3}"
        vr, er = f"{vc}${G0}:{vc}${GLAST}", f"$B${G0}:$B${GLAST}"
        k = f"MATCH({price},{vr},-1)"
        f_ = f'=IFERROR(IF({price}>{vc}${G0},"< 0%",INDEX({er},{k})+(INDEX({vr},{k})-{price})/(INDEX({vr},{k})-INDEX({vr},{k}+1))*0.001),"> 15%")'
        c = cn(2 + j)
        ws.write_formula(f"{c}{r}", f_, F(num_format=PCT2, bold=tot, top=1 if tot else 0, align="right"))
        IMP[(s, g)] = f"'Implied ERP'!${c}${r}"
    ws.write_formula(f"G{r}", f"=History!{HCOL['erp'][g]}{HST['avg']}", F(num_format=PCT2, top=1 if tot else 0))
ws.write("B13", "ERP = cost of equity - 5y real rate, read with each scenario's cash flows and the discount-rate convention in use (RateConv). The History column uses steady-state cash flows.", bk.NOTE)
bk.section(ws, 15, "Grid: equity value (index pts) at each ERP", "B", cn(1 + 5 * len(SCEN)))
for j, s in enumerate(SCEN):
    ws.merge_range(f"{cn(2 + 5 * j)}16:{cn(6 + 5 * j)}16", f"{s} (real rate: DCF {s}!C6)", bk.HDR)
ws.write(f"B{G0 - 4}", "ERP", bk.HDRL)
ws.write(f"B{G0 - 3}", "Index points today", bk.BOLD)
ws.write(f"B{G0 - 2}", "DCF value today (DCF sheet)", bk.NOTE)
ws.write(f"B{G0 - 1}", "Check: closed form at the ERP used", bk.NOTE)
for j, s in enumerate(SCEN):
    sh = q(f"DCF {s}")
    for g in GR:
        vc = gcol[(s, g)]
        ws.write(f"{vc}{G0 - 4}", g, bk.HDR)
        ws.write_formula(f"{vc}{G0 - 3}", f"={sh}!$C${BLK[(s, g)]['pts']}", F(num_format=PTS, bold=True))
        ws.write_formula(f"{vc}{G0 - 2}", f"={sh}!$C${BLK[(s, g)]['val0']}", F(num_format=PTS, font_color=GREY))
        if g == "Ibovespa":
            ws.write_formula(f"{vc}{G0 - 1}", "=" + "+".join(f"{gcol[(s, b)]}{G0 - 1}" for b in BOXES), F(num_format=PTS, font_color=GREY))
        else:
            R = BLK[(s, g)]
            ws.write_formula(f"{vc}{G0 - 1}", "=" + closed_form(sh, R, f"(1+{sh}!$C${R['ke_r']})"), F(num_format=PTS, font_color=GREY))
for k, e in enumerate(ERPS):
    r = G0 + k
    ws.write_number(f"B{r}", e, F(num_format=PCT, align="center"))
    for j, s in enumerate(SCEN):
        sh = q(f"DCF {s}")
        for g in BOXES:
            ws.write_formula(f"{gcol[(s, g)]}{r}", "=" + closed_form(sh, BLK[(s, g)], f"(1+{sh}!$C$6+$B{r})"), F(num_format=PTS))
        ws.write_formula(f"{gcol[(s, 'Ibovespa')]}{r}", "=" + "+".join(f"{gcol[(s, g)]}{r}" for g in BOXES), F(num_format=PTS, bold=True))
ws.freeze_panes(G0, 2)

# =========================================================================== Sensitivity
ws = W["Sensitivity"]
bk.title(ws, "Sensitivity: Ibovespa DCF value vs real rate and ERP (Base cash flows)",
         "One ERP for every box; value depends on real rate + ERP. Cash flows of DCF Base. Edit the pink headers to change the grid points.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 30)
ws.set_column("C:Y", 10.5)
RRS = [0.05, 0.055, 0.06, 0.065, 0.0676, 0.07, 0.075, 0.08, 0.085, 0.09, 0.095]
ERPG = [round(0.035 + 0.005 * i, 3) for i in range(11)]
sh = q("DCF Base")
for blk_i, (lab, kind) in enumerate([("Ibovespa fair value at the horizon (index pts)", "pts"), ("Upside vs today", "up")]):
    top = 5 + blk_i * 17
    bk.section(ws, top, lab, "B", cn(2 + len(ERPG)))
    ws.write(f"B{top + 1}", "5-year real rate (rows)   |   ERP (columns)", bk.NOTE)
    hr = top + 2
    ws.write(f"B{hr}", "Real rate \\ ERP", bk.HDRL)
    for j, e in enumerate(ERPG):
        if blk_i == 0:
            ws.write_number(hr - 1, 2 + j, e, F(**LV, num_format=PCT, align="center"))
        else:
            ws.write_formula(hr - 1, 2 + j, f"={cn(2 + j)}$7", F(bold=True, font_color="#FFFFFF", bg_color=NAVY, num_format=PCT, align="center"))
    for i, rr_ in enumerate(RRS):
        r = hr + 1 + i
        if blk_i == 0:
            ws.write_number(f"B{r}", rr_, F(**LV, num_format=PCT2, align="center"))
        else:
            ws.write_formula(f"B{r}", f"=$B${8 + i}", F(bold=True, num_format=PCT2, align="center"))
        for j in range(len(ERPG)):
            if blk_i == 0:
                kr = f"(1+$B{r}+{cn(2 + j)}${hr})"
                ws.write_formula(r - 1, 2 + j, "=" + "+".join(target_form(sh, BLK[("Base", g)], kr) for g in BOXES), F(num_format=PTS))
            else:
                ws.write_formula(r - 1, 2 + j, f"={cn(2 + j)}{8 + i}/IbovNow-1", F(num_format=UPS))
    last = hr + len(RRS)
    if kind == "up":
        ws.conditional_format(f"C{hr + 1}:{cn(1 + len(ERPG))}{last}", HEAT)
    ws.conditional_format(f"C{hr + 1}:{cn(1 + len(ERPG))}{last}", {"type": "formula",
                          "criteria": f"=AND(ROUND($B{hr + 1}-Assumptions!$E${A['rr']},4)=0,ROUND(C${hr}-Assumptions!$E${A['erpu']},3)=0)",
                          "format": F(bold=True, border=2, border_color=NAVY)})
ws.write("B41", "Framed cell = Base real rate and the Base ERP used (Ibovespa level, rounded to the grid). Bear/Bull also change earnings (rates) and are in their own sheets.", bk.NOTE)

# =========================================================================== Summary
ws = W["Summary"]
bk.title(ws, "DCF by box: bear / base / bull", "Separate from the multiples model. Construction: DCF Bear / Base / Bull / Custom sheets.", legend=False)
ws.set_column("A:A", 2)
ws.set_column("B:B", 48)
ws.set_column("C:L", 13)
ws.write("B3", "Pricing date", bk.NOTE)
ws.write_formula("C3", "=PxDate", F(num_format=DATED))
ws.write("D3", "Ibovespa", bk.NOTE)
ws.write_formula("E3", "=IbovNow", F(num_format=PTS))
ws.write("F3", "Checks", bk.NOTE)
ws.write_formula("G3", "=Checks!C5", F(bold=True))
ws.write("H3", "Horizon", bk.NOTE)
ws.write_formula("I3", '=IF(Horizon=1,"12 months (fair value in 12m)","today")', F(bold=True))
bk.section(ws, 5, "Scenario levers in use (Assumptions)", "B", "G")
bk.header_row(ws, 6, [("B", ""), ("C", "Today"), ("D", "Bear"), ("E", "Base"), ("F", "Bull"), ("G", "Custom")])
for i, (lab, ar, fmt) in enumerate([("Narrative", A["narr"], None), ("Selic at the end of the horizon", A["selic"], PCT2), ("5-year real rate (NTN-B)", A["rr"], PCT2),
                                     ("ERP used, Ibovespa level", A["erpu"], PCT2), ("Cost of equity = real rate + ERP, Ibovespa level", A["ke"], PCT2),
                                     (f"Earnings growth vs consensus path, years 2-{NY} (pp/yr)", A["gadd"], PPF), ("Long-term nominal growth (g)", A["g"], PCT2)]):
    r = 7 + i
    ws.write(f"B{r}", lab, bk.TXT)
    for c in "DEFG":
        ws.write_formula(f"{c}{r}", f"=Assumptions!{c}{ar}", F(num_format=fmt, align="center") if fmt else F(align="center"))
ws.write_formula("C8", "=SelicNow", F(num_format=PCT2, align="center"))
ws.write_formula("C9", "=RR5Now", F(num_format=PCT2, align="center"))
ws.write("B14", "Settings (Assumptions section 4)", bk.NOTE)
ws.write_formula("C14", '="ERP basis "&CHOOSE(ERPBasis,"1: typed","2: Ibovespa 10y average","3: each box\'s 10y average")&"  ·  rate convention "&CHOOSE(RateConv,"1: house","2: Fisher")&"  ·  long-term ROE "&CHOOSE(RoeLT,"1: 10y average","2: consensus")', F(italic=True, font_color=GREY))
bk.section(ws, 15, "DCF fair value by box (index pts, at the horizon in H3) and upside", "B", "L")
bk.header_row(ws, 16, [("B", "Group"), ("C", "Index points"), ("D", "Bear"), ("E", "Base"), ("F", "Bull"), ("G", "Custom"), ("H", ""),
                       ("I", "Upside: Bear"), ("J", "Base"), ("K", "Bull"), ("L", "Custom")])
SUMR = {}
for i, g in enumerate(BOXES + ["Ibovespa", None] + MEMO):
    r = 17 + i
    if g is None:
        ws.write(f"B{r}", "Memo", bk.NOTE)
        continue
    SUMR[g] = r
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", g, F(**b))
    ws.write_formula(f"C{r}", f"={q('DCF Base')}!C{BLK[('Base', g)]['pts']}", F(num_format=PTS, **b))
    for c, uc, s in zip("DEFG", "IJKL", SCEN):
        sh = q(f"DCF {s}")
        ws.write_formula(f"{c}{r}", f"={sh}!C{BLK[(s, g)]['val']}", F(num_format=PTS, **b))
        ws.write_formula(f"{uc}{r}", f"={sh}!C{BLK[(s, g)]['up']}", F(num_format=UPS, **b))
ws.conditional_format("I17:L21", HEAT)
ws.conditional_format("I23:L24", HEAT)
bk.section(ws, 26, "Implied 12m fwd P/E at the horizon (fair value / next 12m earnings) vs today", "B", "G")
bk.header_row(ws, 27, [("B", "Group"), ("C", "Today"), ("D", "Bear"), ("E", "Base"), ("F", "Bull"), ("G", "Custom")])
PER = {}
for i, g in enumerate(BOXES + ["Ibovespa"]):
    r = 28 + i
    PER[g] = r
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", g, F(**b))
    ws.write_formula(f"C{r}", f"={q('DCF Base')}!C{BLK[('Base', g)]['pe_now']}", F(num_format=MULT, **b))
    for c, s in zip("DEFG", SCEN):
        ws.write_formula(f"{c}{r}", f"={q('DCF ' + s)}!C{BLK[(s, g)]['pe_impl']}", F(num_format=MULT, **b))
bk.section(ws, 34, "Implied ERP at today's prices (cost of equity - 5y real rate that makes the DCF = index points)", "B", "G")
bk.header_row(ws, 35, [("B", "Group"), ("C", "ERP used (Base)"), ("D", "Bear"), ("E", "Base"), ("F", "Bull"), ("G", "Custom")])
IMPR = {}
for i, g in enumerate(BOXES + ["Ibovespa"]):
    r = 36 + i
    IMPR[g] = r
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", g, F(**b))
    used = f"=Assumptions!E{A['erpu']}" if tot else f"={q('DCF Base')}!C{BLK[('Base', g)]['erp']}"
    ws.write_formula(f"C{r}", used, F(num_format=PCT2, **b))
    for c, s in zip("DEFG", SCEN):
        ws.write_formula(f"{c}{r}", "=" + IMP[(s, g)], F(num_format=PCT2, align="right", **b))
bk.section(ws, 42, "Implied ERP history in DCF terms (History; 5y real rate, steady state)", "B", "G")
bk.header_row(ws, 43, [("B", "Group"), ("C", "10y average"), ("D", "Std. dev."), ("E", "Now"), ("F", "Now vs average (sd)"), ("G", "")])
for i, g in enumerate(BOXES + ["Ibovespa"]):
    r = 44 + i
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", g, F(**b))
    c_ = HCOL["erp"][g]
    for c, st, fmt in (("C", "avg", PCT2), ("D", "sd", PCT2), ("E", "now", PCT2), ("F", "z", "+0.0;-0.0")):
        ws.write_formula(f"{c}{r}", f"=History!{c_}{HST[st]}", F(num_format=fmt, **b))
bk.section(ws, 50, "Terminal value share and contribution to the Ibovespa upside", "B", "G")
bk.header_row(ws, 51, [("B", ""), ("C", ""), ("D", "Bear"), ("E", "Base"), ("F", "Bull"), ("G", "Custom")])
ws.write("B52", "Terminal value / equity value (Ibovespa)", bk.TXT)
for c, s in zip("DEFG", SCEN):
    ws.write_formula(f"{c}52", f"={q('DCF ' + s)}!C{BLK[(s, 'Ibovespa')]['tvshare']}", F(num_format=PCT))
for i, g in enumerate(BOXES):
    r = 53 + i
    ws.write(f"B{r}", f"Contribution: {g} (pp)", bk.TXT)
    for c, s in zip("DEFG", SCEN):
        ws.write_formula(f"{c}{r}", f"={q('DCF ' + s)}!C{BLK[(s, 'Ibovespa')]['ctb_' + g]}", F(num_format=PP))
bk.section(ws, 58, "Reference (not linked): multiples model and the house DCF", "B", "G")
bk.header_row(ws, 59, [("B", "Ibovespa fair value (index pts)"), ("C", ""), ("D", "Bear"), ("E", "Base"), ("F", "Bull"), ("G", "")])
for i, (lab, vals) in enumerate([("Multiples model, average of methods (Oct-6 table)", [REF["fv"]["Average"][f"Ibovespa|{s}"] for s in ("Bear", "Base", "Bull")]),
                                  ("Multiples model, P/E method", [REF["fv"]["P/E"][f"Ibovespa|{s}"] for s in ("Bear", "Base", "Bull")]),
                                  ("House DCF (DCF Ibov 2026_Out.xlsx), fair value 2026", [None, 202259.2, None])]):
    r = 60 + i
    ws.write(f"B{r}", lab, bk.TXT)
    for c, v in zip("DEF", vals):
        if v is not None:
            ws.write_number(f"{c}{r}", v, F(**bk.IN, num_format=PTS))
ws.write("B63", "House DCF: FCFF of the whole index at 13.4% (5y real rate 7.42% on its date + ERP 6%), g 4.7%, value at end-2026; working capital by level, no net debt, terminal value discounted 5 years. Same premises, different cash flows.", bk.NOTE)
ws.freeze_panes(4, 0)

# =========================================================================== Checks
ws = W["Checks"]
bk.title(ws, "Checks", "'OK' / 'CHECK' drive the status in C5; 'Info' does not.", legend=False)
ws.set_column("A:A", 2)
ws.set_column("B:B", 78)
ws.set_column("C:F", 15)
ws.set_column("G:G", 80)
ws.write("B5", "Overall status", bk.BOLD)
bk.header_row(ws, 7, [("B", "Check"), ("C", "Value"), ("D", "Target"), ("E", "Difference"), ("F", "Status"), ("G", "Note")])
chk = [("Member weights on the pricing date add up to 100%", f"=Members!K{MT}", "=1", 1e-9, "", "0.0000%"),
       ("The four boxes add up to the Ibovespa (index points)", "=SUM('Boxes Now'!F6:F9)", "=IbovNow", 0.01, "", PX),
       ("Memo sub-boxes add up to Cyclicals (index points)", "='Boxes Now'!F12+'Boxes Now'!F13", "='Boxes Now'!F8", 0.01, "", PX),
       ("Every member mapped to a box", f"=SUMPRODUCT(--ISNA(Members!D{MF}:D{ML}))", "=0", 0, "", "0"),
       ("Lowest data coverage among the boxes (% of points)", "=MIN('Boxes Now'!H6:H9)", "=0.9", None, "Info: members without FY26-FY28 P/E or ROE are left out and the box is scaled up.", "0.0%"),
       ("Months of history in the window (Ibovespa ROE)", f"=History!{HCOL['roe']['Ibovespa']}{HST['n']}", "=120", None, "Info.", "0")]
for s in SCEN:
    chk.append((f"{s}: closed-form value today at the ERP used = DCF value today (Ibovespa)", f"='Implied ERP'!{gcol[(s, 'Ibovespa')]}{G0 - 1}",
                f"='Implied ERP'!{gcol[(s, 'Ibovespa')]}{G0 - 2}", 1, "Same cash flows re-discounted by the Implied ERP formula.", PTS))
for s in SCEN:
    for g in BOXES:
        Rg = BLK[(s, g)]
        sh = q(f"DCF {s}")
        chk.append((f"{s}, {g}: long-term Ke minus long-term growth", f"={sh}!{TCOL}{Rg['ke_n']}-{sh}!{TCOL}{Rg['g']}", "=0", None,
                    "Info: must be positive for the terminal value to exist.", PCT2))
chk.append(("Overrides in use (Assumptions section 6)", f"=COUNT(Assumptions!D{A['ovr'][0]}:G{A['ovr'][1]})", "=0", None, "Info.", "0"))
chk.append(("Formula errors in the DCF sheets", "=" + "+".join(f"SUMPRODUCT(--ISERROR({q('DCF ' + s)}!B5:{NC}250))" for s in SCEN), "=0", 0, "", "0"))
chk.append(("Formula errors in Summary, Implied ERP and History", f"=SUMPRODUCT(--ISERROR(Summary!B5:L70))+SUMPRODUCT(--ISERROR('Implied ERP'!C{G0 - 1}:V{GLAST}))+SUMPRODUCT(--ISERROR(History!C8:X{HL}))", "=0", 0, "", "0"))
r = 8
for lab, val, tgt, tol, note, fmt in chk:
    ws.write(f"B{r}", lab, bk.TXT)
    ws.write_formula(f"C{r}", val, F(num_format=fmt))
    ws.write_formula(f"D{r}", tgt, F(num_format=fmt))
    ws.write_formula(f"E{r}", f'=IFERROR(C{r}-D{r},"")', F(num_format=fmt))
    if tol is None:
        st = f'=IF(C{r}>D{r},"Info","Info: not positive")' if "Ke minus" in lab else '="Info"'
        ws.write_formula(f"F{r}", st, F(font_color=GREY, align="center"))
    else:
        ws.write_formula(f"F{r}", f'=IF(ABS(C{r}-D{r})<={tol:.10f},"OK","CHECK")', F(bold=True, align="center"))
    ws.write(f"G{r}", note, bk.NOTE)
    r += 1
ws.write_formula("C5", f'=IF(COUNTIF(F8:F{r},"CHECK")=0,"All checks OK","Some checks need attention")', F(bold=True, font_color=GREEN))
ws.conditional_format(f"F8:F{r}", {"type": "cell", "criteria": "==", "value": '"CHECK"', "format": F(font_color="#C00000", bold=True)})
ws.freeze_panes(7, 0)

# =========================================================================== Read Me
ws = W["Read Me"]
ws.set_column("A:A", 2)
ws.set_column("B:B", 26)
ws.set_column("C:C", 130)
ws.set_row(0, 26)
ws.write("B1", "Ibovespa DCF by box — model", F(bold=True, font_size=16, font_color=NAVY))
ws.write("B2", "XP Research · Equity Strategy · built Oct-7-2026 · prices as of Oct-5-2026, consensus as of Oct-7-2026", bk.NOTE)
rr_ = 4
for head, lines in [
    ("What it does", [("Idea", "Values each box of the Ibovespa (Financials, Defensives, Cyclicals, Commodities) with a free-cash-flow-to-equity DCF and adds them up. Separate from the multiples model (Ibov_valuation_by_box_model_XP.xlsx).")]),
    ("How to use it", [("1. Change", "Assumptions: the pink cells (real rate, ERP and ERP basis, discount-rate convention, long-term growth and ROE, convergence years, IPCA, overrides)."),
                       ("2. Read", "Summary; Implied ERP (what today's prices imply); History (ROE, P/E and implied ERP since 2016); Sensitivity (real rate x ERP)."),
                       ("3. Trace", f"Summary → DCF <scenario> (one block per box, years 1-{NY} + terminal) → Boxes Now / History → Members / ROE Panel.")]),
    ("Method", [("Earnings", "Consensus FY26-FY28 by stock (price / P/E), in Ibovespa points, summed by box; forward years 1-2 blend the fiscal years; growth then converges linearly to the long term by GrowthYear."),
                ("ROE", "Consensus FY26-FY28 (earnings / book value of the box), converging linearly to the box's 10-year average 12m fwd ROE by RoeYear."),
                ("FCFE", "Earnings - increase in book value, with book value = earnings / ROE of the same year. In the long term this is earnings x (1 - g / ROE): growth costs reinvestment, and more so when ROE is low."),
                ("House premises", "As the house Ibovespa model (DCF Ibov 2026_Out.xlsx, Target sheet): cost of equity = 5y real rate (NTN-B) + ERP 6.0%; long-term growth 4.7% nominal; "
                                   "IPCA 4.0%; bear / base / bull real rates 8.5% / today / 6.0%."),
                ("Discount rate", "RateConv 1 (default, house): real rate + ERP discounts the nominal (consensus) cash flows, as Target!B55. "
                                  "RateConv 2 (Fisher): (1 + real rate + ERP) x (1 + IPCA) - 1, about 4 pp higher; with the same 6.0% ERP it lowers the value by roughly 40%."),
                ("ERP", "Basis 1 (default): the house's 6.0%. Basis 2 / 3: the 10-year average implied ERP of the Ibovespa / each box (History): each month, the ERP that makes "
                        "P/E = (1 - g / ROE) / (Ke - g), in the convention in use."),
                ("Terminal value", f"Gordon growth on year-{NY + 1} earnings: E x (1 - g / ROE) / (Ke - g), nominal long-term values."),
                ("Horizon", "Horizon 1 (default): fair value in 12 months = value today x (1 + year-1 discount rate) - year-1 FCFE, as a 12-month target price; upside vs today's index points."),
                ("Scenarios", "Each scenario sets the Selic (earnings, years 1-2), the real rate, the ERP (typed, shift and k x sd of its history, as the multiples model's k), "
                              f"the earnings growth add-on in years 2-{NY} and the long-term growth."),
                ("Implied ERP", "The ERP that makes the DCF equal today's index points, found on a 0.1% grid and interpolated (Implied ERP sheet).")]),
    ("Sources", [("Prices, weights", "Bloomberg PX_LAST (Oct-5); Economatica Ibovespa composition (Sep-30), drifted by price."),
                 ("Consensus", "Bloomberg BEST_PE_RATIO and BEST_ROE (bst), FY26/FY27/FY28, Oct-7-2026."),
                 ("History", "Bloomberg BEST_ROE and BEST_PE_RATIO (BF), monthly, Oct-16..Sep-26, members of each month; 5y NTN-B real rate."),
                 ("Rates and macro", "5y NTN-B real rate (ANBIMA, Oct-5; the house model reads BZRFB5PY Index); ERP, growth and IPCA from the house Ibovespa model (DCF Ibov 2026_Out.xlsx)."),
                 ("Code", "Brazil Bull Case/scripts: dcf_data.py (Bloomberg) and build_dcf.py (this workbook).")]),
    ("Open points", [("Inputs to confirm", "Discount-rate convention (house vs Fisher), convergence years. Premises follow the house model; differences to its DCF come from the cash flows (FCFE by box vs FCFF of the index)."),
                     ("Not yet", "Not integrated into the multiples model's average; banks treated with the same FCFE logic; ITSA4, AURE3, CSAN3, CSNA3, GOAU4, MRVE3, HAPV3 left out for missing or extreme consensus (boxes scaled up).")])]:
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
ws = bk.cover("Ibovespa", " | DCF by box", "VALUATION MODEL  ·  EQUITY STRATEGY",
              [("Summary", "Summary"), ("Assumptions", "Assumptions"), ("DCF Base", "DCF Base"), ("Implied ERP", "Implied ERP")],
              [[("DCF Bear", "DCF Bear"), ("DCF Bull", "DCF Bull"), ("DCF Custom", "DCF Custom"), ("Sensitivity\n(real rate x ERP)", "Sensitivity")],
               [("Boxes Now\n(consensus by box)", "Boxes Now"), ("History\n(ROE, P/E, ERP)", "History"), ("Members\n(76 stocks)", "Members"),
                ("Checks\n(tests)", "Checks")]], readme="Read Me")
for i, (l_, f_, nf) in enumerate([("Ibovespa close", "=IbovNow", PTS), ("Pricing date", "=PxDate", DATED), ("Fair value, base case", f"=Summary!E{SUMR['Ibovespa']}", PTS),
                                   ("Upside, base case", f"=Summary!J{SUMR['Ibovespa']}", UPS), ("Model checks", "=Checks!C5", None)]):
    r = 8 + i
    bb = i == 2
    ws.merge_range(f"K{r}:L{r}", l_, F(font_name="Roboto Light", bold=bb, bottom=4))
    fmt_ = F(align="right", bottom=4, bold=bb, num_format=nf) if nf else F(align="right", bottom=4, bold=True, font_color=GREEN)
    ws.merge_range(f"M{r}:N{r}", "", fmt_)
    ws.write_formula(f"M{r}", f_, fmt_)
ws.write("K14", "Equity Strategy | XP Research", F(bold=True))
ws.write("K16", "Prices Oct-5; consensus Oct-7-2026.", bk.NOTE)
gl = F(bold=True, italic=True, font_size=9, font_color=GREY)
hd = F(bold=True, bottom=2, bottom_color="#8EB3DF", align="center")
ws.write("C19", "DCF summary >>", gl)
ws.merge_range("C20:E20", "Ibovespa", F(bold=True, bottom=2, bottom_color="#8EB3DF"))
for j, s in enumerate(SCEN):
    ws.write(19, 5 + j, s, hd)
crow = [("5-year real rate", "Summary!{c}9", PCT2, False, False), ("ERP used", "Summary!{c}10", PCT2, False, False),
        ("Fair value (index pts)", f"Summary!{{c}}{SUMR['Ibovespa']}", PTS, True, False), ("Upside", f"Summary!{{u}}{SUMR['Ibovespa']}", UPS, False, True),
        ("Implied 12m fwd P/E", f"Summary!{{c}}{PER['Ibovespa']}", MULT, False, False), ("Implied ERP at today's price", f"Summary!{{c}}{IMPR['Ibovespa']}", PCT2, False, True)]
for i, (l_, ref, nf, b, it) in enumerate(crow):
    r = 21 + i
    bt = 1 if i == len(crow) - 1 else 0
    ws.merge_range(f"C{r}:E{r}", ("   " if it else "") + l_, F(bold=b, italic=it, bottom=bt))
    for j, (c_, u_) in enumerate(zip("DEFG", "IJKL")):
        ws.write_formula(r - 1, 5 + j, "=" + ref.format(c=c_, u=u_), F(bold=b, italic=it, num_format=nf, align="center", bottom=bt))
ws.write("C28", "DCF = sum of four boxes, each valued by free cash flow to equity (consensus earnings, ROE converging to its 10-year average). House premises: 5y real rate + ERP 6.0%, g 4.7%.", bk.NOTE)
bk.close()
print("written", OUT)
