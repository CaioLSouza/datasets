"""Builds Ibov_valuation_by_box_model.xlsx: the Ibovespa fair value by box (bear / base / bull / custom), every number a
formula traceable to the raw inputs. Data come from the Oct-5/6 study (Equity Strategy Dashboard/_pedidos/
2026-10-05_ibov_valuation_groups); the pandas reference is scripts/bx_model.py there (approach A, final settings).

Run:  python build_model.py   then   python excel_recalc.py   (recalculates in Excel, checks errors and the tie-out)
"""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.worksheet import Worksheet
from xlsxwriter.utility import xl_col_to_name as cn

SRC = Path(r"C:\Users\Caio\Documents\Documentos\Equity Strategy Dashboard\_pedidos\2026-10-05_ibov_valuation_groups")
S = SRC / "scripts"
HERE = Path(__file__).parent
OUT = HERE.parent / "Ibov_valuation_by_box_model_XP.xlsx"

# --------------------------------------------------------------------------- data
BOXES = ["Financials", "Defensives", "Cyclicals", "Commodities"]
MEMO = ["Cyclicals ex-WEG & Embraer", "WEG & Embraer"]
G6 = BOXES + MEMO
NONFIN = ["Defensives", "Cyclicals", "Commodities"]
SCEN = ["Bear", "Base", "Bull", "Custom"]
SC = ["D", "E", "F", "G"]                       # scenario columns, the same in every sheet
GICS_MAP = [("Financials", "Financials"), ("Consumer Staples", "Defensives"), ("Health Care", "Defensives"),
            ("Utilities", "Defensives"), ("Communication Services", "Defensives"), ("Consumer Discretionary", "Cyclicals"),
            ("Industrials", "Cyclicals"), ("Information Technology", "Cyclicals"), ("Real Estate", "Cyclicals"),
            ("Energy", "Commodities"), ("Materials", "Commodities")]
GLOBAL_IND = ["WEGE3", "EMBJ3"]

m = pd.read_parquet(S / "bx_members.parquet").join(pd.read_parquet(S / "ke_sens_members.parquet"))
m["tp_used"] = m["tp_house"].fillna(m["tp_cons"])
m["bo"] = m["box"].map({b: i for i, b in enumerate(BOXES)})
m = m.sort_values(["bo", "glob", "w30"], ascending=[True, True, False])
d = json.loads((S / "bx_data.json").read_text(encoding="utf-8"))
macro = json.loads((S / "bx_macro.json").read_text(encoding="utf-8"))
rr = pd.Series(macro["real_rate"]).rename(lambda x: pd.Timestamp(x)).sort_index() / 100
p = pd.read_parquet(S / "bx_panel20.parquet").sort_values(["cod_ativo", "date"]).reset_index(drop=True)
pb = json.loads((SRC / "studies/performance_vs_rates/perf_study_results.json").read_text(encoding="utf-8"))["betas_preferred_table"]
reg = [x for x in json.loads((S / "bx_eps_rates2.json").read_text(encoding="utf-8"))
       if x["lag"] == 6 and x["spec"] == "with controls" and x["box"] == "Ibovespa"][0]
REF = json.loads((HERE / "ref_v6.json").read_text(encoding="utf-8"))   # pandas reference with the Oct-6 settings
months = sorted(p["date"].unique())
assert len(months) == 240 and len(rr) == 240

# --------------------------------------------------------------------------- styles (XP model template: template modelo XP.xlsx)
NAVY, LBLUE, YEL, GREY, GREEN = "#1F2F44", "#8EB3DF", "#F9C113", "#74797C", "#008000"
INK, TITLE_C = "#18191A", "#1D1E1F"
INBLUE, PINK, GBAND, KEY = "#0070C0", "#FFE7E7", "#F2F2F2", "#E9EAEC"
NOTE = GREY
PTS, PTS1, MULT, PCT, PCT2 = "#,##0", "#,##0.0", '0.0"x"', "0.0%", "0.00%"
UPS = "+0.0%;-0.0%;0.0%"
PP = '+0.0" pp";-0.0" pp";0.0" pp"'
PPC = '0.00" pp"'
BP = '+0" bp";-0" bp";0" bp"'
PX, DATE, DATED, MN = "#,##0.00", "[$-409]mmm-yy", "[$-409]d-mmm-yy", "#,##0"
MEDIA = HERE / "template_media"                  # logos taken from the template (xl/media/image1.png, image2.png)
LOGO_BLACK, LOGO_WHITE = MEDIA / "xp_research_black_header.png", MEDIA / "xp_research_white.png"

# sheet names as shown (internal keys keep the old names; formulas and texts are rewritten on the fly)
DISPLAY = {"Boxes_Now": "Boxes Now", "Rates_EPS": "Rates & EPS", "BottomUp": "Bottom-up", "README": "Read Me"}
_REN_F = [(f"{k}!", f"'{v}'!") for k, v in DISPLAY.items()]
_REN_T = list(DISPLAY.items())
_wf0, _ws0 = Worksheet._write_formula, Worksheet._write_string


def _wf(self, row, col, formula, cell_format=None, value=0):
    for a_, b_ in _REN_F:
        formula = formula.replace(a_, b_)
    return _wf0(self, row, col, formula, cell_format, value)


def _wstr(self, row, col, string, cell_format=None):
    for a_, b_ in _REN_T:
        string = string.replace(a_, b_)
    return _ws0(self, row, col, string, cell_format)


Worksheet._write_formula, Worksheet._write_string = _wf, _wstr
_sc0 = Worksheet.set_column


def _setcol(self, *args, **kw):
    """Roboto 10 (template) is wider than Roboto Light 9: widen every column but the Cover's by 15%."""
    args = list(args)
    i = 1 if isinstance(args[0], str) else 2
    if self.name != "Cover" and len(args) > i and isinstance(args[i], (int, float)) and args[i] > 3:
        args[i] = round(args[i] * 1.15, 1)
    return _sc0(self, *args, **kw)


Worksheet.set_column = _setcol

wb = xlsxwriter.Workbook(str(OUT))
wb.set_size(1700, 1050)
_fc = {}


def F(**k):
    """Cached format: Roboto 10, as in the XP model template."""
    key = tuple(sorted(k.items()))
    if key not in _fc:
        p_ = {"font_name": "Roboto", "font_size": 10, "valign": "vcenter", "font_color": INK}
        p_.update(k)
        _fc[key] = wb.add_format(p_)
    return _fc[key]


IN = dict(font_color=INBLUE)                                              # hardcoded input (template: blue 0070C0)
LV = dict(font_color=INBLUE, bold=True, bg_color=PINK, border=4)          # assumption to change (template: pink, dotted)
LK = dict()                                                               # links are black, as in the template
HDR = F(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", text_wrap=True)
HDRL = F(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="left", text_wrap=True)
SUB = F(bold=True, bg_color=GBAND, bottom=7)
TXT = F()
TXTW = F(text_wrap=True, valign="top")
NOTEF = F(italic=True, font_color=NOTE, text_wrap=True, valign="top")
NOTE1 = F(italic=True, font_color=NOTE)
BOLD = F(bold=True)
NA = F(font_color=GREY, align="right")
GROUPH = dict(bold=True, bg_color=GBAND, align="center", bottom=7)       # band over a group of columns

ORDER = ["Cover", "Summary", "Assumptions", "Scenarios", "Sensitivity", "Support >", "Boxes_Now", "Rates_EPS",
         "BottomUp", "History", "Charts", "Members", "Panel", "Checks", "README"]
SHEETS = [s for s in ORDER if s not in ("Cover", "Support >")]
TAB = {"Assumptions": YEL, "Support >": YEL}
FOOTER = '&R&"Calibri"&10&K008000[ CLASSIFICAÇÃO: PÚBLICA ]'
W = {}
for s in ORDER:
    W[s] = wb.add_worksheet(DISPLAY.get(s, s))
    if s in TAB:
        W[s].set_tab_color(TAB[s])
    W[s].hide_gridlines(2)
    W[s].set_default_row(14)
    W[s].set_zoom(85)
    if s not in ("Cover", "Support >"):
        W[s].set_header("&L&G", {"image_left": str(LOGO_BLACK)})
        W[s].set_footer(FOOTER)
        W[s].set_margins(top=1.25)


def title(ws, text, sub=None, legend=True):
    ws.set_row(0, 30)
    ws.write("B1", "Ibovespa - " + text, F(bold=True, font_size=18, font_color=TITLE_C))
    if sub:
        ws.write("B2", sub, NOTE1)
    if legend:
        ws.write("B3", "Legend:", NOTE1)
        ws.write("C3", "input", F(**IN, italic=True))
        ws.write("D3", "assumption", F(**LV, italic=True))
        ws.write("E3", "formula", F(italic=True))


def section(ws, row, text, c0="B", c1="H", fill=NAVY, color="#FFFFFF"):
    ws.merge_range(f"{c0}{row}:{c1}{row}", text, F(bold=True, font_color=color, bg_color=fill))


def zebra(ws, rng):
    """The XP template has no banding; kept as a no-op so the call sites stay readable."""
    return None


def nz(x):
    return None if (x is None or (isinstance(x, float) and np.isnan(x))) else x


# =========================================================================== Assumptions
ws = W["Assumptions"]
title(ws, "Assumptions", "The only sheet to edit: change the pink cells (dotted border) and every other sheet recalculates.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 58)
ws.set_column("C:G", 13)
ws.set_column("H:H", 95)
A = {}
NAMES = {}


def name(nm, ref):
    NAMES[nm] = ref
    wb.define_name(nm, f"=Assumptions!{ref}")


section(ws, 5, "1. Market snapshot (data as of the pricing date)")
snap = [("Pricing date", "PxDate", pd.Timestamp("2026-10-05"), DATED, "Close of Oct-5-2026. Prices, 12m fwd multiples and market caps: Bloomberg."),
        ("Ibovespa close (index pts)", "IbovNow", d["ibov_now"], PX, "Bloomberg IBOV Index PX_LAST."),
        ("Selic target today", "SelicNow", 0.1375, PCT2, "BCB (BZSTSETA Index)."),
        ("5-year real rate today (NTN-B)", "RRNow", 0.0676, PCT2, "ANBIMA, Oct-5. Used only for the risk premium (History, Charts)."),
        ("Ibovespa membership date", "MemDate", pd.Timestamp("2026-09-30"), DATED, "Members and weights: Economatica Ibovespa composition on this date, drifted by price to the pricing date (Members)."),
        ("Ibovespa close on the membership date (index pts)", "IbovMem", d["ibov_sep30"], PX, "Bloomberg. Used only in Checks (price-drift test).")]
for i, (lab, nm, v, f, note) in enumerate(snap):
    r = 6 + i
    ws.write(f"B{r}", lab, TXT)
    if isinstance(v, pd.Timestamp):
        ws.write_datetime(f"C{r}", v.to_pydatetime(), F(**IN, num_format=f))
    else:
        ws.write_number(f"C{r}", v, F(**IN, num_format=f))
    ws.write(f"H{r}", note, NOTE1)
    name(nm, f"$C${r}")

section(ws, 13, "2. Scenario levers")
for c, h in zip("BCDEFGH", ["Lever", "Today", "Bear", "Base", "Bull", "Custom", "How it is used"]):
    ws.write(f"{c}14", h, HDR if c not in "BH" else HDRL)
A["narr"], A["selic"], A["cut"], A["k"], A["ke"] = 15, 16, 17, 18, 19
lev = [("Narrative", ["Moderate populist", "Base", "Reformist", "User-defined"], None,
        "Labels only."),
       ("Selic at the end of the horizon", [0.132, 0.115, 0.095, 0.115], PCT2,
        "Drives the rate-driven change in EPS and EBITDA (Rates_EPS). The Oct-5 LTN curve already prices ~11.4–11.9% at end-2027, so the base ≈ today's curve."),
       (None, None, None, None),
       ("Multiples: std. dev. from the history average (k)", [-1, 0, 1, 0], '+0.0;-0.0;0.0',
        "Target 12m fwd P/E and EV/EBITDA of each box = history average + k × std. dev. (History, window in section 3)."),
       ("Bottom-up: shift in the cost of equity, Ke (bp)", [200, 0, -200, 0], BP,
        "Each XP target price is repriced with the analyst's sensitivity of the TP to Ke (BottomUp). +200bp ≈ 0.8σ of 12m moves in the 5y pré.")]
for i, (lab, vals, f, note) in enumerate(lev):
    r = 15 + i
    if lab is None:
        continue
    ws.write(f"B{r}", lab, TXT)
    for c, v in zip(SC, vals):
        if isinstance(v, str):
            ws.write_string(f"{c}{r}", v, F(**LV, align="center"))
        else:
            ws.write_number(f"{c}{r}", v, F(**LV, num_format=f))
    ws.write(f"H{r}", note, NOTE1)
ws.write("C16", "=SelicNow", F(num_format=PCT2))
ws.write("B17", "Selic cut vs today (pp; positive = cut)", TXT)
for c in SC:
    ws.write_formula(f"{c}17", f"=(SelicNow-{c}16)*100", F(num_format=PPC))
ws.write("H17", "Formula: (Selic today − Selic in the scenario) × 100.", NOTE1)
for c in SC:
    ws.data_validation(f"{c}18", {"validate": "decimal", "criteria": "between", "minimum": -3, "maximum": 3})
wb.define_name("BaseSelic", "=Assumptions!$E$16")
wb.define_name("BaseK", "=Assumptions!$E$18")

section(ws, 21, "3. Model settings")
for c, h in zip("BCDEFGH", ["Setting", "Value", "", "", "", "", "Notes"]):
    ws.write(f"{c}22", h, HDR if c == "C" else HDRL)
sets = [("History window: first month", "WinStart", pd.Timestamp("2016-10-31"), DATE,
         "Average and std. dev. of the multiples use the months from first to last (inclusive; any day of the month counts the whole month). Oct-16..Sep-26 = 10 years. History has data from Oct-06."),
        ("History window: last month", "WinEnd", pd.Timestamp("2026-09-30"), DATE, ""),
        ("Ibovespa EPS change per −100bp of Selic", "EPSsens", 0.0207, PCT2,
         "Regression of the 12m change in Ibovespa 12m fwd EPS on the 12m change in Selic (6m lag, controls) — see Rates_EPS."),
        ("EPS change split by box return beta (1 = yes, 0 = same % for all)", "SplitByBeta", 1, "0",
         "1: each box gets the Ibovespa sensitivity × its stock-return beta to the 2y pré relative to the Ibovespa, scaled so the boxes add up to the Ibovespa."),
        ("Bottom-up: cap on the TP change per stock (±)", "KeCap", 0.40, "0%",
         "Perpetuity convexity blows up as k − g approaches the Ke shift; the cap binds for WEGE3 and CSNA3 at ±200bp."),
        ("Bottom-up: haircut on the target prices", "BUhaircut", 0.0, "0%",
         "0% = 12m XP target prices as published. The house YE26 model uses 15% (TP ÷ 1.15)."),
        ("Method weight in the average: P/E", "wPE", 1, "0.0", ""),
        ("Method weight in the average: EV/EBITDA", "wEV", 1, "0.0", "Not applied to Financials (EV/EBITDA does not apply to banks)."),
        ("Method weight in the average: bottom-up", "wBU", 1, "0.0", ""),
        ("P/E: lowest valid value", "PEmin", 1.0, MULT, "Members outside the range are left out of the box P/E (now and history)."),
        ("P/E: highest valid value", "PEmax", 100.0, MULT, ""),
        ("EV/EBITDA: lowest valid value", "EVmin", 1.0, MULT, ""),
        ("EV/EBITDA: highest valid value", "EVmax", 50.0, MULT, ""),
        ("History: drop stale Bloomberg repeats (1 = yes, 0 = no)", "DropStale", 0, "0",
         "Bloomberg's monthly history repeats the last estimate when there is none (NATU3 10.39x for 44 months, GOAU4, BHIA3). 1 keeps only the first month of a run of 3+ identical values. 0 reproduces the Oct-6 numbers.")]
for i, (lab, nm, v, f, note) in enumerate(sets):
    r = 23 + i
    ws.write(f"B{r}", lab, TXT)
    if isinstance(v, pd.Timestamp):
        ws.write_datetime(f"C{r}", v.to_pydatetime(), F(**LV, num_format=f))
    else:
        ws.write_number(f"C{r}", v, F(**LV, num_format=f))
    ws.write(f"H{r}", note, NOTE1)
    name(nm, f"$C${r}")
for nm in ("SplitByBeta", "DropStale"):
    ws.data_validation(NAMES[nm].replace("$", ""), {"validate": "list", "source": [0, 1]})

section(ws, 38, "4. Optional overrides (leave blank to use the model rule)")
ws.write("B39", "A number typed here replaces the model value for that box and scenario. Checks counts the overrides in use.", NOTE1)
for c, h in zip("BCDEFGH", ["Box", "", "Bear", "Base", "Bull", "Custom", ""]):
    ws.write(f"{c}40", h, HDR if c not in "BH" else HDRL)
OVR = {}
for blk, (lab, key, f) in enumerate([("Target 12m fwd P/E (x)", "pe", MULT), ("Target 12m fwd EV/EBITDA (x)", "ev", MULT),
                                      ("EPS / EBITDA change vs today (%)", "eps", UPS)]):
    r0 = 41 + blk * 7
    ws.write(f"B{r0}", lab, SUB)
    for i, g in enumerate(G6):
        r = r0 + 1 + i
        OVR[(key, g)] = r
        ws.write(f"B{r}", ("   " if g in MEMO else "") + g + (" (memo)" if g in MEMO else ""), TXT)
        for c in SC:
            if key == "ev" and g == "Financials":
                ws.write(f"{c}{r}", "n.a.", NA)
            else:
                ws.write_blank(f"{c}{r}", None, F(**LV, num_format=f))
A["ovr_first"], A["ovr_last"] = 42, 61

section(ws, 63, "5. Classification")
ws.write("B64", "XP-adjusted GICS sector", HDRL)
ws.write("C64", "Box", HDR)
for i, (gs, bx) in enumerate(GICS_MAP):
    ws.write(f"B{65 + i}", gs, F(**IN))
    ws.write(f"C{65 + i}", bx, F(**IN))
ws.write("H65", "Communication Services sits in Defensives (Caio, Oct-5). Sector = the XP-adjusted GICS column of the house sector file.", NOTE1)
name("GicsMap", "$B$65:$C$75")
ws.write("B77", "Global industrials: kept in Cyclicals, shown apart as a memo", HDRL)
for i, t in enumerate(GLOBAL_IND):
    ws.write(f"B{78 + i}", t, F(**IN))
ws.write_blank("B80", None, F(**LV))
ws.write_blank("B81", None, F(**LV))
ws.write("H78", "WEG and Embraer trade like exporters (positive USDBRL beta, lowest market betas of the box). Two spare rows to add tickers.", NOTE1)
name("GlobalList", "$B$78:$B$81")
ws.freeze_panes(4, 0)

# =========================================================================== Members
ws = W["Members"]
MF, ML = 5, 5 + len(m) - 1          # data rows
MT = ML + 2                          # totals row
title(ws, "Members: the Ibovespa constituents on the pricing date",
      "One row per member. Inputs in blue (Bloomberg, Economatica, XP comp sheet); every other column is a formula. Box aggregates live in Boxes_Now.")
cols = [  # letter, header, width, kind
    ("A", "Ticker", 8), ("B", "Company", 18), ("C", "XP-adjusted GICS sector", 20), ("D", "Box", 12), ("E", "Global industrial (1/0)", 9),
    ("F", "Sub-box", 22), ("G", "Ibovespa weight on the membership date", 10), ("H", "Price on the membership date (R$)", 10),
    ("I", "Price on the pricing date (R$)", 10), ("J", "Price-drifted weight (unscaled)", 10), ("K", "Weight on the pricing date", 9),
    ("L", "Index points", 10), ("M", "12m fwd P/E (x)", 8), ("N", "P/E used (1/0)", 7), ("O", "Weight in P/E", 9),
    ("P", "Weight ÷ P/E", 9), ("Q", "12m fwd EV/EBITDA (x)", 9), ("R", "EV/EBITDA used (1/0)", 8), ("S", "Weight in EV/EBITDA", 9),
    ("T", "Weight ÷ EV/EBITDA", 9), ("U", "Market cap (R$ mn)", 11), ("V", "Enterprise value (R$ mn)", 11), ("W", "EV ÷ market cap", 8),
    ("X", "Weight with EV/mkt cap", 9), ("Y", "Weight × EV/mkt cap", 9), ("Z", "XP target price (R$)", 9), ("AA", "Target price source", 20),
    ("AB", "Target price date", 10), ("AC", "XP rating", 10), ("AD", "Upside to target", 9), ("AE", "Data flags", 26)]
groups = [("A", "C", "Identification"), ("D", "F", "Classification (Assumptions §5)"), ("G", "L", "Weight and index points"),
          ("M", "P", "P/E"), ("Q", "Y", "EV/EBITDA and net debt"), ("Z", "AD", "Bottom-up (XP targets)"), ("AE", "AE", "")]
for a, b, t in groups:
    if a == b:
        ws.write(f"{a}4", t, HDR)
    else:
        ws.merge_range(f"{a}4:{b}4", t, F(**GROUPH))
ws.set_row(4, 52)
for c, h, wdt in cols:
    ws.write(f"{c}5", h, HDR)
    ws.set_column(f"{c}:{c}", wdt)
MF, ML = 6, 6 + len(m) - 1
MT = ML + 2
rng = lambda c: f"Members!${c}${MF}:${c}${ML}"
for i, (t, x) in enumerate(m.iterrows()):
    r = MF + i
    ws.write_string(f"A{r}", t, F(**IN, bold=True))
    ws.write_string(f"B{r}", x["name"], F(**IN))
    ws.write_string(f"C{r}", x["gics"], F(**IN))
    ws.write_formula(f"D{r}", f"=VLOOKUP(C{r},GicsMap,2,FALSE)", F(**LK))
    ws.write_formula(f"E{r}", f"=IF(COUNTIF(GlobalList,A{r})>0,1,0)", F(**LK, align="center"))
    ws.write_formula(f"F{r}", f'=IF(D{r}="Cyclicals",IF(E{r}=1,"WEG & Embraer","Cyclicals ex-WEG & Embraer"),D{r})', TXT)
    ws.write_number(f"G{r}", x["w30"], F(**IN, num_format="0.00%"))
    ws.write_number(f"H{r}", x["px30"], F(**IN, num_format=PX))
    ws.write_number(f"I{r}", x["px_now"], F(**IN, num_format=PX))
    ws.write_formula(f"J{r}", f"=G{r}*I{r}/H{r}", F(num_format="0.000%"))
    ws.write_formula(f"K{r}", f"=J{r}/SUM($J${MF}:$J${ML})", F(num_format="0.00%"))
    ws.write_formula(f"L{r}", f"=K{r}*IbovNow", F(num_format=PTS))
    if nz(x["pebf"]) is not None:
        ws.write_number(f"M{r}", x["pebf"], F(**IN, num_format=MULT))
    ws.write_formula(f"N{r}", f"=IF(AND(ISNUMBER(M{r}),M{r}>=PEmin,M{r}<=PEmax),1,0)", F(align="center"))
    ws.write_formula(f"O{r}", f"=N{r}*K{r}", F(num_format="0.00%"))
    ws.write_formula(f"P{r}", f"=IF(N{r}=1,K{r}/M{r},0)", F(num_format="0.0000"))
    if x["box"] != "Financials" and nz(x["evbf"]) is not None:
        ws.write_number(f"Q{r}", x["evbf"], F(**IN, num_format=MULT))
    ws.write_formula(f"R{r}", f'=IF(AND(D{r}<>"Financials",ISNUMBER(Q{r}),Q{r}>=EVmin,Q{r}<=EVmax),1,0)', F(align="center"))
    ws.write_formula(f"S{r}", f"=R{r}*K{r}", F(num_format="0.00%"))
    ws.write_formula(f"T{r}", f"=IF(R{r}=1,K{r}/Q{r},0)", F(num_format="0.0000"))
    if nz(x["mcap"]) is not None:
        ws.write_number(f"U{r}", x["mcap"], F(**IN, num_format=MN))
    if nz(x["ev_val"]) is not None:
        ws.write_number(f"V{r}", x["ev_val"], F(**IN, num_format=MN))
    ws.write_formula(f"W{r}", f'=IF(AND(D{r}<>"Financials",ISNUMBER(U{r}),ISNUMBER(V{r})),V{r}/U{r},"")', F(num_format="0.00"))
    ws.write_formula(f"X{r}", f"=IF(ISNUMBER(W{r}),K{r},0)", F(num_format="0.00%"))
    ws.write_formula(f"Y{r}", f"=IF(ISNUMBER(W{r}),K{r}*W{r},0)", F(num_format="0.0000"))
    ws.write_number(f"Z{r}", x["tp_used"], F(**IN, num_format=PX))
    ws.write_string(f"AA{r}", x["tp_src"], F(**IN))
    if isinstance(x.get("tp_pdate"), str) and x["tp_pdate"]:
        ws.write_datetime(f"AB{r}", pd.Timestamp(x["tp_pdate"]).to_pydatetime(), F(**IN, num_format=DATED))
    if isinstance(x.get("rec"), str):
        ws.write_string(f"AC{r}", x["rec"], F(**IN))
    ws.write_formula(f"AD{r}", f'=IF(AND(ISNUMBER(Z{r}),I{r}>0),Z{r}/I{r}-1,"")', F(num_format=UPS))
    ws.write_formula(f"AE{r}", f'=TRIM(IF(N{r}=0,"No valid P/E. ","")&IF(AND(D{r}<>"Financials",R{r}=0),"No valid EV/EBITDA. ","")&IF(AD{r}="","No target. ",""))', F(font_color=GREY))
ws.write(f"A{MT}", "Total / check", BOLD)
for c, f in (("G", "0.00%"), ("K", "0.00%"), ("L", PTS), ("O", "0.00%"), ("S", "0.00%")):
    ws.write_formula(f"{c}{MT}", f"=SUM({c}{MF}:{c}{ML})", F(bold=True, num_format=f, top=1))
ws.write(f"A{MT + 2}", "Notes", BOLD)
notes = ["Weight on the pricing date = membership weight × price change since the membership date, rescaled to 100%. Index points = weight × Ibovespa close.",
         "P/E and EV/EBITDA: Bloomberg BEST_PE_RATIO and BEST_EV_TO_BEST_EBITDA, 12m blended forward (BEST_FPERIOD_OVERRIDE = BF), live price on the pricing date. Box multiples are harmonic means: Σ weight ÷ Σ (weight ÷ multiple) over the members with a valid value (Boxes_Now).",
         "EV ÷ market cap (non-financials) turns the box market value into enterprise value: net debt (index pts) = index points × (EV/mkt cap − 1). Financials have no EV/EBITDA.",
         "Target prices: XP comp sheet (Oct-2); Bloomberg consensus (BEST_TARGET_PRICE, Oct-5) where XP has no coverage. Upside to target is the bottom-up input (BottomUp)."]
for i, t in enumerate(notes):
    ws.write(f"A{MT + 3 + i}", t, NOTE1)
zebra(ws, f"A{MF}:AE{ML}")
ws.freeze_panes(5, 2)
ws.autofilter(f"A5:AE{ML}")

# =========================================================================== BottomUp
ws = W["BottomUp"]
title(ws, "Bottom-up: XP target prices, repriced for a shift in the cost of equity",
      "Per member: the analyst's TP change for Ke −100bp gives the implied perpetuity spread k − g; the TP is repriced for the scenario's Ke shift (Assumptions §2), capped at ±KeCap.")
KE_LIST = list(range(-300, 301, 50))
bcols = [("A", "Ticker", 8), ("B", "Box", 12), ("C", "Sub-box", 22), ("D", "Weight", 8), ("E", "Upside to XP target", 9),
         ("F", "Has target (1/0)", 7), ("G", "Analyst: TP change for Ke −100bp", 10), ("H", "Source of the sensitivity", 14),
         ("I", "Weight × sensitivity (provided)", 10), ("J", "Weight with sensitivity", 9), ("K", "Sub-box weighted-average sensitivity", 11),
         ("L", "Sensitivity used", 9), ("M", "Filled from", 14), ("N", "Implied k − g", 8), ("O", "Weight with target", 9),
         ("P", "Weight × (1 + upside)", 9)]
for j, s in enumerate(SCEN):
    bcols.append((cn(16 + j), f"TP change: {s}", 9))
for j, s in enumerate(SCEN):
    bcols.append((cn(20 + j), f"Weight × (1+upside) × (1+TP change): {s}", 11))
KEG0 = 25   # column Z: helper for the Sensitivity grid
for j in range(len(KE_LIST)):
    bcols.append((cn(KEG0 + j), "", 9))
ws.merge_range("A4:F4", "Member (links to Members)", F(**GROUPH))
ws.merge_range("G4:N4", "Analyst sensitivity to Ke and implied k − g", F(**GROUPH))
ws.merge_range("O4:X4", "Scenario repricing (Ke shift from Assumptions row 19)", F(**GROUPH))
ws.merge_range(f"{cn(KEG0)}4:{cn(KEG0 + len(KE_LIST) - 1)}4", "Helper for Sensitivity grid 3: weight × (1+upside) × (1+TP change) at each Ke shift (bp) in row 5",
               F(**GROUPH))
ws.set_row(4, 52)
for c, h, wdt in bcols:
    ws.write(f"{c}5", h, HDR)
    ws.set_column(f"{c}:{c}", wdt)
BF, BL = MF, ML       # same rows as Members
SEN_KE_ROW0 = 53      # Sensitivity: first row of grid 3 (Ke list)
for j in range(len(KE_LIST)):
    ws.write_formula(f"{cn(KEG0 + j)}5", f"=Sensitivity!$B${SEN_KE_ROW0 + j}", F(bold=True, font_color="#FFFFFF", bg_color=NAVY, num_format=BP, align="center"))
ws.set_column("Y:Y", 2)


def tpchg(r, shift):
    return (f'IF($N{r}="",0,IF($N{r}+{shift}/10000<=0,KeCap,MAX(-KeCap,MIN(KeCap,$N{r}/($N{r}+{shift}/10000)-1))))')


for i, (t, x) in enumerate(m.iterrows()):
    r = BF + i
    ws.write_formula(f"A{r}", f"=Members!A{r}", F(**LK, bold=True))
    ws.write_formula(f"B{r}", f"=Members!D{r}", F(**LK))
    ws.write_formula(f"C{r}", f"=Members!F{r}", F(**LK))
    ws.write_formula(f"D{r}", f"=Members!K{r}", F(**LK, num_format="0.00%"))
    ws.write_formula(f"E{r}", f"=Members!AD{r}", F(**LK, num_format=UPS))
    ws.write_formula(f"F{r}", f"=IF(ISNUMBER(E{r}),1,0)", F(align="center"))
    if nz(x["ke_sens"]) is not None:
        ws.write_number(f"G{r}", x["ke_sens"], F(**IN, num_format=UPS))
    if isinstance(x.get("ke_src"), str):
        ws.write_string(f"H{r}", x["ke_src"], F(**IN))
    ws.write_formula(f"I{r}", f"=IF(ISNUMBER(G{r}),D{r}*G{r},0)", F(num_format="0.0000"))
    ws.write_formula(f"J{r}", f"=IF(ISNUMBER(G{r}),D{r},0)", F(num_format="0.00%"))
    ws.write_formula(f"K{r}", f'=IFERROR(SUMIFS($I${BF}:$I${BL},$C${BF}:$C${BL},C{r})/SUMIFS($J${BF}:$J${BL},$C${BF}:$C${BL},C{r}),"")', F(num_format=UPS))
    ws.write_formula(f"L{r}", f"=IF(ISNUMBER(G{r}),G{r},K{r})", F(num_format=UPS))
    ws.write_formula(f"M{r}", f'=IF(ISNUMBER(G{r}),"Analyst","Sub-box average")', TXT)
    ws.write_formula(f"N{r}", f'=IF(AND(ISNUMBER(L{r}),L{r}>0),0.01*(1+L{r})/L{r},"")', F(num_format=PCT2))
    ws.write_formula(f"O{r}", f"=D{r}*F{r}", F(num_format="0.00%"))
    ws.write_formula(f"P{r}", f"=IF(F{r}=1,D{r}*(1+E{r}),0)", F(num_format="0.0000"))
    for j, c in enumerate(SC):
        ws.write_formula(f"{cn(16 + j)}{r}", "=" + tpchg(r, f"Assumptions!{c}$19"), F(num_format=UPS))
        ws.write_formula(f"{cn(20 + j)}{r}", f"=IF($F{r}=1,$D{r}*(1+$E{r})*(1+{cn(16 + j)}{r}),0)", F(num_format="0.0000"))
    for j in range(len(KE_LIST)):
        c = cn(KEG0 + j)
        ws.write_formula(f"{c}{r}", f"=IF($F{r}=1,$D{r}*(1+$E{r})*(1+{tpchg(r, f'{c}$5')}),0)", F(num_format="0.0000", font_color=GREY))
BT = BL + 2
ws.write(f"A{BT}", "Total", BOLD)
for c in ["D", "J", "O"]:
    ws.write_formula(f"{c}{BT}", f"=SUM({c}{BF}:{c}{BL})", F(bold=True, num_format="0.00%", top=1))
ws.write(f"A{BT + 2}", "How the repricing works", BOLD)
for i, t in enumerate([
        "1. The analyst reports the % change in the target price for Ke −100bp (s). In a perpetuity, P ∝ 1 ÷ (k − g), so s = (k−g) ÷ (k−g − 1%) − 1, which gives k − g = 1% × (1 + s) ÷ s (column N).",
        "2. For a Ke shift of x bp, the TP changes by (k−g) ÷ (k−g + x) − 1, capped at ±KeCap. If k − g + x ≤ 0 the change is set to +KeCap.",
        "3. Members without an analyst sensitivity take the weighted average of their sub-box (column K): real estate and five Financials/Defensives names on Oct-6.",
        "4. Box bottom-up upside = Σ weight × (1 + upside) × (1 + TP change) ÷ Σ weight − 1, over members with a target (used in Scenarios and Boxes_Now).",
        "Source of the sensitivities: Sensitivity.xlsx in the research bucket (Oct-6) and the Financials/TMT table sent on Oct-6; PETR3 = PETR4, ITSA4 = ITUB4, BBDC3 = BBDC4."]):
    ws.write(f"A{BT + 3 + i}", t, NOTE1)
zebra(ws, f"A{BF}:X{BL}")
ws.freeze_panes(5, 1)

# =========================================================================== Panel
ws = W["Panel"]
ws.hide_gridlines(0)
title(ws, "Panel: monthly history of the members (Bloomberg, Economatica weights)",
      "One row per member and month, sorted by ticker and date. Inputs in blue; flags and weights are formulas. History aggregates this sheet with SUMIFS.", legend=False)
pcols = [("A", "Month-end", 9), ("B", "Ticker", 8), ("C", "XP-adjusted GICS sector", 20), ("D", "Box", 12), ("E", "Sub-box", 22),
         ("F", "Ibovespa weight", 9), ("G", "12m fwd P/E (x)", 8), ("H", "12m fwd EV/EBITDA (x)", 9),
         ("I", "Next month of the same stock (1/0)", 9), ("J", "Stale P/E repeat (1/0)", 8), ("K", "Stale EV/EBITDA repeat (1/0)", 9),
         ("L", "P/E used (1/0)", 7), ("M", "Weight in P/E", 9), ("N", "Weight ÷ P/E", 9), ("O", "EV/EBITDA used (1/0)", 8),
         ("P", "Weight in EV/EBITDA", 9), ("Q", "Weight ÷ EV/EBITDA", 9)]
ws.set_row(3, 40)
for c, h, wdt in pcols:
    ws.write(f"{c}4", h, HDR)
    ws.set_column(f"{c}:{c}", wdt)
PF = 5
PL = PF + len(p) - 1
fin_ = F(num_format=DATE, font_color=INBLUE)
fpe = F(num_format="0.00", font_color=INBLUE)
fw = F(num_format="0.000%", font_color=INBLUE)
f0 = F(align="center")
fwt = F(num_format="0.000%")
f4 = F(num_format="0.00000")
for i, x in enumerate(p.itertuples(index=False)):
    r = PF + i
    ws.write_datetime(r - 1, 0, pd.Timestamp(x.date).to_pydatetime(), fin_)
    ws.write_string(r - 1, 1, x.cod_ativo, F(**IN))
    ws.write_string(r - 1, 2, x.gics, F(**IN))
    ws.write_formula(r - 1, 3, f"=VLOOKUP(C{r},GicsMap,2,FALSE)", F(**LK))
    ws.write_formula(r - 1, 4, f'=IF(D{r}="Cyclicals",IF(COUNTIF(GlobalList,B{r})>0,"WEG & Embraer","Cyclicals ex-WEG & Embraer"),D{r})', TXT)
    ws.write_number(r - 1, 5, x.w, fw)
    if not np.isnan(x.pe):
        ws.write_number(r - 1, 6, x.pe, fpe)
    if x.box != "Financials" and not np.isnan(x.evx):
        ws.write_number(r - 1, 7, x.evx, fpe)
    if i == 0:
        ws.write_formula(r - 1, 8, "=0", f0)
    else:
        ws.write_formula(r - 1, 8, f"=IF(AND(B{r}=B{r - 1},A{r}=EOMONTH(A{r - 1},1)),1,0)", f0)
    for col, v in ((9, "G"), (10, "H")):
        ws.write_formula(r - 1, col, f"=IF(AND(ISNUMBER({v}{r}),I{r}=1,{v}{r}={v}{r - 1}),IF(OR(AND(I{r - 1}=1,{v}{r - 1}={v}{r - 2}),AND(I{r + 1}=1,{v}{r + 1}={v}{r})),1,0),0)", f0)
    ws.write_formula(r - 1, 11, f"=IF(AND(ISNUMBER(G{r}),G{r}>=PEmin,G{r}<=PEmax,OR(DropStale=0,J{r}=0)),1,0)", f0)
    ws.write_formula(r - 1, 12, f"=L{r}*F{r}", fwt)
    ws.write_formula(r - 1, 13, f"=IF(L{r}=1,F{r}/G{r},0)", f4)
    ws.write_formula(r - 1, 14, f'=IF(AND(D{r}<>"Financials",ISNUMBER(H{r}),H{r}>=EVmin,H{r}<=EVmax,OR(DropStale=0,K{r}=0)),1,0)', f0)
    ws.write_formula(r - 1, 15, f"=O{r}*F{r}", fwt)
    ws.write_formula(r - 1, 16, f"=IF(O{r}=1,F{r}/H{r},0)", f4)
ws.freeze_panes(4, 2)
ws.autofilter(f"A4:Q{PL}")
PR = lambda c: f"Panel!${c}${PF}:${c}${PL}"

# =========================================================================== Boxes_Now
ws = W["Boxes_Now"]
title(ws, "Boxes now: current size, multiples, earnings and net debt of each box",
      "Aggregates Members with SUMIFS on the Box and Sub-box criteria in columns C:D. Multiples are weighted harmonic means; EPS, EBITDA and net debt are in Ibovespa points.")
BN = {"Financials": 6, "Defensives": 7, "Cyclicals": 8, "Commodities": 9, "Ibovespa": 10,
      "Cyclicals ex-WEG & Embraer": 12, "WEG & Embraer": 13, "Ibovespa ex-Financials": 14}
CRIT = {"Financials": ("Financials", "*"), "Defensives": ("Defensives", "*"), "Cyclicals": ("Cyclicals", "*"),
        "Commodities": ("Commodities", "*"), "Ibovespa": ("*", "*"), "Cyclicals ex-WEG & Embraer": ("Cyclicals", "Cyclicals ex-WEG & Embraer"),
        "WEG & Embraer": ("Cyclicals", "WEG & Embraer"), "Ibovespa ex-Financials": ("<>Financials", "*")}
bnh = [("B", "Group", 26), ("C", "Box criterion", 12), ("D", "Sub-box criterion", 22), ("E", "Members", 7), ("F", "Weight", 8),
       ("G", "Index points", 10), ("H", "12m fwd P/E (x)", 8), ("I", "P/E coverage (% of weight)", 9), ("J", "EPS (index pts)", 9),
       ("K", "12m fwd EV/EBITDA (x)", 9), ("L", "EV/EBITDA coverage", 9), ("M", "EV ÷ market cap", 8), ("N", "Enterprise value (index pts)", 10),
       ("O", "Net debt (index pts)", 9), ("P", "EBITDA (index pts)", 9), ("Q", "Upside to XP targets (weighted)", 9),
       ("R", "Target coverage", 8), ("S", "Earnings yield", 8), ("T", "Risk premium (EY − 5y real rate)", 10)]
ws.set_column("A:A", 2)
ws.set_row(4, 52)
for c, h, wdt in bnh:
    ws.write(f"{c}5", h, HDR if c != "B" else HDRL)
    ws.set_column(f"{c}:{c}", wdt)
ws.write("B11", "Memo (not added to the Ibovespa)", NOTE1)


def sif(col, r, sheet="Members", f=MF, l=ML, boxc="D", subc="F"):
    return f"SUMIFS({sheet}!${col}${f}:${col}${l},{sheet}!${boxc}${f}:${boxc}${l},$C{r},{sheet}!${subc}${f}:${subc}${l},$D{r})"


for g, r in BN.items():
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{r}", g, F(**b) if tot else TXT)
    ws.write_string(f"C{r}", CRIT[g][0], F(**IN, **b))
    ws.write_string(f"D{r}", CRIT[g][1], F(**IN, **b))
    ws.write_formula(f"E{r}", f"=COUNTIFS(Members!$D${MF}:$D${ML},$C{r},Members!$F${MF}:$F${ML},$D{r})", F(num_format="0", **b))
    ws.write_formula(f"F{r}", "=" + sif("K", r), F(num_format=PCT, **b))
    ws.write_formula(f"G{r}", "=" + sif("L", r), F(num_format=PTS, **b))
    ws.write_formula(f"H{r}", f"={sif('O', r)}/{sif('P', r)}", F(num_format="0.00x" if False else '0.00"x"', **b))
    ws.write_formula(f"I{r}", f"={sif('O', r)}/F{r}", F(num_format=PCT, **b))
    ws.write_formula(f"J{r}", f"=G{r}/H{r}", F(num_format=PTS1, **b))
    noev = g in ("Financials", "Ibovespa")
    if noev:
        for c in "KLMNOP":
            ws.write(f"{c}{r}", "n.a.", F(font_color=GREY, align="right", **b))
    else:
        ws.write_formula(f"K{r}", f"={sif('S', r)}/{sif('T', r)}", F(num_format='0.00"x"', **b))
        ws.write_formula(f"L{r}", f"={sif('S', r)}/F{r}", F(num_format=PCT, **b))
        ws.write_formula(f"M{r}", f"={sif('Y', r)}/{sif('X', r)}", F(num_format="0.000", **b))
        ws.write_formula(f"N{r}", f"=G{r}*M{r}", F(num_format=PTS, **b))
        ws.write_formula(f"O{r}", f"=G{r}*(M{r}-1)", F(num_format=PTS, **b))
        ws.write_formula(f"P{r}", f"=N{r}/K{r}", F(num_format=PTS1, **b))
    ws.write_formula(f"Q{r}", f"={sif('P', r, 'BottomUp', BF, BL, 'B', 'C')}/{sif('O', r, 'BottomUp', BF, BL, 'B', 'C')}-1", F(num_format=UPS, **b))
    ws.write_formula(f"R{r}", f"={sif('O', r, 'BottomUp', BF, BL, 'B', 'C')}/F{r}", F(num_format=PCT, **b))
    ws.write_formula(f"S{r}", f"=1/H{r}", F(num_format=PCT, **b))
    ws.write_formula(f"T{r}", f"=S{r}-RRNow", F(num_format=PCT, **b))
ws.write("B16", "How to read", BOLD)
for i, t in enumerate([
        "Criteria: SUMIFS on Members!D (Box) and Members!F (Sub-box); \"*\" = any, \"<>Financials\" = all but Financials.",
        "12m fwd P/E = Σ weight ÷ Σ (weight ÷ P/E) over members with a valid P/E (the index-style harmonic mean). EPS (index pts) = index points ÷ P/E.",
        "EV/EBITDA: same harmonic mean over non-financials. EV (index pts) = index points × weighted EV/mkt cap; net debt = EV − index points; EBITDA = EV ÷ EV/EBITDA.",
        "Upside to XP targets = Σ weight × (1 + upside) ÷ Σ weight − 1 over members with a target (BottomUp). Risk premium = 1 ÷ P/E − 5y real rate (Assumptions).",
        "Ibovespa has no EV/EBITDA because Financials have none; use the 'Ibovespa ex-Financials' memo row."]):
    ws.write(f"B{17 + i}", t, NOTE1)
ws.write_comment("H5", "Harmonic mean = Σw ÷ Σ(w ÷ P/E). It is the P/E of a portfolio holding the members at index weights, so EPS in index points = points ÷ P/E adds up across boxes.",
                 {"x_scale": 2.2, "y_scale": 1.4})
ws.freeze_panes(5, 2)

# =========================================================================== History
ws = W["History"]
title(ws, "History: monthly 12m fwd multiples and risk premium of each box",
      "Rows 17+ aggregate Panel by month with SUMIFS (harmonic means). Rows 8–14: statistics over the window in Assumptions (only months with 'In window' = 1).")
HG_PE = ["Financials", "Defensives", "Cyclicals", "Commodities", "Ibovespa", "Cyclicals ex-WEG & Embraer", "WEG & Embraer"]
HG_EV = ["Defensives", "Cyclicals", "Commodities", "Ibovespa ex-Financials", "Cyclicals ex-WEG & Embraer", "WEG & Embraer"]
HC = {}
c = 3
for g in HG_PE:
    HC[("pe", g)] = cn(c); c += 1
for g in HG_EV:
    HC[("ev", g)] = cn(c); c += 1
for g in HG_PE:
    HC[("erp", g)] = cn(c); c += 1
LASTC = c - 1
HF, HLm = 17, 17 + 239          # monthly rows
HNOW = HLm + 1                  # pricing-date row
ST = {"avg": 8, "sd": 9, "p1": 10, "m1": 11, "n": 12, "now": 13, "z": 14}
ws.set_column("A:A", 10)
ws.set_column("B:B", 8)
ws.set_column("C:C", 9)
ws.set_column(f"D:{cn(LASTC)}", 10)
ws.merge_range(f"{HC[('pe', HG_PE[0])]}4:{HC[('pe', HG_PE[-1])]}4", "12m fwd P/E (x)", F(**GROUPH))
ws.merge_range(f"{HC[('ev', HG_EV[0])]}4:{HC[('ev', HG_EV[-1])]}4", "12m fwd EV/EBITDA (x) — non-financials", F(**GROUPH))
ws.merge_range(f"{HC[('erp', HG_PE[0])]}4:{HC[('erp', HG_PE[-1])]}4", "Risk premium = 12m fwd earnings yield − 5y real rate", F(**GROUPH))
ws.write("A5", "Group", HDRL); ws.write("A6", "Box criterion", NOTE1); ws.write("A7", "Sub-box criterion", NOTE1)
ws.set_row(4, 40)
lab = {"avg": "Average (window)", "sd": "Std. dev. (window)", "p1": "Average + 1 std. dev.", "m1": "Average − 1 std. dev.",
       "n": "Months in window", "now": "Now (pricing date)", "z": "Now vs average (std. dev.)"}
for k_, r in ST.items():
    ws.write(f"A{r}", lab[k_], BOLD if k_ in ("avg", "now") else TXT)
ws.write(f"A{HF - 1}", "Month-end", HDR); ws.write(f"B{HF - 1}", "In window (1/0)", HDR); ws.write(f"C{HF - 1}", "5y real rate (NTN-B)", HDR)
ws.set_row(HF - 2, 30)
crit_h = {"Ibovespa ex-Financials": ("<>Financials", "*")}
for (kind, g), col in HC.items():
    ws.write(f"{col}5", g, HDR)
    ws.write(f"{col}{HF - 1}", g, HDR)
    if kind in ("pe", "ev"):
        bc, sc = crit_h.get(g, CRIT.get(g))
        ws.write_string(f"{col}6", bc, F(**IN, align="center"))
        ws.write_string(f"{col}7", sc, F(**IN, align="center", font_size=8))
    else:
        ws.write_string(f"{col}6", "1 ÷ P/E", NOTE1)
        ws.write(f"{col}7", "− real rate", NOTE1)
    rng_ = f"{col}${HF}:{col}${HLm}"
    fmt = F(num_format=MULT) if kind != "erp" else F(num_format=PCT)
    fmtb = F(num_format=MULT, bold=True) if kind != "erp" else F(num_format=PCT, bold=True)
    ws.write_formula(f"{col}{ST['avg']}", f"=AVERAGEIFS({rng_},$B${HF}:$B${HLm},1)", fmtb)
    ws.write_formula(f"{col}{ST['n']}", f"=SUMPRODUCT($B${HF}:$B${HLm},--ISNUMBER({rng_}))", F(num_format="0"))
    ws.write_formula(f"{col}{ST['sd']}", f"=SQRT((SUMPRODUCT($B${HF}:$B${HLm},{rng_},{rng_})-{col}{ST['n']}*{col}{ST['avg']}^2)/({col}{ST['n']}-1))",
                     F(num_format="0.00" if kind != "erp" else PCT2))
    ws.write_formula(f"{col}{ST['p1']}", f"={col}{ST['avg']}+{col}{ST['sd']}", fmt)
    ws.write_formula(f"{col}{ST['m1']}", f"={col}{ST['avg']}-{col}{ST['sd']}", fmt)
    ws.write_formula(f"{col}{ST['now']}", f"={col}{HNOW}", fmtb)
    ws.write_formula(f"{col}{ST['z']}", f"=({col}{ST['now']}-{col}{ST['avg']})/{col}{ST['sd']}", F(num_format="+0.00;-0.00"))
fm_pe, fm_pct = F(num_format=MULT), F(num_format=PCT)
for i, dt in enumerate(months):
    r = HF + i
    ws.write_datetime(f"A{r}", pd.Timestamp(dt).to_pydatetime(), F(**IN, num_format=DATE))
    ws.write_formula(f"B{r}", f"=IF(AND(A{r}>EOMONTH(WinStart,-1),A{r}<=EOMONTH(WinEnd,0)),1,0)", F(align="center"))
    ws.write_number(f"C{r}", float(rr[pd.Timestamp(dt)]), F(**IN, num_format=PCT2))
    for (kind, g), col in HC.items():
        if kind == "pe":
            ws.write_formula(f"{col}{r}", f'=IFERROR(SUMIFS({PR("M")},{PR("A")},$A{r},{PR("D")},{col}$6,{PR("E")},{col}$7)/SUMIFS({PR("N")},{PR("A")},$A{r},{PR("D")},{col}$6,{PR("E")},{col}$7),"")', fm_pe)
        elif kind == "ev":
            ws.write_formula(f"{col}{r}", f'=IFERROR(SUMIFS({PR("P")},{PR("A")},$A{r},{PR("D")},{col}$6,{PR("E")},{col}$7)/SUMIFS({PR("Q")},{PR("A")},$A{r},{PR("D")},{col}$6,{PR("E")},{col}$7),"")', fm_pe)
        else:
            pc = HC[("pe", g)]
            ws.write_formula(f"{col}{r}", f'=IF(ISNUMBER({pc}{r}),1/{pc}{r}-$C{r},"")', fm_pct)
r = HNOW
ws.write_formula(f"A{r}", "=PxDate", F(**LK, num_format=DATE, bold=True))
ws.write_number(f"B{r}", 0, F(align="center"))
ws.write_formula(f"C{r}", "=RRNow", F(**LK, num_format=PCT2))
for (kind, g), col in HC.items():
    if kind == "pe":
        ws.write_formula(f"{col}{r}", f"=Boxes_Now!H{BN[g]}", F(**LK, num_format=MULT, bold=True))
    elif kind == "ev":
        ws.write_formula(f"{col}{r}", f"=Boxes_Now!K{BN[g]}", F(**LK, num_format=MULT, bold=True))
    else:
        ws.write_formula(f"{col}{r}", f"=1/{HC[('pe', g)]}{r}-$C{r}", F(num_format=PCT, bold=True))
ws.write(f"A{HNOW + 2}", "The last row is the pricing date (Boxes_Now, live prices); it is outside the window. Real rate: NTN-B 5y constant maturity (Bloomberg BZRFB5PY to Aug-21, Tesouro Direto spliced after).", NOTE1)
# chart helpers: average and ±1 std. dev. as constant columns (Charts reads them)
CH = []
for kind in ("pe", "erp"):
    for g in ["Ibovespa"] + BOXES:
        CH.append((kind, g))
HX0 = LASTC + 2
ws.merge_range(f"{cn(HX0)}4:{cn(HX0 + 3 * len(CH) - 1)}4", "Chart helpers (Charts sheet): window average and ±1 std. dev. repeated on every row",
               F(**GROUPH))
CHC = {}
for j, (kind, g) in enumerate(CH):
    src = HC[(kind, g)]
    for q, (stat, labq) in enumerate((("avg", "avg"), ("p1", "+1sd"), ("m1", "−1sd"))):
        col = cn(HX0 + 3 * j + q)
        CHC[(kind, g, stat)] = col
        ws.write(f"{col}{HF - 1}", f"{'P/E' if kind == 'pe' else 'ERP'} {g} {labq}", F(bold=True, font_color="#FFFFFF", bg_color=GREY, text_wrap=True, align="center"))
        for r in range(HF, HNOW + 1):
            ws.write_formula(f"{col}{r}", f"={src}${ST[stat]}", F(num_format=MULT if kind == "pe" else PCT, font_color=GREY))
ws.set_column(f"{cn(HX0)}:{cn(HX0 + 3 * len(CH) - 1)}", 8)
zebra(ws, f"A{HF}:{cn(LASTC)}{HLm}")
ws.freeze_panes(HF - 1, 1)

# =========================================================================== Rates_EPS
ws = W["Rates_EPS"]
title(ws, "Rates → earnings: EPS and EBITDA change by box in each Selic scenario",
      "Step 1 sets the Ibovespa sensitivity, step 2 splits it across boxes by their stock-return beta to rates, step 3 applies the Selic cut of each scenario.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 52)
ws.set_column("C:I", 13)
ws.set_column("J:J", 70)
section(ws, 5, "Step 1 — Ibovespa EPS sensitivity to the Selic", "B", "J")
s1 = [("Ibovespa EPS change per −100bp of Selic (used)", "=EPSsens", PCT2, LK, "Set in Assumptions §3 (rounded from the regression)."),
      ("Regression beta (% change in EPS per +1pp of Selic)", reg["beta"] / 100, PCT2, IN, "12m log change of the Ibovespa 12m fwd EPS index on the 12m change in Selic."),
      ("t-stat (Newey-West)", reg["t"], "0.00", IN, ""),
      ("R²", reg["r2"], "0.00", IN, ""),
      ("Observations (monthly)", reg["n"], "0", IN, ""),
      ("Specification", "Selic (BZSTSETA) 12m change lagged 6 months; controls: BCOM and USDBRL 12m changes; Newey-West s.e.", None, IN, ""),
      ("Source", "bx_eps_rates2.py, Oct-5-2026 (Equity Strategy). Only Defensives is robust at the box level, hence the split in step 2.", None, IN, "")]
for i, (lab_, v, f, kind, note) in enumerate(s1):
    r = 6 + i
    ws.write(f"B{r}", lab_, TXT)
    if isinstance(v, str) and v.startswith("="):
        ws.write_formula(f"C{r}", v, F(**kind, num_format=f, bold=True))
    elif isinstance(v, str):
        ws.write_string(f"C{r}", v, F(**kind))
    else:
        ws.write_number(f"C{r}", v, F(**kind, num_format=f))
    ws.write(f"J{r}", note, NOTE1)
section(ws, 14, "Step 2 — Split across boxes by stock-return beta to rates", "B", "J")
ws.write("B15", "Weekly box returns on the change in the 2y pré, since 2012, controls BCOM (DJP) and USDBRL (PTAX), Newey-West s.e. (perf_study.py, Oct-6). Return beta used as a proxy for each box's earnings sensitivity.", NOTE1)
h2 = ["Group", "Return beta to the 2y pré (% per −100bp)", "t-stat", "R²", "Beta relative to the Ibovespa", "EPS now (index pts)",
      "EPS × relative beta", "EPS sensitivity (% per −100bp Selic)"]
for c, h in zip("BCDEFGHI", h2):
    ws.write(f"{c}16", h, HDR if c != "B" else HDRL)
ws.set_row(15, 40)
RE = {"Financials": 17, "Defensives": 18, "Cyclicals": 19, "Commodities": 20, "Ibovespa": 21,
      "Cyclicals ex-WEG & Embraer": 23, "WEG & Embraer": 24}
ws.write("B22", "Memo", NOTE1)
for g, r in RE.items():
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    x = pb[f"abs|pre2y|{g}"]
    ws.write(f"B{r}", g, F(**b) if tot else TXT)
    ws.write_number(f"C{r}", x["beta"] / 100, F(**IN, num_format=PCT2, **b))
    ws.write_number(f"D{r}", x["t"], F(**IN, num_format="0.0", **b))
    ws.write_number(f"E{r}", x["r2"], F(**IN, num_format="0.00", **b))
    if tot:
        ws.write_number(f"F{r}", 1, F(num_format="0.00", **b))
        ws.write_formula(f"G{r}", f"=SUM(G17:G20)", F(num_format=PTS1, **b))
        ws.write_formula(f"H{r}", f"=SUM(H17:H20)", F(num_format=PTS1, **b))
        ws.write_formula(f"I{r}", "=EPSsens", F(**LK, num_format=PCT2, **b))
    else:
        ws.write_formula(f"F{r}", f"=IF(SplitByBeta=1,C{r}/$C$21,1)", F(num_format="0.00"))
        ws.write_formula(f"G{r}", f"=Boxes_Now!J{BN[g]}", F(**LK, num_format=PTS1))
        if g in BOXES:
            ws.write_formula(f"H{r}", f"=G{r}*F{r}", F(num_format=PTS1))
        ws.write_formula(f"I{r}", f"=F{r}*$C$26", F(num_format=PCT2, bold=True))
ws.write("B26", "Scale factor = Ibovespa sensitivity ÷ (EPS-weighted average relative beta)", TXT)
ws.write_formula("C26", "=EPSsens/(H21/G21)", F(num_format="0.0000"))
ws.write("J26", "Makes the EPS-weighted average of the four boxes equal the Ibovespa sensitivity.", NOTE1)
ws.write("B27", "Check: EPS-weighted average of the four boxes' sensitivities", TXT)
ws.write_formula("C27", "=SUMPRODUCT(G17:G20,I17:I20)/G21", F(num_format=PCT2))
section(ws, 29, "Step 3 — EPS change by scenario (also applied to EBITDA)", "B", "J")
for c, h in zip("BCDEFG", ["Group", "Sensitivity (% per −100bp)", "Bear", "Base", "Bull", "Custom"]):
    ws.write(f"{c}30", h, HDR if c != "B" else HDRL)
ws.write("B31", "Selic cut vs today (pp)", BOLD)
for c in SC:
    ws.write_formula(f"{c}31", f"=Assumptions!{c}17", F(**LK, num_format=PPC, bold=True))
ws.write("B32", "Model rule: sensitivity × Selic cut", SUB)
RM, RU = {}, {}
for i, g in enumerate(G6):
    r = 33 + i
    RM[g] = r
    ws.write(f"B{r}", ("   " if g in MEMO else "") + g, TXT)
    ws.write_formula(f"C{r}", f"=I{RE[g]}", F(num_format=PCT2))
    for c in SC:
        ws.write_formula(f"{c}{r}", f"=$C{r}*{c}$31", F(num_format=UPS))
ws.write("B39", "Used in the scenarios (an override in Assumptions §4 replaces the model rule)", SUB)
order_u = BOXES + ["Ibovespa"] + MEMO
for i, g in enumerate(order_u):
    r = 40 + i
    RU[g] = r
    tot = g == "Ibovespa"
    ws.write(f"B{r}", ("   " if g in MEMO else "") + g + (" (implied, EPS-weighted)" if tot else ""), F(bold=True, top=1) if tot else TXT)
    for c in SC:
        if tot:
            ws.write_formula(f"{c}{r}", f"=SUMPRODUCT($G$17:$G$20,1+{c}40:{c}43)/$G$21-1", F(num_format=UPS, bold=True, top=1))
        else:
            o = f"Assumptions!{c}{OVR[('eps', g)]}"
            ws.write_formula(f"{c}{r}", f"=IF(ISNUMBER({o}),{o},{c}{RM[g]})", F(num_format=UPS, bold=True))
ws.freeze_panes(4, 0)

# =========================================================================== Scenarios
ws = W["Scenarios"]
title(ws, "Scenarios: fair value of each box, method by method",
      "Read top-down: the Ibovespa block adds up the four boxes below it. Each box block shows the multiple, the earnings it applies to and the fair value in index points.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 50)
ws.set_column("C:G", 13)
ws.set_column("H:H", 92)
SR = {}


def hdr_row(r, first="Now"):
    for c, h in zip("BCDEFGH", ["", first, "Bear", "Base", "Bull", "Custom", "How it is calculated"]):
        ws.write(f"{c}{r}", h, HDR if c not in "BH" else HDRL)


def box_block(r0, g):
    fin = g == "Financials"
    memo = g in MEMO
    R = {}
    section(ws, r0, f"{g}{' (memo, not added to the Ibovespa)' if memo else ''}", fill=YEL, color=INK)
    hdr_row(r0 + 1)
    bn = BN[g]
    pe_c, ev_c = HC[("pe", g)], HC.get(("ev", g))
    R["pts"], R["w"], R["eps_chg"] = r0 + 2, r0 + 3, r0 + 4
    ws.write(f"B{r0 + 2}", "Index points", TXT)
    ws.write_formula(f"C{r0 + 2}", f"=Boxes_Now!G{bn}", F(**LK, num_format=PTS))
    ws.write(f"H{r0 + 2}", "Members' weight × Ibovespa close (Boxes_Now).", NOTE1)
    ws.write(f"B{r0 + 3}", "Weight in the Ibovespa", TXT)
    ws.write_formula(f"C{r0 + 3}", f"=Boxes_Now!F{bn}", F(**LK, num_format=PCT))
    ws.write(f"B{r0 + 4}", "EPS / EBITDA change from the Selic scenario", TXT)
    for c in SC:
        ws.write_formula(f"{c}{r0 + 4}", f"=Rates_EPS!{c}{RU[g]}", F(**LK, num_format=UPS))
    ws.write(f"H{r0 + 4}", "Box sensitivity × Selic cut, or the override (Rates_EPS step 3).", NOTE1)
    # P/E
    ws.write(f"B{r0 + 5}", "P/E method", SUB)
    R["pe"], R["pe_z"], R["eps"], R["fv_pe"], R["up_pe"] = r0 + 6, r0 + 7, r0 + 8, r0 + 9, r0 + 10
    avg, sd = f"History!${pe_c}${ST['avg']}", f"History!${pe_c}${ST['sd']}"
    ws.write(f"B{r0 + 6}", "12m fwd P/E (x)", TXT)
    ws.write_formula(f"C{r0 + 6}", f"=Boxes_Now!H{bn}", F(**LK, num_format=MULT))
    for c in SC:
        o = f"Assumptions!{c}{OVR[('pe', g)]}"
        ws.write_formula(f"{c}{r0 + 6}", f"=IF(ISNUMBER({o}),{o},{avg}+Assumptions!{c}$18*{sd})", F(num_format=MULT, bold=True))
    ws.write(f"H{r0 + 6}", "Now: harmonic mean of members (Boxes_Now). Scenario: history average + k × std. dev. (History rows 8–9), unless overridden.", NOTE1)
    ws.write(f"B{r0 + 7}", "   vs history average (std. dev.)", F(font_color=GREY))
    for c in ["C"] + SC:
        ws.write_formula(f"{c}{r0 + 7}", f"=({c}{r0 + 6}-{avg})/{sd}", F(num_format="+0.0;-0.0;0.0", font_color=GREY))
    ws.write(f"B{r0 + 8}", "EPS (index pts)", TXT)
    ws.write_formula(f"C{r0 + 8}", f"=Boxes_Now!J{bn}", F(**LK, num_format=PTS1))
    for c in SC:
        ws.write_formula(f"{c}{r0 + 8}", f"=$C{r0 + 8}*(1+{c}{r0 + 4})", F(num_format=PTS1))
    ws.write(f"H{r0 + 8}", "EPS now × (1 + EPS change).", NOTE1)
    ws.write(f"B{r0 + 9}", "Fair value (index pts)", TXT)
    for c in SC:
        ws.write_formula(f"{c}{r0 + 9}", f"={c}{r0 + 6}*{c}{r0 + 8}", F(num_format=PTS))
    ws.write(f"H{r0 + 9}", "Target P/E × scenario EPS.", NOTE1)
    ws.write(f"B{r0 + 10}", "Upside", BOLD)
    for c in SC:
        ws.write_formula(f"{c}{r0 + 10}", f"={c}{r0 + 9}/$C${r0 + 2}-1", F(num_format=UPS, bold=True))
    # EV/EBITDA
    ws.write(f"B{r0 + 11}", "EV/EBITDA method" + (" — does not apply to Financials" if fin else ""), SUB)
    R["ev"], R["ev_z"], R["ebitda"], R["evv"], R["nd"], R["fv_ev"], R["up_ev"] = range(r0 + 12, r0 + 19)
    labs = ["12m fwd EV/EBITDA (x)", "   vs history average (std. dev.)", "EBITDA (index pts)", "Enterprise value (index pts)",
            "Less: net debt (index pts)", "Fair value (index pts)", "Upside"]
    for i, l_ in enumerate(labs):
        ws.write(f"B{r0 + 12 + i}", l_, BOLD if l_ == "Upside" else (F(font_color=GREY) if "vs history" in l_ else TXT))
    if fin:
        for rr_ in range(r0 + 12, r0 + 19):
            for c in ["C"] + SC:
                ws.write(f"{c}{rr_}", "n.a.", NA)
    else:
        eavg, esd = f"History!${ev_c}${ST['avg']}", f"History!${ev_c}${ST['sd']}"
        ws.write_formula(f"C{r0 + 12}", f"=Boxes_Now!K{bn}", F(**LK, num_format=MULT))
        ws.write_formula(f"C{r0 + 14}", f"=Boxes_Now!P{bn}", F(**LK, num_format=PTS1))
        ws.write_formula(f"C{r0 + 15}", f"=Boxes_Now!N{bn}", F(**LK, num_format=PTS))
        ws.write_formula(f"C{r0 + 16}", f"=Boxes_Now!O{bn}", F(**LK, num_format=PTS))
        for c in ["C"] + SC:
            ws.write_formula(f"{c}{r0 + 13}", f"=({c}{r0 + 12}-{eavg})/{esd}", F(num_format="+0.0;-0.0;0.0", font_color=GREY))
        for c in SC:
            o = f"Assumptions!{c}{OVR[('ev', g)]}"
            ws.write_formula(f"{c}{r0 + 12}", f"=IF(ISNUMBER({o}),{o},{eavg}+Assumptions!{c}$18*{esd})", F(num_format=MULT, bold=True))
            ws.write_formula(f"{c}{r0 + 14}", f"=$C{r0 + 14}*(1+{c}{r0 + 4})", F(num_format=PTS1))
            ws.write_formula(f"{c}{r0 + 15}", f"={c}{r0 + 12}*{c}{r0 + 14}", F(num_format=PTS))
            ws.write_formula(f"{c}{r0 + 16}", f"=$C{r0 + 16}", F(num_format=PTS))
            ws.write_formula(f"{c}{r0 + 17}", f"={c}{r0 + 15}-{c}{r0 + 16}", F(num_format=PTS))
            ws.write_formula(f"{c}{r0 + 18}", f"={c}{r0 + 17}/$C${r0 + 2}-1", F(num_format=UPS, bold=True))
        ws.write(f"H{r0 + 12}", "Same rule as the P/E, on the EV/EBITDA history of the box.", NOTE1)
        ws.write(f"H{r0 + 14}", "EBITDA now (Boxes_Now) × (1 + the same change as EPS).", NOTE1)
        ws.write(f"H{r0 + 15}", "Target EV/EBITDA × scenario EBITDA.", NOTE1)
        ws.write(f"H{r0 + 16}", "Net debt held at today's level (index pts).", NOTE1)
        ws.write(f"H{r0 + 17}", "Equity value = EV − net debt.", NOTE1)
    # bottom-up
    ws.write(f"B{r0 + 19}", "Bottom-up (XP target prices)", SUB)
    R["ke"], R["bu_up"], R["fv_bu"], R["up_bu"] = r0 + 20, r0 + 21, r0 + 22, r0 + 23
    ws.write(f"B{r0 + 20}", "Ke shift (bp)", TXT)
    for c in SC:
        ws.write_formula(f"{c}{r0 + 20}", f"=Assumptions!{c}$19", F(**LK, num_format=BP))
    ws.write(f"B{r0 + 21}", "Weighted upside to XP targets after the Ke shift", TXT)
    ws.write_formula(f"C{r0 + 21}", f"=Boxes_Now!Q{bn}", F(**LK, num_format=UPS))
    crit = f"BottomUp!$B${BF}:$B${BL},Boxes_Now!$C${bn},BottomUp!$C${BF}:$C${BL},Boxes_Now!$D${bn}"
    for j, c in enumerate(SC):
        tc = cn(20 + j)
        ws.write_formula(f"{c}{r0 + 21}", f"=SUMIFS(BottomUp!${tc}${BF}:${tc}${BL},{crit})/SUMIFS(BottomUp!$O${BF}:$O${BL},{crit})-1", F(**LK, num_format=UPS))
    ws.write(f"H{r0 + 21}", "Σ weight × (1 + upside) × (1 + TP change for the Ke shift) ÷ Σ weight − 1 (BottomUp). 'Now' = no shift.", NOTE1)
    ws.write(f"B{r0 + 22}", "Fair value (index pts)", TXT)
    for c in SC:
        ws.write_formula(f"{c}{r0 + 22}", f"=$C${r0 + 2}*(1+{c}{r0 + 21})/(1+BUhaircut)", F(num_format=PTS))
    ws.write(f"H{r0 + 22}", "Index points × (1 + upside) ÷ (1 + haircut).", NOTE1)
    ws.write(f"B{r0 + 23}", "Upside", BOLD)
    for c in SC:
        ws.write_formula(f"{c}{r0 + 23}", f"={c}{r0 + 22}/$C${r0 + 2}-1", F(num_format=UPS, bold=True))
    # average
    ws.write(f"B{r0 + 24}", "Average of the methods", SUB)
    R["fv_avg"], R["up_avg"], R["pe_impl"] = r0 + 25, r0 + 26, r0 + 27
    ws.write(f"B{r0 + 25}", "Fair value (index pts)", TXT)
    for c in SC:
        if fin:
            f_ = f"=(wPE*{c}{r0 + 9}+wBU*{c}{r0 + 22})/(wPE+wBU)"
        else:
            f_ = f"=(wPE*{c}{r0 + 9}+wEV*{c}{r0 + 17}+wBU*{c}{r0 + 22})/(wPE+wEV+wBU)"
        ws.write_formula(f"{c}{r0 + 25}", f_, F(num_format=PTS, bold=True))
    ws.write(f"H{r0 + 25}", "Weighted average of the methods (weights in Assumptions §3)" + ("; Financials: P/E and bottom-up only." if fin else "."), NOTE1)
    ws.write(f"B{r0 + 26}", "Upside", F(bold=True, top=1))
    ws.write_blank(f"C{r0 + 26}", None, F(top=1))
    ws.write_blank(f"H{r0 + 26}", None, F(top=1))
    for c in SC:
        ws.write_formula(f"{c}{r0 + 26}", f"={c}{r0 + 25}/$C${r0 + 2}-1", F(num_format=UPS, bold=True, top=1, bg_color=KEY))
    ws.write(f"B{r0 + 27}", "   Implied 12m fwd P/E on scenario EPS (x)", F(font_color=GREY))
    ws.write_formula(f"C{r0 + 27}", f"=C{r0 + 6}", F(num_format=MULT, font_color=GREY))
    for c in SC:
        ws.write_formula(f"{c}{r0 + 27}", f"={c}{r0 + 25}/{c}{r0 + 8}", F(num_format=MULT, font_color=GREY))
    SR[g] = R
    return r0 + 29


IB0 = 5
r = IB0 + 27
for g in BOXES + MEMO:
    r = box_block(r, g)
# Ibovespa block (top)
r0 = IB0
section(ws, r0, "Ibovespa = sum of the four boxes")
hdr_row(r0 + 1)
I = {}
ws.write(f"B{r0 + 2}", "Index points", TXT)
ws.write_formula(f"C{r0 + 2}", f"=Boxes_Now!G{BN['Ibovespa']}", F(**LK, num_format=PTS))
I["pts"] = r0 + 2
ws.write(f"B{r0 + 3}", "EPS (index pts, sum of the boxes)", TXT)
for c in ["C"] + SC:
    ws.write_formula(f"{c}{r0 + 3}", "=" + "+".join(f"{c}{SR[g]['eps']}" for g in BOXES), F(num_format=PTS1))
I["eps"] = r0 + 3
ws.write(f"B{r0 + 4}", "EPS change", TXT)
for c in SC:
    ws.write_formula(f"{c}{r0 + 4}", f"={c}{r0 + 3}/$C{r0 + 3}-1", F(num_format=UPS))
ws.write(f"B{r0 + 5}", "Fair value by method (index pts)", SUB)
mrows = [("P/E", "fv_pe", None), ("EV/EBITDA (Financials at their P/E value)", "fv_ev", "fv_pe"), ("Bottom-up", "fv_bu", None), ("Average of the methods", "fv_avg", None)]
for i, (l_, key, finkey) in enumerate(mrows):
    rr_ = r0 + 6 + i
    I["fv_" + key] = rr_
    ws.write(f"B{rr_}", l_, BOLD if key == "fv_avg" else TXT)
    for c in SC:
        terms = [f"{c}{SR[g][finkey if (g == 'Financials' and finkey) else key]}" for g in BOXES]
        ws.write_formula(f"{c}{rr_}", "=" + "+".join(terms), F(num_format=PTS, bold=key == "fv_avg"))
ws.write(f"H{r0 + 6}", "Sum of the four box blocks below (P/E fair value rows).", NOTE1)
ws.write(f"H{r0 + 7}", "Financials have no EV/EBITDA, so their P/E fair value is used in the sum (as in the Oct-6 table).", NOTE1)
ws.write(f"B{r0 + 10}", "Upside", SUB)
for i, (l_, key, _) in enumerate(mrows):
    rr_ = r0 + 11 + i
    I["up_" + key] = rr_
    last = key == "fv_avg"
    ws.write(f"B{rr_}", l_, F(bold=True, top=1) if last else TXT)
    if last:
        ws.write_blank(f"C{rr_}", None, F(top=1))
    for c in SC:
        ws.write_formula(f"{c}{rr_}", f"={c}{I['fv_' + key]}/$C${r0 + 2}-1", F(num_format=UPS, bold=last, top=1 if last else 0, bg_color=KEY if last else "#FFFFFF"))
ws.write(f"B{r0 + 15}", "Implied 12m fwd P/E on scenario EPS (x)", SUB)
I["pe_impl_pe"], I["pe_impl_avg"] = r0 + 16, r0 + 17
ws.write(f"B{r0 + 16}", "P/E method", TXT)
ws.write_formula(f"C{r0 + 16}", f"=Boxes_Now!H{BN['Ibovespa']}", F(**LK, num_format=MULT))
ws.write(f"B{r0 + 17}", "Average of the methods", TXT)
for c in SC:
    ws.write_formula(f"{c}{r0 + 16}", f"={c}{I['fv_fv_pe']}/{c}{r0 + 3}", F(num_format=MULT))
    ws.write_formula(f"{c}{r0 + 17}", f"={c}{I['fv_fv_avg']}/{c}{r0 + 3}", F(num_format=MULT))
ws.write(f"H{r0 + 16}", "Fair value ÷ scenario EPS. 'Now' is the harmonic P/E of all members (Boxes_Now).", NOTE1)
ws.write(f"B{r0 + 18}", "Contribution to the Ibovespa upside, average of the methods (pp)", SUB)
for i, g in enumerate(BOXES):
    rr_ = r0 + 19 + i
    I["ctb_" + g] = rr_
    ws.write(f"B{rr_}", g, TXT)
    for c in SC:
        ws.write_formula(f"{c}{rr_}", f"=({c}{SR[g]['fv_avg']}-$C{SR[g]['pts']})/$C${r0 + 2}*100", F(num_format=PP))
I["ctb_tot"] = r0 + 23
ws.write(f"B{r0 + 23}", "Ibovespa", F(bold=True, top=1))
ws.write_blank(f"C{r0 + 23}", None, F(top=1))
for c in SC:
    ws.write_formula(f"{c}{r0 + 23}", f"=SUM({c}{r0 + 19}:{c}{r0 + 22})", F(num_format=PP, bold=True, top=1))
ws.write(f"H{r0 + 19}", "(Box fair value − box index points) ÷ Ibovespa points. Adds up to the Ibovespa upside.", NOTE1)
ws.freeze_panes(4, 2)

# =========================================================================== Summary
ws = W["Summary"]
title(ws, "Fair value by box: bear / base / bull",
      "Upside vs the close on the pricing date. Change the levers in Assumptions; this page updates. Construction: Scenarios.", legend=False)
ws.set_column("A:A", 2)
ws.set_column("B:B", 44)
ws.set_column("C:L", 13)
ws.set_column("M:N", 3)
ws.write("B3", "Pricing date", NOTE1); ws.write_formula("C3", "=PxDate", F(**LK, num_format=DATED))
ws.write("D3", "Ibovespa", NOTE1); ws.write_formula("E3", "=IbovNow", F(**LK, num_format=PTS))
ws.write("F3", "Checks", NOTE1); ws.write_formula("G3", "=Checks!C5", F(**LK, bold=True))
section(ws, 5, "Scenario levers in use (Assumptions)", "B", "G")
for c, h in zip("BCDEFG", ["", "Today", "Bear", "Base", "Bull", "Custom"]):
    ws.write(f"{c}6", h, HDR if c != "B" else HDRL)
lv = [("Narrative", 15, None), ("Selic at the end of the horizon", 16, PCT2), ("Multiples: std. dev. from the history average", 18, "+0.0;-0.0;0.0"),
      ("Bottom-up: Ke shift", 19, BP)]
for i, (l_, ar, f) in enumerate(lv):
    rr_ = 7 + i
    ws.write(f"B{rr_}", l_, TXT)
    for c in SC:
        ws.write_formula(f"{c}{rr_}", f"=Assumptions!{c}{ar}", F(**LK, num_format=f, align="center") if f else F(**LK, align="center"))
ws.write_formula("C8", "=SelicNow", F(**LK, num_format=PCT2, align="center"))
ws.write_formula("C9", f"=History!{HC[('pe', 'Ibovespa')]}{ST['z']}", F(**LK, num_format="+0.0;-0.0;0.0", align="center"))
ws.write("B11", "Ibovespa P/E now vs its own history (std. dev.) is shown under 'Today' in the k row.", NOTE1)
section(ws, 13, "Ibovespa fair value", "B", "L")
for c, h in zip("BCDEFGHIJKL", ["Method", "Now", "Bear", "Base", "Bull", "Custom", "", "Upside: Bear", "Base", "Bull", "Custom"]):
    if h:
        ws.write(f"{c}14", h, HDR if c != "B" else HDRL)
SUMR = {}
srows = [("P/E", "fv_pe"), ("EV/EBITDA (Financials at P/E)", "fv_ev"), ("Bottom-up", "fv_bu"), ("Average of the methods", "fv_avg")]
for i, (l_, key) in enumerate(srows):
    rr_ = 15 + i
    last = key == "fv_avg"
    b = dict(bold=True, top=1) if last else {}
    ws.write(f"B{rr_}", l_, F(**b) if last else TXT)
    ws.write_formula(f"C{rr_}", "=IbovNow", F(**LK, num_format=PTS, **b))
    for c, uc in zip(SC, "IJKL"):
        ws.write_formula(f"{c}{rr_}", f"=Scenarios!{c}{I['fv_' + key]}", F(**LK, num_format=PTS, **b))
        ws.write_formula(f"{uc}{rr_}", f"=Scenarios!{c}{I['up_' + key]}", F(**LK, num_format=UPS, **b))
    SUMR[key] = rr_
ws.write("B19", "Implied 12m fwd P/E on scenario EPS (average)", TXT)
ws.write_formula("C19", f"=Scenarios!C{I['pe_impl_pe']}", F(**LK, num_format=MULT))
ws.write("B20", "Ibovespa EPS change (rates)", TXT)
for c in SC:
    ws.write_formula(f"{c}19", f"=Scenarios!{c}{I['pe_impl_avg']}", F(**LK, num_format=MULT))
    ws.write_formula(f"{c}20", f"=Scenarios!{c}{IB0 + 4}", F(**LK, num_format=UPS))
section(ws, 22, "Upside by box (average of the methods)", "B", "L")
for c, h in zip("BCDEFGHIJKL", ["Box", "Weight", "Bear", "Base", "Bull", "Custom", "", "Fair value (pts): Bear", "Base", "Bull", "Custom"]):
    if h:
        ws.write(f"{c}23", h, HDR if c != "B" else HDRL)
ws.set_row(22, 26)
brows = BOXES + ["Ibovespa", None] + MEMO
for i, g in enumerate(brows):
    rr_ = 24 + i
    if g is None:
        ws.write(f"B{rr_}", "Memo", NOTE1)
        continue
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{rr_}", g, F(**b) if tot else TXT)
    if tot:
        ws.write_formula(f"C{rr_}", "=1", F(num_format="0%", **b))
        for c, uc in zip(SC, "IJKL"):
            ws.write_formula(f"{c}{rr_}", f"=Scenarios!{c}{I['up_fv_avg']}", F(**LK, num_format=UPS, **b))
            ws.write_formula(f"{uc}{rr_}", f"=Scenarios!{c}{I['fv_fv_avg']}", F(**LK, num_format=PTS, **b))
    else:
        ws.write_formula(f"C{rr_}", f"=Scenarios!C{SR[g]['w']}", F(**LK, num_format="0%"))
        for c, uc in zip(SC, "IJKL"):
            ws.write_formula(f"{c}{rr_}", f"=Scenarios!{c}{SR[g]['up_avg']}", F(**LK, num_format=UPS))
            ws.write_formula(f"{uc}{rr_}", f"=Scenarios!{c}{SR[g]['fv_avg']}", F(**LK, num_format=PTS))
SUMBOX0 = 24
ws.conditional_format("D24:G27", {"type": "3_color_scale", "min_color": "#F8696B", "mid_type": "num", "mid_value": 0, "mid_color": "#FFEB84", "max_color": "#63BE7B"})
ws.conditional_format("D30:G31", {"type": "3_color_scale", "min_color": "#F8696B", "mid_type": "num", "mid_value": 0, "mid_color": "#FFEB84", "max_color": "#63BE7B"})
section(ws, 33, "Contribution to the Ibovespa upside (pp)", "B", "G")
for c, h in zip("BCDEFG", ["Box", "", "Bear", "Base", "Bull", "Custom"]):
    ws.write(f"{c}34", h, HDR if c != "B" else HDRL)
for i, g in enumerate(BOXES + ["Ibovespa"]):
    rr_ = 35 + i
    tot = g == "Ibovespa"
    b = dict(bold=True, top=1) if tot else {}
    ws.write(f"B{rr_}", g, F(**b) if tot else TXT)
    for c in SC:
        ws.write_formula(f"{c}{rr_}", f"=Scenarios!{c}{I['ctb_tot' if tot else 'ctb_' + g]}", F(**LK, num_format=PP, **b))
# detail by box and method
section(ws, 41, "Detail by box and method (XP-Ray table layout)", "B", "J")
ws.write("B42", "", HDRL)
for j, s in enumerate(SCEN):
    ws.write(f"{cn(2 + 2 * j)}42", f"{s}: multiple", HDR)
    ws.write(f"{cn(3 + 2 * j)}42", f"{s}: upside", HDR)
ws.set_row(41, 26)
rr_ = 43
DET0 = rr_
for g in ["Ibovespa"] + BOXES + MEMO:
    ws.write(f"B{rr_}", g + (" (memo)" if g in MEMO else ""), F(bold=True, font_color=NAVY, bottom=1, bottom_color=NAVY))
    for cc in range(2, 10):
        ws.write_blank(rr_ - 1, cc, None, F(bottom=1, bottom_color=NAVY))
    rr_ += 1
    if g == "Ibovespa":
        lines = [("P/E (implied on scenario EPS)", "pe_impl_pe", "up_fv_pe", MULT),
                 ("EV/EBITDA (sum of the boxes)", None, "up_fv_ev", None),
                 ("Bottom-up, Ke shift (bp)", "ke", "up_fv_bu", BP),
                 ("Average", None, "up_fv_avg", None)]
        for l_, mk, uk, mf in lines:
            avg = l_ == "Average"
            ws.write(f"B{rr_}", "   " + l_, F(bold=avg))
            for j, c in enumerate(SC):
                if mk == "ke":
                    ws.write_formula(f"{cn(2 + 2 * j)}{rr_}", f"=Assumptions!{c}19", F(**LK, num_format=mf))
                elif mk:
                    ws.write_formula(f"{cn(2 + 2 * j)}{rr_}", f"=Scenarios!{c}{I[mk]}", F(**LK, num_format=mf))
                ws.write_formula(f"{cn(3 + 2 * j)}{rr_}", f"=Scenarios!{c}{I[uk]}", F(**LK, num_format=UPS, bold=avg))
            rr_ += 1
    else:
        R = SR[g]
        lines = [("P/E", "pe", "up_pe", MULT), ("EV/EBITDA", "ev", "up_ev", MULT), ("Bottom-up, Ke shift (bp)", "ke", "up_bu", BP), ("Average", None, "up_avg", None)]
        for l_, mk, uk, mf in lines:
            avg = l_ == "Average"
            ws.write(f"B{rr_}", "   " + l_, F(bold=avg))
            for j, c in enumerate(SC):
                if mk:
                    ws.write_formula(f"{cn(2 + 2 * j)}{rr_}", f"=Scenarios!{c}{R[mk]}", F(**LK, num_format=mf, align="right"))
                ws.write_formula(f"{cn(3 + 2 * j)}{rr_}", f"=Scenarios!{c}{R[uk]}", F(**LK, num_format=UPS, bold=avg, align="right"))
            rr_ += 1
    rr_ += 1
# reference: house model
REFR = rr_ + 1
section(ws, REFR, "Reference: XP house Ibovespa model, YE26 (DCF Ibov 2026_Out.xlsx, Target sheet)", "B", "J")
for c, h in zip("BCDEF", ["Method (index pts)", "", "Bear", "Base", "Bull"]):
    if h:
        ws.write(f"{c}{REFR + 1}", h, HDR if c != "B" else HDRL)
house = [("P/E (7.5x / 9.5x / 12.0x × 2026E EPS)", [157185.0, 199101.0, 251496.0]),
         ("EV/EBITDA (4.5x / 5.2x / 6.0x × 2026E − net debt)", [162215.6, 196742.4, 236201.6]),
         ("Bottom-up (house TPs ÷ 1.15, ∓20%)", [164550.7, 205688.4, 246826.1])]
for i, (l_, vals) in enumerate(house):
    ws.write(f"B{REFR + 2 + i}", l_, TXT)
    for c, v in zip("DEF", vals):
        ws.write_number(f"{c}{REFR + 2 + i}", v, F(**IN, num_format=PTS))
ws.write(f"B{REFR + 5}", "Not comparable 1:1: the house model applies YE26 multiples to 2026E earnings; this workbook applies history-based multiples to 12m fwd earnings. Published YE26 target: 200,000.", NOTE1)
# chart: upside by box and scenario
ch = wb.add_chart({"type": "column"})
for j, (s, col) in enumerate(zip(["Bear", "Base", "Bull"], ["D", "E", "F"])):
    ch.add_series({"name": s, "categories": f"=Summary!$B$24:$B$28", "values": f"=Summary!${col}$24:${col}$28",
                   "fill": {"color": [NAVY, LBLUE, YEL][j]}, "border": {"none": True}, "gap": 60, "overlap": -10,
                   "data_labels": {"value": True, "num_format": "0%", "font": {"name": "Roboto Light", "size": 8}, "position": "outside_end"}})
ch.set_x_axis({"num_font": {"name": "Roboto Light", "size": 9}, "line": {"color": "#D9D9D9", "width": 0.75}, "label_position": "low"})
ch.set_y_axis({"visible": False, "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9", "width": 0.5, "dash_type": "dash"}}})
ch.set_legend({"position": "top", "font": {"name": "Roboto Light", "size": 9}})
ch.set_title({"none": True})
ch.set_chartarea({"border": {"none": True}, "fill": {"color": "#FFFFFF"}})
ch.set_plotarea({"fill": {"none": True}, "border": {"none": True}})
ch.set_size({"width": 590, "height": 300})
ws.write("O5", "Upside by box, average of the methods (bear / base / bull)", F(bold=True, font_size=11, font_name="Roboto"))
ws.insert_chart("O6", ch, {"object_position": 3})
ws.freeze_panes(4, 0)

# =========================================================================== Sensitivity
ws = W["Sensitivity"]
title(ws, "Sensitivity: Ibovespa fair value across Selic and multiple scenarios",
      "Edit the pink headers to change the grid points. The grids use the model rules (overrides in Assumptions §4 are ignored here).")
ws.set_column("A:A", 2)
ws.set_column("B:B", 44)
ws.set_column("C:R", 10.5)
ws.write("B5", "Ke shift used for the bottom-up in grid 1 (bp)", TXT)
ws.write_number("C5", 0, F(**LV, num_format=BP))
ws.data_validation("C5", {"validate": "list", "source": f"=$B${SEN_KE_ROW0}:$B${SEN_KE_ROW0 + len(KE_LIST) - 1}"})
ws.write("D5", "Must be one of the Ke values of grid 3.", NOTE1)
# helper block
section(ws, 8, "Helper: box inputs used by the grids (links)", "B", "M")
hh = ["Box", "Index points", "P/E: average", "P/E: std. dev.", "EPS (index pts)", "EPS change per 1pp of Selic cut", "Has EV/EBITDA",
      "EV/EBITDA: average", "EV/EBITDA: std. dev.", "EBITDA (index pts)", "Net debt (index pts)", "Bottom-up upside at the Ke of C5"]
for j, h in enumerate(hh):
    ws.write(8, 1 + j, h, HDR if j else HDRL)
ws.set_row(8, 40)
for i, g in enumerate(BOXES):
    rr_ = 10 + i
    fin = g == "Financials"
    ws.write(f"B{rr_}", g, TXT)
    ws.write_formula(f"C{rr_}", f"=Boxes_Now!G{BN[g]}", F(**LK, num_format=PTS))
    ws.write_formula(f"D{rr_}", f"=History!{HC[('pe', g)]}{ST['avg']}", F(**LK, num_format="0.00"))
    ws.write_formula(f"E{rr_}", f"=History!{HC[('pe', g)]}{ST['sd']}", F(**LK, num_format="0.00"))
    ws.write_formula(f"F{rr_}", f"=Boxes_Now!J{BN[g]}", F(**LK, num_format=PTS1))
    ws.write_formula(f"G{rr_}", f"=Rates_EPS!C{RM[g]}", F(**LK, num_format=PCT2))
    ws.write_number(f"H{rr_}", 0 if fin else 1, F(num_format="0", align="center"))
    if fin:
        for c in "IJKL":
            ws.write_number(f"{c}{rr_}", 0, F(num_format="0", font_color=GREY))
    else:
        ws.write_formula(f"I{rr_}", f"=History!{HC[('ev', g)]}{ST['avg']}", F(**LK, num_format="0.00"))
        ws.write_formula(f"J{rr_}", f"=History!{HC[('ev', g)]}{ST['sd']}", F(**LK, num_format="0.00"))
        ws.write_formula(f"K{rr_}", f"=Boxes_Now!P{BN[g]}", F(**LK, num_format=PTS1))
        ws.write_formula(f"L{rr_}", f"=Boxes_Now!O{BN[g]}", F(**LK, num_format=PTS))
    gc = cn(2 + i)   # grid 3 column of the box
    ws.write_formula(f"M{rr_}", f"=INDEX({gc}${SEN_KE_ROW0}:{gc}${SEN_KE_ROW0 + len(KE_LIST) - 1},MATCH($C$5,$B${SEN_KE_ROW0}:$B${SEN_KE_ROW0 + len(KE_LIST) - 1},0))", F(num_format=UPS))
SELIC_LIST = [0.085 + 0.005 * i for i in range(13)]
K_LIST = [-1.5, -1.0, -0.5, 0.0, 0.5, 1.0, 1.5]
H_ = "$C$10:$C$13"


def arr(c):
    return f"${c}$10:${c}$13"


def grid(top, label, kind):
    """kind: 'avg' (all methods) or 'pe' (P/E only). Points on the left, upside on the right."""
    section(ws, top, label, "B", "R")
    ws.write(f"B{top + 1}", "Selic at the end of the horizon ↓   |   k (std. dev. from the history average) →", NOTE1)
    hr = top + 2
    ws.write(f"B{hr}", "Index points", HDRL)
    ws.write(f"J{hr}", "Upside", HDRL)
    for j, k in enumerate(K_LIST):
        if top == G1:
            ws.write_number(hr - 1, 2 + j, k, F(**LV, num_format="+0.0;-0.0;0.0", align="center"))
        else:
            ws.write_formula(hr - 1, 2 + j, f"={cn(2 + j)}${G1 + 2}", F(bold=True, font_color="#FFFFFF", bg_color=NAVY, num_format="+0.0;-0.0;0.0", align="center"))
        ws.write_formula(hr - 1, 10 + j, f"={cn(2 + j)}${hr}", F(bold=True, font_color="#FFFFFF", bg_color=NAVY, num_format="+0.0;-0.0;0.0", align="center"))
    for i, s in enumerate(SELIC_LIST):
        r_ = hr + 1 + i
        if top == G1:
            ws.write_number(f"B{r_}", s, F(**LV, num_format=PCT2, align="center"))
        else:
            ws.write_formula(f"B{r_}", f"=$B${G1 + 3 + i}", F(bold=True, num_format=PCT2, align="center"))
        ws.write_formula(f"J{r_}", f"=$B{r_}", F(bold=True, num_format=PCT2, align="center"))
        cut = f"(SelicNow-$B{r_})*100"
        for j in range(len(K_LIST)):
            kc = f"{cn(2 + j)}${hr}"
            pe = f"({arr('D')}+{kc}*{arr('E')})*{arr('F')}*(1+{arr('G')}*{cut})"
            if kind == "pe":
                f_ = f"=SUMPRODUCT({pe})"
            else:
                ev = f"{arr('H')}*(({arr('I')}+{kc}*{arr('J')})*{arr('K')}*(1+{arr('G')}*{cut})-{arr('L')})"
                bu = f"{arr('C')}*(1+{arr('M')})/(1+BUhaircut)"
                f_ = f"=SUMPRODUCT((wPE*{pe}+wEV*{ev}+wBU*{bu})/(wPE+wEV*{arr('H')}+wBU))"
            ws.write_formula(r_ - 1, 2 + j, f_, F(num_format=PTS))
            ws.write_formula(r_ - 1, 10 + j, f"={cn(2 + j)}{r_}/IbovNow-1", F(num_format=UPS))
    last = hr + len(SELIC_LIST)
    ws.conditional_format(f"K{hr + 1}:Q{last}", {"type": "3_color_scale", "min_color": "#F8696B", "mid_type": "num", "mid_value": 0,
                                                    "mid_color": "#FFEB84", "max_color": "#63BE7B"})
    for rng_ in (f"C{hr + 1}:I{last}", f"K{hr + 1}:Q{last}"):
        c0 = rng_.split(":")[0]
        colL = "C" if c0.startswith("C") else "K"
        kref = f"{colL}${hr}"
        ws.conditional_format(rng_, {"type": "formula", "criteria": f"=AND(ROUND($B{hr + 1}-BaseSelic,6)=0,ROUND({kref}-BaseK,6)=0)",
                                     "format": F(bold=True, border=2, border_color=NAVY)})
    return hr, last


G1 = 15
g1h, g1l = grid(G1, "Grid 1 — Ibovespa fair value, average of the methods (bottom-up at the Ke shift in C5)", "avg")
G2 = g1l + 3
g2h, g2l = grid(G2, "Grid 2 — Ibovespa fair value, P/E method only", "pe")
G3 = g2l + 3
assert G3 + 2 == SEN_KE_ROW0, (G3, SEN_KE_ROW0)
section(ws, G3, "Grid 3 — bottom-up upside by box vs the shift in Ke (repricing of the XP targets, BottomUp)", "B", "R")
gh = G3 + 1
ws.write(f"B{gh}", "Ke shift (bp)", HDRL)
g3 = [("Financials", "B"), ("Defensives", "B"), ("Cyclicals", "B"), ("Commodities", "B"), ("Ibovespa", None),
      ("Cyclicals ex-WEG & Embraer", "C"), ("WEG & Embraer", "C")]
for j, (g, _) in enumerate(g3):
    ws.write(gh - 1, 2 + j, g, HDR)
ws.set_row(gh - 1, 40)
for i, k in enumerate(KE_LIST):
    r_ = SEN_KE_ROW0 + i
    ws.write_number(f"B{r_}", k, F(**LV, num_format=BP, align="center"))
    hc = cn(KEG0 + i)
    for j, (g, crit_col) in enumerate(g3):
        if crit_col is None:
            f_ = f"=SUM(BottomUp!${hc}${BF}:${hc}${BL})/SUM(BottomUp!$O${BF}:$O${BL})-1"
        else:
            f_ = (f"=SUMIFS(BottomUp!${hc}${BF}:${hc}${BL},BottomUp!${crit_col}${BF}:${crit_col}${BL},{cn(2 + j)}${gh})"
                  f"/SUMIFS(BottomUp!$O${BF}:$O${BL},BottomUp!${crit_col}${BF}:${crit_col}${BL},{cn(2 + j)}${gh})-1")
        ws.write_formula(r_ - 1, 2 + j, f_, F(num_format=UPS, bold=g == "Ibovespa"))
ws.conditional_format(f"C{SEN_KE_ROW0}:I{SEN_KE_ROW0 + len(KE_LIST) - 1}", {"type": "3_color_scale", "min_color": "#F8696B", "mid_type": "num",
                                                                          "mid_value": 0, "mid_color": "#FFEB84", "max_color": "#63BE7B"})
ws.write(f"B{SEN_KE_ROW0 + len(KE_LIST) + 1}", "Grid cells framed in navy = the Base scenario (Selic and k in Assumptions column E). Grids 1–2 use the model's EPS rule; overrides are not applied.", NOTE1)
ws.freeze_panes(4, 0)

# =========================================================================== Charts
ws = W["Charts"]
title(ws, "Charts: 12m fwd P/E and risk premium, Ibovespa and boxes, Oct-16 to the pricing date",
      "Native charts on History (last point = pricing date). Average and ±1 std. dev. follow the window in Assumptions.", legend=False)
ws.set_column("A:A", 2)
C0 = HF + 120            # Oct-16 row in History
assert months[120] == pd.Timestamp("2016-10-31")
fnt = {"name": "Roboto Light", "size": 9}


def pe_minmax(g):
    q = p[p["date"] >= "2016-10-31"].copy()
    q["bx"] = q["gics"].map(dict(GICS_MAP))
    if g != "Ibovespa":
        q = q[q["bx"] == g]
    q = q[(q["pe"] >= 1) & (q["pe"] <= 100)]
    s_ = q.groupby("date").apply(lambda x: x["w"].sum() / (x["w"] / x["pe"]).sum())
    e_ = 1 / s_ - rr.reindex(s_.index).values
    return float(s_.min()), float(s_.max()), float(e_.min()), float(e_.max())
for j, (kind, g) in enumerate(CH):
    colx = 1 if kind == "pe" else 11
    rowx = 6 + (j % 5) * 22
    lab_ = f"{g}: {'12m fwd P/E (x)' if kind == 'pe' else 'risk premium (12m fwd earnings yield − 5y real rate)'}"
    ws.write(rowx - 2, colx, lab_, F(bold=True, font_size=11, font_name="Roboto"))
    ch = wb.add_chart({"type": "line"})
    cat = f"=History!$A${C0}:$A${HNOW}"
    sc_ = HC[(kind, g)]
    ser = [(f"{g} {'12m fwd P/E' if kind == 'pe' else 'risk premium'}", sc_, {"color": NAVY, "width": 1.5}),
           ("History average", CHC[(kind, g, "avg")], {"color": YEL, "width": 1.25, "dash_type": "dash"}),
           ("±1 std. dev.", CHC[(kind, g, "p1")], {"color": "#A5A5A5", "width": 1.0, "dash_type": "dash"}),
           ("−1 std. dev.", CHC[(kind, g, "m1")], {"color": "#A5A5A5", "width": 1.0, "dash_type": "dash"})]
    for nm, col, ln in ser:
        ch.add_series({"name": nm, "categories": cat, "values": f"=History!${col}${C0}:${col}${HNOW}", "line": ln,
                       "marker": {"type": "none"}, "smooth": False})
    ch.set_x_axis({"date_axis": True, "num_format": DATE, "major_unit": 2, "major_unit_type": "years", "base_unit": "months",
                   "num_font": fnt, "line": {"color": "#D9D9D9", "width": 0.75}, "major_tick_mark": "outside", "label_position": "low"})
    ya = {"num_format": '0"x"' if kind == "pe" else "0%", "num_font": fnt, "line": {"none": True}, "major_tick_mark": "none",
          "crossing": "min", "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9", "width": 0.5, "dash_type": "dash"}}}
    mn, mx, emn, emx = pe_minmax(g)
    if kind == "pe":
        step = 2 if mx - mn < 15 else 5
        ya["min"] = max(0, int(np.floor((mn - 0.1 * (mx - mn)) / step) * step))
        ya["major_unit"] = step
    elif emx - emn < 0.06:
        ya["num_format"] = "0.0%"
    ch.set_y_axis(ya)
    ch.set_legend({"position": "top", "font": fnt, "delete_series": [3]})
    ch.set_title({"none": True})
    ch.set_chartarea({"border": {"none": True}, "fill": {"color": "#FFFFFF"}})
    ch.set_plotarea({"fill": {"none": True}, "border": {"none": True}})
    ch.show_blanks_as("gap")
    ch.set_size({"width": 586, "height": 280})
    ws.insert_chart(rowx - 1, colx, ch, {"object_position": 3})

# =========================================================================== Checks
ws = W["Checks"]
title(ws, "Checks", "Integrity tests and the tie-out to the Oct-6 table. 'OK' / 'CHECK' drive the status in C5; 'Info' and 'Differs' do not.", legend=False)
ws.set_column("A:A", 2)
ws.set_column("B:B", 70)
ws.set_column("C:F", 14)
ws.set_column("G:G", 70)
ws.write("B5", "Overall status", BOLD)
for c, h in zip("BCDEFG", ["Check", "Value", "Target", "Difference", "Status", "Note"]):
    ws.write(f"{c}7", h, HDR if c not in "BG" else HDRL)
chk = []


def addc(label, val, tgt, tol, note="", kind="test", fmt="0.0000"):
    chk.append((label, val, tgt, tol, note, kind, fmt))


addc("Member weights on the pricing date add up to 100%", f"=Members!K{MT}", "=1", 1e-9, "", fmt="0.0000%")
addc("Members' index points add up to the Ibovespa close", f"=Members!L{MT}", "=IbovNow", 0.01, "", fmt=PX)
addc("The four boxes add up to the Ibovespa (index points)", "=SUM(Boxes_Now!G6:G9)", "=Boxes_Now!G10", 0.01, "", fmt=PX)
addc("Memo sub-boxes add up to Cyclicals (index points)", "=Boxes_Now!G12+Boxes_Now!G13", "=Boxes_Now!G8", 0.01, "", fmt=PX)
addc("Every member mapped to a box", f"=SUMPRODUCT(--ISNA(Members!D{MF}:D{ML}))", "=0", 0, "Add the sector to Assumptions §5 if not.", fmt="0")
addc("Every panel row mapped to a box", f"=SUMPRODUCT(--ISNA(Panel!D{PF}:D{PL}))", "=0", 0, "", fmt="0")
addc("Price drift vs the Ibovespa move since the membership date", f"=SUMPRODUCT(Members!G{MF}:G{ML},Members!I{MF}:I{ML}/Members!H{MF}:H{ML})", "=IbovNow/IbovMem", None,
     "Info: the gap (~0.7%) comes from intra-period index adjustments; weights are rescaled to 100%.", "info", "0.0000")
addc("EPS-weighted box sensitivities = Ibovespa sensitivity", "=Rates_EPS!C27", "=EPSsens", 1e-9, "", fmt=PCT2)
addc("Lowest P/E coverage among the boxes (% of weight)", "=MIN(Boxes_Now!I6:I9)", "=0.9", None, "Info: AURE3 and CSNA3 have no 12m fwd P/E on Oct-5.", "info", PCT)
addc("Lowest EV/EBITDA coverage among the non-financial boxes", "=MIN(Boxes_Now!L7:L9)", "=0.9", None, "Info.", "info", PCT)
addc("Members with the Ke sensitivity filled from the sub-box average (count)", f'=COUNTIF(BottomUp!M{BF}:M{BL},"Sub-box average")', "=0", None,
     "Info: real estate (left out by design) and BBSE3, PSSA3, CSAN3, HAPV3, BEEF3 pending from the analysts.", "info", "0")
addc("... their weight in the Ibovespa", f'=SUMIFS(BottomUp!D{BF}:D{BL},BottomUp!M{BF}:M{BL},"Sub-box average")', "=0", None, "Info.", "info", PCT)
addc("Overrides in use (Assumptions §4)", f"=COUNT(Assumptions!D{A['ovr_first']}:G{A['ovr_last']})", "=0", None, "Info: a number other than 0 means some scenario values are typed, not modelled.", "info", "0")
addc("Stale Bloomberg repeats flagged in the panel (P/E)", f"=SUM(Panel!J{PF}:J{PL})", "=0", None, "Info: dropped from the history only if DropStale = 1.", "info", "0")
addc("Formula errors in Scenarios", "=SUMPRODUCT(--ISERROR(Scenarios!B5:H260))", "=0", 0, "", fmt="0")
addc("Formula errors in Summary", "=SUMPRODUCT(--ISERROR(Summary!B5:L120))", "=0", 0, "", fmt="0")
addc("Formula errors in Sensitivity", "=SUMPRODUCT(--ISERROR(Sensitivity!B1:R70))", "=0", 0, "", fmt="0")
addc("Formula errors in History (multiples and premia)", f"=SUMPRODUCT(--ISERROR(History!A8:{cn(LASTC)}{HNOW}))", "=0", 0, "", fmt="0")
addc("Sensitivity grid 1 at the Base levers = Base average fair value",
     f"=IFERROR(INDEX(Sensitivity!C{g1h + 1}:I{g1l},MATCH(Assumptions!E16,Sensitivity!B{g1h + 1}:B{g1l},0),MATCH(Assumptions!E18,Sensitivity!C{g1h}:I{g1h},0)),\"n.a.\")",
     f"=Scenarios!E{I['fv_fv_avg']}", None, "Info: equal when the Base Ke = C5 of Sensitivity, Base levers are on the grid and no overrides.", "info", PTS)
# tie-out
ref_pairs = []
for mt, key in (("P/E", "fv_pe"), ("EV/EBITDA", "fv_ev"), ("Bottom-up", "fv_bu"), ("Average", "fv_avg")):
    for s, c in zip(["Bear", "Base", "Bull"], ["D", "E", "F"]):
        ref_pairs.append((f"Tie-out to the Oct-6 table: Ibovespa {mt}, {s} (index pts)", f"=Scenarios!{c}{I['fv_' + key]}", REF["fv"][mt][f"Ibovespa|{s}"]))
for g in BOXES + MEMO:
    for s, c in zip(["Bear", "Base", "Bull"], ["D", "E", "F"]):
        ref_pairs.append((f"Tie-out: {g}, average upside, {s}", f"=Scenarios!{c}{SR[g]['up_avg']}",
                          REF["fv"]["Average"][f"{g}|{s}"] / REF["bx"][g]["pts"] - 1))
r = 8
for label, val, tgt, tol, note, kind, fmt in chk:
    ws.write(f"B{r}", label, TXT)
    ws.write_formula(f"C{r}", val, F(num_format=fmt))
    ws.write_formula(f"D{r}", tgt, F(num_format=fmt))
    ws.write_formula(f"E{r}", f'=IFERROR(C{r}-D{r},"")', F(num_format=fmt))
    if kind == "info":
        ws.write(f"F{r}", "Info", F(font_color=GREY, align="center"))
    else:
        ws.write_formula(f"F{r}", f'=IF(ABS(C{r}-D{r})<={tol:.10f},"OK","CHECK")', F(bold=True, align="center"))
    ws.write(f"G{r}", note, NOTE1)
    r += 1
r += 1
ws.write(f"B{r}", "Tie-out to the Oct-6 table (Ray_table_by_box_2026-10-06_v6): matches only with the default settings", SUB)
r += 1
for label, val, tgt in ref_pairs:
    pts = "index pts" in label
    ws.write(f"B{r}", label, TXT)
    ws.write_formula(f"C{r}", val, F(num_format=PTS if pts else UPS))
    ws.write_number(f"D{r}", tgt, F(**IN, num_format=PTS if pts else UPS))
    ws.write_formula(f"E{r}", f"=C{r}-D{r}", F(num_format="0.00" if pts else "0.000%"))
    ws.write_formula(f"F{r}", f'=IF(ABS(C{r}-D{r})<={1 if pts else 0.00005},"OK","Differs")', F(bold=True, align="center"))
    r += 1
ws.write_formula("C5", f'=IF(COUNTIF(F8:F{r},"CHECK")=0,"All checks OK","Some checks need attention")', F(bold=True, font_color=GREEN))
ws.conditional_format(f"F8:F{r}", {"type": "cell", "criteria": "==", "value": '"CHECK"', "format": F(font_color="#C00000", bold=True)})
ws.conditional_format(f"F8:F{r}", {"type": "cell", "criteria": "==", "value": '"Differs"', "format": F(font_color="#C55A11", bold=True)})
ws.conditional_format("C5", {"type": "text", "criteria": "begins with", "value": "Some", "format": F(font_color="#C00000", bold=True)})
ws.freeze_panes(7, 0)

# =========================================================================== README
ws = W["README"]
ws.set_column("A:A", 2)
ws.set_column("B:B", 24)
ws.set_column("C:C", 120)
ws.set_row(0, 26)
ws.write("B1", "Ibovespa fair value by box — model", F(bold=True, font_size=16, font_color=NAVY))
ws.write("B2", "XP Research · Equity Strategy · built Oct-7-2026 from the Oct-5/6 study · prices as of Oct-5-2026", NOTE1)
rr_ = 4


def para(head, lines):
    global rr_
    ws.merge_range(f"B{rr_}:C{rr_}", head, F(bold=True, font_color="#FFFFFF", bg_color=NAVY, font_size=10))
    rr_ += 1
    for ln in lines:
        if isinstance(ln, tuple):
            ws.write(f"B{rr_}", ln[0], F(bold=True, valign="top"))
            ws.write(f"C{rr_}", ln[1], TXTW)
            n = -(-len(ln[1]) // 165)
        else:
            ws.merge_range(f"B{rr_}:C{rr_}", ln, TXTW)
            n = -(-len(ln) // 190)
        if n > 1:
            ws.set_row(rr_ - 1, 13 * n + 2)
        rr_ += 1
    rr_ += 1


para("What this workbook does", [
    "Splits the Ibovespa into four boxes — Financials, Defensives (utilities, staples, health care, telecom), Cyclicals (discretionary, industrials, IT, real estate) and Commodities (energy, materials) — "
    "and values each one under bear / base / bull Selic scenarios with three methods: 12m fwd P/E, 12m fwd EV/EBITDA and bottom-up (XP target prices). The Ibovespa fair value is the sum of the boxes. "
    "A fourth 'Custom' scenario is free for testing."])
para("How to use it", [
    ("1. Change", "Assumptions: the pink cells with a dotted border only (Selic per scenario, k = std. devs. from the history average, Ke shift for the bottom-up, history window, method weights, optional overrides)."),
    ("2. Read", "Summary (results), Sensitivity (grids of Selic × k and Ke), Charts (P/E and risk premium vs history)."),
    ("3. Trace", "Any number: Summary → Scenarios (the box blocks show every step) → Boxes_Now / History / Rates_EPS / BottomUp → Members / Panel (raw data)."),
    ("4. Verify", "Checks: integrity tests and the tie-out to the Oct-6 table (Ibovespa 174,414 / 229,232 / 294,832 with the default settings).")])
para("Sheets", [
    ("Cover", "Navigation buttons to every tab, the valuation summary and the multiples at a glance."),
    ("Assumptions", "Inputs and levers. Named cells (IbovNow, SelicNow, WinStart, KeCap, wPE …) are used across the workbook."),
    ("Summary", "Ibovespa fair value by method and scenario, upside and contribution by box, the XP-Ray table layout, the house-model reference."),
    ("Sensitivity", "Grid 1: Ibovespa fair value (average of methods) for Selic × k. Grid 2: P/E only. Grid 3: bottom-up upside by box vs the Ke shift."),
    ("Charts", "12m fwd P/E and risk premium of the Ibovespa and the four boxes, Oct-16 to date, with the window average ±1 std. dev."),
    ("Scenarios", "One block per box: index points → EPS change → target P/E × EPS; target EV/EBITDA × EBITDA − net debt; bottom-up; average. The Ibovespa block sums the boxes."),
    ("Boxes_Now", "Current box aggregates from Members: weight, index points, harmonic-mean P/E and EV/EBITDA, EPS/EBITDA/net debt in index points, upside to targets, risk premium."),
    ("Rates_EPS", "EPS change by box and scenario: Ibovespa EPS sensitivity to the Selic, split by each box's return beta to the 2y pré, × the Selic cut."),
    ("BottomUp", "Member-level repricing of the XP targets for a Ke shift, using the analysts' TP sensitivity to Ke (implied perpetuity k − g)."),
    ("Members", "76 Ibovespa members on the pricing date: classification, weights, prices, multiples, EV, targets (raw inputs in blue)."),
    ("History", "Monthly box P/E, EV/EBITDA and risk premium, Oct-06 to Sep-26 (from Panel), window statistics and the pricing-date point."),
    ("Panel", "17,176 member-months of Bloomberg 12m fwd P/E and EV/EBITDA with Ibovespa weights; stale-repeat flags."),
    ("Checks", "Tests, coverage, overrides in use, formula errors, tie-out."),
    ("Support >", "Divider, as in the XP model template: main tabs before it, support tabs after it.")])
para("Colour code", [
    ("Blue", "Hardcoded input (data from Bloomberg, Economatica, XP, studies)."),
    ("Pink, dotted border", "Assumption meant to be changed (Assumptions, Sensitivity grid headers)."),
    ("Black", "Formula, links to other sheets included, as in the XP model template."),
    ("Navy / yellow bars", "Navy = section; yellow = block (one per box in Scenarios); grey band = sub-block."),
    ("Tabs", "Main tabs first (Summary, Assumptions, Scenarios, Sensitivity); support tabs after the yellow Support > divider.")])
para("Method", [
    ("Boxes", "XP-adjusted GICS sector → box (Assumptions §5). WEG and Embraer stay in Cyclicals and are shown apart as memo rows with Cyclicals ex-WEG & Embraer."),
    ("Weights", "Economatica Ibovespa composition on Sep-30, drifted by price to Oct-5 and rescaled. Index points = weight × Ibovespa close."),
    ("Box multiples", "Weighted harmonic mean of the members' Bloomberg 12m fwd (blended) multiples; P/E valid 1–100x, EV/EBITDA 1–50x. The same rule builds the monthly history."),
    ("Target multiples", "History average + k × std. dev. over the window (default Oct-16..Sep-26, k = −1 / 0 / +1)."),
    ("Rates → earnings", "Ibovespa EPS +2.07% per −100bp of Selic, split by box return beta to the 2y pré (normalised to the Ibovespa) and applied to EPS and EBITDA."),
    ("P/E fair value", "Target P/E × EPS now × (1 + EPS change)."),
    ("EV/EBITDA fair value", "Target EV/EBITDA × EBITDA now × (1 + change) − net debt now. Not applied to Financials."),
    ("Bottom-up", "Weighted upside to the XP 12m targets; bear / bull reprice each target for Ke +200 / −200bp with the analyst's sensitivity, capped at ±40%."),
    ("Average", "Equal-weighted average of the methods (Financials: P/E and bottom-up). Ibovespa = sum of the four boxes.")])
para("Sources", [
    ("Prices, multiples", "Bloomberg PX_LAST, BEST_PE_RATIO and BEST_EV_TO_BEST_EBITDA (BF), CUR_MKT_CAP, CURR_ENTP_VAL, BEST_TARGET_PRICE — Oct-5-2026."),
    ("History", "Bloomberg monthly BDH of the same fields, Oct-06..Sep-26, for the members of each month (bx_panel20.parquet)."),
    ("Weights", "Economatica Ibovespa composition (eco2-index_composition)."),
    ("Targets", "XP comp sheet (raw_data.xlsx, Oct-2); consensus for 4 names without XP coverage."),
    ("Ke sensitivities", "Analysts, Oct-6 (Sensitivity.xlsx in the bucket + Financials/TMT table)."),
    ("Rates", "BCB Selic; NTN-B 5y (Bloomberg BZRFB5PY, Tesouro Direto after Aug-21; ANBIMA for Oct-5); studies perf_study.py and bx_eps_rates2.py."),
    ("Code", "Equity Strategy Dashboard/_pedidos/2026-10-05_ibov_valuation_groups/scripts (bx_model.py = pandas reference). This workbook: Brazil Bull Case/scripts/build_model.py.")])
para("Open points", [
    "Prices are as of Oct-5. Refreshing = new values in the blue columns of Members, the snapshot in Assumptions §1 and new months in Panel/History.",
    "Five members still lack an analyst Ke sensitivity (BBSE3, PSSA3, CSAN3, HAPV3, BEEF3); they take their sub-box average.",
    "Stale Bloomberg repeats in the history are kept by default (DropStale = 0, as on Oct-6). Switching to 1 moves the 10y averages slightly (e.g. Commodities 7.37x → ~7.49x).",
    "The Cyclicals 10y average (17.9x) includes the 2019–21 re-rating; the history window in Assumptions §3 tests how much of the upside depends on it.",
    "New numbers — need Fernando / compliance before going to clients."])

for s_ in ("README", "Assumptions", "Summary", "Sensitivity", "Charts", "Scenarios", "Boxes_Now", "Rates_EPS", "Checks"):
    W[s_].set_landscape()
    W[s_].set_paper(9)
    W[s_].fit_to_pages(1, 0)
    W[s_].set_margins(0.4, 0.4, 1.25, 0.5)

# =========================================================================== Cover (as the template's Cover sheet)
ws = W["Cover"]
ws.set_column("A:A", 2.2)
ws.set_column("B:B", 1.4)
ws.set_column("C:N", 11.5)
ws.set_column("O:O", 1.4)
nav = F(bg_color=NAVY)
for r_ in range(2, 6):
    for c_ in "BCDEFGHIJKLMNO":
        ws.write_blank(f"{c_}{r_}", None, nav)
ws.set_row(1, 8)
ws.set_row(2, 34)
ws.set_row(3, 18)
ws.set_row(4, 8)
ws.write_rich_string("C3", F(bold=True, font_size=18, font_color="#FFFFFF"), "Ibovespa",
                     F(font_name="Roboto Light", font_size=18, font_color="#FFFFFF"), " | Fair value by box", nav)
ws.write("C4", "VALUATION MODEL  ·  EQUITY STRATEGY", F(font_name="Roboto Light", font_color="#FFFFFF", bg_color=NAVY))
ws.insert_image("L2", str(LOGO_WHITE), {"x_scale": 0.62, "y_scale": 0.62, "x_offset": 40, "y_offset": 6, "object_position": 3})
lab = F(bold=True, italic=True, font_size=9, font_color=NAVY)
ws.write("C7", "Main tabs >>", lab)
ws.write("C11", "Auxiliary tabs >>", lab)


def button(rng, text, target, fill, color):
    first = rng.split(":")[0]
    f_ = F(bold=True, font_size=11, font_color=color, bg_color=fill, align="center", text_wrap=True, border=5, border_color="#FFFFFF")
    ws.merge_range(rng, "", f_)
    ws.write_url(first, f"internal:'{target}'!A1", f_, text, tip=f"Go to {target}")


for i, (t, tgt) in enumerate([("Summary", "Summary"), ("Assumptions", "Assumptions"), ("Scenarios", "Scenarios"), ("Sensitivity", "Sensitivity")]):
    button(f"{cn(2 + 2 * i)}8:{cn(3 + 2 * i)}9", t, tgt, NAVY, "#FFFFFF")
aux = [[("Boxes Now\n(current aggregates)", "Boxes Now"), ("Rates & EPS\n(Selic to earnings)", "Rates & EPS"),
        ("Bottom-up\n(XP targets x Ke)", "Bottom-up"), ("History\n(multiples since 2006)", "History")],
       [("Charts\n(P/E and risk premium)", "Charts"), ("Members\n(76 stocks)", "Members"),
        ("Panel\n(Bloomberg history)", "Panel"), ("Checks\n(tests and tie-out)", "Checks")]]
for k_, row_ in enumerate(aux):
    r_ = 12 + 3 * k_
    for i, (t, tgt) in enumerate(row_):
        button(f"{cn(2 + 2 * i)}{r_}:{cn(3 + 2 * i)}{r_ + 1}", t, tgt, YEL, INK)
for r_ in (8, 9, 12, 13, 15, 16):
    ws.set_row(r_ - 1, 19)
kv = [("Ibovespa close", "=IbovNow", PTS), ("Pricing date", "=PxDate", DATED), ("Base-case fair value", "=Summary!E18", PTS),
      ("Upside, base case", "=Summary!J18", UPS), ("Model checks", "=Checks!C5", None)]
for i, (l_, f_, nf) in enumerate(kv):
    r_ = 8 + i
    bb = i == 2
    ws.merge_range(f"K{r_}:L{r_}", l_, F(font_name="Roboto Light", bold=bb, bottom=4))
    fmt_ = F(align="right", bottom=4, bold=bb, num_format=nf) if nf else F(align="right", bottom=4, bold=True, font_color=GREEN)
    ws.merge_range(f"M{r_}:N{r_}", "", fmt_)
    ws.write_formula(f"M{r_}", f_, fmt_)
ws.write("K14", "Equity Strategy | XP Research", F(bold=True))
ws.write_url("K15", "internal:'Read Me'!A1", F(italic=True, underline=1, font_color=INBLUE), "How this model works >>")
ws.write("K16", "Prices as of Oct-5-2026.", NOTE1)
gl = F(bold=True, italic=True, font_size=9, font_color=GREY)
ws.write("C19", "Valuation summary >>", gl)
hd = F(bold=True, bottom=2, bottom_color=LBLUE, align="center")
ws.merge_range("C20:E20", "Ibovespa (index pts)", F(bold=True, bottom=2, bottom_color=LBLUE))
for j, s_ in enumerate(["Bear", "Base", "Bull", "Custom"]):
    ws.write(19, 5 + j, s_, hd)
rows_ = [("Selic at the end of the horizon", "Assumptions!{c}16", PCT2, False, False),
         ("Fair value, average of the methods", "Summary!{c}18", PTS, True, False),
         ("Upside", "Summary!{u}18", UPS, False, True),
         ("Implied 12m fwd P/E", "Summary!{c}19", MULT, False, False),
         ("EPS change from rates", "Summary!{c}20", UPS, False, True)]
for i, (l_, ref, nf, b, it) in enumerate(rows_):
    r_ = 21 + i
    bt = 1 if i == len(rows_) - 1 else 0
    ws.merge_range(f"C{r_}:E{r_}", ("   " if it else "") + l_, F(bold=b, italic=it, bottom=bt))
    for j, (c_, u_) in enumerate(zip("DEFG", "IJKL")):
        ws.write_formula(r_ - 1, 5 + j, "=" + ref.format(c=c_, u=u_), F(bold=b, italic=it, num_format=nf, align="center", bottom=bt))
ws.write("K19", "Multiples >>", gl)
for j, h_ in enumerate(["12m fwd P/E", "Now", "10y avg.", "vs avg. (sd)"]):
    ws.write(19, 10 + j, h_, hd if j else F(bold=True, bottom=2, bottom_color=LBLUE))
for i, g_ in enumerate(["Ibovespa", "Financials", "Defensives", "Cyclicals", "Commodities"]):
    r_ = 21 + i
    bt = 1 if i == 4 else 0
    c_ = HC[("pe", g_)]
    b = g_ == "Ibovespa"
    ws.write(f"K{r_}", g_, F(bold=b, bottom=bt))
    ws.write_formula(f"L{r_}", f"=History!{c_}{ST['now']}", F(bold=b, num_format=MULT, align="center", bottom=bt))
    ws.write_formula(f"M{r_}", f"=History!{c_}{ST['avg']}", F(bold=b, num_format=MULT, align="center", bottom=bt))
    ws.write_formula(f"N{r_}", f"=History!{c_}{ST['z']}", F(bold=b, num_format="+0.0;-0.0;0.0", align="center", bottom=bt))
ws.write("C27", "Fair value = sum of four boxes (Financials, Defensives, Cyclicals, Commodities), each valued by 12m fwd P/E, EV/EBITDA and XP target prices under Selic scenarios.", NOTE1)
ws.set_landscape()
ws.set_paper(9)
ws.fit_to_pages(1, 1)
ws.activate()
wb.close()
print("written", OUT)
