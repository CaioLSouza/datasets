"""Financials box: FCFE / bank DDM with the logic of the XP Financials analysts' models and the content of the consensus.

Actuals and consensus years (2026-28) are summed from the members (Bloomberg BEST net income and ROE, in index points);
the projection 2029-37 is done for the box only: box ROE held at FY28; net income growth fades from the box's 2028
growth to g by 2037; cash to shareholders = net income - the increase in book value (the analysts' equity roll-forward);
Gordon perpetuity on the 2037 flow, discounted with the 2037 factor; value at end-2027.
Ke = nominal risk free ((NTN-B + IPCA) x 0.85, required by the nominal flows) + beta x 5.5% (one ERP for every box).
Section 2 of the main sheet sets each remaining variable to the house or the analysts' choice (beta, g, explicit years,
terminal-value timing).
The analysts' own projections are a memo (Analyst Models) and can replace the box path (switch).
Usage: python build_fin_ddm.py [output file name]"""
from pathlib import Path

from xlsxwriter.utility import xl_col_to_name as cn

import openpyxl
import pandas as pd

from xp_template import Book, PCT, PCT2, PTS1, UPS, NAVY

HERE = Path(__file__).parent
import sys
OUT = HERE.parent / (sys.argv[1] if len(sys.argv) > 1 else "Financials_DDM_XP.xlsx")
YRS = list(range(2026, 2038))
Y = [cn(2 + i) for i in range(12)]          # C .. N
YC = dict(zip(YRS, Y))
NOTE_COL = "P"

# ----------------------------------------------------------------------------- analysts' models (R$ mn), read on Oct-8-2026
# (company, file, ni source, div source, ni list or None, div list, payout list or None, index pts, market cap R$ mn)
FOLDER = "inputs/analyst_models/Financials/"
MODELS = [
    ("Itaú Unibanco (ITUB4)", "ITUB Model Official 2Q26_vCSStrategy.xlsx", "Income Statement r61 Recurring Net Income",
     "Income Statement r66 Dividends/IoC Provisioned",
     [50922, 56028, 61318, 66407, 71853, 76890, 81937, 87124, 92891, 99086, 105569, 112362],
     [29433, 36310, 42954, 47031, 52633, 58205, 62018, 65938, 72614, 77433, 83534, 90010], None, 19089.265090629367, 564225.2276462),
    ("Bradesco (BBDC4 + BBDC3)", "BBDC Model Official 2Q26_vCSStrategy.xlsx", "Income Statement r73 Recurring Net Income",
     "Income Statement r78 Dividends/IoC Provisioned - Net",
     [28221, 31178, 34548, 37971, 41533, 44699, 47964, 51319, 54749, 58243, 61788, 65372],
     [13727, 15095, 16726, 19903, 23431, 26111, 28978, 31517, 34171, 36934, 39801, 42763], None,
     8002.357755462218 + 2037.3454886493955, 226447.75534607),
    ("B3 (B3SA3)", "B3SA3 Model Official 3Q25_vCSStrategy.xlsx", "Income Statement r52 Recurring Net Income",
     "Income Statement r58 Dividends/IoC",
     [6537, 7215, 7971, 8876, 9881, 10905, 11848, 13062],
     [6203, 6859, 7612, 8513, 9413, 10358, 11364, 12526], None, 8337.403451434404, 112688.345),
    ("BTG Pactual (BPAC11)", "BPAC Model Official 3Q25_vCSStrategy.xlsx", "Income Statement r50 Net Income",
     "Income Statement r56 Dividends/IoC",
     [19103, 21961, 25043, 28704, 32446, 36374, 40448, 44253, 48504, 52818, 57507],
     [6686, 9882, 13774, 17222, 21090, 25462, 30336, 35403, 41228, 47536, 54632], None, 6925.218886140881, 305584.144477),
    ("Banco do Brasil (BBAS3)", "BBAS Model Official 2Q26_vCSStrategy.xlsx", "= dividends / payout",
     "Valuation r5 Total Return; payout: Model r26",
     None, [4060, 7114, 8935, 9990, 10846, 13882, 17093, 20522, 24056, 27771, 29140, 30564],
     [0.24, 0.30, 0.30, 0.30, 0.30, 0.35, 0.40, 0.45, 0.50, 0.55, 0.55, 0.55], 5397.318043638426, 138571.5670872),
    ("Caixa Seguridade (CXSE3)", "CXSE3 Model Official 3Q25_vCSStrategy.xlsx", "Income Statement r54 Managerial Net Income (R$ thousand / 1,000)",
     "Valuation r2 Dividend",
     [4565.538, 4903.673, 5256.803, 5652.253, 6018.714, 6421.658, 6858.000, 7330.546, 7842.343, 8396.709, 8997.253],
     [4208, 4520, 4845, 5210, 5547, 5919, 6321, 6756, 7228, 7739, 8293], None, 895.1191230151861, 63090),
    ("Santander (SANB11)", "SANB Model Official 4Q25_vCSStrategy.xlsx", "IS r95 Recurring Net Income", "IS r103 Dividends/IoC Provisioned",
     [17399, 19672, 20985, 22473, 23918, 25510, 27201, 28935, 30589, 32224, 33739],
     [11309, 12787, 14689, 15731, 17939, 19132, 20400, 21702, 22942, 24168, 25304], None, 735.1595758186286, 109308.4474447),
]
ITSA = dict(pts=7020.62358634801, mc=179573.1348910488, ni28=23060)
# Valuation-sheet parameters: ticker, Ke, NTN-B, IPCA, risk free, ERP, beta, g, perpetuity share, last explicit year, value date, index pts
PARAMS = [
    ("ITUB4 + ITSA4", 0.1513, 0.076, 0.04, 0.0986, 0.055, 0.958, 0.06, 0.4623, 2037, "End-2027", 19089.265 + 7020.624),
    ("BBDC4 + BBDC3", 0.1619, 0.076, 0.04, 0.0986, 0.055, 1.15, 0.06, 0.4371, 2037, "End-2027", 8002.358 + 2037.345),
    ("B3SA3", 0.1583, 0.0732, 0.04, 0.0962, 0.055, 1.129, 0.06, 0.4358, 2033, "n/a", 8337.403),
    ("BPAC11", 0.1637, 0.0741, 0.04, 0.0970, 0.055, 1.213, 0.06, 0.5183, 2036, "End-2026", 6925.219),
    ("BBAS3", 0.1809, 0.081, 0.04, 0.1029, 0.06, 1.30, 0.06, 0.4185, 2037, "End-2027", 5397.318),
    ("CXSE3", 0.1386, 0.0751, 0.04, 0.0978, 0.055, 0.742, 0.06, 0.5003, 2036, "End-2026", 895.119),
    ("SANB11", 0.1490, 0.0749, 0.04, 0.0977, 0.05, 1.027, 0.06, 0.4541, 2036, "End-2026", 735.160),
]

bk = Book(OUT, ["DDM Financials", "Box Consensus", "Analyst Models", "Model Parameters"], prefix="Financials - ")
F, LV, IN, NOTE = bk.F, bk.LV, bk.IN, bk.NOTE


def setup(ws):
    ws.set_column("A:A", 2)
    ws.set_column("B:B", 64)
    ws.set_column("C:N", 10.5)
    ws.set_column("O:O", 2)
    ws.set_column("P:P", 90)


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


# ============================================================================= Analyst Models
ws = bk.W["Analyst Models"]
setup(ws)
bk.title(ws, "analysts' projections, aggregated (index points)",
         "Net income and dividends of the XP Financials models (R$ mn, files in " + FOLDER + "). After a model's last year: "
         "net income grows at g and the payout stays at its last value, as the model's own perpetuity. Index points = R$ x "
         "the member's index points / market cap (Members sheet of the main model, Oct-5-2026).")
r = 5
NIP, DVP = [], []                      # rows with net income and dividends in index points
BLK = {}
for name, fn, nisrc, dvsrc, ni, dv, po, pts, mc in MODELS:
    bk.section(ws, r, f"{name}  |  {fn}", "B", NOTE_COL)
    years_header(ws, r + 1)
    rn, rd, rp, rf, rnp, rdp = r + 2, r + 3, r + 4, r + 5, r + 6, r + 7
    BLK[name] = dict(ni=rn, dv=rd, po=rp)
    last = len(dv)                      # number of model years
    # dividends
    cells = {}
    for i, c in enumerate(Y):
        cells[c] = dv[i] if i < last else f"={c}{rn}*{Y[last - 1]}{rp}"
    # net income
    ncells = {}
    for i, c in enumerate(Y):
        if ni is None:
            ncells[c] = f"={c}{rd}/{c}{rp}"
        else:
            ncells[c] = ni[i] if i < len(ni) else f"={Y[i - 1]}{rn}*(1+gAn)"
    pcells = {}
    for i, c in enumerate(Y):
        if po is not None:
            pcells[c] = po[i]
        else:
            pcells[c] = f"={c}{rd}/{c}{rn}"
    put(ws, rn, "Net income (R$ mn)", ncells, "#,##0", f"Model: {nisrc}. " + ("" if ni is None or len(ni) == 12 else f"After {YRS[len(ni) - 1]}: x (1 + g)."))
    put(ws, rd, "Dividends + IoC (R$ mn)", cells, "#,##0", f"Model: {dvsrc}. " + ("" if last == 12 else f"After {YRS[last - 1]}: net income x the {YRS[last - 1]} payout."))
    put(ws, rp, "Payout", pcells, PCT, "Model: Model r26 (input)." if po is not None else "Dividends / net income.")
    ws.write(f"B{rf}", "Index points / market cap (R$ mn)", F())
    ws.write_number(f"C{rf}", pts, F(**IN, num_format=PTS1))
    ws.write_number(f"D{rf}", mc, F(**IN, num_format="#,##0"))
    ws.write_formula(f"E{rf}", f"=C{rf}/D{rf}", F(num_format="0.000000"))
    ws.write(f"{NOTE_COL}{rf}", "C: index points (Members); D: market cap (Members); E: factor that turns R$ mn into index points.", NOTE)
    put(ws, rnp, "Net income (index pts)", {c: f"={c}{rn}*$E${rf}" for c in Y}, PTS1)
    put(ws, rdp, "Dividends (index pts)", {c: f"={c}{rd}*$E${rf}" for c in Y}, PTS1)
    NIP.append(rnp)
    DVP.append(rdp)
    r += 10

# Itaúsa: Itaú's dynamics
it = BLK["Itaú Unibanco (ITUB4)"]
bk.section(ws, r, "Itaúsa (ITSA4)  |  no model: Itaú's net income path and payout (holding of Itaú)", "B", NOTE_COL)
years_header(ws, r + 1)
rn, rd, rp, rf, rnp, rdp = r + 2, r + 3, r + 4, r + 5, r + 6, r + 7
put(ws, rn, "Net income (R$ mn) = consensus FY28 x Itaú's net income / Itaú's 2028",
    {c: f"=$C${rf + 1}*{c}{it['ni']}/$E${it['ni']}" for c in Y}, "#,##0", "Itaú's path scaled to Itaúsa's FY28 consensus (Members).")
put(ws, rd, "Dividends (R$ mn) = net income x Itaú's payout", {c: f"={c}{rn}*{c}{it['po']}" for c in Y}, "#,##0")
put(ws, rp, "Payout", {c: f"={c}{rd}/{c}{rn}" for c in Y}, PCT)
ws.write(f"B{rf}", "Index points / market cap (R$ mn)", F())
ws.write_number(f"C{rf}", ITSA["pts"], F(**IN, num_format=PTS1))
ws.write_number(f"D{rf}", ITSA["mc"], F(**IN, num_format="#,##0"))
ws.write_formula(f"E{rf}", f"=C{rf}/D{rf}", F(num_format="0.000000"))
ws.write(f"B{rf + 1}", "Itaúsa net income FY28, Bloomberg consensus (R$ mn)", F())
ws.write_number(f"C{rf + 1}", ITSA["ni28"], F(**IN, num_format="#,##0"))
put(ws, rnp + 1, "Net income (index pts)", {c: f"={c}{rn}*$E${rf}" for c in Y}, PTS1)
put(ws, rdp + 1, "Dividends (index pts)", {c: f"={c}{rd}*$E${rf}" for c in Y}, PTS1)
NIP.append(rnp + 1)
DVP.append(rdp + 1)
r += 11

# aggregate
bk.section(ws, r, "Aggregate of the modeled members (index points): what the box takes from the analysts", "B", NOTE_COL)
years_header(ws, r + 1)
AG = dict(ni=r + 2, dv=r + 3, gr=r + 4, po=r + 5, cov=r + 6)
put(ws, AG["ni"], "Net income (index pts)", {c: "=" + "+".join(f"{c}{x}" for x in NIP) for c in Y}, PTS1)
put(ws, AG["dv"], "Dividends (index pts)", {c: "=" + "+".join(f"{c}{x}" for x in DVP) for c in Y}, PTS1)
put(ws, AG["gr"], "Net income growth", {c: f"={c}{AG['ni']}/{p}{AG['ni']}-1" for p, c in zip(Y, Y[1:])}, PCT, "Used for the box from 2029.", bold=True)
put(ws, AG["po"], "Payout = dividends / net income", {c: f"={c}{AG['dv']}/{c}{AG['ni']}" for c in Y}, PCT, "Used for the box in every year.", bold=True)
ws.write(f"B{AG['cov']}", "Coverage: index points of the modeled members / Financials box", F())
ws.write_formula(f"C{AG['cov']}", "=(" + "+".join(f"C{x - 1}" for x in NIP[:-1]) + f"+C{rf})/'DDM Financials'!$C$7",
                 F(num_format=PCT))
ws.write(f"{NOTE_COL}{AG['cov']}", "BB Seguridade and Porto Seguro have no model: they take the aggregate path.", NOTE)
ws.freeze_panes(4, 2)

# ============================================================================= Model Parameters
ws = bk.W["Model Parameters"]
bk.title(ws, "valuation parameters of the analysts' models",
         "Valuation sheet of each model. Ke = (NTN-B + long-term IPCA) x 0.85 + beta x ERP; g = long-term IPCA + real GDP (Macro sheet).")
ws.set_column("A:A", 2)
ws.set_column("B:B", 18)
ws.set_column("C:N", 11)
hd = ["Member", "Ke", "NTN-B", "IPCA LT", "Risk free = (NTN-B + IPCA) x 0.85", "ERP", "Beta", "g", "Perpetuity share of value",
      "Last explicit year", "Value date", "Index points"]
for i, h in enumerate(hd):
    ws.write(4, 1 + i, h, bk.HDRL if i == 0 else bk.HDR)
ws.set_row(4, 44)
P0 = 6
for k, p in enumerate(PARAMS):
    rr = P0 + k
    ws.write(f"B{rr}", p[0], F())
    for j, (v, fm) in enumerate(zip(p[1:], [PCT2, PCT2, PCT, PCT2, PCT, "0.000", PCT, PCT, "0", "@", PTS1])):
        c = cn(2 + j)
        if isinstance(v, str):
            ws.write_string(f"{c}{rr}", v, F(**IN, align="center"))
        else:
            ws.write_number(f"{c}{rr}", v, F(**IN, num_format=fm))
PL = P0 + len(PARAMS) - 1
rw = PL + 1
ws.write(f"B{rw}", "Weighted by index points", F(bold=True, top=1))
for c, fm in (("C", PCT2), ("D", PCT2), ("F", PCT2), ("G", PCT2), ("H", "0.000"), ("I", PCT), ("J", PCT)):
    ws.write_formula(f"{c}{rw}", f"=SUMPRODUCT({c}{P0}:{c}{PL},$M${P0}:$M${PL})/SUM($M${P0}:$M${PL})", F(num_format=fm, bold=True, top=1))
for c in ("E", "K", "L", "M"):
    ws.write_blank(f"{c}{rw}", None, F(top=1))
ws.write(f"B{rw + 2}", "Itaúsa takes Itaú's parameters; Bradesco's two share classes are added. Source: Valuation sheet of each file "
         "(Ke, NTN-B, risk free, ERP, beta, g, 'Perpetuity' = PV of the terminal value / fair equity value).", NOTE)
bk.name("BetaAn", "Model Parameters", f"$H${rw}")
bk.name("ERPAn", "Model Parameters", f"$G${rw}")

# ============================================================================= consensus by member (Bloomberg BEST, Oct-7-2026)
BOXF = HERE.parent / "Ibov_house_2027_by_box_XP.xlsx"
_mem = [r for r in openpyxl.load_workbook(BOXF, read_only=True, data_only=True)["Members"].iter_rows(min_row=6, values_only=True)
        if r[0] and r[2] == "Financials"]
_roe = pd.read_parquet(HERE / "dcf_members.parquet")
CONS = []          # name, ticker, pts, mcap, ni[3], roe[3] or None
for r in _mem:
    t = r[0]
    if t == "BBDC3":
        continue
    pts = r[9] + (next(x[9] for x in _mem if x[0] == "BBDC3") if t == "BBDC4" else 0)
    name = f"{r[1]} ({t}{' + BBDC3' if t == 'BBDC4' else ''})"
    roe = None if pd.isna(_roe.loc[t, "roe28"]) else [float(_roe.loc[t, f"roe{y}"]) / 100 for y in (26, 27, 28)]
    CONS.append((name, t, pts, r[10], [r[20], r[21], r[22]], roe))
BOX_NI = [5708.118100817073, 6447.843510049624, 7167.107435011599]     # Box Shares!M7:O7 of the main model
ITUB_ROE = next(x for x in CONS if x[1] == "ITUB4")[5]

# ============================================================================= Box Consensus
ws = bk.W["Box Consensus"]
bk.title(ws, "consensus 2026-28 of the members, summed into the box (index points)",
         "Bloomberg BEST consensus (Oct-7-2026): net income (BEST_NET_INCOME, Members sheet of the main model) and ROE "
         "(BEST_ROE, dcf_members.parquet). Book value = net income / ROE. Index points = R$ mn x points / market cap. "
         "Only the actuals and consensus years are built from the members; the projection is done for the box.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 34)
ws.set_column("C:U", 10.5)
hd = ["Member", "Index points", "Market cap (R$ mn)", "Points / market cap", "Net income 2026 (R$ mn)", "Net income 2027 (R$ mn)",
      "Net income 2028 (R$ mn)", "ROE 2026", "ROE 2027", "ROE 2028", "Net income 2026 (pts)", "Net income 2027 (pts)",
      "Net income 2028 (pts)", "Book value 2026 (pts)", "Book value 2027 (pts)", "Book value 2028 (pts)"]
for i, h in enumerate(hd):
    ws.write(5, 1 + i, h, bk.HDRL if i == 0 else bk.HDR)
ws.set_row(5, 44)
C0 = 7
for k, (name, t, pts, mc, ni, roe) in enumerate(CONS):
    rr = C0 + k
    ws.write(f"B{rr}", name, F())
    ws.write_number(f"C{rr}", pts, F(**IN, num_format=PTS1))
    ws.write_number(f"D{rr}", mc, F(**IN, num_format="#,##0"))
    ws.write_formula(f"E{rr}", f"=C{rr}/D{rr}", F(num_format="0.000000"))
    for j, c in enumerate("FGH"):
        ws.write_number(f"{c}{rr}", ni[j], F(**IN, num_format="#,##0"))
    for j, c in enumerate("IJK"):
        if roe is None:
            ws.write_number(f"{c}{rr}", ITUB_ROE[j], F(**LV, num_format=PCT))
        else:
            ws.write_number(f"{c}{rr}", roe[j], F(**IN, num_format=PCT))
    for j, (c, src) in enumerate(zip("LMN", "FGH")):
        ws.write_formula(f"{c}{rr}", f"={src}{rr}*$E{rr}", F(num_format=PTS1))
    for j, (c, n, ro) in enumerate(zip("OPQ", "LMN", "IJK")):
        ws.write_formula(f"{c}{rr}", f"={n}{rr}/{ro}{rr}", F(num_format=PTS1))
CL = C0 + len(CONS) - 1
CT = CL + 1
ws.write(f"B{CT}", "Financials box", F(bold=True, top=1))
ws.write_formula(f"C{CT}", f"=SUM(C{C0}:C{CL})", F(num_format=PTS1, bold=True, top=1))
for c in "DEFGH":
    ws.write_blank(f"{c}{CT}", None, F(top=1))
for c, n, b in zip("IJK", "LMN", "OPQ"):
    ws.write_formula(f"{c}{CT}", f"={n}{CT}/{b}{CT}", F(num_format=PCT, bold=True, top=1))
for c in "LMNOPQ":
    ws.write_formula(f"{c}{CT}", f"=SUM({c}{C0}:{c}{CL})", F(num_format=PTS1, bold=True, top=1))
ws.write(f"B{CT + 1}", "Check: net income vs. Box Shares!M7:O7 of the main model", F())
for c, v in zip("LMN", BOX_NI):
    ws.write_formula(f"{c}{CT + 1}", f"={c}{CT}-{v}", F(num_format=PTS1))
notes = ["Box ROE = sum of net income / sum of book value (row total).",
         "Itaúsa: no BEST_ROE for the holding; Itaú's consensus ROE as a proxy (pink, editable).",
         "Bradesco: BBDC4 and BBDC3 points added (same company figures)."]
for i, n in enumerate(notes):
    ws.write(f"B{CT + 3 + i}", n, NOTE)
ws.freeze_panes(6, 2)
BC = "'Box Consensus'!"

# ============================================================================= DDM Financials
ws = bk.W["DDM Financials"]
setup(ws)
ws.set_column("C:F", 13)
bk.title(ws, "FCFE / bank DDM, the analysts' logic on the box consensus (end-2027)",
         "Consensus 2026-28 summed from the members (Box Consensus); the projection is done for the box: ROE held, growth "
         "fading to g, cash to shareholders = net income - the increase in book value; Gordon perpetuity on the last explicit year.")

# ---- 1. Inputs
bk.section(ws, 5, "1. Inputs", "B", NOTE_COL)
bk.header_row(ws, 6, [("B", "Item"), ("C", "Value"), ("D", ""), (NOTE_COL, "Source / note")])
put(ws, 7, "Financials, index points today", {"C": f"={BC}C{CT}"}, PTS1,
    "Members' weight x Ibovespa close of Oct-5-2026 (Box Consensus; Box Shares!C7 of the main model).")
put(ws, 8, "Long-term IPCA", {"C": 0.04}, PCT, "Analysts' Macro sheet and the house (Target!B63).", style=LV)
put(ws, 9, "Long-term real GDP growth", {"C": 0.02}, PCT, "Analysts' Macro sheet.", style=LV)
put(ws, 10, "Share of the risk free kept after tax", {"C": 0.85}, "0.00", "As the analysts: risk free = (NTN-B + IPCA) x 0.85 (15% income tax on fixed income).", style=LV)
put(ws, 11, "Earnings beta x normalization (domestic cycle)", {"C": 0.732791, "D": 1.1868587775150152}, "0.000",
    "Box Earnings!I7 and I16 of the main model: shock of Financials = C x D x the Ibovespa ex-Commodities shock.")
put(ws, 12, "Projection: 1 = box consensus with the analysts' logic; 2 = the analysts' own growth and payout (memo)", {"C": 1}, "0",
    "2 applies the analysts' aggregate growth and payout (Analyst Models) to the box consensus of 2026-28.", style=LV)
put(ws, 13, "Box ROE from 2029 (blank = consensus FY28 of the box)", {}, PCT, "Typed value replaces the FY28 box ROE after 2028.", style=LV)
ws.write_blank("C13", None, F(**LV, num_format=PCT))
put(ws, 14, "Equity risk premium, the same for every box", {"C": 0.055}, PCT2,
    "Caio, Oct-8-2026: 5.5%, the standard of most XP Financials models. Implied ERP of the Ibovespa in this convention: 4.8% (10y average), 4.2% now.", style=LV)

# ---- 2. House vs. analysts
bk.section(ws, 15, "2. House vs. analysts, variable by variable (choice: 1 = house, 2 = analysts)", "B", NOTE_COL)
bk.header_row(ws, 16, [("B", "Variable"), ("C", "House"), ("D", "Analysts"), ("E", "Choice"), ("F", "Used"), (NOTE_COL, "House / analysts")])
ws.set_row(15, 30)
HV = dict(beta=17, g=18, n=19, tv=20)
CHOICE = dict(beta=1, g=2, n=1, tv=2)     # Caio, Oct-8-2026: g and the TV discount as the analysts, the rest as the house
put(ws, HV["beta"], "Beta", {"C": 1.149063658384, "D": "=BetaAn"}, "0.000",
    "House: the box beta, weekly 5 years vs. the Ibovespa from the local box index (Box Betas). Analysts: each model's beta, weighted.")
put(ws, HV["g"], "Perpetuity growth (g)", {"C": 0.047, "D": "=C8+C9"}, PCT, "House: 4.7% (Target!B54). Analysts: long-term IPCA + real GDP.")
put(ws, HV["n"], "Explicit years after the value date", {"C": 4, "D": 10}, "0",
    "House: 2028-31 (4 years after end-2027). Analysts: 10-11 years (to 2036-37). Growth fades to g by the last explicit year.")
put(ws, HV["tv"], "Extra periods in the discount of the terminal value", {"C": 1, "D": 0}, "0",
    "House: the TV is discounted one period more than the last flow (Target sheet, NPV layout). Analysts: the last year's factor.")
for k, r in HV.items():
    ws.write_number(f"E{r}", CHOICE[k], F(**LV, num_format="0", align="center"))
    fmt = {"beta": "0.000", "g": PCT, "n": "0", "tv": "0"}[k]
    ws.write_formula(f"F{r}", f"=IF(E{r}=1,C{r},D{r})", F(num_format=fmt, bold=True))
# the house / analysts inputs that are typed numbers should look like assumptions
for r, c, v, fm in ((HV["beta"], "C", 1.149063658384, "0.000"), (HV["g"], "C", 0.047, PCT),
                    (HV["n"], "C", 4, "0"), (HV["n"], "D", 10, "0"), (HV["tv"], "C", 1, "0"), (HV["tv"], "D", 0, "0")):
    ws.write_number(f"{c}{r}", v, F(**LV, num_format=fm))
bk.name("g_", "DDM Financials", f"$F${HV['g']}")
bk.name("gAn", "DDM Financials", f"$D${HV['g']}")
bk.name("NExp", "DDM Financials", f"$F${HV['n']}")
bk.name("TVx", "DDM Financials", f"$F${HV['tv']}")

# ---- 3. Ke
bk.section(ws, 29, "3. Cost of equity by scenario", "B", NOTE_COL)
bk.header_row(ws, 30, [("B", "Item"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), (NOTE_COL, "How it is calculated")])
V = "CDE"
put(ws, 31, "5-year real rate (NTN-B)", {"C": 0.085, "D": 0.0676, "E": 0.055}, PCT2, "Same scenarios as the other boxes.", style=LV)
put(ws, 32, "Risk free, nominal = (NTN-B + long-term IPCA) x share kept after tax", {c: f"=({c}31+$C$8)*$C$10" for c in V}, PCT2,
    "Nominal, because the flows are nominal (consensus in current R$, growth with inflation). Not a choice.")
put(ws, 33, "Ke = risk free + beta x ERP", {c: f"={c}32+$F${HV['beta']}*$C$14" for c in V}, PCT2, bold=True)
put(ws, 34, "Ke - g", {c: f"=({c}33-g_)*100" for c in V}, '0.0" pp"')

# ---- 4. Projection
bk.section(ws, 36, "4. Box projection (index points): consensus 2026-28, box-level projection after 2028", "B", NOTE_COL)
years_header(ws, 37)
E28, N37 = YC[2028], YC[2037]
P = dict(t=38, ni=39, gr=40, roe=41, bv=42, dv=43, po=44, ang=45, anp=46, anni=47, andv=48, use=49)
put(ws, P["t"], "Years after end-2027 (1 = explicit, grey = after the explicit years)",
    {c: f"={YRS[i]}-2027" for i, c in enumerate(Y) if YRS[i] >= 2028}, "0", "Flows after the last explicit year are not used.")
ws.conditional_format(f"{E28}{P['t']}:{N37}{P['t']}", {"type": "formula", "criteria": f"={E28}${P['t']}>NExp",
                                                         "format": bk.wb.add_format({"font_color": "#A6A6A6"})})
ni = {c: f"={BC}{s}{CT}" for c, s in zip(Y[:3], "LMN")}
ni.update({c: f"={p}{P['ni']}*(1+{c}{P['gr']})" for p, c in zip(Y[2:], Y[3:])})
put(ws, P["ni"], "Net income: 2026-28 consensus (sum of the members); then previous x (1 + growth)", ni, PTS1, "Box Consensus.", bold=True)
gr = {c: f"={c}{P['ni']}/{p}{P['ni']}-1" for p, c in zip(Y[:2], Y[1:3])}
gr.update({c: f"=IF({c}{P['t']}>=NExp,g_,${E28}${P['gr']}+(g_-${E28}${P['gr']})*({c}{P['t']}-1)/(NExp-1))" for c in Y[3:]})
put(ws, P["gr"], "Net income growth: 2027-28 consensus; then linear from the 2028 growth to g in the last explicit year", gr, PCT,
    "As the analysts: growth slows to about the perpetuity growth by the last explicit year.")
roe = {c: f"={BC}{s}{CT}" for c, s in zip(Y[:3], "IJK")}
roe.update({c: f"=IF(ISNUMBER($C$13),$C$13,${E28}${P['roe']})" for c in Y[3:]})
put(ws, P["roe"], "ROE: 2026-28 consensus of the box; then held at FY28", roe, PCT,
    "Box ROE = sum of net income / sum of book value. The analysts keep ROE about flat (ITUB 25%, BBDC 17%).")
bv = {c: f"={BC}{s}{CT}" for c, s in zip(Y[:3], "OPQ")}
bv.update({c: f"={c}{P['ni']}/{c}{P['roe']}" for c in Y[3:]})
put(ws, P["bv"], "Book value: 2026-28 consensus (net income / ROE, summed); then net income / ROE", bv, PTS1)
put(ws, P["dv"], "Cash to shareholders = net income - increase in book value (dividends + buybacks)",
    {c: f"={c}{P['ni']}-({c}{P['bv']}-{p}{P['bv']})" for p, c in zip(Y, Y[1:])}, PTS1,
    "FCFE of a bank; as the analysts' equity roll-forward (equity = previous + net income - dividends).", bold=True)
put(ws, P["po"], "Payout", {c: f"={c}{P['dv']}/{c}{P['ni']}" for c in Y[1:]}, PCT, "Converges to about 1 - g / ROE as growth reaches g.")
put(ws, P["ang"], "Memo: analysts' aggregate net income growth", {c: f"='Analyst Models'!{c}{AG['gr']}" for c in Y[1:]}, PCT, "Analyst Models.")
put(ws, P["anp"], "Memo: analysts' aggregate payout", {c: f"='Analyst Models'!{c}{AG['po']}" for c in Y}, PCT, "Analyst Models.")
an = {c: f"={c}{P['ni']}" for c in Y[:3]}
an.update({c: f"={p}{P['anni']}*(1+{c}{P['ang']})" for p, c in zip(Y[2:], Y[3:])})
put(ws, P["anni"], "Memo: net income on the analysts' path (consensus 2026-28, then their growth)", an, PTS1)
put(ws, P["andv"], "Memo: dividends on the analysts' path", {c: f"={c}{P['anni']}*{c}{P['anp']}" for c in Y}, PTS1)
put(ws, P["use"], "Cash to shareholders used (switch in row 12)", {c: f"=IF($C$12=1,{c}{P['dv']},{c}{P['andv']})" for c in Y[1:]}, PTS1, bold=True)
U, T = P["use"], P["t"]
FL = f"${E28}${U}:${N37}${U}"
TT = f"${E28}${T}:${N37}${T}"


def pv(ke, flows=FL):
    return f"SUMPRODUCT({flows},--({TT}<=NExp),1/(1+{ke})^{TT})"


def last(flows=FL):
    return f"INDEX({flows},NExp)"


# ---- 5. Valuation
bk.section(ws, 51, "5. Valuation at the end of 2027", "B", NOTE_COL)
bk.header_row(ws, 52, [("B", "Item"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), (NOTE_COL, "How it is calculated")])
put(ws, 53, "Ke", {c: f"={c}33" for c in V}, PCT2)
put(ws, 54, "Earnings shock of the Ibovespa ex-Commodities", {"C": -0.2, "D": 0, "E": 0.2}, UPS, "Symmetric +/-20%, as the other boxes.", style=LV)
put(ws, 55, "Earnings shock of Financials = beta x normalization x shock", {c: f"=$C$11*$D$11*{c}54" for c in V}, UPS,
    "Moves the cash to shareholders of every year.")
put(ws, 56, "PV of the explicit flows = (1 + shock) x sum of flow / (1 + Ke) ^ t", {c: f"=(1+{c}55)*{pv(c + '53')}" for c in V}, PTS1,
    "t = 1 for 2028 (value at end-2027); only the explicit years.")
put(ws, 57, "Terminal value = last explicit flow x (1 + shock) x (1 + g) / (Ke - g)", {c: f"={last()}*(1+{c}55)*(1+g_)/({c}53-g_)" for c in V},
    PTS1, "Gordon on the last explicit flow.")
put(ws, 58, "PV of the terminal value = TV / (1 + Ke) ^ (explicit years + extra periods)", {c: f"={c}57/(1+{c}53)^(NExp+TVx)" for c in V}, PTS1)
put(ws, 59, "Fair value of Financials at the end of 2027 (index pts)", {c: f"={c}56+{c}58" for c in V}, PTS1, bold=True)
put(ws, 60, "Upside vs. today", {c: f"={c}59/$C$7-1" for c in V}, UPS, bold=True)
put(ws, 61, "Perpetuity share of the value", {c: f"={c}58/{c}59" for c in V}, PCT, "Analysts' models: 42-52% (Model Parameters).")
A = f"${E28}${P['andv']}:${N37}${P['andv']}"
put(ws, 62, "Memo: upside with the analysts' own growth and payout (row 48), same choices",
    {c: f"=((1+{c}55)*{pv(c + '53', A)}+{last(A)}*(1+{c}55)*(1+g_)/({c}53-g_)/(1+{c}53)^(NExp+TVx))/$C$7-1" for c in V}, UPS)
put(ws, 63, "Memo: upside in the unified DCF by box (main model, v7b)", {"C": -0.3543475494311503, "D": -0.0610, "E": 0.2880808349631869},
    UPS, "Ibov_house_2027_by_box_XP.xlsx, DCF by Box row 63 (FCFE, 4 explicit years, house g and rate).")

# ---- 6. Sensitivity
bk.section(ws, 65, "6. Base upside (no earnings shock): Ke (rows) x perpetuity g (columns); flows as projected", "B", NOTE_COL)
gs = [0.04, 0.047, 0.05, 0.055, 0.06]
kes = [0.13, 0.14, 0.15, 0.1547, 0.16, 0.17]
ws.write("B66", "Ke \\ g", bk.HDRL)
for c, gg in zip(Y, gs):
    ws.write_number(f"{c}66", gg, F(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", num_format=PCT))
for i, k in enumerate(kes):
    rr = 67 + i
    ws.write_number(f"B{rr}", k, F(bold=True, num_format=PCT2, align="left"))
    for c in Y[:len(gs)]:
        ws.write_formula(f"{c}{rr}", f"=({pv(f'$B{rr}')}+{last()}*(1+{c}$66)/($B{rr}-{c}$66)/(1+$B{rr})^(NExp+TVx))/$C$7-1", F(num_format=UPS))
ws.conditional_format(f"C67:{Y[len(gs) - 1]}72", {"type": "3_color_scale", "min_color": "#F8696B", "mid_type": "num", "mid_value": 0,
                                                  "mid_color": "#FFEB84", "max_color": "#63BE7B"})
ws.write("B74", "15.47% = Ke in the base ((6.76% + 4%) x 0.85 + 1.149 x 5.5%). 4.7% = house g; 6.0% = analysts' g.", NOTE)
r0 = 78
bk.section(ws, r0, "7. What is fixed (not switchable) and where it comes from", "B", NOTE_COL)
fixed = [("Risk free nominal: (NTN-B + IPCA) x 0.85", "Required by the nominal flows; the tax haircut as the analysts"),
         ("Equity risk premium 5.5%, the same for every box", "Unified (Caio); standard of most XP Financials models"),
         ("Content of 2026-28: Bloomberg consensus of the members, summed", "Strategy (unified with the other boxes)"),
         ("Projection for the box only: ROE held at FY28, growth fading to g", "Analysts' logic"),
         ("Cash to shareholders = net income - increase in book value", "Analysts' equity roll-forward (= house FCFE)"),
         ("Gordon perpetuity on the last explicit flow", "Both"),
         ("Scenario real rates 8.5% / today / 5.5%", "House (unified with the other boxes)"),
         ("Earnings shock +/-20% x earnings beta", "Strategy (unified with the other boxes)"),
         ("Value at end-2027, compared with today's points", "House")]
for i, (a, b) in enumerate(fixed):
    ws.write(f"B{r0 + 1 + i}", a, F())
    ws.write(f"{NOTE_COL}{r0 + 1 + i}", b, NOTE)
ws.freeze_panes(4, 2)
ws.activate()
bk.close()
print("saved", OUT)
