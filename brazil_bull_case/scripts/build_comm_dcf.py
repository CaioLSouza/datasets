"""Commodities box: FCFE with the logic of the XP commodity analysts' models and the content of the consensus.

Cash to shareholders = net income + D&A + capex (negative) - lease principal + change in working capital, as the analysts'
FCFF / FCFE bridges (Vale, Suzano, Petrobras: NOPAT + D&A - capex - NWC, minus leases for Petrobras), with net debt held
constant (FCFE at the unified Ke instead of FCFF at a WACC). Consensus 2026-28 summed from the members (Bloomberg BEST
sales, EBITDA, EBIT, net income, capex; Members sheet of the main model); lease principal of Petrobras and Suzano from the
analysts' models. The projection is done for the box only: every line grows with net income, whose growth goes from the
box's 2028 growth to g by the last explicit year; g = (1 + IPCA) x (1 + real growth) - 1 with zero real growth (Caio,
Oct-8-2026), as the analysts' perpetuities (Suzano g = IPCA, Vale g = US CPI). Ke: (NTN-B + IPCA) x 0.85 + beta x 5.5%,
unified with the other boxes. Scenarios: real rate 8.5% / today / 5.5%; commodity prices +/-20% x the earnings elasticity
(1.26), applied to net income (capex and leases do not move with prices).
Usage: python build_comm_dcf.py [output file name]
"""
import sys
from pathlib import Path

import openpyxl
from xlsxwriter.utility import xl_col_to_name as cn

from xp_template import Book, PCT, PCT2, PTS1, UPS, NAVY

HERE = Path(__file__).parent
OUT = HERE.parent / (sys.argv[1] if len(sys.argv) > 1 else "Commodities_DCF_XP.xlsx")
BOXF = HERE.parent / "Ibov_house_2027_by_box_XP.xlsx"
YRS = list(range(2026, 2038))
Y = [cn(2 + i) for i in range(12)]          # C .. N
YC = dict(zip(YRS, Y))
E28, N37 = YC[2028], YC[2037]
NOTE_COL = "P"

# ----------------------------------------------------------------------------- consensus by member (R$ mn)
_mem = [r for r in openpyxl.load_workbook(BOXF, read_only=True, data_only=True)["Members"].iter_rows(min_row=6, values_only=True)
        if r[0] and r[2] == "Commodities"]
byt = {r[0]: r for r in _mem}
MET = ["sales", "ebitda", "ebit", "ni", "capex", "lease"]
IDX = {"sales": 11, "ebit": 14, "ebitda": 17, "ni": 20, "capex": 23}
# lease principal (R$ mn, negative): Petrobras = lease payments - lease interest (Consolidated r250 / r162, US$ mn, x FX);
# Suzano = '(-) Leasing' of its FCFE bridge (DCF & Sensitivity r15)
PETR_LEASE_USD = [-(10129 - 2806), -(10127 - 2724), -(10639 - 2873)]
SUZB_LEASE = [-1389.9043516091242, -1320.655140407495, -1269.770748497096]
HOLD = {"GOAU4": "GGBR4", "BRAP4": "VALE3"}      # holdings: look-through at the subsidiary's ratios to net income


def line(r, m):
    return [r[IDX[m] + i] for i in range(3)]


CONS = []          # name, ticker, pts, mcap, {metric: [3]}, {metric: [3 flags]} (flag = proxy / assumption)
for r in sorted(_mem, key=lambda x: -x[9]):
    t = r[0]
    if t == "PETR3":
        continue
    pts = r[9] + (byt["PETR3"][9] if t == "PETR4" else 0)
    name = f"{r[1]} ({t}{' + PETR3' if t == 'PETR4' else ''})"
    d = {m: (line(r, m) if m in IDX else [0.0, 0.0, 0.0]) for m in MET}
    fl = {m: [False] * 3 for m in MET}
    if t == "PETR4":
        d["lease"] = ["=PETRFX*" + str(x) for x in PETR_LEASE_USD]
    if t == "SUZB3":
        d["lease"] = SUZB_LEASE
    if t in HOLD:
        s = byt[HOLD[t]]
        ni = line(r, "ni")
        for i in range(3):
            if ni[i] is None:                      # GOAU 2028: Gerdau's growth
                ni[i] = ni[i - 1] * s[IDX["ni"] + i] / s[IDX["ni"] + i - 1]
                fl["ni"][i] = True
        d["ni"] = ni
        for m in ("sales", "ebitda", "ebit", "capex"):
            d[m] = [ni[i] * s[IDX[m] + i] / s[IDX["ni"] + i] for i in range(3)]
            fl[m] = [True] * 3
    CONS.append((name, t, pts, r[10], d, fl))
BOX_NI = [10671.972549948501, 9633.489419041834, 9381.246397941843]     # Box Shares!M10:O10 of the main model

bk = Book(OUT, ["DCF Commodities", "Box Consensus", "Analyst Models"], prefix="Commodities - ")
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


# ============================================================================= Box Consensus
ws = bk.W["Box Consensus"]
bk.title(ws, "consensus 2026-28 of the members, summed into the box",
         "Bloomberg BEST consensus (Oct-7-2026, company totals in R$ mn; Members sheet of the main model). Lease principal: "
         "the analysts' models. Index points = R$ mn x points / market cap. Only 2026-28 are built from the members.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 30)
ws.set_column("C:W", 9.5)
LAB = {"sales": "Sales", "ebitda": "EBITDA", "ebit": "EBIT", "ni": "Net income", "capex": "Capex", "lease": "Lease principal"}
MC = {m: [cn(5 + 3 * k + i) for i in range(3)] for k, m in enumerate(MET)}      # F.. W
ws.write("B5", "FX for Petrobras' lease (R$ / US$)", F())
ws.write_number("C5", 5.15, F(**LV, num_format="0.00"))
ws.write("D5", "Analysts' 2027 average (Vale, Suzano Macro sheets).", NOTE)
bk.name("PETRFX", "Box Consensus", "$C$5")


def table(r_hdr, title, pts_mode):
    bk.section(ws, r_hdr - 1, title, "B", "W")
    hd = [("B", "Member"), ("C", "Index points"), ("D", "Market cap (R$ mn)"), ("E", "Points / market cap")]
    for m in MET:
        for i, c in enumerate(MC[m]):
            hd.append((c, f"{LAB[m]} {2026 + i}"))
    bk.header_row(ws, r_hdr, hd)
    ws.set_row(r_hdr - 1, 30)
    ws.set_row(r_hdr - 1 + 1, 44)


R1 = 8                           # R$ table header
table(R1, "1. Consensus by member (R$ mn)", False)
r = R1 + 1
ROWS1 = {}
for name, t, pts, mc, d, fl in CONS:
    ws.write(f"B{r}", name, F())
    ws.write_number(f"C{r}", pts, F(**IN, num_format=PTS1))
    ws.write_number(f"D{r}", mc, F(**IN, num_format="#,##0"))
    ws.write_formula(f"E{r}", f"=C{r}/D{r}", F(num_format="0.000000"))
    for m in MET:
        for i, c in enumerate(MC[m]):
            v = d[m][i]
            sty = LV if fl[m][i] or (m == "lease" and t not in ("PETR4", "SUZB3")) else IN
            if isinstance(v, str):
                ws.write_formula(f"{c}{r}", v, F(num_format="#,##0"))
            else:
                ws.write_number(f"{c}{r}", v if v is not None else 0, F(**sty, num_format="#,##0"))
    ROWS1[t] = r
    r += 1
L1 = r - 1
R2 = r + 3
table(R2, "2. In index points (= R$ mn x points / market cap) and the box total", True)
r = R2 + 1
for name, t, pts, mc, d, fl in CONS:
    r1 = ROWS1[t]
    ws.write(f"B{r}", name, F())
    ws.write_formula(f"C{r}", f"=C{r1}", F(num_format=PTS1))
    for m in MET:
        for c in MC[m]:
            ws.write_formula(f"{c}{r}", f"={c}{r1}*$E${r1}", F(num_format=PTS1))
    r += 1
L2 = r - 1
CT = r
ws.write(f"B{CT}", "Commodities box", F(bold=True, top=1))
ws.write_formula(f"C{CT}", f"=SUM(C{R2 + 1}:C{L2})", F(num_format=PTS1, bold=True, top=1))
for m in MET:
    for c in MC[m]:
        ws.write_formula(f"{c}{CT}", f"=SUM({c}{R2 + 1}:{c}{L2})", F(num_format=PTS1, bold=True, top=1))
ws.write(f"B{CT + 1}", "Check: net income vs. Box Shares!M10:O10", F())
for c, v in zip(MC["ni"], BOX_NI):
    ws.write_formula(f"{c}{CT + 1}", f"={c}{CT}-{v}", F(num_format=PTS1))
notes = ["Pink = assumption / proxy. Holdings (Metalurgica Gerdau, Bradespar): sales, EBITDA, EBIT and capex at the subsidiary's ratios "
         "to net income (Gerdau, Vale); Metalurgica Gerdau's 2028 net income grows as Gerdau's (the main model had 0, hence the check).",
         "Lease principal: Petrobras = lease payments - interest on lease liabilities (analyst model, Consolidated r250 and r162, US$ mn x FX); "
         "Suzano = '(-) Leasing' of the analyst's FCFE bridge. Other members: 0 until the Bloomberg pull (Vibra, Ultrapar, Brava and PRIO have IFRS 16 leases).",
         "Petrobras: PETR4 and PETR3 points added (same company figures)."]
for i, n in enumerate(notes):
    ws.write(f"B{CT + 3 + i}", n, NOTE)
ws.freeze_panes(R1, 2)
BC = "'Box Consensus'!"


def box(m, i):
    return f"{BC}{MC[m][i]}{CT}"


# ============================================================================= Analyst Models (memo)
ws = bk.W["Analyst Models"]
bk.title(ws, "how the XP commodity analysts treat the long term (memo)",
         "Files in inputs/analyst_models/Commodities (DCF & Sensitivity, Valuation, Control Panel sheets), read on Oct-8-2026.")
ws.set_column("A:A", 2)
ws.set_column("B:B", 22)
ws.set_column("C:K", 15)
hd = ["Model", "Currency", "Method", "Value date", "Explicit to", "Perpetuity g", "Risk free", "Beta x ERP", "Ke / WACC",
      "Long-term price"]
for i, h in enumerate(hd):
    ws.write(4, 1 + i, h, bk.HDRL if i == 0 else bk.HDR)
ws.set_row(4, 30)
rows = [("Vale (VALE3)", "USD", "FCFF - net debt, SOTP", "End-2027", "2035", "2.4% (US CPI, real 0%)",
         "(real 70% BR 7.8% + 30% US 2.55%, + US CPI) x 0.85 = 7.5%", "1.0 x 5.5%", "13.0% / 11.3%",
         "Iron ore US$100 (26YE) > 97.5 > 93.75 > 90/t from 2029; copper 14k > 15k (2030)"),
        ("Petrobras (PETR4)", "USD", "FCFF - net debt, by business", "(older vintage)", "2031 (EBITDA to 2060)",
         "2.4% on an adjusted 2031 FCFF", "Treasury 4.0% + CDS 2.1%", "1.53 (levered) x 5.0%", "13.7% / 12.1%",
         "Brent US$63/bbl; EBITDA falls with reserves: -6% a year real 2028-35"),
        ("Suzano (SUZB3)", "BRL", "FCFE (FCFF check)", "End-2027", "2035", "4.0% (IPCA, real 0%)",
         "(real 90% BR 7.5% + 10% US, + IPCA) x 0.85 = 9.6%", "1.0 x 5.5%", "15.1% / 11.2%", "BHKP real US$590 > 570 > 560/t from 2028"),
        ("Aura (AURA33)", "USD", "FCFF, SOTP by mine", "End-2026", "Mine lives (to ~2047)", "None (finite reserves)",
         "Same rule, USD = 8.5%", "0.5 x 5.5%", "11.2% / 10.5%", "Gold US$4,500 (26-27) > 3,500 by 2035; copper 14k")]
for k, row in enumerate(rows):
    for i, v in enumerate(row):
        ws.write(5 + k, 1 + i, v, F(text_wrap=True, valign="top", **({} if i else {"bold": True})))
    ws.set_row(5 + k, 44)
r0 = 11
ws.write(f"B{r0}", "Real growth after 2028 in the analysts' models (a year, 2028-35)", F(bold=True))
growth = [("Vale", "NOPLAT, USD", 0.063, 0.038, "Copper and nickel growth; 28.8% of the box"),
          ("Petrobras", "EBITDA, USD", -0.038, -0.061, "Reserve depletion; 43.0% of the box"),
          ("Suzano", "NOPLAT, BRL", 0.047, 0.007, "Real pulp price flat from 2028; 2.8% of the box")]
for i, h in enumerate(["Model", "Line", "Nominal", "Real", "Note"]):
    ws.write(r0, 1 + i, h, bk.HDRL if i == 0 else bk.HDR)
for k, (a, b, c, d, e) in enumerate(growth):
    rr = r0 + 2 + k
    ws.write(f"B{rr}", a, F())
    ws.write(f"C{rr}", b, F())
    ws.write_number(f"D{rr}", c, F(**IN, num_format=UPS))
    ws.write_number(f"E{rr}", d, F(**IN, num_format=UPS))
    ws.write(f"F{rr}", e, NOTE)
ws.write(f"B{r0 + 6}", "Weighted by the box points, the three give about -2% a year real: zero real growth for the box is neutral to slightly generous.",
         NOTE)


# ============================================================================= DCF Commodities
ws = bk.W["DCF Commodities"]
setup(ws)
ws.set_column("C:F", 13)
bk.title(ws, "FCFE with zero real growth, the analysts' logic on the box consensus (end-2027)",
         "Cash to shareholders = net income + D&A + capex - lease principal + change in working capital (net debt constant). "
         "Consensus 2026-28 summed from the members; projection for the box; Gordon perpetuity with zero real growth.")
bk.section(ws, 5, "1. Inputs", "B", NOTE_COL)
bk.header_row(ws, 6, [("B", "Item"), ("C", "Value"), ("D", ""), (NOTE_COL, "Source / note")])
put(ws, 7, "Commodities, index points today", {"C": f"={BC}C{CT}"}, PTS1,
    "Members' weight x Ibovespa close of Oct-5-2026 (Box Shares!C10 of the main model).")
put(ws, 8, "Long-term IPCA", {"C": 0.04}, PCT, "House (Target!B63) and the analysts' Macro sheets.", style=LV)
put(ws, 9, "Long-term real growth of the box", {"C": 0.0}, PCT,
    "Caio, Oct-8-2026: zero, as the analysts' commodity perpetuities (g = inflation) or finite lives.", style=LV)
put(ws, 10, "Share of the risk free kept after tax", {"C": 0.85}, "0.00", "As the analysts: risk free = (NTN-B + IPCA) x 0.85.", style=LV)
put(ws, 11, "Equity risk premium, the same for every box", {"C": 0.055}, PCT2, "Unified (Caio, Oct-8-2026).", style=LV)
put(ws, 12, "Earnings elasticity to commodity prices", {"C": 1.262}, "0.000",
    "Box Earnings!J10 of the main model: 12-month change of the box 12m fwd EPS vs. the DJP commodity ETN, Oct-07 to Sep-26.")
put(ws, 13, "Working capital, % of sales", {"C": 0.04}, PCT, "House assumption (Assumptions!C15 of the main model).", style=LV)

bk.section(ws, 15, "2. House vs. analysts, variable by variable (choice: 1 = house, 2 = analysts)", "B", NOTE_COL)
bk.header_row(ws, 16, [("B", "Variable"), ("C", "House"), ("D", "Analysts"), ("E", "Choice"), ("F", "Used"), (NOTE_COL, "House / analysts")])
ws.set_row(15, 30)
HV = dict(beta=17, g=18, n=19, tv=20)
CHOICE = dict(beta=1, g=2, n=1, tv=2)
put(ws, HV["beta"], "Beta", {"C": 0.872, "D": 1.0}, "0.000",
    "House: the box beta, weekly 5 years vs. the Ibovespa from the local box index (Box Betas). Analysts: 1.0 (Vale, Suzano).", style=LV)
put(ws, HV["g"], "Perpetuity growth (g)", {"C": 0.047, "D": "=(1+C8)*(1+C9)-1"}, PCT,
    "House: 4.7% (Target!B54). Analysts (commodities): inflation x (1 + real growth), real growth 0%.")
ws.write_number(f"C{HV['g']}", 0.047, F(**LV, num_format=PCT))
put(ws, HV["n"], "Explicit years after the value date", {"C": 4, "D": 8}, "0",
    "House: 2028-31. Analysts: to 2035 (Vale, Suzano). Growth reaches g by the last explicit year.", style=LV)
put(ws, HV["tv"], "Extra periods in the discount of the terminal value", {"C": 1, "D": 0}, "0",
    "House: the TV is discounted one period more than the last flow. Analysts: the last year's factor.", style=LV)
for k, r in HV.items():
    ws.write_number(f"E{r}", CHOICE[k], F(**LV, num_format="0", align="center"))
    ws.write_formula(f"F{r}", f"=IF(E{r}=1,C{r},D{r})", F(num_format={"beta": "0.000", "g": PCT, "n": "0", "tv": "0"}[k], bold=True))
bk.name("g_", "DCF Commodities", f"$F${HV['g']}")
bk.name("NExp", "DCF Commodities", f"$F${HV['n']}")
bk.name("TVx", "DCF Commodities", f"$F${HV['tv']}")

bk.section(ws, 22, "3. Cost of equity by scenario", "B", NOTE_COL)
bk.header_row(ws, 23, [("B", "Item"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), (NOTE_COL, "How it is calculated")])
V = "CDE"
put(ws, 24, "5-year real rate (NTN-B)", {"C": 0.085, "D": 0.0676, "E": 0.055}, PCT2, "Same scenarios as the other boxes.", style=LV)
put(ws, 25, "Risk free, nominal = (NTN-B + long-term IPCA) x share kept after tax", {c: f"=({c}24+$C$8)*$C$10" for c in V}, PCT2,
    "Nominal, because the flows are nominal.")
put(ws, 26, "Ke = risk free + beta x ERP", {c: f"={c}25+$F${HV['beta']}*$C$11" for c in V}, PCT2, bold=True)
put(ws, 27, "Ke - g", {c: f"=({c}26-g_)*100" for c in V}, '0.0" pp"')

bk.section(ws, 29, "4. Box projection (index points): consensus 2026-28, box-level projection after 2028", "B", NOTE_COL)
years_header(ws, 30)
P = dict(t=31, gr=32, sales=33, ni=34, da=35, capex=36, lease=37, nwc=38, fcfe=39, po=40, cxda=41)
put(ws, P["t"], "Years after end-2027 (grey = after the explicit years, not used)",
    {c: f"={YRS[i]}-2027" for i, c in enumerate(Y) if YRS[i] >= 2028}, "0")
ws.conditional_format(f"{E28}{P['t']}:{N37}{P['t']}", {"type": "formula", "criteria": f"={E28}${P['t']}>NExp",
                                                         "format": bk.wb.add_format({"font_color": "#A6A6A6"})})
gr = {c: f"={c}{P['ni']}/{p}{P['ni']}-1" for p, c in zip(Y[:2], Y[1:3])}
gr.update({c: f"=IF({c}{P['t']}>=NExp,g_,${E28}${P['gr']}+(g_-${E28}${P['gr']})*({c}{P['t']}-1)/(NExp-1))" for c in Y[3:]})
put(ws, P["gr"], "Growth: 2027-28 net income consensus; then linear from the 2028 growth to g in the last explicit year", gr, PCT,
    "Every line grows at this rate after 2028: prices at their long-term level, volumes flat in real terms.")


def grow_row(key, label, metric, note="", bold=False, sign=1):
    cells = {c: f"={sign}*{box(metric, i)}" if sign == -1 else f"={box(metric, i)}" for i, c in enumerate(Y[:3])}
    cells.update({c: f"={p}{P[key]}*(1+{c}{P['gr']})" for p, c in zip(Y[2:], Y[3:])})
    put(ws, P[key], label, cells, PTS1, note, bold)


grow_row("sales", "Sales", "sales", "Box Consensus; used only for the working capital.")
grow_row("ni", "Net income", "ni", "Box Consensus (sum of the members).", bold=True)
da = {c: f"={box('ebitda', i)}-{box('ebit', i)}" for i, c in enumerate(Y[:3])}
da.update({c: f"={p}{P['da']}*(1+{c}{P['gr']})" for p, c in zip(Y[2:], Y[3:])})
put(ws, P["da"], "(+) D&A = EBITDA - EBIT", da, PTS1, "Includes the depreciation of leased assets (IFRS 16), hence the lease line below.")
grow_row("capex", "(-) Capex", "capex", "Consensus 2026-28 (about 1.1x D&A in 2028: maintenance level); then with inflation.")
grow_row("lease", "(-) Lease principal", "lease", "Petrobras and Suzano (analysts' models); others pending. The lease interest is already in net income.")
put(ws, P["nwc"], "(+/-) Change in working capital = -NWC % x change in sales",
    {c: f"=-$C$13*({c}{P['sales']}-{p}{P['sales']})" for p, c in zip(Y, Y[1:])}, PTS1)
put(ws, P["fcfe"], "Cash to shareholders (FCFE, net debt constant)",
    {c: f"={c}{P['ni']}+{c}{P['da']}+{c}{P['capex']}+{c}{P['lease']}+{c}{P['nwc']}" for c in Y[1:]}, PTS1,
    "As the analysts' bridges (NOPAT + D&A - capex - NWC - leases), at the equity level with constant net debt.", bold=True)
put(ws, P["po"], "FCFE / net income", {c: f"={c}{P['fcfe']}/{c}{P['ni']}" for c in Y[1:]}, PCT)
put(ws, P["cxda"], "Memo: capex / D&A", {c: f"=-{c}{P['capex']}/{c}{P['da']}" for c in Y}, "0.00x")
U, T, NI_ = P["fcfe"], P["t"], P["ni"]
FL = f"${E28}${U}:${N37}${U}"
NL = f"${E28}${NI_}:${N37}${NI_}"
TT = f"${E28}${T}:${N37}${T}"


def pv(ke, flows):
    return f"SUMPRODUCT({flows},--({TT}<=NExp),1/(1+{ke})^{TT})"


def last(flows):
    return f"INDEX({flows},NExp)"


bk.section(ws, 43, "5. Valuation at the end of 2027", "B", NOTE_COL)
bk.header_row(ws, 44, [("B", "Item"), ("C", "Bear"), ("D", "Base"), ("E", "Bull"), (NOTE_COL, "How it is calculated")])
put(ws, 45, "Ke", {c: f"={c}26" for c in V}, PCT2)
put(ws, 46, "Commodity price shock", {"C": -0.2, "D": 0, "E": 0.2}, UPS, "Symmetric +/-20%, as Assumptions row 62 of the main model.", style=LV)
put(ws, 47, "Earnings shock = elasticity x price shock", {c: f"=$C$12*{c}46" for c in V}, UPS,
    "Applied to net income only: capex and leases do not move with prices, so the cash flow moves more than earnings.")
put(ws, 48, "PV of the explicit flows = PV(FCFE) + shock x PV(net income)", {c: f"={pv(c + '45', FL)}+{c}47*{pv(c + '45', NL)}" for c in V},
    PTS1, "t = 1 for 2028 (value at end-2027); only the explicit years.")
put(ws, 49, "Terminal value = (last FCFE + shock x last net income) x (1 + g) / (Ke - g)",
    {c: f"=({last(FL)}+{c}47*{last(NL)})*(1+g_)/({c}45-g_)" for c in V}, PTS1, "Gordon, nominal growth = inflation (zero real).")
put(ws, 50, "PV of the terminal value = TV / (1 + Ke) ^ (explicit years + extra periods)", {c: f"={c}49/(1+{c}45)^(NExp+TVx)" for c in V}, PTS1)
put(ws, 51, "Fair value of Commodities at the end of 2027 (index pts)", {c: f"={c}48+{c}50" for c in V}, PTS1, bold=True)
put(ws, 52, "Upside vs. today", {c: f"={c}51/$C$7-1" for c in V}, UPS, bold=True)
put(ws, 53, "Perpetuity share of the value", {c: f"={c}50/{c}51" for c in V}, PCT)
put(ws, 54, "Memo: upside in the unified DCF by box (main model, v7b)", {"C": -0.10, "D": 0.50, "E": 1.28}, UPS,
    "Ibov_house_2027_by_box_XP.xlsx, DCF by Box (house FCFF, CAPM rate, g 4.7%).")
put(ws, 55, "Memo: upside in the multiples model (Oct-6)", {"C": -0.17, "D": 0.17, "E": 0.55}, UPS, "Ibov_valuation_by_box_model_XP.xlsx.")

bk.section(ws, 57, "6. Base upside (no price shock): Ke (rows) x long-term real growth (columns); explicit flows as projected", "B", NOTE_COL)
rg = [-0.02, -0.01, 0.0, 0.01, 0.02]
kes = [0.12, 0.13, 0.1394, 0.15, 0.16, 0.17]
ws.write("B58", "Ke \\ real growth", bk.HDRL)
for c, x in zip(Y, rg):
    ws.write_number(f"{c}58", x, F(bold=True, font_color="#FFFFFF", bg_color=NAVY, align="center", num_format=UPS))
for i, k in enumerate(kes):
    rr = 59 + i
    ws.write_number(f"B{rr}", k, F(bold=True, num_format=PCT2, align="left"))
    for c in Y[:len(rg)]:
        gg = f"((1+$C$8)*(1+{c}$58)-1)"
        ws.write_formula(f"{c}{rr}", f"=({pv(f'$B{rr}', FL)}+{last(FL)}*(1+{gg})/($B{rr}-{gg})/(1+$B{rr})^(NExp+TVx))/$C$7-1",
                         F(num_format=UPS))
ws.conditional_format(f"C59:{Y[len(rg) - 1]}64", {"type": "3_color_scale", "min_color": "#F8696B", "mid_type": "num", "mid_value": 0,
                                                  "mid_color": "#FFEB84", "max_color": "#63BE7B"})
ws.write("B66", "Real growth enters the perpetuity only. 13.94% = Ke in the base ((6.76% + 4%) x 0.85 + 0.872 x 5.5%).", NOTE)

r0 = 68
bk.section(ws, r0, "7. What is fixed (not switchable) and where it comes from", "B", NOTE_COL)
fixed = [("Risk free nominal: (NTN-B + IPCA) x 0.85", "Required by the nominal flows; the tax haircut as the analysts"),
         ("Equity risk premium 5.5%, the same for every box", "Unified (Caio)"),
         ("Content of 2026-28: Bloomberg consensus of the members, summed", "Strategy (unified with the other boxes)"),
         ("No commodity price model: the consensus carries the analysts' price decks to 2028", "Analysts reach long-term prices around 2028-29"),
         ("Cash flow = net income + D&A - capex - lease principal - change in working capital", "Analysts' FCFF / FCFE bridges, net debt constant"),
         ("Projection for the box only; every line grows with net income; zero real growth in the long term", "Analysts' commodity logic"),
         ("Scenario real rates 8.5% / today / 5.5%", "House (unified with the other boxes)"),
         ("Commodity price shock +/-20% x earnings elasticity 1.26, on net income", "Main model (Box Earnings)"),
         ("Value at end-2027, compared with today's points", "House")]
for i, (a, b) in enumerate(fixed):
    ws.write(f"B{r0 + 1 + i}", a, F())
    ws.write(f"{NOTE_COL}{r0 + 1 + i}", b, NOTE)
ws.freeze_panes(4, 2)
ws.activate()
bk.close()
print("saved", OUT)
