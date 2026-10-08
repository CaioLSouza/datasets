"""Bloomberg data for the Financials DDM (XP analysts' structure with consensus numbers). One API session.

Snapshot ('bst' consensus, FY26 / FY27 / FY28): EPS, DPS, BVPS, ROE and net income of the Financials members of the box
model (Members sheet of Ibov_house_2027_by_box_XP.xlsx), plus price and P/E for the checks.
History: daily total-return index (TOT_RETURN_INDEX_GROSS_DVDS) Sep-16..Sep-26, resampled to Friday closes in the
builder, for the betas computed by hand against the Ibovespa of the Box Betas sheet.
Output: fin_ddm_data.json; raw responses cached in fin_raw_*.json.
"""
from __future__ import annotations

import json
import time
from pathlib import Path

import blpapi
import openpyxl

from dcf_data import ref, session, _wait

HERE = Path(__file__).parent
BOX = HERE.parent / "Ibov_house_2027_by_box_XP.xlsx"
PERIODS = {"1FY": "26", "2FY": "27", "3FY": "28"}
SNAP = ["BEST_EPS", "BEST_DPS", "BEST_BPS", "BEST_ROE", "BEST_NET_INCOME", "BEST_PE_RATIO"]


def members():
    ws = openpyxl.load_workbook(BOX, read_only=True, data_only=True)["Members"]
    out = []
    for r in ws.iter_rows(min_row=6, values_only=True):
        if r[0] and r[2] == "Financials":
            out.append({"ticker": r[0], "name": r[1], "px_sep30": r[5], "px_oct5": r[6], "pts": r[9], "mcap": r[10]})
    return out


def hist_daily(s, svc, tickers, field, start, end, cache):
    if cache.exists():
        return json.loads(cache.read_text(encoding="utf-8"))
    req = svc.createRequest("HistoricalDataRequest")
    for t in tickers:
        req.getElement("securities").appendValue(t)
    req.getElement("fields").appendValue(field)
    req.set("startDate", start)
    req.set("endDate", end)
    req.set("periodicitySelection", "DAILY")
    out = {}

    def h(msg):
        if msg.hasElement("securityData"):
            sd = msg.getElement("securityData")
            tk = sd.getElementAsString("security")
            fd = sd.getElement("fieldData")
            for i in range(fd.numValues()):
                row = fd.getValueAsElement(i)
                if row.hasElement(field):
                    out.setdefault(tk, {})[str(row.getElementAsDatetime("date"))[:10]] = row.getElementAsFloat(field)

    s.sendRequest(req)
    assert _wait(s, h, 900), "historical request timed out"
    cache.write_text(json.dumps(out), encoding="utf-8")
    return out


def main():
    m = members()
    tk = [f"{x['ticker']} BZ Equity" for x in m]
    s, svc = session()
    t0 = time.time()
    px = ref(s, svc, tk, ["PX_LAST"], cache=HERE / "fin_raw_px.json")
    snap = {y: ref(s, svc, tk, SNAP, {"BEST_DATA_SOURCE_OVERRIDE": "bst", "BEST_FPERIOD_OVERRIDE": p},
                   cache=HERE / f"fin_raw_snap_{y}.json") for p, y in PERIODS.items()}
    print("snapshot done", round(time.time() - t0), "s", flush=True)
    tr = hist_daily(s, svc, tk, "TOT_RETURN_INDEX_GROSS_DVDS", "20160901", "20260930", HERE / "fin_raw_tr.json")
    s.stop()
    print("history done", round(time.time() - t0), "s", flush=True)
    for x, t in zip(m, tk):
        x["px_last"] = px.get(t, {}).get("PX_LAST")
        for y in PERIODS.values():
            for f in SNAP:
                x[f"{f}_{y}"] = snap[y].get(t, {}).get(f)
        x["tr_days"] = len(tr.get(t, {}))
    (HERE / "fin_ddm_data.json").write_text(json.dumps({"pulled": time.strftime("%Y-%m-%d %H:%M"), "members": m},
                                                       indent=1, default=str), encoding="utf-8")
    for x in m:
        print({k: (round(v, 3) if isinstance(v, float) else v) for k, v in x.items()})


if __name__ == "__main__":
    main()
