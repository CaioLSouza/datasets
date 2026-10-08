"""Bloomberg data for the DCF by box (FCFE, value-driver form). One API session, batched requests.

Snapshot (today, BDP, 'bst' consensus): price, forward P/E and ROE for FY26 / FY27 / FY28 and 12m blended (BF) for the 76
Ibovespa members of the valuation model (Sep-30 composition). Forward EPS levels are taken as price / P/E, so the currency
of reporting (Embraer in USD) does not matter.
History (Oct-16..Sep-26, monthly): 12m fwd ROE (BEST_ROE, BF) of every member of each month, with the same ticker map as the
P/E panel -> long-term ROE per box.
Output: dcf_members.parquet, dcf_roe_hist.parquet, dcf_data.json; raw responses cached in dcf_raw_*.json.
"""
from __future__ import annotations

import json
import time
from pathlib import Path

import blpapi
import pandas as pd

HERE = Path(__file__).parent
SRC = Path(r"C:\Users\Caio\Documents\Documentos\Equity Strategy Dashboard\_pedidos")
H0, H1 = "2016-10-01", "2026-09-30"
PERIODS = {"1FY": "26", "2FY": "27", "3FY": "28", "BF": "bf"}


def session():
    o = blpapi.SessionOptions()
    o.setServerHost("localhost")
    o.setServerPort(8194)
    s = blpapi.Session(o)
    if not s.start() or not s.openService("//blp/refdata"):
        raise RuntimeError("Bloomberg unavailable")
    return s, s.getService("//blp/refdata")


def _overrides(req, ov):
    el = req.getElement("overrides")
    for k, v in (ov or {}).items():
        o = el.appendElement()
        o.setElement("fieldId", k)
        o.setElement("value", v)


def _wait(s, handle, limit):
    t0 = time.time()
    while time.time() - t0 < limit:
        ev = s.nextEvent(5000)
        for msg in ev:
            handle(msg)
        if ev.eventType() == blpapi.Event.RESPONSE:
            return True
    return False


def ref(s, svc, tickers, fields, ov=None, cache=None):
    if cache and cache.exists():
        return json.loads(cache.read_text(encoding="utf-8"))
    req = svc.createRequest("ReferenceDataRequest")
    for t in tickers:
        req.getElement("securities").appendValue(t)
    for f in fields:
        req.getElement("fields").appendValue(f)
    _overrides(req, ov)
    out = {}

    def h(msg):
        if msg.hasElement("securityData"):
            arr = msg.getElement("securityData")
            for i in range(arr.numValues()):
                sd = arr.getValueAsElement(i)
                fd = sd.getElement("fieldData")
                out[sd.getElementAsString("security")] = {f: fd.getElementAsFloat(f) for f in fields if fd.hasElement(f)}

    s.sendRequest(req)
    assert _wait(s, h, 180), "reference request timed out"
    if cache:
        cache.write_text(json.dumps(out), encoding="utf-8")
    return out


def hist(s, svc, tickers, field, start, end, ov=None, cache=None):
    if cache and cache.exists():
        return json.loads(cache.read_text(encoding="utf-8"))
    req = svc.createRequest("HistoricalDataRequest")
    for t in tickers:
        req.getElement("securities").appendValue(t)
    req.getElement("fields").appendValue(field)
    req.set("startDate", start)
    req.set("endDate", end)
    req.set("periodicitySelection", "MONTHLY")
    req.set("periodicityAdjustment", "CALENDAR")
    req.set("nonTradingDayFillOption", "ALL_CALENDAR_DAYS")
    req.set("nonTradingDayFillMethod", "PREVIOUS_VALUE")
    _overrides(req, ov)
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
    if cache:
        cache.write_text(json.dumps(out), encoding="utf-8")
    return out


def main():
    m = pd.read_parquet(SRC / "2026-10-05_ibov_valuation_groups/scripts/bx_members.parquet")
    tk = [f"{t} BZ Equity" for t in m.index]
    s, svc = session()
    t0 = time.time()
    px = ref(s, svc, tk + ["IBOV Index"], ["PX_LAST"], cache=HERE / "dcf_raw_px.json")
    snap = {}
    for p, y in PERIODS.items():
        snap[y] = ref(s, svc, tk + ["IBOV Index"], ["BEST_PE_RATIO", "BEST_ROE", "BEST_EPS"],
                      {"BEST_DATA_SOURCE_OVERRIDE": "bst", "BEST_FPERIOD_OVERRIDE": p}, cache=HERE / f"dcf_raw_snap_{y}.json")
    print("snapshot done", round(time.time() - t0), "s", flush=True)
    out = pd.DataFrame(index=m.index)
    g = lambda d, t, f: d.get(f"{t} BZ Equity", {}).get(f)
    out["px_today"] = [g(px, t, "PX_LAST") for t in m.index]
    for y in PERIODS.values():
        out[f"pe{y}"] = [g(snap[y], t, "BEST_PE_RATIO") for t in m.index]
        out[f"roe{y}"] = [g(snap[y], t, "BEST_ROE") for t in m.index]
    idx = {"PX_LAST": px.get("IBOV Index", {}).get("PX_LAST")}
    for y in PERIODS.values():
        for f in ("BEST_PE_RATIO", "BEST_ROE", "BEST_EPS"):
            idx[f"{f}_{y}"] = snap[y].get("IBOV Index", {}).get(f)
    # ---------------------------------------------------------------- 10y history of the 12m fwd ROE (one request)
    p = pd.read_parquet(SRC / "2026-10-05_ibov_valuation_groups/scripts/bx_panel20.parquet")
    p = p[(p["date"] >= H0) & (p["date"] <= H1)]
    used = json.loads((SRC / "2026-10-05_brazil_ex_commodities/scripts/fourgroups2.json").read_text(encoding="utf-8"))["panel_used"]
    names = {t: used.get(t, f"{t} BZ Equity") for t in p["cod_ativo"].unique()}
    bbg = sorted({v for v in names.values() if not v.startswith("(")})
    h = hist(s, svc, bbg, "BEST_ROE", "20161001", "20260930", {"BEST_DATA_SOURCE_OVERRIDE": "bst", "BEST_FPERIOD_OVERRIDE": "BF"},
             cache=HERE / "dcf_raw_roe_hist.json")
    s.stop()
    print("history done", round(time.time() - t0), "s", flush=True)
    rows = [{"cod_ativo": t, "p": pd.Timestamp(d).to_period("M"), "roe": v} for t, bt in names.items() for d, v in h.get(bt, {}).items()]
    roe = pd.DataFrame(rows)
    p = p.assign(p=p["date"].dt.to_period("M")).merge(roe, on=["cod_ativo", "p"], how="left")
    hist_df = p[["date", "cod_ativo", "gics", "box", "glob", "w", "pe", "roe"]]
    out.to_parquet(HERE / "dcf_members.parquet")
    hist_df.to_parquet(HERE / "dcf_roe_hist.parquet")
    info = {"pulled": time.strftime("%Y-%m-%d %H:%M"), "index": idx,
            "no_roe_history": sorted(t for t, v in names.items() if v.startswith("(") or v not in h),
            "n_hist": len(hist_df), "n_roe": int(hist_df["roe"].notna().sum()),
            "missing_snapshot": {c: list(out.index[out[c].isna()]) for c in out.columns}}
    (HERE / "dcf_data.json").write_text(json.dumps(info, indent=1, default=str), encoding="utf-8")
    print(json.dumps(info, indent=1, default=str))
    print(out.round(2).to_string())


if __name__ == "__main__":
    main()
