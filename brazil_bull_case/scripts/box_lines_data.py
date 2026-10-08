"""Bloomberg data for the house DCF by box: consensus sales, EBIT, EBITDA, net income and capex (company totals, BRL) for
FY26 / FY27 / FY28 (BEST_FPERIOD_OVERRIDE 1FY/2FY/3FY) and the company market cap, for the 76 Ibovespa members of the
valuation model. In the workbook each line becomes index points (points x metric / market cap) and then box shares, which
split the house's IBOV Index lines among the boxes. One API session, batched requests; raw responses cached.

Output: box_lines.parquet (one row per member), box_lines.json (pull info, coverage).
"""
from __future__ import annotations

import json
import time
from pathlib import Path

import pandas as pd

from dcf_data import session, ref

HERE = Path(__file__).parent
SRC = Path(r"C:\Users\Caio\Documents\Documentos\Equity Strategy Dashboard\_pedidos\2026-10-05_ibov_valuation_groups\scripts")
FIELDS = {"BEST_SALES": "sales", "BEST_EBIT": "ebit", "BEST_EBITDA": "ebitda", "BEST_NET_INCOME": "ni", "BEST_CAPEX": "capex"}
PERIODS = {"1FY": "26", "2FY": "27", "3FY": "28"}


def main():
    m = pd.read_parquet(SRC / "bx_members.parquet")
    tk = [f"{t} BZ Equity" for t in m.index]
    s, svc = session()
    t0 = time.time()
    mc = ref(s, svc, tk + ["IBOV Index"], ["CUR_MKT_CAP"], cache=HERE / "box_raw_mcap.json")
    snap = {}
    for p, y in PERIODS.items():
        snap[y] = ref(s, svc, tk + ["IBOV Index"], list(FIELDS), {"BEST_FPERIOD_OVERRIDE": p, "EQY_FUND_CRNCY": "BRL"},
                      cache=HERE / f"box_raw_{y}.json")
    s.stop()
    print("pulled in", round(time.time() - t0), "s", flush=True)
    out = pd.DataFrame(index=m.index)
    out["mcap"] = [mc.get(f"{t} BZ Equity", {}).get("CUR_MKT_CAP") for t in m.index]
    for y, d in snap.items():
        for f, k in FIELDS.items():
            out[f"{k}{y}"] = [d.get(f"{t} BZ Equity", {}).get(f) for t in m.index]
    out.to_parquet(HERE / "box_lines.parquet")
    idx = {y: d.get("IBOV Index", {}) for y, d in snap.items()}
    info = {"pulled": time.strftime("%Y-%m-%d %H:%M"), "ibov_index_lines": idx,
            "missing": {c: list(out.index[out[c].isna()]) for c in out.columns}}
    (HERE / "box_lines.json").write_text(json.dumps(info, indent=1, default=str), encoding="utf-8")
    print(json.dumps(info, indent=1, default=str)[:4000])
    print(out.join(m[["box"]]).round(0).to_string())


if __name__ == "__main__":
    main()
