import sys, json
from pathlib import Path
S = Path(r"C:\Users\Caio\Documents\Documentos\Equity Strategy Dashboard\_pedidos\2026-10-05_ibov_valuation_groups\scripts")
sys.path.insert(0, str(S))
import pandas as pd
import bx_model as M
pb = json.loads((S.parent / "studies/performance_vs_rates/perf_study_results.json").read_text(encoding="utf-8"))["betas_preferred_table"]
P = {"perf_beta": {g: pb[f"abs|pre2y|{g}"]["beta"] for g in M.BOXES + M.MEMO + ["Ibovespa"]}}
R = M.run(P)
bx, d = R["bx"], R["d"]
fv, tgt, cal = R["res"]["A"]
out = {"bx": {g: {k: (float(v) if v == v else None) for k, v in r.items()} for g, r in bx.items()},
       "fv": {mt: {f"{g}|{s}": float(v) for (g, s), v in fv[mt].items()} for mt in fv},
       "tgt": {mt: {f"{g}|{s}": (float(v) if v == v else None) for (g, s), v in tgt[mt].items()} for mt in tgt},
       "sens": {g: float(v) for g, v in cal["eps_sens"].items()}}
H = R["H"]; HA = H[H.index >= "2016-10-01"]
out["hist_stats"] = {f"{k}|{g}": [float(HA[(k, g)].mean()), float(HA[(k, g)].std())] for (k, g) in H.columns if k in ("pe", "ev")}
m = R["m"]
out["members_sample"] = m[["key2", "wn", "pts", "pe_sel", "ev_sel", "evmc", "up", "ke_used", "c_bear", "c_bull"]].head(5).round(6).to_dict()
Path(sys.argv[1]).write_text(json.dumps(out, indent=1), encoding="utf-8")
for mt in ["P/E", "EV/EBITDA", "Bottom-up", "Average"]:
    print(mt, {s: round(fv[mt][("Ibovespa", s)]) for s in M.SCEN})
print("sens", {g: round(v, 3) for g, v in cal["eps_sens"].items()})
print(pd.DataFrame(bx).T[["w", "pts", "pe", "ev", "evmc", "bu_up", "bu_bear", "bu_bull"]].round(4).to_string())
print(m[["key2","ke_used","c_bear","c_bull"]].groupby("key2").agg(["count","mean"]).round(3))
