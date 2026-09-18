"""
Bottom-up Valuation — raw data da cobertura XP
==============================================

Lê os target prices da cobertura (COMP SHEET/raw_data.xlsx) e a composição mais
recente do Ibovespa (Economatica) e grava `bottom_up_raw.xlsx` na pasta output.
Esse arquivo é lido via Power Query pela planilha `Implied Ibovespa.xlsx`, que
completa com o consenso Bloomberg e calcula o implied target price do índice.

Uso:  python bottom_up_raw.py
      (o arquivo bottom_up_raw.xlsx precisa estar fechado)

Nenhuma regra de exclusão (Under Review, idade do modelo) é aplicada aqui: o raw
leva tudo, e a decisão fica na planilha final, onde dá para mexer sem rodar código.
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------- #
# CAMINHOS
# --------------------------------------------------------------------------- #
RAW_DATA = Path(r"\\xpdocs\Research\Equities\COMP SHEET\raw_data.xlsx")
INDEX_COMPOSITION = Path(r"\\xpdocs\Research\Equities\Quant\_Cross Data\economatica-index_composition.parquet")
BDR_MARKET_DATA = Path(r"\\xpdocs\Research\Equities\Quant\_Cross Data\economatica-bdr_market_data.parquet")
OUTPUT_DIR = Path(r"\\xpdocs\Research\Equities\Estrategia\Banco de dados\Bottom up Valuation\output")
OUTPUT_FILE = "bottom_up_raw.xlsx"   # nome fixo: é o que o Power Query procura

INDEX = "IBOV"                        # coluna da composição a usar
BBG_SUFFIX = " BZ Equity"

# --------------------------------------------------------------------------- #
# LEITURA
# --------------------------------------------------------------------------- #
def clean_ticker(s: pd.Series) -> pd.Series:
    """'PETR4<XBSP>' -> 'PETR4'."""
    return s.astype(str).str.replace(r"<.*?>", "", regex=True).str.strip().str.upper()


def read_coverage(path: Path) -> pd.DataFrame:
    keep = ["TICKER", "NAME", "SECTOR_XP", "LEAD_ANALYST", "PDATE", "RECOMMENDATION",
            "RESTRICTED", "MODEL_CURRENCY", "PRICE_CURRENCY", "TARGET", "KE", "WACC"]
    df = pd.read_excel(path, usecols=lambda c: c in keep)
    df["TICKER"] = df["TICKER"].astype(str).str.strip().str.upper()
    df["PDATE"] = pd.to_datetime(df["PDATE"], errors="coerce")
    df["TARGET"] = pd.to_numeric(df["TARGET"], errors="coerce")
    df["RESTRICTED"] = df["RESTRICTED"].fillna(0).astype(bool)
    df["RECOMMENDATION"] = df["RECOMMENDATION"].fillna("n.a.").astype(str).str.strip()
    df = df.dropna(subset=["TICKER"]).drop_duplicates("TICKER", keep="last")
    return df


def read_composition(path: Path, index: str = INDEX) -> tuple[pd.DataFrame, pd.Timestamp]:
    """Última composição disponível do índice. Aceita o parquet da rede ou o CSV da Economatica."""
    cols = ["Ativo", "Data", index]
    if path.suffix.lower() == ".parquet":
        df = pd.read_parquet(path, columns=cols)
    else:
        parts = [c[c[index] != "-"] for c in pd.read_csv(path, usecols=cols, dtype=str, chunksize=2_000_000)]
        df = pd.concat(parts)
    df[index] = pd.to_numeric(df[index], errors="coerce")
    df = df[df[index] > 0]
    df["Data"] = pd.to_datetime(df["Data"])
    last = df["Data"].max()
    comp = df[df["Data"] == last].copy()
    comp["Ticker"] = clean_ticker(comp["Ativo"])
    comp["Weight"] = comp[index] / comp[index].sum()          # fração, soma 1
    comp = comp.sort_values("Weight", ascending=False)[["Ticker", "Weight"]].reset_index(drop=True)
    return comp, last


def read_bdr_tickers(path: Path) -> set[str]:
    """Tickers de BDR negociados na B3 — só para classificar a cobertura. Opcional."""
    try:
        df = pd.read_parquet(path) if path.suffix.lower() == ".parquet" else pd.read_csv(path, nrows=0)
        col = next(c for c in df.columns if c.lower() in ("ativo", "ticker", "cod_ativo"))
        return set(clean_ticker(pd.Series(df[col].unique())))
    except Exception as e:  # noqa: BLE001 — arquivo auxiliar, não pode travar o processo
        print(f"  aviso: não li os BDRs ({e.__class__.__name__}); sigo sem essa marcação")
        return set()


# --------------------------------------------------------------------------- #
# TRATAMENTO
# --------------------------------------------------------------------------- #
def listing(row, bdr: set[str]) -> str:
    t = row["TICKER"]
    if row["PRICE_CURRENCY"] != "BRL":
        return "US (ADR/NYSE/Nasdaq)"
    if t in bdr or t[-2:] in ("32", "33", "34", "35", "39"):
        return "B3 BDR"
    return "B3"


def build_ibov(comp: pd.DataFrame, cov: pd.DataFrame) -> pd.DataFrame:
    """Casa cada papel do índice com um TP da cobertura: direto ou pela outra classe da mesma empresa."""
    local = cov[cov["PRICE_CURRENCY"] == "BRL"].set_index("TICKER")
    rec_rank = {"Buy": 0, "Neutral": 0, "Sell": 0}              # prefere classe com recomendação ativa

    def match(t: str):
        if t in local.index:
            return t, "Direct"
        sib = [s for s in local.index if s[:4] == t[:4] and s != t]
        if not sib:
            return None, "None"
        sib.sort(key=lambda s: (rec_rank.get(local.at[s, "RECOMMENDATION"], 1), s))
        return sib[0], "Other class"

    m = comp["Ticker"].map(match)
    out = comp.copy()
    out["XP Ticker"] = [x[0] for x in m]
    out["Coverage"] = [x[1] for x in m]
    info = local.reindex(out["XP Ticker"])
    out["Name"] = info["NAME"].values
    out["Sector"] = info["SECTOR_XP"].values
    out["Analyst"] = info["LEAD_ANALYST"].values
    out["Recommendation"] = info["RECOMMENDATION"].values
    out["Restricted"] = info["RESTRICTED"].values
    out["Model Date"] = info["PDATE"].values
    out["XP Target"] = info["TARGET"].values
    out["BBG Ticker"] = out["Ticker"] + BBG_SUFFIX
    out["XP BBG Ticker"] = out["XP Ticker"].where(out["XP Ticker"].isna(), out["XP Ticker"] + BBG_SUFFIX)
    # colunas auxiliares (tickers Bloomberg) no fim: na planilha final elas ficam ocultas
    return out[["Ticker", "Weight", "Coverage", "XP Ticker", "Name", "Sector", "Analyst",
                "Recommendation", "Restricted", "Model Date", "XP Target", "BBG Ticker", "XP BBG Ticker"]]


# --------------------------------------------------------------------------- #
# SAÍDA
# --------------------------------------------------------------------------- #
def write_output(ibov: pd.DataFrame, cov: pd.DataFrame, meta: pd.DataFrame, path: Path) -> None:
    with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as xw:
        wb = xw.book
        hdr = wb.add_format({"font_name": "Roboto Light", "font_size": 9, "bold": True,
                             "font_color": "#FFFFFF", "bg_color": "#1F2F44"})
        body = wb.add_format({"font_name": "Roboto Light", "font_size": 9})
        fmts = {"pct": wb.add_format({"font_name": "Roboto Light", "font_size": 9, "num_format": "0.00%"}),
                "num": wb.add_format({"font_name": "Roboto Light", "font_size": 9, "num_format": "#,##0.00"}),
                "date": wb.add_format({"font_name": "Roboto Light", "font_size": 9, "num_format": "dd/mm/yyyy"}),
                "datetime": wb.add_format({"font_name": "Roboto Light", "font_size": 9, "num_format": "dd/mm/yyyy hh:mm"})}
        colfmt = {"Weight": "pct", "Ibov Weight": "pct", "XP Target": "num", "TARGET": "num",
                  "KE": "pct", "WACC": "pct", "Model Date": "date", "PDATE": "date",
                  "Model Age (months)": "num", "Composition Date": "date", "Run At": "datetime"}
        for name, df in (("ibov", ibov), ("coverage", cov), ("meta", meta)):
            df = df.astype(object).where(df.notna(), None)
            df.to_excel(xw, sheet_name=name, index=False, startrow=1, header=False)
            ws = xw.sheets[name]
            ws.add_table(0, 0, len(df), len(df.columns) - 1, {
                "name": name, "style": None,
                "columns": [{"header": c, "header_format": hdr} for c in df.columns]})
            for j, c in enumerate(df.columns):
                ws.set_column(j, j, max(10, min(34, len(c) + 4)), fmts[colfmt[c]] if c in colfmt else body)
            ws.freeze_panes(1, 0)
            ws.hide_gridlines(2)


def main() -> int:
    run_at = datetime.now()
    out_path = OUTPUT_DIR / OUTPUT_FILE
    print(f"Bottom-up raw  |  {run_at:%d/%m/%Y %H:%M}")

    print("  lendo cobertura ...")
    cov = read_coverage(RAW_DATA)
    print(f"    {len(cov)} papéis")

    print(f"  lendo composição {INDEX} ...")
    comp, comp_date = read_composition(INDEX_COMPOSITION)
    print(f"    {len(comp)} papéis em {comp_date:%d/%m/%Y}, peso somado {comp['Weight'].sum():.4f}")

    bdr = read_bdr_tickers(BDR_MARKET_DATA)

    ibov = build_ibov(comp, cov)
    ibov["Composition Date"] = comp_date
    ibov["Run At"] = run_at

    cov = cov.copy()
    cov["Listing"] = cov.apply(listing, axis=1, bdr=bdr)
    w = comp.set_index("Ticker")["Weight"]
    cov["In Ibov"] = cov["TICKER"].isin(w.index)
    cov["Ibov Weight"] = cov["TICKER"].map(w)
    cov["Model Age (months)"] = ((run_at - cov["PDATE"]).dt.days / 30.4375).round(1)

    meta = pd.DataFrame({"Run At": [run_at], "Composition Date": [comp_date], "Index": [INDEX],
                         "Coverage Source": [str(RAW_DATA)], "Composition Source": [str(INDEX_COMPOSITION)]})

    n = ibov["Coverage"].value_counts()
    wt = ibov.groupby("Coverage")["Weight"].sum()
    print("  casamento com o índice:")
    for k in ("Direct", "Other class", "None"):
        print(f"    {k:<12} {n.get(k, 0):>3} papéis  {wt.get(k, 0):>7.2%}")
    if n.get("None", 0):
        print("    sem TP XP (vão para o consenso): " + ", ".join(ibov.loc[ibov["Coverage"] == "None", "Ticker"]))

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    try:
        write_output(ibov, cov, meta, out_path)
    except PermissionError:
        print(f"\nERRO: {out_path} está aberto. Feche o arquivo e rode de novo.")
        return 1
    print(f"  gravado: {out_path}")
    print("\nAgora abra 'Implied Ibovespa.xlsx' e clique em Dados > Atualizar Tudo.")
    return 0


if __name__ == "__main__":
    # para testar fora da rede:  python bottom_up_raw.py <raw_data.xlsx> <composição .parquet/.csv> <pasta output>
    if len(sys.argv) == 4:
        RAW_DATA, INDEX_COMPOSITION, OUTPUT_DIR = Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3])
    sys.exit(main())
