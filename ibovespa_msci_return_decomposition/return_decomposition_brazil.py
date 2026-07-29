"""
Return decomposition for the Ibovespa and MSCI Brazil sectors.

Components
----------
1. Dividends
2. Earnings growth (implied residual)
3. Valuation / multiple expansion

The daily methodology follows the supplied global script:
    total return = dividends + earnings growth + valuation

Valuation is inferred from the 21-day moving average of P/E versus its level
21 trading days earlier and converted into a daily rate. Earnings growth is
the residual. Multi-period contributions are linked with the Carino method so
that the three components reconcile exactly to the compounded total return.

Data modes
----------
* bloomberg: pulls Ibovespa and MSCI Brazil sector data with xbbg.
* bloomberg-xlsx: reads all sheets from the Bloomberg-formula workbook after
  the user refreshes and saves it in Bloomberg Excel.
* supplied-xlsx: reproduces the analysis for one sheet in the workbook
  supplied by the user (useful for validation).
"""

from __future__ import annotations

import argparse
import json
import math
import re
import warnings
from dataclasses import dataclass, asdict
from datetime import date
from pathlib import Path
from typing import Dict, Iterable, Optional

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


COMPONENTS = ["dividends", "earnings_growth", "valuation"]
COLORS = {
    "total_return": "#202124",
    "dividends": "#2E86AB",
    "earnings_growth": "#3A9D5D",
    "valuation": "#E67E22",
}
LABELS = {
    "total_return": "Total return",
    "dividends": "Dividends",
    "earnings_growth": "Earnings growth",
    "valuation": "Valuation",
}


@dataclass(frozen=True)
class SeriesSpec:
    name: str
    total_ticker: str
    price_ticker: Optional[str] = None
    dividend_method: str = "total_minus_price"
    family: str = "MSCI Brazil"


SERIES_SPECS = [
    SeriesSpec(
        name="Ibovespa",
        total_ticker="IBOV Index",
        dividend_method="index_dividend_points",
        family="B3",
    ),
    SeriesSpec("MSCI Brazil", "M2BR Index", "MXBR Index", family="MSCI Brazil"),
    SeriesSpec("Energy", "M2BR0EN Index", "MXBR0EN Index"),
    SeriesSpec("Materials", "M2BR0MT Index", "MXBR0MT Index"),
    SeriesSpec("Industrials", "M2BR0IN Index", "MXBR0IN Index"),
    SeriesSpec(
        "Consumer Discretionary", "M2BR0CD Index", "MXBR0CD Index"
    ),
    SeriesSpec("Consumer Staples", "M2BR0CS Index", "MXBR0CS Index"),
    SeriesSpec("Health Care", "M2BR0HC Index", "MXBR0HC Index"),
    SeriesSpec("Financials", "M2BR0FN Index", "MXBR0FN Index"),
    SeriesSpec(
        "Information Technology", "M2BR0IT Index", "MXBR0IT Index"
    ),
    SeriesSpec(
        "Communication Services", "M2BR0TC Index", "MXBR0TC Index"
    ),
    SeriesSpec("Utilities", "M2BR0UT Index", "MXBR0UT Index"),
]

WORKBOOK_SHEET_MAP = {
    "IBOV": "Ibovespa",
    "MSCI Brazil": "MSCI Brazil",
    "Energy": "Energy",
    "Materials": "Materials",
    "Industrials": "Industrials",
    "Cons Discretionary": "Consumer Discretionary",
    "Cons Staples": "Consumer Staples",
    "Health Care": "Health Care",
    "Financials": "Financials",
    "Info Tech": "Information Technology",
    "Comm Services": "Communication Services",
    "Utilities": "Utilities",
}


def slugify(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9]+", "_", value.strip())
    return value.strip("_").lower()


def _pick_bdh_column(df: pd.DataFrame, ticker: str, field: str) -> pd.Series:
    """Return one xbbg BDH column despite minor column-layout differences."""
    if df.empty:
        return pd.Series(dtype=float)
    if isinstance(df.columns, pd.MultiIndex):
        candidates = [
            (ticker, field),
            (ticker, field.lower()),
            (ticker, field.upper()),
        ]
        for candidate in candidates:
            if candidate in df.columns:
                return pd.to_numeric(df[candidate], errors="coerce")
        for col in df.columns:
            col_text = "|".join(map(str, col)).upper()
            if ticker.upper() in col_text and field.upper() in col_text:
                return pd.to_numeric(df[col], errors="coerce")
    else:
        for col in df.columns:
            if str(col).upper() == field.upper():
                return pd.to_numeric(df[col], errors="coerce")
    return pd.Series(index=df.index, dtype=float)


def _bdh_one(blp, ticker: str, field: str, start: str, end: str) -> pd.Series:
    try:
        frame = blp.bdh(
            tickers=ticker,
            flds=field,
            start_date=start,
            end_date=end,
            Per="D",
            Days="A",
            Fill="P",
        )
        result = _pick_bdh_column(frame, ticker, field)
        result.index = pd.to_datetime(result.index)
        return result.sort_index()
    except Exception as exc:  # Bloomberg entitlement/field errors vary.
        warnings.warn(f"Bloomberg failed for {ticker} / {field}: {exc}")
        return pd.Series(dtype=float)


def pull_bloomberg_series(
    spec: SeriesSpec, start: str, end: str
) -> pd.DataFrame:
    try:
        from xbbg import blp
    except ImportError as exc:
        raise RuntimeError(
            "Bloomberg mode requires xbbg and a running Bloomberg Desktop API. "
            "Install with: pip install xbbg blpapi"
        ) from exc

    total_level = _bdh_one(blp, spec.total_ticker, "PX_LAST", start, end)
    pe = _bdh_one(blp, spec.price_ticker or spec.total_ticker, "PE_RATIO", start, end)

    if spec.dividend_method == "total_minus_price":
        if not spec.price_ticker:
            raise ValueError(f"{spec.name}: price_ticker is required")
        price_level = _bdh_one(blp, spec.price_ticker, "PX_LAST", start, end)
        raw = pd.concat(
            {
                "total_level": total_level,
                "price_level": price_level,
                "pe": pe,
            },
            axis=1,
        )
        raw["dividend_method"] = "total_return_minus_price_return"
    elif spec.dividend_method == "index_dividend_points":
        dividend_points = _bdh_one(
            blp, spec.total_ticker, "INDX_GROSS_DAILY_DIV", start, end
        )
        raw = pd.concat(
            {
                "total_level": total_level,
                "pe": pe,
                "dividend_points": dividend_points,
            },
            axis=1,
        )
        if raw["dividend_points"].notna().any():
            raw["dividend_method"] = "INDX_GROSS_DAILY_DIV"
        else:
            # Fallback is deliberately transparent and quality-flagged.
            div_yield = _bdh_one(
                blp, spec.total_ticker, "DIV_YIELD", start, end
            )
            raw["dividend_yield"] = div_yield
            raw["dividend_method"] = "DIV_YIELD/252_APPROXIMATION"
    else:
        raise ValueError(f"Unknown dividend method: {spec.dividend_method}")

    raw.index.name = "date"
    raw = raw[~raw.index.duplicated(keep="last")].sort_index()
    return raw


def read_supplied_workbook(
    workbook: Path, sheet: str, series_name: str
) -> Dict[str, pd.DataFrame]:
    data = pd.read_excel(workbook, header=5, sheet_name=sheet)
    data["Dates"] = pd.to_datetime(data["Dates"], errors="coerce")
    data = data.dropna(subset=["Dates"]).set_index("Dates").sort_index()
    raw = pd.DataFrame(index=data.index)
    raw["total_level"] = pd.to_numeric(data["LAST_PRICE"], errors="coerce")
    raw["price_level"] = pd.to_numeric(data["LAST_PRICE.1"], errors="coerce")
    raw["pe"] = pd.to_numeric(data["PE_RATIO"], errors="coerce")
    raw["dividend_method"] = "total_return_minus_price_return"
    raw.index.name = "date"
    return {series_name: raw}


def read_bloomberg_formula_workbook(
    workbook: Path,
) -> Dict[str, pd.DataFrame]:
    """
    Read cached values from the workbook after Bloomberg Excel refresh + save.

    MSCI sheets use two Bloomberg blocks:
      A:E = gross total return index
      F:J = price return index

    IBOV uses one block:
      A:E = date, total-return level, P/E, daily dividend points, dividend yield
    """
    raw_by_series: Dict[str, pd.DataFrame] = {}
    available_sheets = set(pd.ExcelFile(workbook).sheet_names)
    expected_sheets = set(WORKBOOK_SHEET_MAP)
    missing = expected_sheets - available_sheets
    if missing:
        raise ValueError(
            "The Bloomberg input workbook is missing sheets: "
            + ", ".join(sorted(missing))
        )

    for sheet, series_name in WORKBOOK_SHEET_MAP.items():
        data = pd.read_excel(workbook, header=5, sheet_name=sheet)
        if "Dates" not in data.columns:
            raise ValueError(
                f"{sheet}: no cached Bloomberg dates found. "
                "Open the workbook in Bloomberg Excel, refresh it, wait for "
                "all BDH formulas to finish and save the workbook."
            )
        dates = pd.to_datetime(data["Dates"], errors="coerce")
        valid = dates.notna()
        data = data.loc[valid].copy()
        data.index = dates.loc[valid]
        data.index.name = "date"
        raw = pd.DataFrame(index=data.index)

        if sheet == "IBOV":
            required = [
                "LAST_PRICE",
                "PE_RATIO",
                "INDX_GROSS_DAILY_DIV",
                "DIV_YIELD",
            ]
            absent = [field for field in required if field not in data.columns]
            if absent:
                raise ValueError(f"IBOV: missing Bloomberg fields {absent}")
            raw["total_level"] = pd.to_numeric(
                data["LAST_PRICE"], errors="coerce"
            )
            raw["pe"] = pd.to_numeric(data["PE_RATIO"], errors="coerce")
            raw["dividend_points"] = pd.to_numeric(
                data["INDX_GROSS_DAILY_DIV"], errors="coerce"
            )
            raw["dividend_yield"] = pd.to_numeric(
                data["DIV_YIELD"], errors="coerce"
            )
            if raw["dividend_points"].notna().any():
                raw["dividend_method"] = "INDX_GROSS_DAILY_DIV"
            elif raw["dividend_yield"].notna().any():
                raw["dividend_method"] = "DIV_YIELD/252_APPROXIMATION"
            else:
                raise ValueError(
                    "IBOV: neither INDX_GROSS_DAILY_DIV nor DIV_YIELD "
                    "returned valid cached data."
                )
        else:
            required = ["LAST_PRICE", "LAST_PRICE.1", "PE_RATIO"]
            absent = [field for field in required if field not in data.columns]
            if absent:
                raise ValueError(f"{sheet}: missing Bloomberg fields {absent}")
            raw["total_level"] = pd.to_numeric(
                data["LAST_PRICE"], errors="coerce"
            )
            raw["price_level"] = pd.to_numeric(
                data["LAST_PRICE.1"], errors="coerce"
            )
            raw["pe"] = pd.to_numeric(data["PE_RATIO"], errors="coerce")
            raw["dividend_method"] = "total_return_minus_price_return"

        raw = raw[~raw.index.duplicated(keep="last")].sort_index()
        raw_by_series[series_name] = raw
    return raw_by_series


def calculate_daily_components(
    raw: pd.DataFrame, pe_window: int = 21, pe_lag: int = 21
) -> pd.DataFrame:
    data = raw.copy()
    data["total_return"] = data["total_level"].pct_change(fill_method=None)

    method_values = (
        data["dividend_method"].dropna().astype(str).unique().tolist()
        if "dividend_method" in data
        else []
    )
    method = method_values[0] if method_values else ""

    if method == "total_return_minus_price_return":
        data["price_return"] = data["price_level"].pct_change(fill_method=None)
        data["dividends"] = data["total_return"] - data["price_return"]
    elif method == "INDX_GROSS_DAILY_DIV":
        data["dividends"] = data["dividend_points"].fillna(0).div(
            data["total_level"].shift(1)
        )
        data["price_return"] = data["total_return"] - data["dividends"]
    elif method == "DIV_YIELD/252_APPROXIMATION":
        # Bloomberg dividend-yield fields are normally percentage points.
        dy = pd.to_numeric(data["dividend_yield"], errors="coerce")
        if dy.dropna().median() > 1:
            dy = dy / 100.0
        data["dividends"] = (1.0 + dy.clip(lower=0.0)) ** (1.0 / 252.0) - 1.0
        data["price_return"] = data["total_return"] - data["dividends"]
    else:
        raise ValueError(f"Unsupported dividend calculation method: {method}")

    pe = pd.to_numeric(data["pe"], errors="coerce")
    pe_smooth = pe.rolling(pe_window, min_periods=pe_window).mean()
    pe_ratio = pe_smooth.div(pe_smooth.shift(pe_lag))
    pe_ratio = pe_ratio.where(pe_ratio > 0)
    data["valuation"] = pe_ratio.pow(1.0 / pe_lag) - 1.0
    data["earnings_growth"] = (
        data["total_return"] - data["dividends"] - data["valuation"]
    )

    required = ["total_return", *COMPONENTS]
    data = data.dropna(subset=required).copy()
    data["daily_reconciliation_error"] = data["total_return"] - data[
        COMPONENTS
    ].sum(axis=1)
    data["dividend_method"] = method
    return data


def _carino_k(return_series: pd.Series) -> pd.Series:
    r = pd.to_numeric(return_series, errors="coerce")
    k = pd.Series(index=r.index, dtype=float)
    near_zero = r.abs() < 1e-12
    k.loc[near_zero] = 1.0
    valid = (~near_zero) & (r > -1.0)
    k.loc[valid] = np.log1p(r.loc[valid]) / r.loc[valid]
    return k


def carino_period(data: pd.DataFrame) -> pd.Series:
    data = data.dropna(subset=["total_return", *COMPONENTS])
    if data.empty:
        return pd.Series(
            {"total_return": np.nan, **{c: np.nan for c in COMPONENTS}}
        )
    r = data["total_return"]
    total = (1.0 + r).prod() - 1.0
    daily_k = _carino_k(r)
    if abs(total) < 1e-12:
        period_k = 1.0
    elif total > -1.0:
        period_k = math.log1p(total) / total
    else:
        return pd.Series(
            {"total_return": total, **{c: np.nan for c in COMPONENTS}}
        )
    contributions = data[COMPONENTS].mul(daily_k, axis=0).sum() / period_k
    result = pd.Series({"total_return": total, **contributions.to_dict()})
    result["reconciliation_error"] = total - contributions.sum()
    return result


def carino_history(data: pd.DataFrame) -> pd.DataFrame:
    data = data.dropna(subset=["total_return", *COMPONENTS])
    daily_k = _carino_k(data["total_return"])
    cumulative_total = (1.0 + data["total_return"]).cumprod() - 1.0
    period_k = pd.Series(1.0, index=data.index)
    nonzero = cumulative_total.abs() >= 1e-12
    valid = nonzero & (cumulative_total > -1.0)
    period_k.loc[valid] = (
        np.log1p(cumulative_total.loc[valid])
        / cumulative_total.loc[valid]
    )
    weighted_components = data[COMPONENTS].mul(daily_k, axis=0).cumsum()
    contributions = weighted_components.div(period_k, axis=0)
    output = pd.concat(
        [cumulative_total.rename("total_return"), contributions], axis=1
    )
    output["reconciliation_error"] = output["total_return"] - output[
        COMPONENTS
    ].sum(axis=1)
    return output


def build_annual_table(
    daily_by_series: Dict[str, pd.DataFrame]
) -> pd.DataFrame:
    rows = []
    for name, data in daily_by_series.items():
        for year, group in data.groupby(data.index.year):
            result = carino_period(group)
            row = {"series": name, "year": int(year), **result.to_dict()}
            row["start_date"] = group.index.min()
            row["end_date"] = group.index.max()
            row["observations"] = len(group)
            rows.append(row)
    return pd.DataFrame(rows).sort_values(["series", "year"])


def build_period_table(
    daily_by_series: Dict[str, pd.DataFrame]
) -> pd.DataFrame:
    rows = []
    for name, data in daily_by_series.items():
        if data.empty:
            continue
        end = data.index.max()
        periods = {
            "YTD": data.loc[data.index.year == end.year],
            "Trailing 12M": data.loc[data.index > end - pd.DateOffset(years=1)],
        }
        for period_name, group in periods.items():
            result = carino_period(group)
            rows.append(
                {
                    "series": name,
                    "period": period_name,
                    "as_of_date": end,
                    "start_date": group.index.min(),
                    "end_date": group.index.max(),
                    "observations": len(group),
                    **result.to_dict(),
                }
            )
    return pd.DataFrame(rows).sort_values(["period", "series"])


def build_history_table(
    daily_by_series: Dict[str, pd.DataFrame]
) -> pd.DataFrame:
    frames = []
    for name, data in daily_by_series.items():
        history = carino_history(data)
        history = history.resample("ME").last()
        history.insert(0, "series", name)
        frames.append(history.reset_index())
    return pd.concat(frames, ignore_index=True)


def _percent_axis(ax):
    from matplotlib.ticker import PercentFormatter

    ax.yaxis.set_major_formatter(PercentFormatter(1.0, decimals=0))
    ax.axhline(0, color="#A8A8A8", linewidth=0.8)
    ax.grid(axis="y", color="#E8E8E8", linewidth=0.7)


def plot_history(
    name: str, history: pd.DataFrame, output_dir: Path
) -> None:
    frame = history.loc[history["series"] == name].set_index("date")
    if frame.empty:
        return
    fig, ax = plt.subplots(figsize=(12, 6.5))
    for col in ["total_return", *COMPONENTS]:
        ax.plot(
            frame.index,
            frame[col],
            label=LABELS[col],
            color=COLORS[col],
            linewidth=2.2 if col == "total_return" else 1.8,
        )
    ax.set_title(
        f"{name}: cumulative return decomposition",
        loc="left",
        fontsize=15,
        fontweight="bold",
        pad=14,
    )
    _percent_axis(ax)
    ax.legend(ncol=4, frameon=False, loc="upper left")
    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    fig.tight_layout()
    fig.savefig(
        output_dir / f"{slugify(name)}_history.png",
        dpi=180,
        bbox_inches="tight",
    )
    plt.close(fig)


def _stacked_bar_with_negatives(ax, x, frame: pd.DataFrame):
    positive_base = np.zeros(len(frame))
    negative_base = np.zeros(len(frame))
    for component in COMPONENTS:
        values = frame[component].fillna(0).to_numpy()
        bottoms = np.where(values >= 0, positive_base, negative_base)
        ax.bar(
            x,
            values,
            bottom=bottoms,
            color=COLORS[component],
            label=LABELS[component],
            width=0.72,
        )
        positive_base += np.where(values >= 0, values, 0)
        negative_base += np.where(values < 0, values, 0)


def plot_annual(name: str, annual: pd.DataFrame, output_dir: Path) -> None:
    frame = annual.loc[annual["series"] == name].copy()
    if frame.empty:
        return
    x = np.arange(len(frame))
    fig, ax = plt.subplots(figsize=(12, 6.5))
    _stacked_bar_with_negatives(ax, x, frame)
    ax.plot(
        x,
        frame["total_return"],
        color=COLORS["total_return"],
        marker="o",
        linewidth=1.5,
        markersize=4,
        label=LABELS["total_return"],
    )
    ax.set_xticks(x)
    ax.set_xticklabels(frame["year"].astype(str), rotation=45, ha="right")
    ax.set_title(
        f"{name}: annual return decomposition",
        loc="left",
        fontsize=15,
        fontweight="bold",
    )
    _percent_axis(ax)
    ax.legend(ncol=4, frameon=False, loc="upper left")
    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    fig.tight_layout()
    fig.savefig(
        output_dir / f"{slugify(name)}_annual.png",
        dpi=180,
        bbox_inches="tight",
    )
    plt.close(fig)


def plot_period_summary(periods: pd.DataFrame, output_file: Path) -> None:
    if periods.empty:
        return
    period_order = ["YTD", "Trailing 12M"]
    fig, axes = plt.subplots(
        1, 2, figsize=(16, max(6.5, 0.46 * periods["series"].nunique()))
    )
    for ax, period_name in zip(axes, period_order):
        frame = periods.loc[periods["period"] == period_name].copy()
        frame = frame.sort_values("total_return")
        y = np.arange(len(frame))
        positive_base = np.zeros(len(frame))
        negative_base = np.zeros(len(frame))
        for component in COMPONENTS:
            values = frame[component].fillna(0).to_numpy()
            left = np.where(values >= 0, positive_base, negative_base)
            ax.barh(
                y,
                values,
                left=left,
                color=COLORS[component],
                label=LABELS[component],
                height=0.7,
            )
            positive_base += np.where(values >= 0, values, 0)
            negative_base += np.where(values < 0, values, 0)
        ax.scatter(
            frame["total_return"],
            y,
            color=COLORS["total_return"],
            s=22,
            zorder=4,
            label=LABELS["total_return"],
        )
        ax.set_yticks(y)
        ax.set_yticklabels(frame["series"])
        ax.set_title(period_name, loc="left", fontweight="bold")
        ax.axvline(0, color="#999999", linewidth=0.8)
        ax.grid(axis="x", color="#E8E8E8", linewidth=0.7)
        from matplotlib.ticker import PercentFormatter

        ax.xaxis.set_major_formatter(PercentFormatter(1.0, decimals=0))
        for spine in ["top", "right", "left"]:
            ax.spines[spine].set_visible(False)
    handles, labels = axes[0].get_legend_handles_labels()
    fig.legend(
        handles,
        labels,
        ncol=4,
        frameon=False,
        loc="upper center",
        bbox_to_anchor=(0.5, 0.91),
    )
    as_of = pd.to_datetime(periods["as_of_date"]).max().strftime("%d %b %Y")
    fig.suptitle(
        f"Brazil equities: return decomposition by index and sector\nAs of {as_of}",
        x=0.02,
        y=0.99,
        ha="left",
        fontsize=16,
        fontweight="bold",
    )
    fig.tight_layout(rect=(0, 0, 1, 0.84))
    fig.savefig(output_file, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_excel(
    output_file: Path,
    history: pd.DataFrame,
    annual: pd.DataFrame,
    periods: pd.DataFrame,
    daily_by_series: Dict[str, pd.DataFrame],
    metadata: pd.DataFrame,
) -> None:
    """Save tabular results; charts remain as high-resolution PNG files."""
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        periods.to_excel(writer, sheet_name="YTD_LTM", index=False)
        annual.to_excel(writer, sheet_name="Annual", index=False)
        history.to_excel(writer, sheet_name="History_Monthly", index=False)
        metadata.to_excel(writer, sheet_name="Methodology", index=False)
        daily_long = []
        for name, frame in daily_by_series.items():
            keep = frame[
                [
                    "total_return",
                    "dividends",
                    "earnings_growth",
                    "valuation",
                    "daily_reconciliation_error",
                ]
            ].copy()
            keep.insert(0, "series", name)
            daily_long.append(keep.reset_index())
        pd.concat(daily_long, ignore_index=True).to_excel(
            writer, sheet_name="Daily", index=False
        )


def validate_outputs(
    daily_by_series: Dict[str, pd.DataFrame],
    history: pd.DataFrame,
    annual: pd.DataFrame,
    periods: pd.DataFrame,
) -> dict:
    checks = {}
    checks["series_count"] = len(daily_by_series)
    checks["max_daily_reconciliation_error"] = float(
        max(
            data["daily_reconciliation_error"].abs().max()
            for data in daily_by_series.values()
        )
    )
    checks["max_history_reconciliation_error"] = float(
        history["reconciliation_error"].abs().max()
    )
    checks["max_annual_reconciliation_error"] = float(
        annual["reconciliation_error"].abs().max()
    )
    checks["max_period_reconciliation_error"] = float(
        periods["reconciliation_error"].abs().max()
    )
    checks["all_period_totals_finite"] = bool(
        np.isfinite(periods["total_return"]).all()
    )
    return checks


def run(args: argparse.Namespace) -> None:
    output_dir = Path(args.output).resolve()
    chart_history_dir = output_dir / "charts" / "history"
    chart_annual_dir = output_dir / "charts" / "annual"
    chart_history_dir.mkdir(parents=True, exist_ok=True)
    chart_annual_dir.mkdir(parents=True, exist_ok=True)

    if args.source == "supplied-xlsx":
        if not args.input:
            raise ValueError("--input is required for supplied-xlsx mode")
        raw_by_series = read_supplied_workbook(
            Path(args.input), args.sheet, args.series_name
        )
        specs_used = [
            SeriesSpec(
                args.series_name,
                "ticker in supplied workbook",
                "ticker in supplied workbook",
            )
        ]
    elif args.source == "bloomberg-xlsx":
        if not args.input:
            raise ValueError("--input is required for bloomberg-xlsx mode")
        raw_by_series = read_bloomberg_formula_workbook(Path(args.input))
        requested = set(args.series) if args.series else None
        if requested:
            missing = requested - set(raw_by_series)
            if missing:
                raise ValueError(f"Unknown series: {sorted(missing)}")
            raw_by_series = {
                name: frame
                for name, frame in raw_by_series.items()
                if name in requested
            }
        specs_used = [
            spec for spec in SERIES_SPECS if spec.name in raw_by_series
        ]
    else:
        requested = set(args.series) if args.series else None
        specs_used = [
            spec
            for spec in SERIES_SPECS
            if requested is None or spec.name in requested
        ]
        if requested:
            missing = requested - {spec.name for spec in specs_used}
            if missing:
                raise ValueError(f"Unknown series: {sorted(missing)}")
        raw_by_series = {
            spec.name: pull_bloomberg_series(spec, args.start, args.end)
            for spec in specs_used
        }

    daily_by_series = {}
    failures = {}
    for name, raw in raw_by_series.items():
        try:
            daily = calculate_daily_components(
                raw, pe_window=args.pe_window, pe_lag=args.pe_lag
            )
            if len(daily) < args.min_observations:
                raise ValueError(
                    f"only {len(daily)} valid observations; "
                    f"minimum is {args.min_observations}"
                )
            daily_by_series[name] = daily
        except Exception as exc:
            failures[name] = str(exc)

    if not daily_by_series:
        raise RuntimeError(f"No series could be calculated. Failures: {failures}")

    history = build_history_table(daily_by_series)
    annual = build_annual_table(daily_by_series)
    periods = build_period_table(daily_by_series)

    history.to_csv(output_dir / "history_monthly.csv", index=False)
    annual.to_csv(output_dir / "annual_decomposition.csv", index=False)
    periods.to_csv(output_dir / "period_decomposition.csv", index=False)

    for name in daily_by_series:
        plot_history(name, history, chart_history_dir)
        plot_annual(name, annual, chart_annual_dir)
    plot_period_summary(periods, output_dir / "charts" / "ytd_ltm_summary.png")

    metadata_rows = [
        {
            "item": "Methodology",
            "value": (
                "Daily methodology adapted from the supplied global code. "
                "P/E uses a moving average and 21-trading-day change; "
                "earnings growth is the implied residual."
            ),
        },
        {
            "item": "Multi-period linking",
            "value": (
                "Carino linking. Dividends + earnings growth + valuation "
                "reconcile exactly to compounded total return."
            ),
        },
        {
            "item": "P/E moving-average window",
            "value": args.pe_window,
        },
        {"item": "P/E lag", "value": args.pe_lag},
        {
            "item": "Ibovespa dividend source",
            "value": (
                "Bloomberg INDX_GROSS_DAILY_DIV. If unavailable, the code "
                "falls back to DIV_YIELD/252 and flags the approximation."
            ),
        },
        {
            "item": "MSCI dividend source",
            "value": "Gross total return index minus price return index.",
        },
        {
            "item": "Failed series",
            "value": json.dumps(failures, ensure_ascii=False),
        },
    ]
    metadata = pd.DataFrame(metadata_rows)
    save_excel(
        output_dir / "return_decomposition_brazil.xlsx",
        history,
        annual,
        periods,
        daily_by_series,
        metadata,
    )

    checks = validate_outputs(daily_by_series, history, annual, periods)
    checks["failures"] = failures
    checks["latest_date_by_series"] = {
        name: frame.index.max().strftime("%Y-%m-%d")
        for name, frame in daily_by_series.items()
    }
    checks["dividend_method_by_series"] = {
        name: frame["dividend_method"].iloc[0]
        for name, frame in daily_by_series.items()
    }
    (output_dir / "validation.json").write_text(
        json.dumps(checks, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print(json.dumps(checks, indent=2, ensure_ascii=False))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Ibovespa and MSCI Brazil return decomposition"
    )
    parser.add_argument(
        "--source",
        choices=["bloomberg", "bloomberg-xlsx", "supplied-xlsx"],
        default="bloomberg",
    )
    parser.add_argument("--input", help="Input workbook for supplied-xlsx mode")
    parser.add_argument("--sheet", default="Brazil")
    parser.add_argument("--series-name", default="MSCI Brazil")
    parser.add_argument("--start", default="2000-01-01")
    parser.add_argument("--end", default=date.today().isoformat())
    parser.add_argument(
        "--series",
        nargs="*",
        help="Optional exact series names; default pulls all configured series",
    )
    parser.add_argument("--pe-window", type=int, default=21)
    parser.add_argument("--pe-lag", type=int, default=21)
    parser.add_argument("--min-observations", type=int, default=126)
    parser.add_argument(
        "--output",
        default="outputs/ibovespa_msci_return_decomposition/results",
    )
    return parser.parse_args()


if __name__ == "__main__":
    run(parse_args())

