"""
Ibovespa and MSCI Brazil return decomposition.

The historical series follows the methodology of the first code supplied by
the user:

    daily total return
      = daily dividend contribution
      + daily implied earnings growth
      + daily multiple expansion

Multiple expansion is calculated from a 21-day moving average of P/E versus
the same moving average 21 trading days earlier, converted into a daily rate.
Earnings growth is the residual.

For annual, YTD and trailing-12-month tables, valuation is instead measured
directly from the P/E at the period boundaries. This makes the reported
re-rating auditable against Bloomberg. The other period components are
residualized so that they add exactly to total return.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd


COMPONENTS = ["dividends", "earnings_growth", "valuation"]
ALL_RETURNS = ["total_return", *COMPONENTS]
COLORS = {
    "total_return": "#202124",
    "dividends": "#2E86AB",
    "earnings_growth": "#3A9D5D",
    "valuation": "#E67E22",
}

DEFAULT_INPUT = Path(
    r"\\xpdocs\Research\Equities\Estrategia\Raio-XP da Bolsa\2026"
    r"\8. Agosto\return_decomposition"
    r"\bloomberg_input_ibovespa_msci_sectors.xlsx"
)
DEFAULT_OUTPUT = DEFAULT_INPUT.parent / "results_ibov_msci_brazil"
LABELS = {
    "total_return": "Total return",
    "dividends": "Dividends",
    "earnings_growth": "Earnings growth",
    "valuation": "Valuation",
}


def slugify(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", value.lower()).strip("_")


def _read_sheet(workbook: Path, sheet: str) -> pd.DataFrame:
    data = pd.read_excel(workbook, sheet_name=sheet, header=5)
    if "Dates" not in data.columns:
        raise ValueError(
            f"{sheet}: Bloomberg dates not found. Refresh, wait, save and close "
            "the workbook before running Python."
        )
    dates = pd.to_datetime(data["Dates"], errors="coerce")
    data = data.loc[dates.notna()].copy()
    data.index = dates.loc[dates.notna()]
    data.index.name = "date"
    return data


def read_bloomberg_workbook(workbook: Path) -> dict[str, pd.DataFrame]:
    sheets = set(pd.ExcelFile(workbook).sheet_names)
    required = {"IBOV", "MSCI Brazil"}
    missing = required - sheets
    if missing:
        raise ValueError(f"Missing sheets: {sorted(missing)}")

    output: dict[str, pd.DataFrame] = {}

    msci = _read_sheet(workbook, "MSCI Brazil")
    required_msci = ["LAST_PRICE", "PE_RATIO"]
    missing_msci = [c for c in required_msci if c not in msci.columns]
    if missing_msci:
        raise ValueError(f"MSCI Brazil: missing fields {missing_msci}")
    msci_raw = pd.DataFrame(index=msci.index)

    # Accept both supported Bloomberg workbook layouts:
    # 1) simplified workbook: M2BR total-return block + MXBR price block;
    # 2) original sector workbook: MXBR price index with Bloomberg's
    #    TOT_RETURN_INDEX_GROSS_DVDS field in the same block.
    if "LAST_PRICE.1" in msci.columns:
        msci_raw["total_level"] = pd.to_numeric(
            msci["LAST_PRICE"], errors="coerce"
        )
        msci_raw["price_level"] = pd.to_numeric(
            msci["LAST_PRICE.1"], errors="coerce"
        )
    elif "TOT_RETURN_INDEX_GROSS_DVDS" in msci.columns:
        msci_raw["total_level"] = pd.to_numeric(
            msci["TOT_RETURN_INDEX_GROSS_DVDS"], errors="coerce"
        )
        msci_raw["price_level"] = pd.to_numeric(
            msci["LAST_PRICE"], errors="coerce"
        )
    else:
        raise ValueError(
            "MSCI Brazil: the workbook must contain either LAST_PRICE.1 "
            "(separate total/price blocks) or TOT_RETURN_INDEX_GROSS_DVDS "
            "(the original sector-workbook layout)."
        )

    if (
        msci_raw["total_level"].notna().sum() < 2
        or msci_raw["price_level"].notna().sum() < 2
    ):
        raise ValueError(
            "MSCI Brazil: insufficient refreshed Bloomberg data. Open the "
            "workbook in Bloomberg Excel, refresh, wait, save and close it."
        )
    msci_raw["pe"] = pd.to_numeric(msci["PE_RATIO"], errors="coerce")
    msci_raw["dividend_method"] = "total_return_minus_price_return"
    output["MSCI Brazil"] = msci_raw

    ibov = _read_sheet(workbook, "IBOV")
    required_ibov = [
        "LAST_PRICE",
        "PE_RATIO",
        "INDX_GROSS_DAILY_DIV",
        "DIV_YIELD",
    ]
    missing_ibov = [c for c in required_ibov if c not in ibov.columns]
    if missing_ibov:
        raise ValueError(f"IBOV: missing fields {missing_ibov}")
    ibov_raw = pd.DataFrame(index=ibov.index)
    ibov_raw["total_level"] = pd.to_numeric(
        ibov["LAST_PRICE"], errors="coerce"
    )
    ibov_raw["pe"] = pd.to_numeric(ibov["PE_RATIO"], errors="coerce")
    ibov_raw["dividend_points"] = pd.to_numeric(
        ibov["INDX_GROSS_DAILY_DIV"], errors="coerce"
    )
    ibov_raw["dividend_yield"] = pd.to_numeric(
        ibov["DIV_YIELD"], errors="coerce"
    )
    if (
        ibov_raw["dividend_points"].notna().sum() >= 126
        and ibov_raw["dividend_points"].fillna(0).clip(lower=0).sum() > 0
    ):
        ibov_raw["dividend_method"] = "INDX_GROSS_DAILY_DIV"
    elif ibov_raw["dividend_yield"].notna().sum() >= 126:
        ibov_raw["dividend_method"] = "DIV_YIELD/252_APPROXIMATION"
    else:
        raise ValueError(
            "IBOV: neither daily dividend points nor dividend yield returned "
            "sufficient historical data."
        )
    output["Ibovespa"] = ibov_raw

    return {
        name: frame[~frame.index.duplicated(keep="last")].sort_index()
        for name, frame in output.items()
    }


def decompose_daily(
    raw: pd.DataFrame, pe_window: int = 21, pe_lag: int = 21
) -> pd.DataFrame:
    data = raw.copy()
    data["total_return"] = data["total_level"].pct_change(fill_method=None)
    method = data["dividend_method"].dropna().iloc[0]

    if method == "total_return_minus_price_return":
        data["price_return"] = data["price_level"].pct_change(fill_method=None)
        data["dividends"] = data["total_return"] - data["price_return"]
    elif method == "INDX_GROSS_DAILY_DIV":
        data["dividends"] = data["dividend_points"].fillna(0).div(
            data["total_level"].shift(1)
        )
        data["price_return"] = data["total_return"] - data["dividends"]
    elif method == "DIV_YIELD/252_APPROXIMATION":
        dy = pd.to_numeric(data["dividend_yield"], errors="coerce")
        if dy.dropna().median() > 1:
            dy = dy / 100.0
        data["dividends"] = (
            1.0 + dy.shift(1).clip(lower=0)
        ) ** (1.0 / 252.0) - 1.0
        data["price_return"] = data["total_return"] - data["dividends"]
    else:
        raise ValueError(f"Unsupported dividend method: {method}")

    pe = pd.to_numeric(data["pe"], errors="coerce")
    pe_ma = pe.rolling(pe_window, min_periods=pe_window).mean()
    data["pe_21d_ma"] = pe_ma
    pe_change = pe_ma.div(pe_ma.shift(pe_lag)) - 1.0
    pe_change = pe_change.where(1.0 + pe_change > 0)
    data["valuation"] = (1.0 + pe_change) ** (1.0 / pe_lag) - 1.0
    data["earnings_growth"] = (
        data["total_return"] - data["dividends"] - data["valuation"]
    )

    data = data.dropna(subset=ALL_RETURNS).copy()
    data["daily_reconciliation_error"] = data["total_return"] - data[
        COMPONENTS
    ].sum(axis=1)
    data["dividend_method"] = method
    return data


def independently_compound(data: pd.DataFrame) -> pd.Series:
    result = {
        column: (1.0 + data[column].dropna()).prod() - 1.0
        for column in ALL_RETURNS
    }
    result["component_sum"] = sum(result[column] for column in COMPONENTS)
    result["compounding_residual"] = (
        result["total_return"] - result["component_sum"]
    )
    return pd.Series(result)


def endpoint_period_decomposition(
    full_data: pd.DataFrame,
    group: pd.DataFrame,
    endpoint_window: int = 1,
) -> pd.Series:
    """Build an intuitive, exactly additive period attribution.

    Period returns start on the first date in ``group`` and therefore use the
    latest P/E observation strictly before that date as the initial multiple.
    With endpoint_window=1, valuation is exactly PE_end / PE_start - 1.
    """
    if group.empty:
        raise ValueError("Cannot decompose an empty period")
    if endpoint_window < 1:
        raise ValueError("endpoint_window must be at least 1")

    first_return_date = group.index.min()
    end_date = group.index.max()
    pe = pd.to_numeric(full_data["pe"], errors="coerce").dropna()
    start_sample = pe.loc[pe.index < first_return_date].tail(endpoint_window)
    end_sample = pe.loc[pe.index <= end_date].tail(endpoint_window)
    if start_sample.empty or end_sample.empty:
        raise ValueError(
            f"Insufficient P/E endpoint data for {first_return_date:%Y-%m-%d} "
            f"to {end_date:%Y-%m-%d}"
        )

    pe_start = float(start_sample.mean())
    pe_end = float(end_sample.mean())
    if pe_start <= 0 or pe_end <= 0:
        raise ValueError("P/E endpoint values must be positive")

    total_return = float((1.0 + group["total_return"]).prod() - 1.0)
    price_return = float((1.0 + group["price_return"]).prod() - 1.0)
    dividends = total_return - price_return
    valuation = pe_end / pe_start - 1.0
    earnings_growth = price_return - valuation
    component_sum = dividends + earnings_growth + valuation

    return pd.Series(
        {
            "total_return": total_return,
            "price_return": price_return,
            "dividends": dividends,
            "earnings_growth": earnings_growth,
            "valuation": valuation,
            "component_sum": component_sum,
            "compounding_residual": total_return - component_sum,
            "pe_start": pe_start,
            "pe_end": pe_end,
            "pe_direct_change": pe_end / pe_start - 1.0,
            "valuation_audit_error": valuation - (pe_end / pe_start - 1.0),
            "pe_start_date": start_sample.index.max(),
            "pe_end_date": end_sample.index.max(),
            "endpoint_window": endpoint_window,
        }
    )


def cumulative_history(
    data: pd.DataFrame, analysis_start: str
) -> pd.DataFrame:
    frame = data.loc[analysis_start:, ALL_RETURNS].fillna(0)
    output = (1.0 + frame).cumprod() - 1.0
    output["earnings_growth_63dma"] = output["earnings_growth"].rolling(63).mean()
    output["valuation_63dma"] = output["valuation"].rolling(63).mean()
    return output


def build_history(
    daily: dict[str, pd.DataFrame], analysis_start: str
) -> pd.DataFrame:
    frames = []
    for name, data in daily.items():
        frame = cumulative_history(data, analysis_start).resample("ME").last()
        frame.insert(0, "series", name)
        frames.append(frame.reset_index())
    return pd.concat(frames, ignore_index=True)


def build_annual(
    daily: dict[str, pd.DataFrame],
    analysis_start: str,
    endpoint_window: int = 1,
) -> pd.DataFrame:
    rows = []
    for name, data in daily.items():
        eligible = data.loc[analysis_start:]
        for year, group in eligible.groupby(eligible.index.year):
            try:
                row = endpoint_period_decomposition(
                    data, group, endpoint_window=endpoint_window
                )
            except ValueError as exc:
                if "Insufficient P/E endpoint data" in str(exc):
                    # The first partial calendar year can lack a prior
                    # observation after the historical warm-up period.
                    continue
                raise
            rows.append(
                {
                    "series": name,
                    "year": int(year),
                    "start_date": group.index.min(),
                    "end_date": group.index.max(),
                    "observations": len(group),
                    **row.to_dict(),
                }
            )
    return pd.DataFrame(rows).sort_values(["series", "year"])


def build_periods(
    daily: dict[str, pd.DataFrame], endpoint_window: int = 1
) -> pd.DataFrame:
    rows = []
    for name, data in daily.items():
        end = data.index.max()
        periods = {
            "YTD": data.loc[data.index.year == end.year],
            "Trailing 12M": data.loc[data.index > end - pd.DateOffset(years=1)],
        }
        for period, group in periods.items():
            row = endpoint_period_decomposition(
                data, group, endpoint_window=endpoint_window
            )
            rows.append(
                {
                    "series": name,
                    "period": period,
                    "as_of_date": end,
                    "start_date": group.index.min(),
                    "end_date": group.index.max(),
                    "observations": len(group),
                    **row.to_dict(),
                }
            )
    return pd.DataFrame(rows).sort_values(["period", "series"])


def build_valuation_audit(
    annual: pd.DataFrame, periods: pd.DataFrame
) -> pd.DataFrame:
    annual_audit = annual.copy()
    annual_audit.insert(1, "scope", "Annual")
    annual_audit["period"] = annual_audit["year"].astype(str)

    period_audit = periods.copy()
    period_audit.insert(1, "scope", "Current period")

    audit = pd.concat([annual_audit, period_audit], ignore_index=True)
    columns = [
        "series",
        "scope",
        "period",
        "start_date",
        "end_date",
        "pe_start_date",
        "pe_end_date",
        "endpoint_window",
        "pe_start",
        "pe_end",
        "pe_direct_change",
        "valuation",
        "valuation_audit_error",
        "total_return",
        "price_return",
        "dividends",
        "earnings_growth",
        "component_sum",
        "compounding_residual",
    ]
    return audit[columns].sort_values(["series", "scope", "period"])


def _format_percent_axis(ax, orientation: str = "y"):
    from matplotlib.ticker import PercentFormatter

    formatter = PercentFormatter(1.0, decimals=0)
    if orientation == "y":
        ax.yaxis.set_major_formatter(formatter)
        ax.axhline(0, color="#999999", linewidth=0.8)
        ax.grid(axis="y", color="#E8E8E8", linewidth=0.7)
    else:
        ax.xaxis.set_major_formatter(formatter)
        ax.axvline(0, color="#999999", linewidth=0.8)
        ax.grid(axis="x", color="#E8E8E8", linewidth=0.7)


def plot_history(
    name: str, history: pd.DataFrame, output_dir: Path
) -> None:
    frame = history.loc[history["series"] == name].set_index("date")
    fig, ax = plt.subplots(figsize=(12, 6.5))
    for column in ALL_RETURNS:
        ax.plot(
            frame.index,
            frame[column],
            label=LABELS[column],
            color=COLORS[column],
            linewidth=2.2 if column == "total_return" else 1.8,
        )
    ax.set_title(
        f"{name}: cumulative return decomposition",
        loc="left",
        fontsize=15,
        fontweight="bold",
        pad=12,
    )
    _format_percent_axis(ax)
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


def plot_annual(name: str, annual: pd.DataFrame, output_dir: Path) -> None:
    frame = annual.loc[annual["series"] == name]
    x = np.arange(len(frame))
    fig, ax = plt.subplots(figsize=(12, 6.5))
    width = 0.22
    for offset, component in zip([-1, 0, 1], COMPONENTS):
        ax.bar(
            x + offset * width,
            frame[component],
            width=width,
            color=COLORS[component],
            label=LABELS[component],
        )
    ax.plot(
        x,
        frame["total_return"],
        color=COLORS["total_return"],
        marker="o",
        linewidth=1.7,
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
    _format_percent_axis(ax)
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


def plot_periods(periods: pd.DataFrame, output_file: Path) -> None:
    fig, axes = plt.subplots(1, 2, figsize=(14, 5.5))
    for ax, period in zip(axes, ["YTD", "Trailing 12M"]):
        frame = periods.loc[periods["period"] == period].copy()
        y = np.arange(len(frame))
        height = 0.20
        for offset, component in zip([-1, 0, 1], COMPONENTS):
            ax.barh(
                y + offset * height,
                frame[component],
                height=height,
                color=COLORS[component],
                label=LABELS[component],
            )
        ax.scatter(
            frame["total_return"],
            y,
            color=COLORS["total_return"],
            label=LABELS["total_return"],
            zorder=4,
        )
        ax.set_yticks(y)
        ax.set_yticklabels(frame["series"])
        ax.set_title(period, loc="left", fontweight="bold")
        _format_percent_axis(ax, "x")
        for spine in ["top", "right", "left"]:
            ax.spines[spine].set_visible(False)
    handles, labels = axes[0].get_legend_handles_labels()
    as_of = pd.to_datetime(periods["as_of_date"]).max().strftime("%d %b %Y")
    fig.suptitle(
        f"Ibovespa and MSCI Brazil: return decomposition\nAs of {as_of}",
        x=0.03,
        y=0.98,
        ha="left",
        fontsize=15,
        fontweight="bold",
    )
    fig.legend(
        handles,
        labels,
        ncol=4,
        frameon=False,
        loc="upper left",
        bbox_to_anchor=(0.03, 0.86),
    )
    fig.tight_layout(rect=(0, 0, 1, 0.76))
    fig.savefig(output_file, dpi=180, bbox_inches="tight")
    plt.close(fig)


def save_workbook(
    output_file: Path,
    daily: dict[str, pd.DataFrame],
    history: pd.DataFrame,
    annual: pd.DataFrame,
    periods: pd.DataFrame,
    valuation_audit: pd.DataFrame,
) -> None:
    daily_frames = []
    for name, frame in daily.items():
        keep = frame[
            [
                "total_return",
                "price_return",
                "pe",
                "pe_21d_ma",
                "dividends",
                "earnings_growth",
                "valuation",
                "daily_reconciliation_error",
            ]
        ].copy()
        keep.insert(0, "series", name)
        daily_frames.append(keep.reset_index())
    methodology = pd.DataFrame(
        [
            {
                "item": "Daily identity",
                "value": "Total return = dividends + implied earnings growth + valuation",
            },
            {
                "item": "Historical valuation",
                "value": "21-day average P/E change versus 21 trading days earlier, converted to daily rate",
            },
            {
                "item": "Annual / YTD / trailing 12M valuation",
                "value": "P/E at the end of the period divided by P/E immediately before the first return date, minus one",
            },
            {
                "item": "Period dividends",
                "value": "Compounded total return minus compounded price return",
            },
            {
                "item": "Period earnings growth",
                "value": "Compounded price return minus endpoint valuation contribution; includes the valuation/earnings interaction",
            },
            {
                "item": "Period aggregation",
                "value": "Dividends + earnings growth + valuation equals total return exactly",
            },
        ]
    )
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        periods.to_excel(writer, sheet_name="YTD_LTM", index=False)
        annual.to_excel(writer, sheet_name="Annual", index=False)
        valuation_audit.to_excel(
            writer, sheet_name="Valuation_Audit", index=False
        )
        history.to_excel(writer, sheet_name="History_Monthly", index=False)
        pd.concat(daily_frames, ignore_index=True).to_excel(
            writer, sheet_name="Daily", index=False
        )
        methodology.to_excel(writer, sheet_name="Methodology", index=False)


def run(args: argparse.Namespace) -> None:
    output_dir = Path(args.output).resolve()
    history_chart_dir = output_dir / "charts" / "history"
    annual_chart_dir = output_dir / "charts" / "annual"
    history_chart_dir.mkdir(parents=True, exist_ok=True)
    annual_chart_dir.mkdir(parents=True, exist_ok=True)

    raw = read_bloomberg_workbook(Path(args.input))
    daily = {
        name: decompose_daily(
            frame, pe_window=args.pe_window, pe_lag=args.pe_lag
        )
        for name, frame in raw.items()
    }
    failures = {
        name: f"only {len(frame)} valid observations"
        for name, frame in daily.items()
        if len(frame) < args.min_observations
    }
    if failures:
        raise RuntimeError(
            "Insufficient refreshed data. Check Bloomberg workbook: "
            + json.dumps(failures)
        )

    history = build_history(daily, args.analysis_start)
    annual = build_annual(
        daily,
        args.analysis_start,
        endpoint_window=args.endpoint_window,
    )
    periods = build_periods(
        daily, endpoint_window=args.endpoint_window
    )
    valuation_audit = build_valuation_audit(annual, periods)

    history.to_csv(output_dir / "history_monthly.csv", index=False)
    annual.to_csv(output_dir / "annual_decomposition.csv", index=False)
    periods.to_csv(output_dir / "period_decomposition.csv", index=False)
    valuation_audit.to_csv(
        output_dir / "valuation_audit.csv", index=False
    )
    save_workbook(
        output_dir / "return_decomposition_ibov_msci_brazil.xlsx",
        daily,
        history,
        annual,
        periods,
        valuation_audit,
    )

    for name in daily:
        plot_history(name, history, history_chart_dir)
        plot_annual(name, annual, annual_chart_dir)
    plot_periods(periods, output_dir / "charts" / "ytd_ltm_summary.png")

    validation = {
        "series": list(daily),
        "latest_date": {
            name: frame.index.max().strftime("%Y-%m-%d")
            for name, frame in daily.items()
        },
        "dividend_method": {
            name: frame["dividend_method"].iloc[0]
            for name, frame in daily.items()
        },
        "max_daily_reconciliation_error": max(
            float(frame["daily_reconciliation_error"].abs().max())
            for frame in daily.values()
        ),
        "max_period_reconciliation_error": float(
            max(
                annual["compounding_residual"].abs().max(),
                periods["compounding_residual"].abs().max(),
            )
        ),
        "max_valuation_audit_error": float(
            valuation_audit["valuation_audit_error"].abs().max()
        ),
        "endpoint_window": args.endpoint_window,
        "note": (
            "Historical lines retain the original smoothed daily methodology. "
            "Annual, YTD and trailing-12-month valuation use P/E endpoints and "
            "period components reconcile exactly to total return."
        ),
    }
    (output_dir / "validation.json").write_text(
        json.dumps(validation, indent=2), encoding="utf-8"
    )
    print(json.dumps(validation, indent=2))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Ibovespa and MSCI Brazil legacy return decomposition"
    )
    parser.add_argument(
        "--input",
        default=str(DEFAULT_INPUT),
        help=f"Bloomberg input workbook (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT),
        help=f"Output directory (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument("--analysis-start", default="2005-01-01")
    parser.add_argument("--pe-window", type=int, default=21)
    parser.add_argument("--pe-lag", type=int, default=21)
    parser.add_argument(
        "--endpoint-window",
        type=int,
        default=1,
        help=(
            "Number of P/E observations averaged at each period endpoint. "
            "Default 1 matches the Bloomberg point-to-point re-rating."
        ),
    )
    parser.add_argument("--min-observations", type=int, default=126)
    return parser.parse_args()


if __name__ == "__main__":
    run(parse_args())
