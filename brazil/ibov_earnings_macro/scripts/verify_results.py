"""Valida dados e principais regressões publicadas neste diretório."""

from __future__ import annotations

import json
import math
from pathlib import Path

import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
RESULTS = ROOT / "results"

XS = [
    "d_focus_pib3",
    "d_bcom_usd3",
    "d_usdbrl3",
    "d_ipca12_3_available",
    "d_swap360_3",
]


def ols_hac(y: np.ndarray, x: np.ndarray, lags: int = 6) -> dict:
    y = np.asarray(y, dtype=float)
    x = np.asarray(x, dtype=float)
    n, k = x.shape
    beta = np.linalg.lstsq(x, y, rcond=None)[0]
    residual = y - x @ beta
    xtx_inverse = np.linalg.inv(x.T @ x)
    u = x * residual[:, None]
    covariance_sum = u.T @ u
    for lag in range(1, lags + 1):
        weight = 1 - lag / (lags + 1)
        gamma = u[lag:].T @ u[:-lag]
        covariance_sum += weight * (gamma + gamma.T)
    covariance = xtx_inverse @ covariance_sum @ xtx_inverse * n / max(n - k, 1)
    standard_error = np.sqrt(np.diag(covariance))
    t_stat = beta / standard_error
    r2 = 1 - residual.var() / y.var()
    return {"beta": beta, "se": standard_error, "t": t_stat, "r2": r2, "n": n}


def fit(frame: pd.DataFrame, target: str, variables: list[str]) -> dict:
    data = frame[[target] + variables].dropna()
    design = np.column_stack(
        [np.ones(len(data))] + [data[variable].to_numpy(float) for variable in variables]
    )
    result = ols_hac(data[target].to_numpy(float), design)
    result["p"] = np.array(
        [2 * (1 - 0.5 * (1 + math.erf(abs(value) / math.sqrt(2)))) for value in result["t"]]
    )
    standardized = data[[target] + variables].apply(
        lambda series: (series - series.mean()) / series.std(ddof=0)
    )
    standardized_design = np.column_stack(
        [np.ones(len(data))]
        + [standardized[variable].to_numpy(float) for variable in variables]
    )
    result["beta_std"] = np.linalg.lstsq(
        standardized_design, standardized[target].to_numpy(float), rcond=None
    )[0]
    return result


def log_change(series: pd.Series, periods: int = 3) -> pd.Series:
    lag = series.shift(periods)
    valid = series.gt(0) & lag.gt(0)
    result = pd.Series(np.nan, index=series.index, dtype=float)
    result.loc[valid] = 100 * np.log(series.loc[valid] / lag.loc[valid])
    return result


def assert_close(actual, expected, label: str, tolerance: float = 1e-9) -> None:
    if not np.allclose(actual, expected, rtol=tolerance, atol=tolerance, equal_nan=True):
        raise AssertionError(f"Falha em {label}: atual={actual}, esperado={expected}")


def main() -> None:
    checks: list[str] = []

    sector = pd.read_csv(DATA / "earnings_12m_fwd_sector_monthly.csv")
    if len(sector) != 225 or sector["month"].duplicated().any():
        raise AssertionError("Cobertura ou unicidade mensal da série setorial")
    periods = pd.PeriodIndex(sector["month"], freq="M")
    expected_periods = pd.period_range("2008-01", "2026-09", freq="M")
    if not periods.equals(expected_periods):
        raise AssertionError("Calendário mensal setorial não é contínuo")
    checks.append("225 meses contínuos de earnings setoriais")

    base100 = pd.read_csv(DATA / "earnings_12m_fwd_sector_base100_2020_01.csv")
    base_row = base100.loc[base100["month"].eq("2020-01")].iloc[0]
    numeric_base = base_row.drop(labels=["month", "date", "source_file"]).astype(float)
    assert_close(numeric_base.to_numpy(), np.full(len(numeric_base), 100.0), "base 100")
    checks.append("base 100 reconciliada em janeiro de 2020")

    revisions = pd.read_csv(DATA / "earnings_revisions_3m_sector_monthly.csv")
    for target in ["IBOV", "Energy", "Financials", "Materials"]:
        calculated = log_change(pd.to_numeric(sector[target], errors="coerce"))
        published = pd.to_numeric(revisions[target], errors="coerce")
        assert_close(calculated.to_numpy(), published.to_numpy(), f"revisão {target}")
    checks.append("revisões de três meses reconciliadas")

    macro = pd.read_csv(DATA / "macro_and_ibov_eps_monthly.csv").set_index("month")
    macro.index = pd.PeriodIndex(macro.index, freq="M")
    calculated_eps_revision = 100 * np.log(macro["eps"] / macro["eps"].shift(3))
    assert_close(calculated_eps_revision.to_numpy(), macro["y_rev3"].to_numpy(), "revisão EPS")

    labels = macro.index.astype(str)
    primary_macro = macro.loc[~((labels >= "2013-01") & (labels <= "2014-12"))]
    ibov_fit = fit(primary_macro, "y_rev3", XS)
    stored_ibov = pd.read_csv(RESULTS / "ibov_contemporaneous_coefficients.csv")
    stored_ibov = stored_ibov.loc[stored_ibov["sample"].eq("ex_2013_14")].set_index("variable")
    if ibov_fit["n"] != int(stored_ibov["n"].iloc[0]):
        raise AssertionError("n do modelo IBOV")
    assert_close(ibov_fit["r2"], float(stored_ibov["r2"].iloc[0]), "R² IBOV")
    for position, variable in enumerate(XS, 1):
        assert_close(ibov_fit["beta"][position], stored_ibov.loc[variable, "beta"], f"beta IBOV {variable}")
        assert_close(ibov_fit["beta_std"][position], stored_ibov.loc[variable, "beta_std"], f"beta std IBOV {variable}")
        assert_close(ibov_fit["p"][position], stored_ibov.loc[variable, "p_hac"], f"p IBOV {variable}")
    checks.append("modelo contemporâneo do BEst EPS reproduzido")

    sector_panel = revisions.set_index("month")
    sector_panel.index = pd.PeriodIndex(sector_panel.index, freq="M")
    sector_panel = sector_panel.join(macro[XS], how="outer")
    stored_sector = pd.read_csv(RESULTS / "sector_regressions.csv")
    stored_sector = stored_sector.loc[stored_sector["sample"].eq("ex_2013_14")].set_index("target")
    for target in stored_sector.index:
        sample = sector_panel.loc[
            ~((sector_panel.index.astype(str) >= "2013-01") & (sector_panel.index.astype(str) <= "2014-12"))
        ]
        sector_fit = fit(sample, target, XS)
        if sector_fit["n"] != int(stored_sector.loc[target, "n"]):
            raise AssertionError(f"n setorial {target}")
        assert_close(sector_fit["r2"], stored_sector.loc[target, "r2"], f"R² setorial {target}")
        for position, variable in enumerate(XS, 1):
            assert_close(
                sector_fit["beta"][position],
                stored_sector.loc[target, f"{variable}_beta"],
                f"beta setorial {target} {variable}",
            )
    checks.append("regressões principais de IBOV e 11 setores reproduzidas")

    matrix = pd.read_csv(RESULTS / "focus_swap_sensitivity_matrix.csv")
    focus_beta = stored_ibov.loc["d_focus_pib3", "beta"]
    swap_beta = stored_ibov.loc["d_swap360_3", "beta"]
    for _, row in matrix.iterrows():
        focus = float(row["focus_delta_pp"])
        for column in matrix.columns[1:]:
            swap = float(column.removeprefix("swap_").removesuffix("pp"))
            assert_close(row[column], focus_beta * focus + swap_beta * swap, f"matriz {focus} {swap}")
    checks.append("matriz Focus–swap reconciliada")

    energy = pd.read_csv(DATA / "energy_company_earnings_12m_fwd_monthly.csv").set_index("month")
    energy.index = pd.PeriodIndex(energy.index, freq="M")
    energy_revisions = energy[["Energy reported", "Petrobras"]].apply(log_change)
    comparison = energy_revisions.loc[
        ~((energy_revisions.index.astype(str) >= "2013-01") & (energy_revisions.index.astype(str) <= "2014-12"))
    ].dropna()
    correlation = comparison.corr().iloc[0, 1]
    summary = json.loads((RESULTS / "energy_petrobras_brent_summary.json").read_text(encoding="utf-8"))
    assert_close(
        correlation,
        summary["revision_correlations"]["primary"]["Petrobras"]["correlation"],
        "correlação Energy Petrobras",
    )
    checks.append("decomposição de Energia reconciliada")

    required_files = [
        ROOT / "workbook" / "earnings_12m_fwd_sector_2008_2026.xlsx",
        ROOT / "charts" / "earnings_12m_fwd_sector_base100.png",
        DATA / "sources.csv",
    ]
    if not all(path.is_file() and path.stat().st_size > 0 for path in required_files):
        raise AssertionError("Arquivos complementares ausentes")
    checks.append("planilha, gráfico e fontes presentes")

    for check in checks:
        print(f"OK - {check}")
    print(f"OK - {len(checks)} verificações concluídas")


if __name__ == "__main__":
    main()
