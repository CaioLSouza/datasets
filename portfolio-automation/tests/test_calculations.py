"""Testes unitários que não dependem dos arquivos corporativos."""

from __future__ import annotations

import unittest
from unittest.mock import patch

import numpy as np
import pandas as pd

from xp_carteiras import email_report as email
from xp_carteiras import performance as portfolio
from xp_carteiras.constants import arq_base100, arq_performance
from xp_carteiras.email_report import EMAIL_FILENAME
from xp_carteiras.settings import Settings


class EmailCalculationsTest(unittest.TestCase):
    def setUp(self) -> None:
        self.index = pd.date_range("2025-01-01", periods=400, freq="D")
        self.series = pd.Series(np.linspace(100.0, 150.0, len(self.index)), index=self.index)

    def test_calcular_janelas_includes_all_expected_periods(self) -> None:
        result = email.calcular_janelas(self.series)

        self.assertEqual(list(result), email.WINDOW_COLUMNS)
        self.assertAlmostEqual(result["Desde Início"], 0.5)
        self.assertGreater(result["1D"], 0)

    def test_calcular_risco_returns_finite_metrics(self) -> None:
        ibov = pd.Series(np.linspace(100.0, 140.0, len(self.index)), index=self.index)
        cdi = pd.Series(np.linspace(100.0, 110.0, len(self.index)), index=self.index)

        result = email.calcular_risco(self.series, ibov, cdi)

        self.assertEqual(list(result), email.RISK_COLUMNS)
        self.assertTrue(np.isfinite(result["VOLATILIDADE"]))
        self.assertTrue(np.isfinite(result["BETA"]))
        self.assertLessEqual(result["DRAWDOWN"], 0)


class PortfolioCalculationsTest(unittest.TestCase):
    def test_daily_return_without_drift_rebalances_to_target(self) -> None:
        returns = pd.DataFrame(
            {
                "cod_ativo": ["AAA", "BBB", "AAA", "BBB"],
                "data": pd.to_datetime(["2026-01-02", "2026-01-02", "2026-01-05", "2026-01-05"]),
                "ret": [0.10, 0.00, 0.00, 0.10],
            }
        )
        composition = pd.DataFrame({"cod_ativo": ["AAA", "BBB"], "peso": [0.5, 0.5]})

        result = portfolio._ret_diario_sem_drift(returns, composition)

        np.testing.assert_allclose(result.to_numpy(), [0.05, 0.05])

    def test_daily_return_with_drift_uses_evolved_weights(self) -> None:
        returns = pd.DataFrame(
            {
                "cod_ativo": ["AAA", "BBB", "AAA", "BBB"],
                "data": pd.to_datetime(["2026-01-02", "2026-01-02", "2026-01-05", "2026-01-05"]),
                "ret": [0.10, 0.00, 0.00, 0.10],
            }
        )
        composition = pd.DataFrame({"cod_ativo": ["AAA", "BBB"], "peso": [0.5, 0.5]})

        result = portfolio._ret_diario_com_drift(returns, composition)

        self.assertAlmostEqual(result.iloc[0], 0.05)
        self.assertAlmostEqual(result.iloc[1], 0.1 * (0.5 / 1.05))

class SettingsTest(unittest.TestCase):
    def test_environment_overrides_output_directory(self) -> None:
        with patch.dict("os.environ", {"XP_OUTPUT_DIR": r"C:\temp\xp-output"}):
            settings = Settings.from_env()

        self.assertEqual(str(settings.output_dir), r"C:\temp\xp-output")


class ArtifactNamesContractTest(unittest.TestCase):
    def test_external_filenames_remain_compatible(self) -> None:
        self.assertEqual(
            set(arq_base100.values()),
            {
                "top_acoes_base_100",
                "top_dividendos_base_100",
                "top_small_caps_base_100",
                "esg_base_100",
            },
        )
        self.assertEqual(
            set(arq_performance.values()),
            {
                "tab_performance_top_acoes.xlsx",
                "tab_performance_top_dividendos.xlsx",
                "tab_performance_top_small_caps.xlsx",
                "tab_performance_esg.xlsx",
            },
        )
        self.assertEqual(EMAIL_FILENAME, "email_carteiras.msg")


if __name__ == "__main__":
    unittest.main()
