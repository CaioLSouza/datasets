"""Configuração central dos scripts de carteiras XP.

Os caminhos corporativos continuam sendo os padrões. Para executar em outro
ambiente, cada caminho pode ser sobrescrito por variável de ambiente, sem
alterar o código-fonte.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


DEFAULT_PORTFOLIO_ROOT = Path(
    r"\\xpdocs\Research\Equities\Estrategia\Carteiras\Carteiras de Ações XP"
)
DEFAULT_CROSS_DATA_DIR = Path(
    r"\\xpdocs\Research\Equities\Quant\_Cross Data"
)
DEFAULT_COMP_SHEET = Path(
    r"\\xpdocs\Research\Equities\COMP SHEET\raw_data.xlsx"
)


def _env_path(name: str, default: Path) -> Path:
    """Lê um caminho do ambiente e mantém ``default`` quando não informado."""
    value = os.getenv(name)
    return Path(value) if value else default


@dataclass(frozen=True)
class Settings:
    """Caminhos usados pelo pipeline.

    As variáveis de ambiente aceitas estão documentadas no README. A classe é
    imutável para impedir mudanças acidentais durante uma execução longa.
    """

    portfolio_root: Path
    cross_data_dir: Path
    comp_sheet_path: Path
    output_dir: Path
    email_dir: Path
    templates_dir: Path
    commercial_deck_dir: Path
    performance_workbook_path: Path
    sector_classification_path: Path
    market_data_path: Path
    bdr_market_data_path: Path
    indices_path: Path
    index_composition_path: Path

    @classmethod
    def from_env(cls) -> "Settings":
        portfolio_root = _env_path("XP_PORTFOLIO_ROOT", DEFAULT_PORTFOLIO_ROOT)
        cross_data_dir = _env_path("XP_CROSS_DATA_DIR", DEFAULT_CROSS_DATA_DIR)
        return cls(
            portfolio_root=portfolio_root,
            cross_data_dir=cross_data_dir,
            comp_sheet_path=_env_path("XP_COMP_SHEET_PATH", DEFAULT_COMP_SHEET),
            output_dir=_env_path("XP_OUTPUT_DIR", portfolio_root / "output"),
            email_dir=_env_path("XP_EMAIL_DIR", portfolio_root),
            templates_dir=_env_path("XP_TEMPLATES_DIR", portfolio_root / "Templates"),
            commercial_deck_dir=_env_path(
                "XP_COMMERCIAL_DECK_DIR", portfolio_root / "Lâmina Comercial"
            ),
            performance_workbook_path=_env_path(
                "XP_PERFORMANCE_WORKBOOK", portfolio_root.parent / "Performance carteiras.xlsm"
            ),
            sector_classification_path=_env_path(
                "XP_SECTOR_CLASSIFICATION_PATH",
                cross_data_dir / "xpqs-sector_classification.xlsx",
            ),
            market_data_path=_env_path(
                "XP_MARKET_DATA_PATH", cross_data_dir / "economatica-market_data.parquet"
            ),
            bdr_market_data_path=_env_path(
                "XP_BDR_MARKET_DATA_PATH", cross_data_dir / "raw" / "bdr_market_data.csv"
            ),
            indices_path=_env_path(
                "XP_INDICES_PATH", cross_data_dir / "economatica-indices.parquet"
            ),
            index_composition_path=_env_path(
                "XP_INDEX_COMPOSITION_PATH",
                cross_data_dir / "economatica-index_composition.csv",
            ),
        )


def load_settings() -> Settings:
    """Ponto único de criação da configuração usada pelos dois scripts."""
    return Settings.from_env()
