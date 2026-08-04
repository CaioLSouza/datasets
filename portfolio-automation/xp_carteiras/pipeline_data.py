"""Carga e preparação compartilhadas pelos fluxos de saída e PowerPoint."""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from .components import calcular_pesos_ibovespa, montar_df_composicao
from .constants import benchmarks_por_carteira, portfolio_names, sector_to_segment
from .performance import calcular_performance, indice_base100
from .settings import Settings


@dataclass
class PipelineContext:
    """Dados preparados uma vez e reutilizados por todas as etapas do pipeline."""

    settings: Settings
    xpqs: pd.DataFrame
    market_data: pd.DataFrame
    composition_dict: dict[str, pd.DataFrame]
    resultado_dfs: dict[str, pd.DataFrame]
    cdi: pd.Series
    mapa_nome: dict
    mapa_setor: dict
    sector_weight_ibovespa: pd.DataFrame
    segment_weight_ibovespa: pd.DataFrame
    segment_sector: pd.DataFrame
    dfs_composicao: dict[str, dict[str, pd.DataFrame]]


def _load_compositions(settings: Settings) -> dict[str, pd.DataFrame]:
    compositions = {}
    for portfolio in portfolio_names:
        frame = pd.read_excel(
            settings.performance_workbook_path,
            sheet_name=portfolio,
            skiprows=5,
        )
        if portfolio == "Carteira - TOP Ações XP":
            frame = frame.iloc[:, 89:]
            frame.rename(columns={"Ticker.1": "cod_ativo"}, inplace=True)
        else:
            frame = frame.iloc[:, 3:]
            frame.rename(columns={"Ticker": "cod_ativo"}, inplace=True)
        compositions[portfolio] = frame
    return compositions


def _load_market_data(settings: Settings) -> pd.DataFrame:
    market_data = pd.read_parquet(settings.market_data_path)
    bdr_market_data = pd.read_csv(settings.bdr_market_data_path)
    bdr_market_data.rename(columns={"Ativo": "cod_ativo", "Data": "data"}, inplace=True)
    bdr_market_data["cod_ativo"] = bdr_market_data["cod_ativo"].str.replace(
        "<XBSP>", "", regex=False
    )
    bdr_market_data["data"] = pd.to_datetime(bdr_market_data["data"])
    valid_dates = market_data["data"].unique()
    bdr_market_data = bdr_market_data[
        bdr_market_data["data"].isin(valid_dates)
    ].copy()
    bdr_market_data["adj_close_price"] = pd.to_numeric(
        bdr_market_data["adj_close_price"], errors="coerce"
    )

    combined = pd.concat([market_data, bdr_market_data], ignore_index=True)
    combined["data"] = pd.to_datetime(combined["data"])
    combined = combined[["cod_ativo", "data", "adj_close_price"]].dropna()
    combined = combined.sort_values(["cod_ativo", "data"])
    combined["ret"] = combined.groupby("cod_ativo")["adj_close_price"].pct_change()
    return combined


def _build_composition_tables(
    settings: Settings,
    xpqs: pd.DataFrame,
    composition_dict: dict[str, pd.DataFrame],
) -> tuple[
    pd.DataFrame,
    pd.DataFrame,
    pd.DataFrame,
    dict[str, dict[str, pd.DataFrame]],
]:
    comp_sheet = pd.read_excel(settings.comp_sheet_path, sheet_name="Sheet1")[[
        "TICKER", "NAME", "TARGET", "RECOMMENDATION"
    ]].rename(columns={
        "TICKER": "Ticker",
        "NAME": "Company",
        "TARGET": "Target Price",
        "RECOMMENDATION": "Rating",
    })
    sector_classification = xpqs[["cod_ativo", "adjusted_GICS_sector"]].rename(
        columns={"cod_ativo": "Ticker", "adjusted_GICS_sector": "Sector"}
    ).drop_duplicates("Ticker")
    sector_weight_ibovespa, segment_weight_ibovespa = calcular_pesos_ibovespa(
        settings.index_composition_path, xpqs
    )
    segment_sector = pd.DataFrame(
        [{"Sector": sector, "Segment": segment} for sector, segment in sector_to_segment.items()]
    )

    dfs_composicao = {}
    for portfolio in portfolio_names:
        args = (
            composition_dict[portfolio],
            comp_sheet,
            sector_classification,
            sector_weight_ibovespa,
            segment_sector,
        )
        dfs_composicao[portfolio] = {
            "EN": montar_df_composicao(*args, idioma="EN"),
            "PT": montar_df_composicao(*args, idioma="PT"),
        }
    return (
        sector_weight_ibovespa,
        segment_weight_ibovespa,
        segment_sector,
        dfs_composicao,
    )


def prepare_pipeline_context(settings: Settings) -> PipelineContext:
    """Lê as fontes corporativas e calcula as séries usadas nas saídas."""
    xpqs = pd.read_excel(settings.sector_classification_path)[[
        "cod_ativo", "name", "adjusted_GICS_sector", "sector_xp"
    ]]
    market_data = _load_market_data(settings)
    indices = pd.read_parquet(settings.indices_path)
    indices["data"] = pd.to_datetime(indices["data"])
    composition_dict = _load_compositions(settings)

    benchmark_data = indices.loc[
        indices["cod_ativo"].isin(["Ibovespa", "SMLL", "ISEE"])
    ].copy()
    resultado_dfs = {}
    for portfolio in portfolio_names:
        curva_carteira = calcular_performance(composition_dict[portfolio], market_data)
        df_port = pd.DataFrame({portfolio: curva_carteira})
        for benchmark in benchmarks_por_carteira[portfolio]:
            df_port[benchmark] = indice_base100(
                benchmark, curva_carteira.index, benchmark_data
            )
        df_port.index.name = "data"
        resultado_dfs[portfolio] = df_port

    cdi = (
        indices.loc[
            indices["cod_ativo"] == "CDI Acumulado", ["data", "close_price"]
        ]
        .drop_duplicates("data")
        .set_index("data")["close_price"]
        .sort_index()
    )
    mapa_info = xpqs.drop_duplicates("cod_ativo").set_index("cod_ativo")
    mapa_nome = mapa_info["name"].to_dict()
    mapa_setor = mapa_info["sector_xp"].to_dict()
    (
        sector_weight_ibovespa,
        segment_weight_ibovespa,
        segment_sector,
        dfs_composicao,
    ) = _build_composition_tables(settings, xpqs, composition_dict)

    return PipelineContext(
        settings=settings,
        xpqs=xpqs,
        market_data=market_data,
        composition_dict=composition_dict,
        resultado_dfs=resultado_dfs,
        cdi=cdi,
        mapa_nome=mapa_nome,
        mapa_setor=mapa_setor,
        sector_weight_ibovespa=sector_weight_ibovespa,
        segment_weight_ibovespa=segment_weight_ibovespa,
        segment_sector=segment_sector,
        dfs_composicao=dfs_composicao,
    )
