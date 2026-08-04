"""Gera somente as Lâminas Comerciais e Prestações de Contas."""

from __future__ import annotations

import os

from .accountability_reports import (
    accountability_output_filename,
    atualizar_prestacao_contas,
    performance_summary,
    reconcile_decomposition_total,
    resolve_accountability_period,
)
from .components import decomposicao_retorno
from .monthly_config import accountability_ppt_config, commercial_ppt_config
from .performance import _df_para_lamina
from .pipeline_data import PipelineContext
from .powerpoint_reports import atualizar_ppt


def _generate_commercial_decks(context: PipelineContext) -> None:
    for portfolio, config in commercial_ppt_config(context.settings).items():
        template = config.get("template", "")
        if not template or not os.path.exists(template):
            print(f"[PULADO] template não encontrado para {portfolio}: {template}")
            continue
        atualizar_ppt(
            caminho_template=template,
            caminho_saida=config["saida"],
            df_port=_df_para_lamina(portfolio, context.resultado_dfs),
            composition=context.composition_dict[portfolio],
            serie_cdi=context.cdi,
            portfolio=portfolio,
        )


def _generate_accountability_decks(context: PipelineContext) -> None:
    for portfolio, config in accountability_ppt_config(context.settings).items():
        template = config.get("template", "")
        if not template or not os.path.exists(template):
            print(
                "[PULADO] template da Prestação de Contas não encontrado "
                f"para {portfolio}: {template}"
            )
            continue

        df_port = _df_para_lamina(portfolio, context.resultado_dfs)
        period = resolve_accountability_period(df_port)
        decomposition = decomposicao_retorno(
            context.composition_dict[portfolio],
            period.reference_year,
            period.reference_month,
            context.market_data,
            context.mapa_nome,
            context.mapa_setor,
        )
        expected_return = performance_summary(
            df_port, period.reference_year, period.reference_month
        ).loc[portfolio, "Mês"]
        decomposition = reconcile_decomposition_total(decomposition, expected_return)
        composition = context.dfs_composicao[portfolio]["PT"][
            ["Setor", "Companhia", "Ticker", "Peso"]
        ]
        output_name = accountability_output_filename(config["output_label"], period)
        output_path = context.settings.accountability_deck_dir / output_name
        atualizar_prestacao_contas(
            caminho_template=template,
            caminho_saida=output_path,
            df_port=df_port,
            df_composicao=composition,
            df_decomposicao=decomposition,
            portfolio_label=config["display_label"],
            reference_year=period.reference_year,
            reference_month=period.reference_month,
        )
        print(f"Prestação de Contas atualizada: {output_path}")


def generate_powerpoints(context: PipelineContext) -> None:
    """Gera os dois tipos de PPT sem regravar os Excel do output."""
    context.settings.commercial_deck_dir.mkdir(parents=True, exist_ok=True)
    context.settings.accountability_deck_dir.mkdir(parents=True, exist_ok=True)
    _generate_commercial_decks(context)
    _generate_accountability_decks(context)
