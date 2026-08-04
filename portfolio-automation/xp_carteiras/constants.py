"""Constantes e mapas estáticos do pipeline."""

from __future__ import annotations

import pandas as pd

portfolio_names = [
    'Carteira - TOP Ações XP',
    'Carteira - TOP DIVIDENDOS XP',
    'Carteira - TOP SMALL CAPS XP',
    'Carteira - ESG XP',
]

DRIFT_INICIO = pd.Timestamp(2026, 5, 29)

benchmarks_por_carteira = {
    'Carteira - TOP Ações XP':        ['Ibovespa'],
    'Carteira - TOP DIVIDENDOS XP':   ['Ibovespa'],
    'Carteira - TOP SMALL CAPS XP':   ['Ibovespa', 'SMLL'],   
    'Carteira - ESG XP':              ['Ibovespa', 'ISEE'],
}

arq_base100 = {
    'Carteira - TOP Ações XP':      'top_acoes_base_100',
    'Carteira - TOP DIVIDENDOS XP': 'top_dividendos_base_100',
    'Carteira - TOP SMALL CAPS XP': 'top_small_caps_base_100',
    'Carteira - ESG XP':            'esg_base_100',
}

arq_fig2 = {
    'Carteira - TOP Ações XP':      'portfolio_metrics_top_acoes',
    'Carteira - TOP DIVIDENDOS XP': 'portfolio_metrics_top_dividendos',
    'Carteira - TOP SMALL CAPS XP': 'portfolio_metrics_top_small_caps',
    'Carteira - ESG XP':            'portfolio_metrics_esg',
}

arq_performance = {
    'Carteira - TOP Ações XP':      'tab_performance_top_acoes.xlsx',
    'Carteira - TOP DIVIDENDOS XP': 'tab_performance_top_dividendos.xlsx',
    'Carteira - TOP SMALL CAPS XP': 'tab_performance_top_small_caps.xlsx',
    'Carteira - ESG XP':            'tab_performance_esg.xlsx',
}

BENCH_LAMINA = {
    'Carteira - TOP SMALL CAPS XP': ['SMLL'],   # remove Ibovespa da lâmina/PPT
}

mapa_arquivo = {
    'Carteira - TOP Ações XP':      'top_acoes',
    'Carteira - TOP DIVIDENDOS XP': 'top_dividendos',
    'Carteira - TOP SMALL CAPS XP': 'top_small_caps',
    'Carteira - ESG XP':            'esg',
}

cols_pct = ['Peso', 'Desempenho desde entrada', 'Desempenho no mês', 'Desempenho YTD']

sector_dict = {
    'Commodities':          ['Energy', 'Materials'],
    'Domestic Defensives':  ['Consumer Staples', 'Health Care', 'Utilities'],
    'Domestic Cyclicals':   ['Consumer Discretionary', 'Industrials',
                             'Information Technology', 'Real Estate', 'Communication Services'],
    'Financials':           ['Financials'],
}

sector_to_segment = {s: grupo for grupo, setores in sector_dict.items() for s in setores}

SETOR_PT = {
    'Energy': 'Energia',
    'Materials': 'Materiais',
    'Industrials': 'Bens Industriais',
    'Consumer Discretionary': 'Consumo Discricionário',
    'Consumer Staples': 'Consumo Básico',
    'Health Care': 'Saúde',
    'Financials': 'Financeiro',
    'Information Technology': 'Tecnologia da Informação',
    'Communication Services': 'Comunicações',
    'Utilities': 'Utilidade Pública',
    'Real Estate': 'Imobiliário',
}

SEGMENTO_PT = {
    'Commodities': 'Commodities',
    'Domestic Defensives': 'Defensivas Domésticas',
    'Domestic Cyclicals': 'Cíclicas Domésticas',
    'Financials': 'Financeiro',
}

ABREV_EN = {
    'Consumer Discretionary': 'Cons. Disc.',
    'Information Technology': 'Info Tech.',
    'Consumer Staples': 'Cons. Staples',
}

ABREV_PT = {
    'Consumo Discricionário': 'Cons. Disc.',
    'Tecnologia da Informação': 'Tec. Info.',
    'Consumo Básico': 'Cons. Básico',
}

nomes_exib = {
    'Carteira - TOP Ações XP':      'Top Ideas',
    'Carteira - TOP DIVIDENDOS XP': 'Top Dividends',
    'Carteira - TOP SMALL CAPS XP': 'Top Small Caps',
    'Carteira - ESG XP':            'Top ESG',
}

arq_comp = {
    'Carteira - TOP Ações XP':      'composicao_top_acoes',
    'Carteira - TOP DIVIDENDOS XP': 'composicao_top_dividendos',
    'Carteira - TOP SMALL CAPS XP': 'composicao_top_small_caps',
    'Carteira - ESG XP':            'composicao_esg',
}


rotulo_carteira_pt = {
    'Carteira - TOP Ações XP':      'Top Ações',
    'Carteira - TOP DIVIDENDOS XP': 'Top Dividendos',
    'Carteira - TOP SMALL CAPS XP': 'Top Small Caps',
    'Carteira - ESG XP':            'ESG',
}

arq_decomp = {
    'Carteira - TOP Ações XP':      'decomposicao_top_acoes',
    'Carteira - TOP DIVIDENDOS XP': 'decomposicao_top_dividendos',
    'Carteira - TOP SMALL CAPS XP': 'decomposicao_top_small_caps',
    'Carteira - ESG XP':            'decomposicao_esg',
}

MES_LBL_PT = {1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun',
              7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'}

cols_pct_decomp = ['Peso', 'Retorno no mês', 'Contribuição']


