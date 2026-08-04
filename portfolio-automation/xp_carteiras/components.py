"""Components do pipeline de carteiras XP."""

from __future__ import annotations

import numpy as np
import pandas as pd

from .constants import ABREV_EN, ABREV_PT, SEGMENTO_PT, SETOR_PT, sector_to_segment

def _primeiro_dia_util_do_mes(data, datas_pregao):
    """
    Mês de entrada para EXIBIÇÃO, considerando TRADING DAYS reais (com feriados):
      - se data_rebal == último TRADING DAY do seu mês  -> entra no mês seguinte
      - se for no meio do mês                           -> entra no próprio mês
    Retorna o 1º dia (1º do mês de entrada).

    `datas_pregao`: array/Series de datas reais de pregão (ex: md['data'].unique()).
    """
    d = pd.to_datetime(data)
    dp = pd.DatetimeIndex(sorted(pd.to_datetime(datas_pregao)))

    # último TRADING DAY real do mês de d (maior pregão <= fim do mês)
    fim_mes = pd.Timestamp(d.year, d.month, 1) + pd.offsets.MonthEnd(0)
    pregoes_do_mes = dp[(dp >= pd.Timestamp(d.year, d.month, 1)) & (dp <= fim_mes)]
    if len(pregoes_do_mes) == 0:
        ultimo_trading = fim_mes   # fallback
    else:
        ultimo_trading = pregoes_do_mes[-1]

    # se rebal == último trading day -> mês seguinte; senão -> mesmo mês
    if d.normalize() == ultimo_trading.normalize():
        return pd.Timestamp(d.year, d.month, 1) + pd.offsets.MonthBegin(1)
    else:
        return pd.Timestamp(d.year, d.month, 1)


def _preco_em(cod_ativo, data_alvo, market_data):
    """Último preço ajustado disponível <= data_alvo para o ativo."""
    s = market_data[
        (market_data['cod_ativo'] == cod_ativo) & (market_data['data'] <= data_alvo)
    ]
    if s.empty:
        return np.nan
    return s.sort_values('data')['adj_close_price'].iloc[-1]


def _rotulo_mes_entrada(data_entrada):
    """
    Rótulo da data de entrada (só exibição, não altera cálculo).
    Se a entrada cair no ÚLTIMO DIA ÚTIL do mês, exibe o mês SUBSEQUENTE.
    """
    data = pd.to_datetime(data_entrada)

    ultimo_dia_util = data + pd.offsets.MonthEnd(0)
    if ultimo_dia_util.weekday() >= 5:
        ultimo_dia_util -= pd.offsets.BDay(1)

    if data.normalize() == ultimo_dia_util.normalize():
        return (data + pd.offsets.MonthBegin(1)).strftime('%b-%y')
    return data.strftime('%b-%y')


def _data_entrada_continua(comp_long, cod_ativo, rebal_dates, rebal_ref):
    """
    Data de entrada na carteira considerando o período contínuo que inclui rebal_ref.
    """
    presentes = set(comp_long.loc[comp_long['cod_ativo'] == cod_ativo, 'data_rebal'])
    idx_ref = rebal_dates.index(rebal_ref)

    entrada = rebal_ref
    for i in range(idx_ref, -1, -1):
        if rebal_dates[i] in presentes:
            entrada = rebal_dates[i]
        else:
            break
    return entrada


def tabela_componentes(
    composition, rebal_ref, ini_mes, fim_mes, ini_ytd, fim_ref,
    market_data, name_map, sector_map,
):
    """Monta a tabela por componente para o rebalanceamento `rebal_ref`."""
    comp = composition.copy()
    comp_long = comp.melt(id_vars='cod_ativo', var_name='data_rebal', value_name='peso')
    comp_long['data_rebal'] = pd.to_datetime(comp_long['data_rebal'])
    comp_long = comp_long.dropna(subset=['peso'])

    rebal_dates = sorted(comp_long['data_rebal'].unique())

    comp_ref = comp_long[comp_long['data_rebal'] == rebal_ref][['cod_ativo', 'peso']].copy()

    linhas = []
    for _, r in comp_ref.iterrows():
        cod = r['cod_ativo']
        peso = r['peso'] / comp_ref['peso'].sum()

        entrada = _data_entrada_continua(comp_long, cod, rebal_dates, rebal_ref)

        p_entrada = _preco_em(cod, entrada, market_data)
        p_fim_ref = _preco_em(cod, fim_ref, market_data)
        desemp_entrada = (p_fim_ref / p_entrada - 1) if (p_entrada and p_fim_ref) else np.nan

        p_ini_mes = _preco_em(cod, ini_mes, market_data)
        p_fim_mes = _preco_em(cod, fim_mes, market_data)
        desemp_mes = (p_fim_mes / p_ini_mes - 1) if (p_ini_mes and p_fim_mes) else np.nan

        p_ini_ytd = _preco_em(cod, ini_ytd, market_data)
        p_fim_ytd = _preco_em(cod, fim_ref, market_data)
        desemp_ytd = (p_fim_ytd / p_ini_ytd - 1) if (p_ini_ytd and p_fim_ytd) else np.nan
        
        linhas.append({
            'Companhia':                 name_map.get(cod, ''),
            'Ticker':                    cod,
            'Setor':                     sector_map.get(cod, ''),
            'Peso':                      peso,
            'Data de entrada':           _primeiro_dia_util_do_mes(entrada, market_data['data'].unique()),  # data real (1º dia útil), exibida como mmm/yy
            'Desempenho desde entrada':  desemp_entrada,   # cálculo usa 'entrada' (último trading day)
            'Desempenho no mês':         desemp_mes,
            'Desempenho YTD':            desemp_ytd,
        })

    tabela = pd.DataFrame(linhas).sort_values('Ticker', ascending=True).reset_index(drop=True)
    return tabela


def _formata_componentes(tabela):
    """Formata pesos e desempenhos em %."""
    t = tabela.copy()
    t['Peso'] = (t['Peso'] * 100).round(1).astype(str) + '%'
    for c in ['Desempenho desde entrada', 'Desempenho no mês', 'Desempenho YTD']:
        t[c] = (tabela[c] * 100).round(1).astype(str) + '%'
        t[c] = t[c].replace('nan%', '')
    return t


def gerar_tabelas_componentes(composition, market_data, name_map, sector_map):
    """Gera as tabelas atual, último rebal e composição anterior com MTD atual."""
    comp = composition.copy()
    comp_long = comp.melt(id_vars='cod_ativo', var_name='data_rebal', value_name='peso')
    comp_long['data_rebal'] = pd.to_datetime(comp_long['data_rebal'])
    comp_long = comp_long.dropna(subset=['peso'])
    rebal_dates = sorted(comp_long['data_rebal'].unique())

    rebal_atual = rebal_dates[-1]
    rebal_anterior = rebal_dates[-2] if len(rebal_dates) >= 2 else rebal_dates[-1]

    data_max = market_data['data'].max()

    ini_mes_atual = pd.Timestamp(data_max.year, data_max.month, 1) - pd.Timedelta(days=1)
    fim_mes_atual = data_max
    ini_ytd_atual = pd.Timestamp(data_max.year, 1, 1) - pd.Timedelta(days=1)

    tab_atual = tabela_componentes(
        comp, rebal_atual,
        ini_mes=ini_mes_atual, fim_mes=fim_mes_atual,
        ini_ytd=ini_ytd_atual, fim_ref=data_max,
        market_data=market_data, name_map=name_map, sector_map=sector_map,
    )

    mes_ant = pd.Timestamp(data_max.year, data_max.month, 1) - pd.Timedelta(days=1)
    ini_mes_ant = pd.Timestamp(mes_ant.year, mes_ant.month, 1) - pd.Timedelta(days=1)
    fim_mes_ant = mes_ant
    ini_ytd_ant = pd.Timestamp(mes_ant.year, 1, 1) - pd.Timedelta(days=1)

    tab_ultimo = tabela_componentes(
        comp, rebal_anterior,
        ini_mes=ini_mes_ant, fim_mes=fim_mes_ant,
        ini_ytd=ini_ytd_ant, fim_ref=fim_mes_ant,
        market_data=market_data, name_map=name_map, sector_map=sector_map,
    )

    tab_comp_anterior_mtd_atual = tabela_componentes(
        comp, rebal_anterior,
        ini_mes=ini_mes_atual, fim_mes=fim_mes_atual,
        ini_ytd=ini_ytd_atual, fim_ref=data_max,
        market_data=market_data, name_map=name_map, sector_map=sector_map,
    )

    return tab_atual, tab_ultimo, tab_comp_anterior_mtd_atual


def _download_ibov_composition(index_composition_path):
    fp = index_composition_path
    df = (pd.read_csv(fp, usecols=['cod_ativo', 'data', 'IBOV'], sep=';')
            .dropna()
            .rename(columns={'IBOV': 'weight'}))
    df['weight'] = df['weight'].str.replace(',', '.').astype(float)
    df.sort_values(['data', 'cod_ativo'], inplace=True)
    return df


def calcular_pesos_ibovespa(index_composition_path, sector_classification):
    """
    Peso setorial (GICS individual) + por grupo do Ibovespa na data mais recente.
    Retorna (individual, grupos), ambos com pesos em fração (0-1).
    """
    members = _download_ibov_composition(index_composition_path)

    setores = (sector_classification[['cod_ativo', 'adjusted_GICS_sector']]
               .drop_duplicates('cod_ativo')
               .set_index('cod_ativo'))

    df = members.merge(setores, on='cod_ativo', how='left')
    df['sector_group'] = df['adjusted_GICS_sector'].map(sector_to_segment)

    ultima = df['data'].max()
    df_u = df[df['data'] == ultima]

    individual = (df_u.groupby('adjusted_GICS_sector')['weight'].sum() / 100).reset_index()
    individual = individual.rename(columns={
        'adjusted_GICS_sector': 'Sector', 'weight': 'Sector Weight (Ibovespa)'
    })

    grupos = (df_u.groupby('sector_group')['weight'].sum() / 100).reset_index()
    grupos = grupos.rename(columns={
        'sector_group': 'Segment', 'weight': 'Segment Weight (Ibovespa)'
    })

    return individual, grupos


def _composicao_vigente(composition):
    """Ticker + Weight do último rebalanceamento (carteira vigente)."""
    comp_long = composition.melt(id_vars='cod_ativo', var_name='data_rebal', value_name='peso')
    comp_long['data_rebal'] = pd.to_datetime(comp_long['data_rebal'])
    comp_long = comp_long.dropna(subset=['peso'])
    ult = comp_long['data_rebal'].max()
    comp = comp_long[comp_long['data_rebal'] == ult][['cod_ativo', 'peso']].copy()
    comp = comp.rename(columns={'cod_ativo': 'Ticker', 'peso': 'Weight'})
    comp['Weight'] = comp['Weight'] / comp['Weight'].sum()
    return comp


def montar_df_composicao(
    composition, comp_sheet_data, sector_classification,
    ibov_sector_weights, segment_sector_map, idioma='EN',
):
    df = _composicao_vigente(composition)
    df = df.merge(comp_sheet_data, on='Ticker', how='left')
    df = df.merge(sector_classification, on='Ticker', how='left')   # setor GICS puro
    df = df.merge(ibov_sector_weights, on='Sector', how='left')      # casa em GICS puro
    df = df.merge(segment_sector_map, on='Sector', how='left')       # casa em GICS puro
    df['Sector Weight (Portfolio)'] = df.groupby('Sector')['Weight'].transform('sum')

    df = df[['Segment', 'Sector', 'Sector Weight (Ibovespa)', 'Sector Weight (Portfolio)',
             'Company', 'Ticker', 'Weight', 'Rating', 'Target Price']]
    df = df.sort_values(['Segment', 'Sector'], ignore_index=True)

    if idioma == 'EN':
        df['Sector'] = df['Sector'].replace(ABREV_EN)   # só abrevia
    else:
        df['Sector']  = df['Sector'].replace(SETOR_PT).replace(ABREV_PT)  # traduz e abrevia
        df['Segment'] = df['Segment'].replace(SEGMENTO_PT)                # traduz segmento
        df['Rating']  = df['Rating'].replace({'Buy': 'Compra', 'Neutral': 'Neutro', 'Sell': 'Venda'})
        df = df.rename(columns={
            'Segment': 'Segmento',
            'Sector': 'Setor',
            'Sector Weight (Ibovespa)': 'Peso do setor (Ibovespa)',
            'Sector Weight (Portfolio)': 'Peso do setor (Carteira)',
            'Company': 'Companhia',
            'Weight': 'Peso',
            'Target Price': 'Preço-Alvo',
        })
    return df


def _retorno_papel_mes(cod_ativo, ano, mes, market_data):
    """Retorno do papel no mês cheio (fim do mês anterior -> fim do mês),
    usando o último preço ajustado disponível <= cada data de corte."""
    ini = pd.Timestamp(ano, mes, 1) - pd.Timedelta(days=1)
    fim = pd.Timestamp(ano, mes, 1) + pd.offsets.MonthEnd(0)
    p_ini = _preco_em(cod_ativo, ini, market_data)
    p_fim = _preco_em(cod_ativo, fim, market_data)
    if not p_ini or not p_fim or np.isnan(p_ini) or np.isnan(p_fim):
        return np.nan
    return p_fim / p_ini - 1


def _rebal_vigente_no_mes(comp_long, rebal_dates, ano, mes):
    """Rebalanceamento vigente durante o mês (último rebal <= fim do mês)."""
    fim_mes = pd.Timestamp(ano, mes, 1) + pd.offsets.MonthEnd(0)
    vigentes = [d for d in rebal_dates if pd.Timestamp(d) <= fim_mes]
    return vigentes[-1] if vigentes else rebal_dates[0]


def decomposicao_retorno(
    composition, ano, mes, market_data, name_map, sector_map,
):
    """
    Return attribution do mês (ano, mes):
      - peso: peso-alvo normalizado do rebal vigente no mês
      - retorno: retorno mensal do papel
      - contribuição: peso * retorno
    A linha 'Carteira (total)' traz a soma das contribuições (retorno do mês).
    """
    comp_long = composition.melt(id_vars='cod_ativo', var_name='data_rebal', value_name='peso')
    comp_long['data_rebal'] = pd.to_datetime(comp_long['data_rebal'])
    comp_long = comp_long.dropna(subset=['peso'])
    rebal_dates = sorted(comp_long['data_rebal'].unique())

    rebal_ref = _rebal_vigente_no_mes(comp_long, rebal_dates, ano, mes)
    comp_ref = comp_long[comp_long['data_rebal'] == rebal_ref][['cod_ativo', 'peso']].copy()
    soma_peso = comp_ref['peso'].sum()

    linhas = []
    for _, r in comp_ref.iterrows():
        cod = r['cod_ativo']
        peso = r['peso'] / soma_peso if soma_peso else np.nan
        ret = _retorno_papel_mes(cod, ano, mes, market_data)
        contrib = peso * ret if (peso is not None and not np.isnan(peso) and not np.isnan(ret)) else np.nan
        linhas.append({
            'Companhia':        name_map.get(cod, ''),
            'Ticker':           cod,
            'Setor':            sector_map.get(cod, ''),
            'Peso':             peso,
            'Retorno no mês':   ret,
            'Contribuição':     contrib,
        })

    tab = pd.DataFrame(linhas).sort_values('Contribuição', ascending=False).reset_index(drop=True)

    # linha de total (retorno da carteira no mês = soma das contribuições)
    total = pd.DataFrame([{
        'Companhia': 'Carteira (total)',
        'Ticker': '',
        'Setor': '',
        'Peso': tab['Peso'].sum(skipna=True),
        'Retorno no mês': np.nan,
        'Contribuição': tab['Contribuição'].sum(skipna=True),
    }])
    return pd.concat([tab, total], ignore_index=True)


def gerar_decomposicoes(composition, market_data, name_map, sector_map):
    """Return attribution do mês atual (aberto) e do mês anterior (fechado)."""
    data_max = market_data['data'].max()

    ano_atual, mes_atual = data_max.year, data_max.month

    ref_ant = pd.Timestamp(ano_atual, mes_atual, 1) - pd.Timedelta(days=1)
    ano_ant, mes_ant = ref_ant.year, ref_ant.month

    dec_atual = decomposicao_retorno(
        composition, ano_atual, mes_atual, market_data, name_map, sector_map
    )
    dec_ant = decomposicao_retorno(
        composition, ano_ant, mes_ant, market_data, name_map, sector_map
    )
    return dec_atual, dec_ant


