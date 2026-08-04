"""Performance do pipeline de carteiras XP."""

from __future__ import annotations

from datetime import date

import numpy as np
import pandas as pd

from .constants import BENCH_LAMINA, DRIFT_INICIO


def recortar_ultimo_mes_fechado(df_port, *, today=None):
    """Remove do dataframe qualquer dado do mês corrente ainda em aberto.

    Quando a base ainda termina em um mês anterior ao corrente, preserva o
    último mês disponível. ``today`` existe para reprocessamentos e testes.
    """
    available = df_port.dropna(how='all')
    if available.empty:
        raise ValueError('df_port não possui dados de performance.')

    latest = pd.Timestamp(available.index.max())
    current = pd.Timestamp(today or date.today())
    latest_period = latest.to_period('M')
    current_period = current.to_period('M')
    reference = current_period - 1 if latest_period >= current_period else latest_period
    cutoff = reference.end_time
    closed = df_port.loc[df_port.index <= cutoff].copy()
    if closed.dropna(how='all').empty:
        raise ValueError('Não há dados até o último mês fechado.')
    return closed

def _ret_diario_sem_drift(ret_periodo, comp_periodo):
    """Pesos renormalizados ao alvo todo dia (metodologia antiga)."""
    rp = ret_periodo.merge(comp_periodo, on='cod_ativo', how='inner')

    def retorno_diario(grupo):
        w = grupo['peso'] / grupo['peso'].sum()
        return (w * grupo['ret']).sum()

    return rp.groupby('data').apply(retorno_diario)


def _ret_diario_com_drift(ret_periodo, comp_periodo):
    """
    Buy & hold dentro da janela: o peso evolui com os preços.
    Para cada dia t:
        ret_carteira_t = Σ_i ( w_i,t * ret_i,t )
    onde w_i,t é o peso DRIFTED no início do dia (soma 1), atualizado por:
        w_i,t = w_i,t-1 * (1 + ret_i,t)  -> renormaliza só para obter a fração do dia
    Ativo sem retorno no dia (NaN) é tratado como ret 0 (carrega o peso).
    """
    # pesos-alvo iniciais (normalizados), indexados por ativo
    pesos = comp_periodo.set_index('cod_ativo')['peso']
    pesos = pesos / pesos.sum()

    # matriz dias x ativos com os retornos da janela
    rp = ret_periodo[ret_periodo['cod_ativo'].isin(pesos.index)].copy()
    ret_mat = (rp.pivot_table(index='data', columns='cod_ativo', values='ret')
                 .reindex(columns=pesos.index)
                 .sort_index())

    datas = ret_mat.index
    out = pd.Series(index=datas, dtype=float)

    # capital relativo de cada ativo; começa nos pesos-alvo
    cap = pesos.copy()

    for dia in datas:
        r = ret_mat.loc[dia].fillna(0.0)      # ret do dia (NaN -> 0)
        w = cap / cap.sum()                   # peso efetivo no início do dia (soma 1)
        out.loc[dia] = float((w * r).sum())   # retorno da carteira no dia
        cap = cap * (1.0 + r)                 # deixa o capital derivar (sem renormalizar)

    return out


def calcular_performance(composition, market_data):
    comp = composition.copy()

    comp_long = comp.melt(id_vars='cod_ativo', var_name='data_rebal', value_name='peso')
    comp_long['data_rebal'] = pd.to_datetime(comp_long['data_rebal'])
    comp_long = comp_long.dropna(subset=['peso'])

    rebal_dates = sorted(comp_long['data_rebal'].unique())
    retornos_carteira = []

    for i, start in enumerate(rebal_dates):
        next_start = (rebal_dates[i + 1] if i + 1 < len(rebal_dates)
                      else market_data['data'].max())

        comp_periodo = comp_long[comp_long['data_rebal'] == start][['cod_ativo', 'peso']]

        mask = (market_data['data'] > start) & (market_data['data'] <= next_start)
        ret_periodo = market_data.loc[mask, ['cod_ativo', 'data', 'ret']]

        # escolhe a metodologia conforme a data do rebal
        if pd.Timestamp(start) >= DRIFT_INICIO:
            ret_dia = _ret_diario_com_drift(ret_periodo, comp_periodo)
        else:
            ret_dia = _ret_diario_sem_drift(ret_periodo, comp_periodo)

        retornos_carteira.append(ret_dia)

    serie_ret = pd.concat(retornos_carteira).sort_index()
    serie_ret = serie_ret[~serie_ret.index.duplicated(keep='first')]

    fatores = (1 + serie_ret.fillna(0)).cumprod()

    data_inicial = pd.Timestamp(rebal_dates[0])

    base100 = pd.concat([
        pd.Series([100.0], index=[data_inicial]),
        100 * fatores
    ]).sort_index()

    return base100


def indice_base100(cod_indice, datas_carteira, indices_data):
    """
    Retorna o índice em base 100, alinhado exatamente às datas da carteira
    (reindex + ffill) e ancorado em 100 no inception da carteira.
    """
    serie = (indices_data.loc[indices_data['cod_ativo'] == cod_indice, ['data', 'close_price']]
                    .drop_duplicates('data')
                    .set_index('data')['close_price']
                    .sort_index())

    inception = datas_carteira.min()
    fim = datas_carteira.max()

    serie = serie[(serie.index >= inception - pd.Timedelta(days=10)) &
                  (serie.index <= fim)]

    if serie.empty:
        return pd.Series(dtype=float)

    serie = serie.reindex(datas_carteira.union(serie.index)).ffill()
    serie = serie.reindex(datas_carteira)

    primeiro_valido = serie.first_valid_index()
    if primeiro_valido is None:
        return pd.Series(dtype=float)

    return 100 * serie / serie.loc[primeiro_valido]


def _ret_diarios(serie):
    return serie.dropna().pct_change().dropna()


def _ultimos_12m(serie):
    """Recorta a série aos últimos 12 meses (a partir da última data disponível)."""
    s = serie.dropna()
    if s.empty:
        return s
    corte = s.index.max() - pd.DateOffset(months=12)
    return s[s.index >= corte]


def indicadores_12m(df_port, serie_cdi):
    """Sharpe, Volatilidade e Beta (últimos 12m) p/ carteira e cada benchmark. Beta vs. Ibovespa."""
    ret_ibov = _ret_diarios(_ultimos_12m(df_port['Ibovespa'])) if 'Ibovespa' in df_port.columns else None
    ret_cdi  = _ret_diarios(_ultimos_12m(serie_cdi))

    def _vol(r):
        return r.std() * np.sqrt(252)

    def _sharpe(r):
        idx = r.index.intersection(ret_cdi.index)
        if len(idx) < 2:
            return np.nan
        v = r.loc[idx].std() * np.sqrt(252)
        return ((r.loc[idx] - ret_cdi.loc[idx]).mean() * 252) / v if v > 0 else np.nan

    def _beta(r, eh_ibov):
        if eh_ibov:
            return 1.0                      
        if ret_ibov is None:
            return np.nan
        idx = r.index.intersection(ret_ibov.index)
        rc, ri = r.loc[idx], ret_ibov.loc[idx]
        var_i = np.var(ri)
        return np.cov(rc, ri)[0, 1] / var_i if var_i > 0 else np.nan

    dados = {col: [_sharpe(r := _ret_diarios(_ultimos_12m(df_port[col]))),
                   _vol(r),
                   _beta(r, col == 'Ibovespa')]
             for col in df_port.columns}

    return pd.DataFrame(dados, index=['Sharpe', 'Volatilidade', 'Beta'])


def _df_para_lamina(portfolio, result_frames):
    """Retorna o df_port para uso na lâmina/PPT, aplicando o filtro de
    benchmarks específico da lâmina quando houver (senão, usa o padrão)."""
    df = result_frames[portfolio]
    benchs_lamina = BENCH_LAMINA.get(portfolio)
    if benchs_lamina is None:
        return df
    nome_cart = _nome_col_carteira(df)
    cols_manter = [nome_cart] + [b for b in benchs_lamina if b in df.columns]
    return df[cols_manter]


def _ret_periodo(serie, data_ini, data_fim):
    """Retorno simples entre dois pontos da série base 100 (usa ffill)."""
    s = serie.dropna()
    if s.empty:
        return np.nan

    s_fim = s[s.index <= data_fim]
    if s_fim.empty:
        return np.nan
    v_fim = s_fim.iloc[-1]

    s_ini = s[s.index <= data_ini]
    if s_ini.empty:
        v_ini = s.iloc[0]
    else:
        v_ini = s_ini.iloc[-1]

    return v_fim / v_ini - 1


def _ret_mes(serie, ano, mes):
    """Retorno do mês cheio (fim do mês anterior -> fim do mês)."""
    s = serie.dropna()
    if s.empty:
        return np.nan

    ini = pd.Timestamp(ano, mes, 1) - pd.Timedelta(days=1)
    fim = pd.Timestamp(ano, mes, 1) + pd.offsets.MonthEnd(0)

    v_ini = s[s.index <= ini]
    v_fim = s[s.index <= fim]
    if v_ini.empty or v_fim.empty:
        return np.nan
    return v_fim.iloc[-1] / v_ini.iloc[-1] - 1


def _ret_ano(serie, ano):
    """
    Retorno do ano cheio (fim do ano anterior -> fim do ano).
    No ano de inception, usa o inception como ponto inicial.
    """
    s = serie.dropna()
    if s.empty:
        return np.nan

    inception = s.index.min()
    ini = pd.Timestamp(ano, 1, 1) - pd.Timedelta(days=1)
    fim = pd.Timestamp(ano, 12, 31)

    if ini < inception:
        v_ini = s.iloc[0]
    else:
        v = s[s.index <= ini]
        if v.empty:
            return np.nan
        v_ini = v.iloc[-1]

    v_fim = s[s.index <= fim]
    if v_fim.empty:
        return np.nan
    return v_fim.iloc[-1] / v_ini - 1


def tabela_retornos(df_port, n_meses=12):
    """
    Gera a tabela de retornos para todas as colunas (carteira + benchmarks):
    Since inception, YTD, LTM, últimos N meses, [coluna vazia], retornos anuais.
    """
    data_fim = df_port.dropna(how='all').index.max()
    ano_atual = data_fim.year

    inception   = df_port.index.min()
    inicio_ytd  = pd.Timestamp(ano_atual, 1, 1) - pd.Timedelta(days=1)
    inicio_ltm  = data_fim - pd.DateOffset(years=1)

    meses = []
    ref = pd.Timestamp(data_fim.year, data_fim.month, 1)
    for i in range(n_meses):
        m = ref - pd.DateOffset(months=i)
        meses.append((m.year, m.month))

    ano_inception = inception.year
    anos = list(range(ano_atual - 1, ano_inception - 1, -1))

    linhas = {}
    for col in df_port.columns:
        s = df_port[col]
        row = {
            'Since inception': _ret_periodo(s, inception, data_fim),
            'YTD':             _ret_periodo(s, inicio_ytd, data_fim),
            'LTM':             _ret_periodo(s, inicio_ltm, data_fim),
        }
        for (ano, mes) in meses:
            label = pd.Timestamp(ano, mes, 1).strftime('%b-%y')
            row[label] = _ret_mes(s, ano, mes)

        row[''] = np.nan

        for ano in anos:
            row[str(ano)] = _ret_ano(s, ano)

        linhas[col] = row

    tabela = pd.DataFrame(linhas).T

    col_meses = [pd.Timestamp(a, m, 1).strftime('%b-%y') for (a, m) in meses]
    col_anos  = [str(a) for a in anos]
    tabela = tabela[['Since inception', 'YTD', 'LTM'] + col_meses + [''] + col_anos]

    tabela_fmt = tabela.copy()
    for c in tabela_fmt.columns:
        if c == '':
            tabela_fmt[c] = ''
        else:
            tabela_fmt[c] = (tabela[c] * 100).round(1).astype(str) + '%'
            tabela_fmt[c] = tabela_fmt[c].replace('nan%', '')

    return tabela, tabela_fmt


def giro_medio_mensal(composition):
    """
    Giro médio mensal (two-way): para cada rebalanceamento,
        giro = Σ_i |peso_alvo_novo_i − peso_alvo_antigo_i|
    comparando os pesos-alvo (normalizados) de rebalanceamentos consecutivos.
    Ativos que entram/saem contam com peso 0 do outro lado.
    Retorna a média dos giros ao longo dos rebalanceamentos (mensais).
    """
    comp_long = composition.melt(id_vars='cod_ativo', var_name='data_rebal', value_name='peso')
    comp_long['data_rebal'] = pd.to_datetime(comp_long['data_rebal'])
    comp_long = comp_long.dropna(subset=['peso'])

    rebal_dates = sorted(comp_long['data_rebal'].unique())
    giros = []
    for i in range(1, len(rebal_dates)):
        ant = (comp_long[comp_long['data_rebal'] == rebal_dates[i - 1]]
               .set_index('cod_ativo')['peso'])
        nov = (comp_long[comp_long['data_rebal'] == rebal_dates[i]]
               .set_index('cod_ativo')['peso'])
        ant = ant / ant.sum()
        nov = nov / nov.sum()
        ativos = ant.index.union(nov.index)
        ant = ant.reindex(ativos, fill_value=0.0)
        nov = nov.reindex(ativos, fill_value=0.0)
        giros.append(float((nov - ant).abs().sum()))   # two-way (sem 0.5)

    return float(np.mean(giros)) if giros else np.nan


def _serie_mensal(serie):
    """
    Nível base 100 no fim de cada mês (último valor observado no mês).
    Usa agrupamento por período mensal (robusto a versões do pandas — evita
    o alias de frequência 'M'/'ME' do resample, que mudou entre versões).
    """
    s = serie.dropna()
    if s.empty:
        return s
    grp = s.groupby(s.index.to_period('M')).last()
    grp.index = grp.index.to_timestamp()   # índice datetime (início do mês), monotônico
    return grp


def retornos_mensais(serie):
    """Retornos mensais (fim de mês -> fim de mês). Exclui o mês parcial de inception."""
    return _serie_mensal(serie).pct_change().dropna()


def estatisticas_carteira(serie):
    """
    Estatísticas da lâmina, sobre os retornos mensais e a curva mensal (base 100):
      Meses positivos/negativos, retorno médio/máximo/mínimo mensal,
      maior drawdown e sua duração (em meses, do topo ao fundo).
    """
    r = retornos_mensais(serie)
    m = _serie_mensal(serie)   # níveis base 100 (para o drawdown)

    if r.empty or m.empty:
        return {
            'Meses positivos': np.nan, 'Meses negativos': np.nan,
            'Retorno médio mensal': np.nan, 'Retorno máximo mensal': np.nan,
            'Retorno mínimo mensal': np.nan, 'Maior drawdown': np.nan,
            'Duração do maior drawdown (meses)': np.nan,
        }

    run_max = m.cummax()
    dd = m / run_max - 1.0
    maior_dd = float(dd.min())
    fundo = dd.idxmin()
    topo = m.loc[:fundo].idxmax()
    duracao = (fundo.year - topo.year) * 12 + (fundo.month - topo.month)

    return {
        'Meses positivos':        int((r > 0).sum()),
        'Meses negativos':        int((r < 0).sum()),
        'Retorno médio mensal':   float(r.mean()),
        'Retorno máximo mensal':  float(r.max()),
        'Retorno mínimo mensal':  float(r.min()),
        'Maior drawdown':         maior_dd,
        'Duração do maior drawdown (meses)': int(duracao),
    }


def _nome_col_carteira(df_port):
    """Nome da coluna da carteira (a que não é benchmark) dentro do df de resultados."""
    for c in df_port.columns:
        if c not in ('Ibovespa', 'SMLL', 'ISEE'):
            return c
    return df_port.columns[0]


def tabela_performance_mensal_ano(df_port, ano=None):
    """
    Tabela 'Retornos <ano>' da lâmina: Jan..Dez do ano-calendário + total do ano (YTD).
    Linhas = carteira + benchmarks; colunas = meses + ano.
    Meses futuros ficam NaN (exibidos como '-' na lâmina).
    """
    data_fim = df_port.dropna(how='all').index.max()
    if ano is None:
        ano = data_fim.year

    ini_ano = pd.Timestamp(ano, 1, 1) - pd.Timedelta(days=1)
    meses_lbl = [pd.Timestamp(ano, mes, 1).strftime('%b') for mes in range(1, 13)]

    linhas = {}
    for col in df_port.columns:
        s = df_port[col]
        row = {}
        for mes, lbl in enumerate(meses_lbl, start=1):
            # mês ainda não iniciado (futuro) -> vazio (exibido como '-' na lâmina)
            if pd.Timestamp(ano, mes, 1) > data_fim:
                row[lbl] = np.nan
            else:
                row[lbl] = _ret_mes(s, ano, mes)
        row[str(ano)] = _ret_periodo(s, ini_ano, data_fim)
        linhas[col] = row

    tabela = pd.DataFrame(linhas).T                 # linhas = carteira/benchs
    return tabela[meses_lbl + [str(ano)]]


def tabela_retorno_anual(df_port):
    """
    'Retornos anos anteriores': retorno de cada ano-calendário fechado
    (do inception até o ano anterior ao atual). Linhas = anos; colunas = carteira + benchs.
    """
    data_fim = df_port.dropna(how='all').index.max()
    inception = df_port.index.min()
    anos = list(range(data_fim.year - 1, inception.year - 1, -1))

    linhas = {}
    for col in df_port.columns:
        s = df_port[col]
        linhas[col] = {str(a): _ret_ano(s, a) for a in anos}

    return pd.DataFrame(linhas)


def tabela_retornos_acumulados(df_port):
    """
    'Retornos acumulados': 12m, 24m, 36m e desde o início (Retorno Acumulado).
    Quando a janela ultrapassa o inception, usa o inception como ponto inicial.
    Linhas = período; colunas = carteira + benchmarks.
    """
    data_fim = df_port.dropna(how='all').index.max()
    inception = df_port.index.min()

    periodos = {
        'Últimos 12 meses': data_fim - pd.DateOffset(months=12),
        'Últimos 24 meses': data_fim - pd.DateOffset(months=24),
        'Últimos 36 meses': data_fim - pd.DateOffset(months=36),
        'Retorno Acumulado': inception,
    }

    linhas = {}
    for col in df_port.columns:
        s = df_port[col]
        linhas[col] = {k: _ret_periodo(s, ini, data_fim) for k, ini in periodos.items()}

    return pd.DataFrame(linhas)


def info_adicionais_lamina(df_port, composition, serie_cdi):
    """
    Bloco 'Informações adicionais' pedido: giro médio mensal, volatilidade
    anualizada (12m) e Índice de Sharpe (12m). Vol/Sharpe vêm de indicadores_12m.
    """
    ind = indicadores_12m(df_port, serie_cdi)
    nome_cart = _nome_col_carteira(df_port)
    return {
        'Giro médio mensal':        giro_medio_mensal(composition),
        'Volatilidade anualizada':  float(ind.loc['Volatilidade', nome_cart]),
        'Índice de Sharpe':         float(ind.loc['Sharpe', nome_cart]),
    }
