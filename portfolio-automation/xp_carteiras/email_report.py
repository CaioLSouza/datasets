"""Gera o e-mail mensal de carteiras em formato ``.msg``.

O módulo pode ser importado sem ler arquivos de rede nem abrir o Outlook. A
execução completa começa apenas em :func:`main`.
"""

from __future__ import annotations

import argparse
import datetime as dt
from html import escape
from pathlib import Path

import numpy as np
import pandas as pd

from .settings import Settings, load_settings


PORTFOLIOS = {
    "TOP Ações XP": {
        "base100": "top_acoes_base_100",
        "components": "componentes_top_acoes_atual",
        "title": "Carteira Top Ações",
    },
    "TOP DIVIDENDOS XP": {
        "base100": "top_dividendos_base_100",
        "components": "componentes_top_dividendos_atual",
        "title": "Carteira Top Dividendos",
    },
    "TOP SMALL CAPS XP": {
        "base100": "top_small_caps_base_100",
        "components": "componentes_top_small_caps_atual",
        "title": "Carteira Small Caps",
    },
    "ESG XP": {
        "base100": "esg_base_100",
        "components": "componentes_esg_atual",
        "title": "Carteira ESG",
    },
}

BENCHMARK_NAMES = {"Ibovespa": "Ibovespa", "SMLL": "SMLL", "ISEE": "ISE"}

SECTOR_OVERVIEW = [
    ("Transportation", ["Rental", "Infrastructure", "Airlines"]),
    ("Financials", []),
    ("Real Estate", ["Homebuilders", "Income Properties"]),
    ("Commodities", ["Metals & Mining", "Pulp & Paper", "Oil, Gas & Petrochemicals"]),
    ("Capital Goods", []),
    ("TMT", []),
    ("Utilities", []),
    ("Retail", []),
    ("Agribusiness, Food & Beverages", ["Frigoríficos", "F&B", "S&E", "Grãos"]),
]

WINDOW_COLUMNS = ["1D", "1M", "3M", "6M", "12M", "24M", "MTD", "YTD", "Desde Início"]
RISK_COLUMNS = [
    "VOLATILIDADE",
    "SHARPE",
    "BETA",
    "CORRELAÇÃO IBOV",
    "TRACKING ERROR",
    "DRAWDOWN",
]

HIGHLIGHT_STYLE = "background-color:#FFFF00;"
DEFAULT_SUBJECT = "Reunião de carteiras | Preview setorial e performance"
EMAIL_FILENAME = "email_carteiras.msg"


# ============================================================================
# Dados e cálculos
# ============================================================================

def load_indices(path: Path) -> tuple[pd.Series, pd.Series]:
    """Carrega as séries de Ibovespa e CDI usadas no relatório."""
    indices = pd.read_parquet(path)
    indices["data"] = pd.to_datetime(indices["data"])

    ibov = (
        indices.loc[indices["cod_ativo"] == "Ibovespa", ["data", "close_price"]]
        .drop_duplicates("data")
        .set_index("data")["close_price"]
        .sort_index()
    )
    cdi = (
        indices.loc[indices["cod_ativo"] == "CDI Acumulado", ["data", "close_price"]]
        .drop_duplicates("data")
        .set_index("data")["close_price"]
        .sort_index()
    )
    return ibov, cdi


def _ultimo_trading_day_do_mes(data_ref, serie_precos):
    """Último trading day do mês de ``data_ref`` presente na série."""
    data = pd.to_datetime(data_ref)
    fim_mes = pd.Timestamp(data.year, data.month, 1) + pd.offsets.MonthEnd(0)
    disponiveis = serie_precos[serie_precos.index <= fim_mes]
    return disponiveis.index[-1] if not disponiveis.empty else None


def _preco_em(serie_precos, data_alvo):
    serie = serie_precos[serie_precos.index <= pd.to_datetime(data_alvo)]
    return serie.iloc[-1] if not serie.empty else np.nan


def _ret_entre(serie, data_ini, data_fim):
    serie = serie.dropna()
    if serie.empty:
        return np.nan
    valores_fim = serie[serie.index <= data_fim]
    if valores_fim.empty:
        return np.nan
    valores_ini = serie[serie.index <= data_ini]
    valor_fim = valores_fim.iloc[-1]
    valor_ini = valores_ini.iloc[-1] if not valores_ini.empty else serie.iloc[0]
    return valor_fim / valor_ini - 1


def calcular_janelas(serie):
    serie = serie.dropna()
    if serie.empty:
        return {}
    data_fim = serie.index.max()
    retorno_1d = serie.iloc[-1] / serie.iloc[-2] - 1 if len(serie) >= 2 else np.nan
    inicio_mtd = pd.Timestamp(data_fim.year, data_fim.month, 1) - pd.Timedelta(days=1)
    inicio_ytd = pd.Timestamp(data_fim.year, 1, 1) - pd.Timedelta(days=1)
    return {
        "1D": retorno_1d,
        "1M": _ret_entre(serie, data_fim - pd.DateOffset(months=1), data_fim),
        "3M": _ret_entre(serie, data_fim - pd.DateOffset(months=3), data_fim),
        "6M": _ret_entre(serie, data_fim - pd.DateOffset(months=6), data_fim),
        "12M": _ret_entre(serie, data_fim - pd.DateOffset(months=12), data_fim),
        "24M": _ret_entre(serie, data_fim - pd.DateOffset(months=24), data_fim),
        "MTD": _ret_entre(serie, inicio_mtd, data_fim),
        "YTD": _ret_entre(serie, inicio_ytd, data_fim),
        "Desde Início": serie.iloc[-1] / serie.iloc[0] - 1,
    }


def _retornos_diarios(serie):
    return serie.dropna().pct_change().dropna()


def _janela_12m(serie):
    serie = serie.dropna()
    if serie.empty:
        return serie
    corte = serie.index.max() - pd.DateOffset(months=12)
    return serie[serie.index >= corte]


def calcular_risco(serie_carteira, serie_ibov, serie_cdi):
    """Calcula risco em 12 meses; beta, correlação e TE usam o Ibovespa."""
    serie = _janela_12m(serie_carteira)
    ibov_12m = _janela_12m(serie_ibov) if serie_ibov is not None else None
    if serie.empty or len(serie) < 5:
        return {}

    retornos = _retornos_diarios(serie)
    volatilidade = retornos.std() * np.sqrt(252)

    if serie_cdi is not None:
        retornos_cdi = _retornos_diarios(_janela_12m(serie_cdi))
        datas = retornos.index.intersection(retornos_cdi.index)
        excesso = retornos.loc[datas] - retornos_cdi.loc[datas]
        sharpe = excesso.mean() * 252 / volatilidade if volatilidade > 0 else np.nan
    else:
        sharpe = np.nan

    beta = correlacao = tracking_error = np.nan
    if ibov_12m is not None and not ibov_12m.empty:
        retornos_ibov = _retornos_diarios(ibov_12m)
        datas = retornos.index.intersection(retornos_ibov.index)
        if len(datas) > 5:
            retorno_carteira = retornos.loc[datas]
            retorno_ibov = retornos_ibov.loc[datas]
            variancia_ibov = np.var(retorno_ibov)
            beta = (
                np.cov(retorno_carteira, retorno_ibov)[0, 1] / variancia_ibov
                if variancia_ibov > 0
                else np.nan
            )
            correlacao = np.corrcoef(retorno_carteira, retorno_ibov)[0, 1]
            tracking_error = (retorno_carteira - retorno_ibov).std() * np.sqrt(252)

    drawdown = serie / serie.cummax() - 1
    return {
        "VOLATILIDADE": volatilidade,
        "SHARPE": sharpe,
        "BETA": beta,
        "CORRELAÇÃO IBOV": correlacao,
        "TRACKING ERROR": tracking_error,
        "DRAWDOWN": drawdown.min(),
    }


def carregar_componentes(arquivo: str, output_dir: Path, ibov: pd.Series) -> pd.DataFrame:
    """Carrega componentes e calcula o Ibovespa desde a entrada de cada papel."""
    tabela = pd.read_excel(output_dir / f"{arquivo}.xlsx", sheet_name="componentes")
    data_fim = ibov.index.max()

    ibov_desde = []
    for _, row in tabela.iterrows():
        data_exibicao = pd.to_datetime(row["Data de entrada"])
        mes_anterior = data_exibicao - pd.offsets.MonthBegin(1)
        data_entrada = _ultimo_trading_day_do_mes(mes_anterior, ibov)
        if data_entrada is None:
            ibov_desde.append(np.nan)
            continue
        preco_inicial = _preco_em(ibov, data_entrada)
        preco_final = _preco_em(ibov, data_fim)
        ibov_desde.append(
            preco_final / preco_inicial - 1 if preco_inicial and preco_final else np.nan
        )

    tabela["Desempenho do Ibov desde entrada"] = ibov_desde
    return tabela.sort_values("Companhia").reset_index(drop=True)


def build_performance_tables(
    output_dir: Path, cdi: pd.Series
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Consolida janelas de retorno e métricas de risco de todas as carteiras."""
    linhas_janelas: dict[str, dict] = {}
    linhas_risco: dict[str, dict] = {}
    ordem: list[str] = []

    for nome_carteira, info in PORTFOLIOS.items():
        tabela = pd.read_excel(
            output_dir / f"{info['base100']}.xlsx",
            sheet_name="base100",
            index_col=0,
        )
        tabela.index = pd.to_datetime(tabela.index)
        coluna_carteira = tabela.columns[0]
        benchmarks = list(tabela.columns[1:])
        ibov_local = tabela["Ibovespa"] if "Ibovespa" in tabela.columns else None

        linhas_janelas[nome_carteira] = calcular_janelas(tabela[coluna_carteira])
        linhas_risco[nome_carteira] = calcular_risco(tabela[coluna_carteira], ibov_local, cdi)
        ordem.append(nome_carteira)

        for benchmark in benchmarks:
            nome = BENCHMARK_NAMES.get(benchmark, benchmark)
            if nome not in linhas_janelas:
                linhas_janelas[nome] = calcular_janelas(tabela[benchmark])
                linhas_risco[nome] = calcular_risco(tabela[benchmark], ibov_local, cdi)
                ordem.append(nome)

    janelas = pd.DataFrame(linhas_janelas).T.reindex(ordem)[WINDOW_COLUMNS]
    risco = pd.DataFrame(linhas_risco).T.reindex(ordem)[RISK_COLUMNS]
    return janelas, risco


# ============================================================================
# HTML
# ============================================================================

def _grifar(texto: str) -> str:
    return f'<span style="{HIGHLIGHT_STYLE}">{texto}</span>'


def _pct(valor, casas: int = 1) -> str:
    return "" if pd.isna(valor) else f"{valor * 100:.{casas}f}%"


def _num(valor, casas: int = 2) -> str:
    return "" if pd.isna(valor) else f"{valor:.{casas}f}"


def _data_br(valor) -> str:
    return "" if valor is None or pd.isna(valor) else pd.to_datetime(valor).strftime("%d/%m/%Y")


def _mes_ano(valor) -> str:
    return "" if valor is None or pd.isna(valor) else pd.to_datetime(valor).strftime("%b/%y")


def tabela_html(df: pd.DataFrame, formatadores: dict, index_label: str = "NOME") -> str:
    th = '<th style="border:1px solid #999;padding:3px;background:#FFC000;color:#000;font-size:10pt;">'
    td = '<td style="border:1px solid #ccc;padding:3px;font-size:10pt;text-align:center;">'
    td_esquerda = '<td style="border:1px solid #ccc;padding:3px;font-size:10pt;text-align:left;">'

    partes = ['<table style="border-collapse:collapse;font-family:Calibri;">']
    partes.append(f"<tr>{th}{escape(index_label)}</th>")
    partes.extend(f"{th}{escape(str(coluna))}</th>" for coluna in df.columns)
    partes.append("</tr>")

    for indice, row in df.iterrows():
        partes.append(f"<tr>{td_esquerda}{escape(str(indice))}</td>")
        for coluna in df.columns:
            valor = formatadores.get(coluna, lambda item: item)(row[coluna])
            partes.append(f"{td}{escape(str(valor))}</td>")
        partes.append("</tr>")
    partes.append("</table>")
    return "".join(partes)


def tabela_componentes_html(df: pd.DataFrame) -> str:
    colunas = [
        "Companhia",
        "Ticker",
        "Setor",
        "Peso",
        "Data de entrada",
        "Desempenho desde entrada",
        "Desempenho do Ibov desde entrada",
        "Desempenho no mês",
        "Desempenho YTD",
    ]
    colunas = [coluna for coluna in colunas if coluna in df.columns]
    formatadores = {
        "Peso": _pct,
        "Desempenho desde entrada": _pct,
        "Desempenho do Ibov desde entrada": _pct,
        "Desempenho no mês": _pct,
        "Desempenho YTD": _pct,
        "Data de entrada": _mes_ano,
    }
    return tabela_html(
        df[colunas].set_index("Companhia"),
        {chave: valor for chave, valor in formatadores.items() if chave in colunas},
        index_label="Companhia",
    )


def build_email_html(
    settings: Settings,
    ibov: pd.Series,
    tabela_janelas: pd.DataFrame,
    tabela_risco: pd.DataFrame,
) -> str:
    """Monta o corpo completo do e-mail em HTML."""
    hoje = dt.date.today().strftime("%d/%m")
    ultimo_dia = _data_br(ibov.index.max())
    formatadores_janelas = {coluna: _pct for coluna in WINDOW_COLUMNS}
    formatadores_risco = {
        "VOLATILIDADE": lambda valor: _pct(valor, 2),
        "SHARPE": lambda valor: _num(valor, 2),
        "BETA": lambda valor: _num(valor, 2),
        "CORRELAÇÃO IBOV": lambda valor: _num(valor, 2),
        "TRACKING ERROR": lambda valor: _num(valor, 2),
        "DRAWDOWN": lambda valor: _pct(valor, 2),
    }

    partes = [
        '<div style="font-family:Calibri;font-size:13px;">',
        "Fala, pessoal, tudo bem?<br><br>",
        f"Segue o e-mail de carteiras para a reunião de hoje ({hoje}) às 17h00.<br><br>",
        f"Consolidei tudo neste e-mail – seções em {_grifar('amarelo')} ainda a serem preenchidas.<br><br>",
        "<b>1) Overview do setor no mês:</b><br>",
        '<ul style="margin-top:2px;">',
    ]
    for setor, subsetores in SECTOR_OVERVIEW:
        partes.append(f"<li>{_grifar(escape(setor))}:")
        if subsetores:
            partes.append("<ul>")
            partes.extend(f"<li>{_grifar(escape(subsetor))}: </li>" for subsetor in subsetores)
            partes.append("</ul>")
        partes.append("</li>")
    partes.append("</ul><br>")

    partes.extend(
        [
            "<b>2) Comentário breve sobre a performance dos papéis nas carteiras:</b><br>",
        ]
    )
    componentes: dict[str, pd.DataFrame] = {}
    for nome_carteira, info in PORTFOLIOS.items():
        tabela = carregar_componentes(info["components"], settings.output_dir, ibov)
        componentes[nome_carteira] = tabela
        partes.extend([f"<br><b>{escape(info['title'])}:</b>", '<ul style="margin-top:2px;">'])
        for _, row in tabela.iterrows():
            papel = f"{row['Companhia']} ({row['Ticker']})"
            partes.append(f"<li>{_grifar(escape(papel))}: </li>")
        partes.append("</ul>")
    partes.append("<br>")

    partes.extend(["<b>3) Sugestões de troca de papéis nas carteiras:</b><br>", '<ul style="margin-top:2px;">'])
    for info in PORTFOLIOS.values():
        partes.append(f"<li><b>{escape(info['title'])}:</b> {_grifar('&nbsp;&nbsp;&nbsp;')}</li>")
    partes.extend(
        [
            "</ul><br>",
            f"<b>Performance com dados de {ultimo_dia}:</b><br><br>",
            tabela_html(tabela_janelas, formatadores_janelas) + "<br><br>",
            tabela_html(tabela_risco, formatadores_risco) + "<br><br>",
        ]
    )

    for nome_carteira, info in PORTFOLIOS.items():
        partes.extend(
            [
                f"<b>{escape(info['title'])}:</b><br>",
                tabela_componentes_html(componentes[nome_carteira]) + "<br><br>",
            ]
        )

    partes.extend(["<br>Qualquer coisa, só nos puxar aqui!<br><br>Obrigado.<br>", "</div>"])
    return "".join(partes)


# ============================================================================
# Outlook e entrada do programa
# ============================================================================

def save_outlook_message(
    html_body: str,
    destination: Path,
    subject: str = DEFAULT_SUBJECT,
    display: bool = False,
) -> Path:
    """Salva o ``.msg`` no Outlook e, opcionalmente, abre o rascunho."""
    try:
        import win32com.client as win32
    except ImportError as exc:
        raise RuntimeError(
            "O pacote pywin32 e o Microsoft Outlook são necessários para gerar o .msg."
        ) from exc

    outlook = win32.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.Subject = subject
    message.HTMLBody = html_body
    message.SaveAs(str(destination))
    if display:
        message.Display()
    return destination


def generate_email(settings: Settings, display: bool = False) -> Path:
    """Executa a geração completa do e-mail e retorna o caminho salvo."""
    ibov, cdi = load_indices(settings.indices_path)
    tabela_janelas, tabela_risco = build_performance_tables(settings.output_dir, cdi)
    html_body = build_email_html(settings, ibov, tabela_janelas, tabela_risco)
    destination = settings.email_dir / EMAIL_FILENAME
    return save_outlook_message(html_body, destination, display=display)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--display",
        action="store_true",
        help="abre o rascunho no Outlook depois de salvar o arquivo .msg",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    destination = generate_email(load_settings(), display=args.display)
    print(f"E-mail (.msg) gerado: {destination}")


if __name__ == "__main__":
    main()
