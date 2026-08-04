"""Geração da Prestação de Contas em PowerPoint.

O template contém gráficos que podem estar vinculados a arquivos corporativos.
Por isso esta rotina usa o próprio PowerPoint, via COM, e grava gráficos com
dados incorporados no arquivo de saída. Os textos editoriais das páginas 1 e 2
não são alterados.
"""

from __future__ import annotations

import shutil
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd

from .excel_reports import _fmt_pct_br
from .performance import _nome_col_carteira, _ret_ano, _ret_mes, _ret_periodo


MESES_EXT_PT = (
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
)
MESES_ABREV_PT = (
    "jan", "fev", "mar", "abr", "mai", "jun",
    "jul", "ago", "set", "out", "nov", "dez",
)

XL_COLUMN_STACKED = 52
XL_LINE = 4
XL_LABEL_POSITION_OUTSIDE_END = 2
MSO_FALSE = 0
MSO_TRUE = -1

# PowerPoint/Office usa BGR nesses inteiros. Estes valores representam:
# amarelo #FFBC00, cinza #7E7E7E e texto #595959.
XP_YELLOW = 0x00BCFF
NEUTRAL_GRAY = 0x7E7E7E
TEXT_GRAY = 0x595959


@dataclass(frozen=True)
class AccountabilityPeriod:
    """Mês fechado de desempenho e mês seguinte exibido no relatório."""

    reference_year: int
    reference_month: int
    report_year: int
    report_month: int


def next_month(year: int, month: int) -> tuple[int, int]:
    """Retorna o mês imediatamente posterior."""
    return year + (month // 12), (month % 12) + 1


def resolve_accountability_period(
    df_port: pd.DataFrame,
    *,
    today: date | pd.Timestamp | None = None,
    reference_year: int | None = None,
    reference_month: int | None = None,
) -> AccountabilityPeriod:
    """Define o mês fechado usado nos números da Prestação de Contas.

    Se a base já contém observações do mês corrente, usa o mês anterior. Se a
    base ainda termina em um mês anterior, usa o próprio último mês disponível.
    O período pode ser fixado explicitamente para reprocessamentos históricos.
    """
    if (reference_year is None) != (reference_month is None):
        raise ValueError("Informe reference_year e reference_month juntos.")
    if reference_year is not None and reference_month is not None:
        if not 1 <= reference_month <= 12:
            raise ValueError("reference_month deve estar entre 1 e 12.")
        year, month = int(reference_year), int(reference_month)
    else:
        available = df_port.dropna(how="all")
        if available.empty:
            raise ValueError("df_port não possui dados para definir o período.")
        latest = pd.Timestamp(available.index.max())
        current = pd.Timestamp(today or date.today())
        latest_period = latest.to_period("M")
        current_period = current.to_period("M")
        reference = latest_period - 1 if latest_period >= current_period else latest_period
        year, month = reference.year, reference.month

    report_year, report_month = next_month(year, month)
    return AccountabilityPeriod(year, month, report_year, report_month)


def performance_summary(
    df_port: pd.DataFrame,
    reference_year: int,
    reference_month: int,
) -> pd.DataFrame:
    """Calcula mês, acumulado no ano e últimos 12 meses até o mês fechado."""
    cutoff = pd.Timestamp(reference_year, reference_month, 1) + pd.offsets.MonthEnd(0)
    closed = df_port.loc[df_port.index <= cutoff].copy()
    if closed.dropna(how="all").empty:
        raise ValueError("Não há dados de performance até o mês de referência.")
    start_12m = cutoff - pd.DateOffset(months=12)
    rows = []
    for column in closed.columns:
        series = closed[column]
        rows.append(
            {
                "Série": column,
                "Mês": _ret_mes(series, reference_year, reference_month),
                "Acumulado": _ret_ano(series, reference_year),
                "12 meses": _ret_periodo(series, start_12m, cutoff),
            }
        )
    return pd.DataFrame(rows).set_index("Série")


def reconcile_decomposition_total(
    decomposition: pd.DataFrame,
    expected_return: float,
    *,
    tolerance: float = 0.0005,
) -> pd.DataFrame:
    """Garante que o total do waterfall coincide com o retorno da 1ª página.

    Diferenças de até 5 bps recebem uma barra explícita de reconciliação.
    Diferenças maiores interrompem a geração para nunca publicar números
    contraditórios.
    """
    if pd.isna(expected_return):
        raise ValueError("O retorno mensal da carteira não está disponível para reconciliação.")
    result = decomposition.copy()
    total_mask = result["Ticker"].astype(str).str.strip() == ""
    details = result.loc[~total_mask]
    calculated = float(details["Contribuição"].sum(skipna=True))
    difference = float(expected_return) - calculated
    if abs(difference) > tolerance:
        raise ValueError(
            "O total da decomposição não confere com o retorno mensal da carteira: "
            f"decomposição={calculated:.4%}, retorno={expected_return:.4%}."
        )

    total_row = result.loc[total_mask].copy()
    result = result.loc[~total_mask].copy()
    if abs(difference) > 1e-10:
        adjustment = {column: np.nan for column in result.columns}
        adjustment.update(
            {
                "Companhia": "Ajuste de reconciliação",
                "Ticker": "Ajuste",
                "Setor": "",
                "Contribuição": difference,
            }
        )
        result = pd.concat([result, pd.DataFrame([adjustment])], ignore_index=True)

    if total_row.empty:
        total = {column: np.nan for column in result.columns}
        total.update({"Companhia": "Carteira (total)", "Ticker": "", "Contribuição": expected_return})
        total_row = pd.DataFrame([total])
    else:
        total_row.loc[:, "Contribuição"] = expected_return
    return pd.concat([result, total_row], ignore_index=True)


def waterfall_arrays(
    tickers: list[str],
    contributions: list[float],
    total: float | None,
    total_label: str = "Carteira",
) -> tuple[list[str], list[float], list[float | None], list[float | None], list[float | None], list[float]]:
    """Monta as quatro séries de um waterfall com colunas empilhadas."""
    clean = [0.0 if pd.isna(value) else float(value) for value in contributions]
    resolved_total = float(np.nansum(clean)) if total is None or pd.isna(total) else float(total)
    positive_region = resolved_total >= 0
    pairs = sorted(zip(tickers, clean), key=lambda item: item[1], reverse=positive_region)

    categories: list[str] = []
    base: list[float] = []
    increases: list[float | None] = []
    decreases: list[float | None] = []
    totals: list[float | None] = []
    ordered: list[float] = []
    cumulative = 0.0

    for ticker, contribution in pairs:
        before = cumulative
        cumulative += contribution
        categories.append(str(ticker))
        if positive_region:
            floating_base = before if contribution >= 0 else cumulative
            height = abs(contribution)
        else:
            floating_base = max(before, cumulative)
            height = -abs(contribution)
        base.append(floating_base)
        increases.append(height if contribution >= 0 else None)
        decreases.append(height if contribution < 0 else None)
        totals.append(None)
        ordered.append(contribution)

    categories.append(total_label)
    base.append(0.0)
    increases.append(None)
    decreases.append(None)
    totals.append(resolved_total)
    return categories, base, increases, decreases, totals, ordered


def accountability_output_filename(portfolio_label: str, period: AccountabilityPeriod) -> str:
    """Nome mensal do arquivo final, no padrão do template fornecido."""
    month = MESES_EXT_PT[period.report_month - 1].capitalize()
    return f"Prestação de Contas - {portfolio_label} - {month} {period.report_year}.pptx"


def _iter_shapes(slide):
    for index in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(index)
        yield shape
        try:
            group_items = shape.GroupItems
            for group_index in range(1, group_items.Count + 1):
                yield group_items(group_index)
        except Exception:
            pass


def _shape_by_name(slide, name: str):
    for shape in _iter_shapes(slide):
        if str(shape.Name) == name:
            return shape
    raise KeyError(f"Shape não encontrado: {name}")


def _set_text(shape, value: str) -> None:
    shape.TextFrame.TextRange.Text = str(value)


def _replace_month_prefix(shape, month_year: str, portfolio_label: str) -> None:
    _set_text(shape, f"{month_year} | Carteira {portfolio_label}")


def _update_portfolio_titles(slides, portfolio_label: str) -> None:
    """Troca apenas títulos exatos, preservando os textos editoriais."""
    for slide_number in (1, 2):
        for shape in _iter_shapes(slides(slide_number)):
            if not getattr(shape, "HasTextFrame", False):
                continue
            current = str(shape.TextFrame.TextRange.Text).strip()
            if current.startswith("Top Ações XP"):
                shape.TextFrame.TextRange.Paragraphs(1).Text = portfolio_label


def _set_table_cell(table, row: int, column: int, value: str) -> None:
    table.Cell(row, column).Shape.TextFrame.TextRange.Text = str(value)


def _find_table(slide, first_header: str):
    for shape in _iter_shapes(slide):
        if not getattr(shape, "HasTable", False):
            continue
        table = shape.Table
        header = str(table.Cell(1, 1).Shape.TextFrame.TextRange.Text).strip()
        if header.startswith(first_header):
            return shape, table
    raise KeyError(f"Tabela com cabeçalho '{first_header}' não encontrada.")


def _benchmark_column(df_port: pd.DataFrame) -> str:
    portfolio = _nome_col_carteira(df_port)
    benchmarks = [column for column in df_port.columns if column != portfolio]
    if not benchmarks:
        raise ValueError("A Prestação de Contas requer uma série de benchmark.")
    return benchmarks[0]


def _update_summary_table(
    slide,
    df_port: pd.DataFrame,
    period: AccountabilityPeriod,
    portfolio_label: str,
) -> None:
    _, table = _find_table(slide, "Desempenho")
    summary = performance_summary(df_port, period.reference_year, period.reference_month)
    portfolio = _nome_col_carteira(df_port)
    benchmark = _benchmark_column(df_port)
    month = MESES_EXT_PT[period.reference_month - 1].capitalize()
    _set_table_cell(table, 1, 2, f"{month} {period.reference_year}")
    _set_table_cell(table, 1, 3, f"Acumulado {period.reference_year}")

    labels_and_series = ((portfolio_label, portfolio), (benchmark, benchmark))
    for row, (label, series) in enumerate(labels_and_series, start=2):
        _set_table_cell(table, row, 1, label)
        _set_table_cell(table, row, 2, _fmt_pct_br(summary.loc[series, "Mês"], 1))
        _set_table_cell(table, row, 3, _fmt_pct_br(summary.loc[series, "Acumulado"], 1))
        _set_table_cell(table, row, 4, _fmt_pct_br(summary.loc[series, "12 meses"], 1))


def _unmerge_composition_table(table) -> None:
    """Desfaz merges verticais do setor sem falhar em células já simples."""
    row = 2
    while row <= table.Rows.Count:
        cell = table.Cell(row, 1)
        try:
            row_height = float(table.Rows(row).Height)
            span = max(1, round(float(cell.Shape.Height) / row_height))
            if span > 1:
                cell.Split(span, 1)
            row += span
        except Exception:
            row += 1


def _resize_composition_table(table, data_rows: int) -> None:
    while table.Rows.Count - 1 < data_rows:
        table.Rows.Add()
    while table.Rows.Count - 1 > data_rows:
        table.Rows(table.Rows.Count).Delete()


def _update_composition_table(slide, composition: pd.DataFrame) -> None:
    _, table = _find_table(slide, "Setor")
    required = ["Setor", "Companhia", "Ticker", "Peso"]
    missing = [column for column in required if column not in composition.columns]
    if missing:
        raise ValueError(f"Colunas ausentes na composição: {', '.join(missing)}")
    compact = composition[required].dropna(subset=["Ticker"]).copy()
    compact = compact.sort_values(["Setor", "Companhia"], kind="stable").reset_index(drop=True)

    _unmerge_composition_table(table)
    _resize_composition_table(table, len(compact))
    if len(compact):
        row_height = table.Parent.Height / (len(compact) + 1)
        for row in range(1, table.Rows.Count + 1):
            try:
                table.Rows(row).Height = row_height
            except Exception:
                pass

    for index, item in compact.iterrows():
        row = index + 2
        _set_table_cell(table, row, 1, item["Setor"])
        _set_table_cell(table, row, 2, item["Companhia"])
        _set_table_cell(table, row, 3, item["Ticker"])
        _set_table_cell(table, row, 4, _fmt_pct_br(item["Peso"], 1))

    start = 2
    while start <= table.Rows.Count:
        sector = str(table.Cell(start, 1).Shape.TextFrame.TextRange.Text).strip()
        end = start
        while end + 1 <= table.Rows.Count:
            next_sector = str(table.Cell(end + 1, 1).Shape.TextFrame.TextRange.Text).strip()
            if next_sector != sector:
                break
            end += 1
        if end > start:
            for row in range(start + 1, end + 1):
                _set_table_cell(table, row, 1, "")
            try:
                table.Cell(start, 1).Merge(table.Cell(end, 1))
            except Exception:
                pass
        start = end + 1


def _safe_delete_chart(shape) -> None:
    try:
        shape.Delete()
    except Exception:
        pass


def _format_series(series, rgb: int, *, transparent: bool = False) -> None:
    series.Format.Line.Visible = MSO_FALSE
    if transparent:
        series.Format.Fill.Visible = MSO_FALSE
    else:
        series.Format.Fill.Visible = MSO_TRUE
        series.Format.Fill.Solid()
        series.Format.Fill.ForeColor.RGB = rgb


def _label_text(value: float) -> str:
    return _fmt_pct_br(value, 1)


def _apply_custom_labels(series, labels: dict[int, str]) -> None:
    if not labels:
        return
    try:
        series.ApplyDataLabels()
    except Exception:
        return
    for point_index in range(1, series.Points().Count + 1):
        try:
            series.Points(point_index).HasDataLabel = False
        except Exception:
            pass
    for point_index, label in labels.items():
        try:
            point = series.Points(point_index)
            point.HasDataLabel = True
            point.DataLabel.Text = label
            point.DataLabel.Position = XL_LABEL_POSITION_OUTSIDE_END
            point.DataLabel.Font.Name = "Roboto Condensed"
            point.DataLabel.Font.Size = 8
            point.DataLabel.Font.Color = TEXT_GRAY
        except Exception:
            pass


def _replace_chart_series(chart, names: list[str], categories: list[object], values: list[list[object]]):
    """Recria as séries com arrays literais, sem abrir uma instância do Excel."""
    collection = chart.SeriesCollection()
    while collection.Count:
        collection(1).Delete()
    created = []
    for name, series_values in zip(names, values):
        series = collection.NewSeries()
        series.Name = name
        series.XValues = tuple(categories)
        series.Values = tuple(
            0.0 if value is None or pd.isna(value) else float(value)
            for value in series_values
        )
        created.append(series)
    return created


def _replace_waterfall(slide, decomposition: pd.DataFrame) -> None:
    total_rows = decomposition[decomposition["Ticker"].astype(str).str.strip() == ""]
    total = float(total_rows["Contribuição"].iloc[0]) if not total_rows.empty else None
    detail = decomposition[decomposition["Ticker"].astype(str).str.strip() != ""].copy()
    arrays = waterfall_arrays(
        detail["Ticker"].astype(str).tolist(),
        detail["Contribuição"].tolist(),
        total,
    )
    categories, base, increases, decreases, totals, ordered = arrays

    source = _shape_by_name(slide, "Chart 3")
    left, top, width, height = source.Left, source.Top, source.Width, source.Height
    _safe_delete_chart(source)
    for shape in list(_iter_shapes(slide)):
        if str(shape.Name) == "Gráfico 45":
            _safe_delete_chart(shape)

    chart_shape = slide.Shapes.AddChart2(201, XL_COLUMN_STACKED, left, top, width, height, True)
    chart_shape.Name = "Gráfico de decomposição"
    chart = chart_shape.Chart
    series = _replace_chart_series(
        chart,
        ["Base", "Alta", "Baixa", "Total"],
        categories,
        [base, increases, decreases, totals],
    )
    chart.ChartType = XL_COLUMN_STACKED
    chart.HasLegend = False
    chart.HasTitle = False

    series[0].Format.Fill.Visible = MSO_FALSE
    series[0].Format.Line.Visible = MSO_FALSE
    _format_series(series[1], NEUTRAL_GRAY)
    _format_series(series[2], NEUTRAL_GRAY)
    _format_series(series[3], XP_YELLOW)
    chart.ChartGroups(1).Overlap = 100
    chart.ChartGroups(1).GapWidth = 55

    high_labels = {index + 1: _label_text(value) for index, value in enumerate(ordered) if value >= 0}
    low_labels = {index + 1: _label_text(value) for index, value in enumerate(ordered) if value < 0}
    _apply_custom_labels(series[1], high_labels)
    _apply_custom_labels(series[2], low_labels)
    _apply_custom_labels(series[3], {len(categories): _label_text(float(totals[-1]))})

    try:
        value_axis = chart.Axes(2)
        value_axis.HasMajorGridlines = False
        value_axis.TickLabelPosition = -4142
        value_axis.Format.Line.Visible = MSO_FALSE
    except Exception:
        pass
    try:
        category_axis = chart.Axes(1)
        category_axis.Format.Line.Visible = MSO_FALSE
        category_axis.TickLabels.Font.Name = "Roboto Condensed"
        category_axis.TickLabels.Font.Size = 8
        category_axis.TickLabels.Orientation = 45
    except Exception:
        pass
    try:
        chart.ChartArea.Format.Fill.Visible = MSO_FALSE
        chart.ChartArea.Format.Line.Visible = MSO_FALSE
        chart.PlotArea.Format.Fill.Visible = MSO_FALSE
        chart.PlotArea.Format.Line.Visible = MSO_FALSE
    except Exception:
        pass
def _update_base100_chart(slide, df_port: pd.DataFrame, period: AccountabilityPeriod) -> None:
    shape = _shape_by_name(slide, "Chart 10")
    chart = shape.Chart
    cutoff = pd.Timestamp(period.reference_year, period.reference_month, 1) + pd.offsets.MonthEnd(0)
    data = df_port.loc[df_port.index <= cutoff].dropna(how="all").copy()
    if data.empty:
        raise ValueError("Não há dados para o gráfico base 100.")
    portfolio = _nome_col_carteira(data)
    benchmark = _benchmark_column(data)
    data = data[[portfolio, benchmark]]
    categories = tuple(pd.Timestamp(index).to_pydatetime() for index in data.index)
    portfolio_values = tuple(None if pd.isna(value) else float(value) for value in data[portfolio])
    benchmark_values = tuple(None if pd.isna(value) else float(value) for value in data[benchmark])
    series = chart.SeriesCollection()
    while series.Count > 2:
        series(series.Count).Delete()
    while series.Count < 2:
        series.NewSeries()
    series(1).Name = portfolio
    series(1).XValues = categories
    series(1).Values = portfolio_values
    series(2).Name = benchmark
    series(2).XValues = categories
    series(2).Values = benchmark_values
    chart.ChartType = XL_LINE
    inception = pd.Timestamp(data.index.min())
    _set_text(
        _shape_by_name(slide, "CaixaDeTexto 5"),
        f"Performance desde o início vs. {benchmark} "
        f"(100 = {MESES_ABREV_PT[inception.month - 1]}. {inception.year})",
    )


def _update_dates(slides, period: AccountabilityPeriod, portfolio_label: str) -> None:
    report = f"{MESES_EXT_PT[period.report_month - 1].capitalize()} de {period.report_year}"
    reference = f"{MESES_EXT_PT[period.reference_month - 1].capitalize()} de {period.reference_year}"
    _replace_month_prefix(_shape_by_name(slides(1), "object 7"), report, portfolio_label)
    _replace_month_prefix(_shape_by_name(slides(2), "object 7"), reference, portfolio_label)
    _set_text(
        _shape_by_name(slides(2), "CaixaDeTexto 24"),
        f"Decomposição do retorno da carteira ({MESES_EXT_PT[period.reference_month - 1].capitalize()}/{str(period.reference_year)[2:]})",
    )
    _set_text(
        _shape_by_name(slides(2), "CaixaDeTexto 9"),
        f"Composição da Carteira – {MESES_EXT_PT[period.report_month - 1].capitalize()} de {str(period.report_year)[2:]}",
    )
    _set_text(
        _shape_by_name(slides(3), "object 7"),
        f"1 de {MESES_EXT_PT[period.report_month - 1].capitalize()} de {period.report_year}",
    )


def atualizar_prestacao_contas(
    caminho_template: str | Path,
    caminho_saida: str | Path,
    df_port: pd.DataFrame,
    df_composicao: pd.DataFrame,
    df_decomposicao: pd.DataFrame,
    *,
    portfolio_label: str = "Top Ações XP",
    reference_year: int | None = None,
    reference_month: int | None = None,
    today: date | pd.Timestamp | None = None,
) -> Path:
    """Gera a Prestação de Contas, preservando os textos editoriais.

    Requer Windows com Microsoft PowerPoint instalado.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "A Prestação de Contas requer Windows, PowerPoint e pywin32 instalado."
        ) from exc

    template = Path(caminho_template).resolve()
    output = Path(caminho_saida).resolve()
    if not template.exists():
        raise FileNotFoundError(f"Template não encontrado: {template}")
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, output)
    period = resolve_accountability_period(
        df_port,
        today=today,
        reference_year=reference_year,
        reference_month=reference_month,
    )

    pythoncom.CoInitialize()
    application = None
    presentation = None
    try:
        print("Prestação de Contas: iniciando PowerPoint...")
        application = win32com.client.DispatchEx("PowerPoint.Application")
        print("Prestação de Contas: PowerPoint iniciado...")
        application.DisplayAlerts = 1  # ppAlertsNone
        print("Prestação de Contas: abrindo template...")
        presentation = application.Presentations.Open(str(output), False, False, True)
        slides = presentation.Slides
        print("Prestação de Contas: atualizando datas e títulos...")
        _update_portfolio_titles(slides, portfolio_label)
        _update_dates(slides, period, portfolio_label)
        print("Prestação de Contas: atualizando tabela de desempenho...")
        _update_summary_table(slides(1), df_port, period, portfolio_label)
        print("Prestação de Contas: atualizando waterfall...")
        _replace_waterfall(slides(2), df_decomposicao)
        print("Prestação de Contas: atualizando composição...")
        _update_composition_table(slides(2), df_composicao)
        print("Prestação de Contas: atualizando gráfico base 100...")
        _update_base100_chart(slides(2), df_port, period)
        print("Prestação de Contas: salvando arquivo...")
        presentation.Save()
        return output
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if application is not None:
            try:
                application.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
