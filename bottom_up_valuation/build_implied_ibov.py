"""
Monta a planilha `Implied Ibovespa.xlsx` (roda uma vez; só precisa rodar de novo se
quiser mudar o layout ou o caminho do raw).

- Power Query lê o bottom_up_raw.xlsx gerado pelo bottom_up_raw.py.
- Na própria tabela do query, colunas com BDP (preço e consenso) e a consolidação.
- Topo: implied target do Ibovespa, quebra por fonte e o filtro de idade dos modelos.

Uso:  python build_implied_ibov.py                         -> tudo na rede
      python build_implied_ibov.py <raw.xlsx> <saída.xlsx>   -> teste local
      python build_implied_ibov.py <raw.xlsx> <saída.xlsx> --rede
            -> monta com um raw local, mas grava o query apontando para a rede
"""
import sys
from pathlib import Path

import win32com.client as win32

RAW_PATH = r"\\xpdocs\Research\Equities\Estrategia\Banco de dados\Bottom up Valuation\output\bottom_up_raw.xlsx"
OUT_PATH = r"\\xpdocs\Research\Equities\Estrategia\Banco de dados\Bottom up Valuation\Implied Ibovespa.xlsx"
DEFAULT_MAX_AGE = 12          # meses

NAVY, LIGHT, YELLOW, GREY, ZEBRA = "1F2F44", "8EB3DF", "F9C113", "7E7E7E", "F2F2F2"
FONT = "Roboto Light"


def rgb(hex_):
    r, g, b = (int(hex_[i:i + 2], 16) for i in (0, 2, 4))
    return r + g * 256 + b * 65536


M_IBOV = '''let
    Source = Excel.Workbook(File.Contents("{path}"), null, true),
    Data = Source{{[Item="ibov",Kind="Table"]}}[Data],
    Typed = Table.TransformColumnTypes(Data, {{
        {{"Ticker", type text}}, {{"BBG Ticker", type text}}, {{"Weight", type number}},
        {{"Coverage", type text}}, {{"XP Ticker", type text}}, {{"XP BBG Ticker", type text}},
        {{"Name", type text}}, {{"Sector", type text}}, {{"Analyst", type text}},
        {{"Recommendation", type text}}, {{"Restricted", type logical}}, {{"Model Date", type date}},
        {{"XP Target", type number}}, {{"Composition Date", type date}}, {{"Run At", type datetime}}}})
in
    Typed'''

M_COV = '''let
    Source = Excel.Workbook(File.Contents("{path}"), null, true),
    Data = Source{{[Item="coverage",Kind="Table"]}}[Data],
    Typed = Table.TransformColumnTypes(Data, {{
        {{"PDATE", type date}}, {{"TARGET", type number}}, {{"KE", type number}}, {{"WACC", type number}},
        {{"RESTRICTED", type logical}}, {{"In Ibov", type logical}}, {{"Ibov Weight", type number}},
        {{"Model Age (months)", type number}}}})
in
    Typed'''

# colunas calculadas da tabela do Ibov: (nome, fórmula, formato, oculta?)
VALID = 'OR([@Recommendation]="Buy",[@Recommendation]="Neutral",[@Recommendation]="Sell")'
CALC = [
    ("Price", '=BDP([@[BBG Ticker]],"PX_LAST")', "#,##0.00", False),
    ("XP Ticker Price", '=IF([@Coverage]="Other class",BDP([@[XP BBG Ticker]],"PX_LAST"),[@Price])', "#,##0.00", True),
    ("Model Age (m)", '=IF(ISNUMBER([@[Model Date]]),(TODAY()-[@[Model Date]])/30.4375,"")', "0.0", False),
    ("XP Valid", f'=AND(ISNUMBER([@[XP Target]]),{VALID},[@Restricted]<>TRUE,'
                 'OR(NOT(AgeFilterOn),N([@[Model Age (m)]])<=MaxAgeMonths))', "General", False),
    ("XP TP", '=IF([@[XP Valid]],IF([@Coverage]="Direct",[@[XP Target]],'
              'IFERROR([@Price]*[@[XP Target]]/[@[XP Ticker Price]],"")),"")', "#,##0.00", False),
    ("Consensus TP", '=BDP([@[BBG Ticker]],"BEST_TARGET_PRICE")', "#,##0.00", False),
    ("TP Used", '=IF(ISNUMBER([@[XP TP]]),[@[XP TP]],IF(ISNUMBER([@[Consensus TP]]),[@[Consensus TP]],""))', "#,##0.00", False),
    ("Source", '=IF(ISNUMBER([@[XP TP]]),IF([@Coverage]="Direct","XP","XP (other class)"),'
               'IF(ISNUMBER([@[Consensus TP]]),"Consensus","No TP"))', "General", False),
    ("Upside", '=IF(AND(ISNUMBER([@[TP Used]]),ISNUMBER([@Price])),[@[TP Used]]/[@Price]-1,"")', "0.0%", False),
    ("Consensus Upside", '=IF(AND(ISNUMBER([@[Consensus TP]]),ISNUMBER([@Price])),[@[Consensus TP]]/[@Price]-1,"")', "0.0%", False),
    ("XP vs Consensus", '=IF(AND(ISNUMBER([@[XP TP]]),ISNUMBER([@[Consensus TP]])),[@[XP TP]]/[@[Consensus TP]]-1,"")', "0.0%", False),
    ("Contribution (p.p.)", '=IF(ISNUMBER([@Upside]),[@Weight]*[@Upside]/SUM([W Covered])*100,"")', "0.00", False),
    ("W Covered", '=IF(ISNUMBER([@Upside]),[@Weight],0)', "General", True),
    ("W x Upside", '=IF(ISNUMBER([@Upside]),[@Weight]*[@Upside],0)', "General", True),
    ("W Cons Covered", '=IF(ISNUMBER([@[Consensus Upside]]),[@Weight],0)', "General", True),
    ("W x Cons Upside", '=IF(ISNUMBER([@[Consensus Upside]]),[@Weight]*[@[Consensus Upside]],0)', "General", True),
]
QUERY_FMT = {"Weight": "0.00%", "XP Target": "#,##0.00", "Model Date": "dd/mm/yyyy",
             "Composition Date": "dd/mm/yyyy", "Run At": "dd/mm/yyyy hh:mm"}
HIDE_QUERY = {"BBG Ticker", "XP BBG Ticker", "Composition Date", "Run At"}

TABLE_ROW = 22   # linha do cabeçalho da tabela principal
COL0 = 2         # coluna B


def xp_table_style(wb):
    ts = wb.TableStyles.Add("XP Table")
    ts.ShowAsAvailableTableStyle = True
    h = ts.TableStyleElements(1)                   # xlHeaderRow
    h.Interior.Color = rgb(NAVY); h.Font.Color = rgb("FFFFFF"); h.Font.Bold = True
    ts.TableStyleElements(5).Interior.Color = rgb(ZEBRA)   # xlRowStripe1
    return ts


def load_query(wb, ws, name, m, cell, table_name):
    wb.Queries.Add(name, m)
    conn = (f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
            f'Location={name};Extended Properties=""')
    lo = ws.ListObjects.Add(0, conn, True, 1, ws.Range(cell))
    qt = lo.QueryTable
    qt.CommandType = 2
    qt.CommandText = [f"SELECT * FROM [{name}]"]
    qt.RowNumbers = False
    qt.PreserveFormatting = True
    qt.AdjustColumnWidth = False
    qt.BackgroundQuery = False
    qt.RefreshStyle = 1                            # xlInsertDeleteCells
    lo.Name = table_name
    qt.Refresh(False)
    return lo


# O Excel desta máquina lê NumberFormat via COM com separador trocado ("0.0%" vira "#.#00%").
# Então gravo um marcador de texto e troco pelo formato real direto no styles.xml depois de salvar.
FORMATS = []


def nf(rng, fmt):
    if fmt == "General":
        return
    if fmt not in FORMATS:
        FORMATS.append(fmt)
    rng.NumberFormat = f'"XPFMT{FORMATS.index(fmt)}"'


def fix_formats(path):
    import shutil, tempfile, zipfile
    from xml.sax.saxutils import escape
    tmp = Path(tempfile.mkdtemp()) / "x.xlsx"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                x = data.decode("utf-8")
                for i, f in enumerate(FORMATS):
                    x = x.replace(f'formatCode="&quot;XPFMT{i}&quot;"', f'formatCode="{escape(f, {chr(34): "&quot;"})}"')
                assert "XPFMT" not in x, "marcador de formato não substituído"
                data = x.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)


def label(ws, cell, text, bold=False, size=9, color=None):
    r = ws.Range(cell)
    r.Value = text
    r.Font.Bold = bold
    r.Font.Size = size
    if color:
        r.Font.Color = rgb(color)
    return r


def build(raw_path: str, out_path: str, query_path: str | None = None):
    """raw_path: arquivo usado para montar; query_path: caminho gravado no Power Query (default = raw_path)."""
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Add()
        while wb.Worksheets.Count > 1:
            wb.Worksheets(wb.Worksheets.Count).Delete()
        ws = wb.Worksheets(1)
        ws.Name = "Implied Ibovespa"
        cv = wb.Worksheets.Add(After=ws)
        cv.Name = "Coverage"
        for sh in (ws, cv):
            sh.Cells.Font.Name = FONT
            sh.Cells.Font.Size = 9
        style = xp_table_style(wb)

        # ---- tabela principal (query + colunas calculadas) --------------------
        # parâmetros primeiro: as fórmulas da tabela dependem dos nomes
        ws.Range("J5").Value = False
        wb.Names.Add("AgeFilterOn", "='Implied Ibovespa'!$J$5")
        ws.Range("I6").Value = DEFAULT_MAX_AGE
        wb.Names.Add("MaxAgeMonths", "='Implied Ibovespa'!$I$6")

        lo = load_query(wb, ws, "ibov_raw", M_IBOV.format(path=raw_path),
                        ws.Cells(TABLE_ROW, COL0).Address, "tbl_ibov")
        for name, f, fmt, hidden in CALC:
            lc = lo.ListColumns.Add()
            lc.Name = name
        for name, f, fmt, hidden in CALC:       # fórmulas depois: umas referenciam colunas à frente
            lc = lo.ListColumns(name)
            lc.DataBodyRange.Formula = f
            nf(lc.DataBodyRange, fmt)
            if hidden:
                lc.Range.EntireColumn.Hidden = True
        for name, fmt in QUERY_FMT.items():
            nf(lo.ListColumns(name).DataBodyRange, fmt)
        for name in HIDE_QUERY:
            lo.ListColumns(name).Range.EntireColumn.Hidden = True
        lo.TableStyle = style
        lo.ShowTableStyleRowStripes = True
        for i in range(1, lo.ListColumns.Count + 1):
            c = lo.ListColumns(i)
            if not c.Range.EntireColumn.Hidden:
                c.Range.EntireColumn.ColumnWidth = max(9, min(22, len(c.Name) + 3))
        lo.ListColumns("Name").Range.EntireColumn.ColumnWidth = 18
        lo.ListColumns("Sector").Range.EntireColumn.ColumnWidth = 20
        lo.ListColumns("Analyst").Range.EntireColumn.ColumnWidth = 18
        lo.HeaderRowRange.WrapText = True
        lo.HeaderRowRange.RowHeight = 24
        lo.HeaderRowRange.VerticalAlignment = -4108

        # ---- topo -------------------------------------------------------------
        ws.Columns("A").ColumnWidth = 2
        label(ws, "B2", "Ibovespa — Bottom-up Implied Target Price", bold=True, size=13, color=NAVY)
        label(ws, "B3", "XP coverage target prices; Bloomberg consensus (BEST_TARGET_PRICE) where XP has no valid target.",
              color=GREY)

        rows = [
            ("Ibovespa", '=BDP("IBOV Index","PX_LAST")', "#,##0"),
            ("Implied target", '=IFERROR(C5*(1+C7),"")', "#,##0"),
            ("Upside", '=IFERROR(SUM(tbl_ibov[W x Upside])/SUM(tbl_ibov[W Covered]),"")', "0.0%"),
            ("Implied target — consensus only", '=IFERROR(C5*(1+C9),"")', "#,##0"),
            ("Upside — consensus only", '=IFERROR(SUM(tbl_ibov[W x Cons Upside])/SUM(tbl_ibov[W Cons Covered]),"")', "0.0%"),
            ("Index weight without TP", '=SUMIFS(tbl_ibov[Weight],tbl_ibov[Source],"No TP")', "0.0%"),
            ("Composition date", '=MAX(tbl_ibov[Composition Date])', "dd/mm/yyyy"),
            ("Raw data run", '=MAX(tbl_ibov[Run At])', "dd/mm/yyyy hh:mm"),
        ]
        for i, (lab, f, fmt) in enumerate(rows):
            r = 5 + i
            label(ws, f"B{r}", lab)
            c = ws.Range(f"C{r}")
            c.Formula = f
            nf(c, fmt)
            c.HorizontalAlignment = -4152
        for r in (6, 7):                               # implied target em destaque
            ws.Range(f"B{r}:C{r}").Font.Bold = True
        ws.Range("B6:C6").Font.Size = 11
        ws.Range("B6:C6").Font.Color = rgb(NAVY)
        b = ws.Range("B6:C6").Borders(8)               # xlEdgeTop
        b.LineStyle = 1; b.Color = rgb(NAVY)
        ws.Range("B11:C12").Font.Color = rgb(GREY)
        ws.Columns("B").ColumnWidth = 28
        ws.Columns("C").ColumnWidth = 15

        # filtro de idade: checkbox + célula de input logo à direita dele
        label(ws, "E5", "Model age filter", bold=True, color=NAVY)
        cb = ws.CheckBoxes().Add(ws.Range("E6").Left, ws.Range("E6").Top - 1, 175, 14)
        cb.Caption = "Exclude XP models older than (months):"
        right = cb.Left + cb.Width
        k = 5
        while ws.Cells(6, k).Left < right:
            k += 1
        inp, link = ws.Cells(6, k), ws.Cells(5, k)
        ws.Range("J5").ClearContents(); ws.Range("I6").ClearContents()   # posições provisórias
        link.Value = False
        nf(link, ";;;")
        inp.Value = DEFAULT_MAX_AGE
        wb.Names("AgeFilterOn").RefersTo = f"='Implied Ibovespa'!{link.Address}"
        wb.Names("MaxAgeMonths").RefersTo = f"='Implied Ibovespa'!{inp.Address}"
        cb.LinkedCell = f"'Implied Ibovespa'!{link.Address}"
        cb.Value = -4146                                # xlOff
        inp.Interior.Color = rgb("FFF2CC")
        inp.Font.Bold = True
        inp.HorizontalAlignment = -4108
        inp.Validation.Add(Type=2, AlertStyle=1, Operator=7, Formula1="0")   # decimal >= 0
        st = ws.Range("E7")
        st.Formula = ('=IF(AgeFilterOn,"ON — models older than "&MaxAgeMonths&" months go to consensus ("'
                      '&COUNTIFS(tbl_ibov[Model Age (m)],">"&MaxAgeMonths,tbl_ibov[Coverage],"<>None")&" models)",'
                      '"OFF — all XP models used")')
        st.Font.Color = rgb(GREY)
        label(ws, "E8", "Under Review, restricted and n.a. recommendations always go to consensus.", color=GREY)
        label(ws, "E9", "Workflow: run bottom_up_raw.py → Data › Refresh All → Bloomberg › Refresh Workbook.", color=GREY)

        # quebra por fonte
        r0 = 14
        heads = ["Source", "# Stocks", "Weight", "Avg Upside", "Contrib. (p.p.)"]
        for j, h in enumerate(heads):
            ws.Cells(r0, COL0 + j).Value = h
        hr = ws.Range(ws.Cells(r0, COL0), ws.Cells(r0, COL0 + len(heads) - 1))
        hr.Interior.Color = rgb(NAVY); hr.Font.Color = rgb("FFFFFF"); hr.Font.Bold = True
        srcs = ["XP", "XP (other class)", "Consensus", "No TP"]
        for i, s in enumerate(srcs):
            r = r0 + 1 + i
            ws.Cells(r, 2).Value = s
            ws.Cells(r, 3).Formula = f'=COUNTIFS(tbl_ibov[Source],B{r})'
            ws.Cells(r, 4).Formula = f'=SUMIFS(tbl_ibov[Weight],tbl_ibov[Source],B{r})'
            ws.Cells(r, 5).Formula = f'=IFERROR(SUMIFS(tbl_ibov[W x Upside],tbl_ibov[Source],B{r})/SUMIFS(tbl_ibov[W Covered],tbl_ibov[Source],B{r}),"")'
            ws.Cells(r, 6).Formula = f'=IFERROR(SUMIFS(tbl_ibov[W x Upside],tbl_ibov[Source],B{r})/SUM(tbl_ibov[W Covered])*100,"")'
            if i % 2:
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Interior.Color = rgb(ZEBRA)
        rt = r0 + 1 + len(srcs)
        ws.Cells(rt, 2).Value = "Total"
        ws.Cells(rt, 3).Formula = f"=SUM(C{r0 + 1}:C{rt - 1})"
        ws.Cells(rt, 4).Formula = f"=SUM(D{r0 + 1}:D{rt - 1})"
        ws.Cells(rt, 5).Formula = "=C7"
        ws.Cells(rt, 6).Formula = f"=SUM(F{r0 + 1}:F{rt - 1})"
        tr = ws.Range(ws.Cells(rt, 2), ws.Cells(rt, 6))
        tr.Font.Bold = True
        tr.Borders(8).LineStyle = 1; tr.Borders(8).Color = 0
        nf(ws.Range(f"C{r0 + 1}:C{rt}"), "0")
        nf(ws.Range(f"D{r0 + 1}:D{rt}"), "0.0%")
        nf(ws.Range(f"E{r0 + 1}:E{rt}"), "0.0%")
        nf(ws.Range(f"F{r0 + 1}:F{rt}"), "0.00")
        ws.Range(ws.Cells(r0, 3), ws.Cells(rt, 6)).HorizontalAlignment = -4152

        # ---- aba Coverage ------------------------------------------------------
        cv.Columns("A").ColumnWidth = 2
        label(cv, "B2", "XP Coverage — raw target prices", bold=True, size=13, color=NAVY)
        lc = load_query(wb, cv, "coverage_raw", M_COV.format(path=raw_path), "$B$4", "tbl_coverage")
        lc.TableStyle = style
        for name, fmt in {"PDATE": "dd/mm/yyyy", "TARGET": "#,##0.00", "KE": "0.0%", "WACC": "0.0%",
                          "Ibov Weight": "0.00%", "Model Age (months)": "0.0"}.items():
            nf(lc.ListColumns(name).DataBodyRange, fmt)
        for i in range(1, lc.ListColumns.Count + 1):
            c = lc.ListColumns(i)
            c.Range.EntireColumn.ColumnWidth = max(9, min(22, len(c.Name) + 3))

        for sh in (ws, cv):
            sh.Activate()
            xl.ActiveWindow.DisplayGridlines = False
            xl.ActiveWindow.Zoom = 100
        ws.Activate()
        ws.Range("C1").Select()
        xl.ActiveWindow.FreezePanes = True
        ws.Range("A1").Select()

        if query_path and query_path != raw_path:
            wb.Queries("ibov_raw").Formula = M_IBOV.format(path=query_path)
            wb.Queries("coverage_raw").Formula = M_COV.format(path=query_path)

        Path(out_path).parent.mkdir(parents=True, exist_ok=True)
        wb.SaveAs(str(Path(out_path)), FileFormat=51)
        wb.Close(False)
    finally:
        xl.Quit()
    fix_formats(out_path)
    print(f"gravado: {out_path}")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--rede"]
    if len(args) == 2:
        build(args[0], args[1], RAW_PATH if "--rede" in sys.argv else None)
    else:
        build(RAW_PATH, OUT_PATH)
