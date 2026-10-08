"""Smoke test of the levers: changes Assumptions in a private Excel instance (not saved) and prints what moves."""
from pathlib import Path
import win32com.client as w32
src = (Path(__file__).parent.parent / "Ibov_valuation_by_box_model_XP.xlsx").resolve()
xl = w32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(str(src), UpdateLinks=0, ReadOnly=True)
    A, S, H, C = (wb.Worksheets(n) for n in ("Assumptions", "Summary", "History", "Checks"))
    def show(tag):
        xl.CalculateFull()
        ib = [S.Range(f"{c}18").Text for c in "DEFG"]
        box = {S.Range(f"B{r}").Text: [S.Range(f"{c}{r}").Text for c in "DEFG"] for r in (24, 25, 26, 27)}
        hist = {H.Range(f"{c}5").Text: H.Range(f"{c}8").Text for c in "DEFGH"}
        print(f"\n== {tag}\n  Ibov avg FV (Bear/Base/Bull/Custom): {ib}\n  box upside: {box}\n  P/E 10y avg: {hist}\n  Checks: {C.Range('C5').Text}")
    show("default")
    A.Range("G16").Value = 0.10; A.Range("G18").Value = 0.5; A.Range("G19").Value = -100
    show("Custom = Selic 10%, k +0.5, Ke -100bp")
    A.Range("C36").Value = 1; A.Range("C32").Value = 2
    show("DropStale = 1 and P/E min 2x (dashboard rule)")
    A.Range("C36").Value = 0; A.Range("C32").Value = 1
    import datetime as dt
    A.Range("C23").Value = dt.datetime(2021, 1, 31)
    show("history window Jan-21..Sep-26 (ex 2019-20 re-rating)")
    A.Range("C23").Value = dt.datetime(2016, 10, 31)
    A.Range("E44").Value = 14.0
    show("override: Cyclicals Base target P/E = 14.0x")
    print("  overrides counted in Checks:", C.Range("C20").Text)
    wb.Close(SaveChanges=False)
finally:
    xl.Quit()
