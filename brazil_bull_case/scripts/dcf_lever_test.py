"""Smoke test of the DCF levers in a private Excel instance (not saved)."""
from pathlib import Path

import win32com.client as w32

src = (Path(__file__).parent.parent / "Ibov_DCF_by_box_XP.xlsx").resolve()
xl = w32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(str(src), UpdateLinks=0, ReadOnly=True)
    A, S, C = wb.Worksheets("Assumptions"), wb.Worksheets("Summary"), wb.Worksheets("Checks")

    def find(label):
        for r in range(1, 120):
            if str(A.Cells(r, 2).Value or "").startswith(label):
                return r

    rc, rb = find("Discount-rate convention"), find("ERP basis")
    rh, rl, rk = find("Valuation horizon"), find("Long-term ROE: 1"), find("ERP move in std")

    def show(tag):
        xl.CalculateFull()
        print(f"\n== {tag}   (upside Bear/Base/Bull | value Bear/Base/Bull)   checks: {C.Range('C5').Value}")
        for r in range(17, 22):
            print(f"  {S.Cells(r, 2).Value:12s}", [S.Cells(r, c).Text for c in (9, 10, 11)], [S.Cells(r, c).Text for c in (4, 5, 6)])
        print("  ERP used (Ibov level) Bear/Base/Bull:", [S.Cells(10, c).Text for c in (4, 5, 6)],
              " implied ERP Ibov Bear/Base/Bull:", [S.Cells(40, c).Text for c in (4, 5, 6)])

    show("default: house premises, 5 years, 12-month horizon, consensus ROE, growth -2/0/+2 pp, g 4.2/4.7/5.2%")
    A.Cells(rh, 3).Value = 0
    show("horizon 0 (value today)")
    A.Cells(rh, 3).Value = 1
    A.Cells(rl, 3).Value = 1
    show("long-term ROE = 10y average")
    A.Cells(rl, 3).Value = 2
    A.Cells(rc, 3).Value = 2
    show("Fisher convention")
    A.Cells(rc, 3).Value = 1
    A.Cells(rb, 3).Value = 3
    show("ERP basis 3 (each box's 10y average)")
    A.Cells(rb, 3).Value = 1
    for c, v in zip("DEF", (-1, 0, 1)):
        A.Range(f"{c}{rk}").Value = v
    show("ERP k -1 / 0 / +1")
    wb.Close(SaveChanges=False)
finally:
    xl.Quit()
