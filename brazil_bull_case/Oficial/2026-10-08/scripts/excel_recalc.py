"""Recalculate Ibov_valuation_by_box_model.xlsx in a private Excel instance, save it with cached values, count formula
errors per sheet and print the Checks sheet and the Summary headline. Never touches other Excel instances."""
import sys
from pathlib import Path

import win32com.client as w32
from openpyxl.utils import get_column_letter as xl_col

ERR = {-2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!", -2146826252: "#NUM!",
       -2146826265: "#REF!", -2146826273: "#VALUE!"}
args = [a for a in sys.argv[1:] if not a.startswith("--")]
src = Path(args[0] if args else Path(__file__).parent.parent / "Ibov_valuation_by_box_model_XP.xlsx").resolve()
save = "--nosave" not in sys.argv
xl = w32.DispatchEx("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
try:
    wb = xl.Workbooks.Open(str(src), UpdateLinks=0)
    xl.CalculateFull()
    for ws in wb.Worksheets:
        ur = ws.UsedRange
        v = ur.Value
        r0, c0 = ur.Row, ur.Column
        errs = []
        if isinstance(v, tuple):
            for i, row in enumerate(v):
                for j, x in enumerate(row):
                    if isinstance(x, int) and x in ERR:
                        errs.append(f"{xl_col(c0 + j)}{r0 + i} {ERR[x]}")
        print(f"{ws.Name:12s} errors: {len(errs)}  {errs[:12]}")
    c = wb.Worksheets("Checks")
    print("\nChecks:", c.Range("C5").Text)
    for r in range(8, 80):
        lab = c.Cells(r, 2).Text
        if lab:
            print(f"  {lab[:78]:78s} | {c.Cells(r, 3).Text:>14s} | {c.Cells(r, 4).Text:>14s} | {c.Cells(r, 6).Text}")
    s = wb.Worksheets("Summary")
    print("\nSummary:")
    for r in range(14, 40):
        print("  " + " | ".join(s.Cells(r, k).Text for k in range(2, 13)))
    if save:
        wb.Save()
    try:
        wb.Close(SaveChanges=False)
    except Exception as e:                 # Excel sometimes rejects the call right after a save; the file is already saved
        print("note: Close raised", type(e).__name__, "- the workbook was saved before")
finally:
    xl.Quit()
