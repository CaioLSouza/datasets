"""Open the workbook in real Excel, force a full rebuild, hunt for errors, dump key values."""
import win32com.client as w32, os, json, sys, pythoncom

PATH = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else
                       'Equity IRR - Consolidated Coverage (Auto).xlsx')
XL_FORMULAS, XL_ERRORS = -4123, 16

pythoncom.CoInitialize()
xl = w32.gencache.EnsureDispatch('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False
out = {}
try:
    wb = xl.Workbooks.Open(PATH)
    xl.Application.CalculateFullRebuild()
    while xl.CalculationState != 0:
        pass

    # ---- 1. error sweep
    errs = {}
    for sh in wb.Worksheets:
        try:
            rng = sh.UsedRange.SpecialCells(XL_FORMULAS, XL_ERRORS)
            cells = [c.Address(False, False) + '=' + str(c.Text) for c in rng][:15]
            errs[sh.Name] = {'count': rng.Count, 'sample': cells}
        except Exception:
            pass
    out['errors'] = errs

    # ---- 2. formula census
    cens = {}
    for sh in wb.Worksheets:
        try:
            cens[sh.Name] = sh.UsedRange.SpecialCells(XL_FORMULAS).Count
        except Exception:
            cens[sh.Name] = 0
    out['formula_counts'] = cens

    # ---- 3. key values
    idx, sm, inp = wb.Worksheets('IRR Index (Monthly)'), wb.Worksheets('Company Summary'), wb.Worksheets('Inputs')
    out['inputs'] = {inp.Cells(r, 2).Value: inp.Cells(r, 3).Text for r in range(5, 10)}
    out['index_tail'] = [[str(idx.Cells(r, 1).Text)] +
                         [idx.Cells(r, c).Value for c in (2, 3, 5, 7, 8, 9)] + [idx.Cells(r, 10).Text]
                         for r in range(88, 94)]
    out['index_head'] = [[str(idx.Cells(r, 1).Text)] +
                         [idx.Cells(r, c).Value for c in (2, 3, 5, 7)] + [idx.Cells(r, 10).Text]
                         for r in range(4, 10)]
    out['summary'] = [[sm.Cells(r, 1).Value, sm.Cells(r, 4).Value] +
                      [sm.Cells(r, c).Value for c in (8, 11, 14, 15, 16, 17, 20)]
                      for r in range(4, 29)]
    out['totals'] = [[sm.Cells(r, 2).Value, sm.Cells(r, 8).Value] for r in range(30, 36)]

    # ---- 4. input responsiveness: flip tenor 10 -> 15 and re-read
    base_sp = sm.Range('Q4').Value
    base_ix = idx.Range('I93').Value
    inp.Range('C6').Value = 15
    xl.Application.CalculateFullRebuild()
    out['tenor15'] = {'RAIL3_spread': sm.Range('Q4').Value, 'index_spread': idx.Range('I93').Value}
    inp.Range('C6').Value = 10
    # ---- 5. exclusion responsiveness: drop LIGT3
    inp.Range('D28').Value = 'No'          # LIGT3 row on Inputs
    xl.Application.CalculateFullRebuild()
    out['excl_LIGT3'] = {'n': idx.Range('B93').Value, 'index': idx.Range('E93').Value,
                         'median_irr': sm.Range('H31').Value}
    inp.Range('D28').Value = 'Yes'
    # ---- 6. window responsiveness: widen to 2019
    inp.Range('C5').Value = '2019-02-01'
    xl.Application.CalculateFullRebuild()
    out['win2019'] = {'first_index': idx.Range('E4').Value, 'first_n': idx.Range('B4').Value,
                      'flag': idx.Range('J4').Text, 'RAIL3_pctile': sm.Range('N4').Value,
                      'RAIL3_months': sm.Range('O4').Value}
    inp.Range('C5').Value = '2022-12-01'
    xl.Application.CalculateFullRebuild()
    out['restored'] = {'spread_RAIL3': sm.Range('Q4').Value, 'index_spread': idx.Range('I93').Value,
                       'match_base': abs(sm.Range('Q4').Value - base_sp) < 1e-12
                                     and abs(idx.Range('I93').Value - base_ix) < 1e-12}
    wb.Close(SaveChanges=False)
finally:
    xl.Quit()
    pythoncom.CoUninitialize()

print(json.dumps(out, indent=1, default=str))
