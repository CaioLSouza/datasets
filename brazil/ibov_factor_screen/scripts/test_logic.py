"""Stub out ONLY the Bloomberg calls, keep every Excel formula, then check the maths in real Excel.
This verifies the aggregation, sector discovery and missing-data guards. It cannot verify BDP/BDH/BDS.
"""
import openpyxl, os, pythoncom, win32com.client as w32, pandas as pd, numpy as np

SRC = 'Ibovespa Factor Screen (BBG).xlsx'
TST = 'zz_test_stubbed.xlsx'
DR0 = 5

#            ticker   sector        shout   price  eps_now eps_3m eps_6m  trr3   trr6
T = [('PETR4 BZ','Energy',        13044.0, 38.50,  10.20,  9.80,  9.00,  12.5,  22.0),
     ('PETR3 BZ','Energy',         7442.0, 41.20,  10.20,  9.80,  9.00,  11.8,  21.0),
     ('PRIO3 BZ','Energy',          865.0, 44.10,   6.10,  6.40,  7.00,  -8.2, -15.5),
     ('ITUB4 BZ','Financials',     5286.0, 36.90,   4.10,  3.95,  3.80,   6.4,  14.2),
     ('BBDC4 BZ','Financials',     5238.0, 17.30,   2.05,  2.10,  2.20,  -2.1,   3.3),
     ('BBAS3 BZ','Financials',     5730.0, 28.40,   7.80,  7.20,  6.90,  15.0,  27.4),
     ('MGLU3 BZ','Retail',        66000.0,  9.15,   0.12,  0.09,  0.05,  33.0,  51.0),
     ('LREN3 BZ','Retail',         1010.0, 18.60,   1.40,  1.55,  1.62,  -9.5, -12.0),
     ('VALE3 BZ','Materials',      4300.0, 61.70,   8.90,  8.40,  8.10,   7.7,   9.9),
     ('ABEV3 BZ','Consumer',      15700.0, 13.05,   0.95,  0.95,  0.92,   1.2,   4.5),
     # edge cases -----------------------------------------------------------------
     ('WEGE3 BZ','Materials',      4200.0, 52.30,   1.80,  0.00,  1.55,   4.0,   None),
     ('XPTO3 BZ', None,            1000.0, 10.00,   1.00,  0.90,  0.80,   5.0,  10.0)]
#  WEGE3: eps_3m = 0  -> rev A must be blank, rev B must still compute; trr6 missing -> out of mom B only
#  XPTO3: no sector   -> must be excluded from every sector row AND from the index total

wb = openpyxl.load_workbook(SRC)
mem, mp, d = wb['Index Members'], wb['Sector Map'], wb['Data']
mem['A4'] = None
for i, row in enumerate(T):
    mem.cell(4 + i, 1, row[0])                       # stub the BDS spill
    if row[1] is not None:
        mp.cell(4 + i, 2, row[1])                    # fill the sector map
    r = DR0 + i
    d.cell(r, 4, row[2] * row[3])                    # D company mkt cap (stub)
    d.cell(r, 5, row[2])                             # E shares out (stub)
    d.cell(r, 6, row[3])                             # F price (stub)
    d.cell(r, 9, row[4])                             # I eps now (stub)
    d.cell(r,10, row[5])                             # J eps at A (stub)
    d.cell(r,11, row[6])                             # K eps at B (stub)
    # the live cells divide the BBG field by 100; stub the POST-division value
    d.cell(r,14, row[7]/100 if row[7] is not None else None)      # N momentum A
    d.cell(r,15, row[8]/100 if row[8] is not None else None)      # O momentum B
for i in range(len(T), 120):                         # clear unused rows' stub targets
    r = DR0 + i
    for c in (4,5,6,9,10,11,14,15): d.cell(r, c, None)
wb.save(TST)

# ---------------- expected values, computed independently -------------------------
df = pd.DataFrame(T, columns=['tk','sec','sh','px','e0','eA','eB','mA','mB'])
df['w'] = df.sh * df.px
df['revA'] = np.where(df.eA > 0, df.e0/df.eA - 1, np.nan)
df['revB'] = np.where(df.eB > 0, df.e0/df.eB - 1, np.nan)
df['momA'] = df.mA / 100.0
df['momB'] = df.mB.astype(float) / 100.0
ok = df.sec.notna()
def wavg(sub, col):
    m = sub[col].notna()
    return np.nan if sub.loc[m,'w'].sum() == 0 else (sub.loc[m,col]*sub.loc[m,'w']).sum()/sub.loc[m,'w'].sum()
exp_sec = {s: {c: wavg(g, c) for c in ('revA','revB','momA','momB')} | {'n': len(g), 'w': g.w.sum()}
           for s, g in df[ok].groupby('sec')}
exp_tot = {c: wavg(df[ok], c) for c in ('revA','revB','momA','momB')}
exp_tot['w'] = df[ok].w.sum(); exp_tot['n'] = int(ok.sum())

# ---------------- read Excel back ------------------------------------------------
pythoncom.CoInitialize()
xl = w32.gencache.EnsureDispatch('Excel.Application'); xl.Visible=False; xl.DisplayAlerts=False
try:
    bk = xl.Workbooks.Open(os.path.abspath(TST)); xl.Application.CalculateFullRebuild()
    ag, da = bk.Worksheets('Sector Aggregation'), bk.Worksheets('Data')
    nerr = 0
    for sh in bk.Worksheets:
        try:
            rg = sh.UsedRange.SpecialCells(-4123, 16)
        except Exception:
            continue                       # no error cells on this sheet
        nerr += rg.Count
        locs = []
        try:
            for ar in rg.Areas:
                for c in ar:
                    locs.append(c.Address(False, False) + '=' + str(c.Text))
                    if len(locs) >= 8: break
                if len(locs) >= 8: break
        except Exception as e:
            locs = [f'(could not enumerate: {e})']
        print(f'  ERRORS on "{sh.Name}": {rg.Count} -> {locs}')
    print(f'error cells (BBG calls are stubbed, so these are mine): {nerr}\n')

    print('--- per-ticker guards ---')
    for i in (10, 11):
        r = DR0 + i
        print(f'  {da.Cells(r,1).Value:9s} sector={str(da.Cells(r,2).Value):14s} '
              f'revA={da.Cells(r,12).Text:>8s} revB={da.Cells(r,13).Text:>8s} '
              f'momB={da.Cells(r,15).Text:>8s} flag={da.Cells(r,16).Value}')

    print('\n--- sector rows discovered (order of first appearance) ---')
    print(f'  {"sector":12s} {"n":>3s} {"weight":>14s}   revA        revB        momA        momB')
    worst = 0.0
    for r in range(4, 34):
        s = ag.Cells(r,1).Value
        if not s: continue
        got = {c: ag.Cells(r, 5+k).Value for k, c in enumerate(('revA','revB','momA','momB'))}
        e = exp_sec[s]
        line = f'  {s:12s} {ag.Cells(r,2).Value:3.0f} {ag.Cells(r,3).Value:14,.0f}  '
        for c in ('revA','revB','momA','momB'):
            gv, ev = got[c], e[c]
            if gv in ('', None) or (isinstance(ev, float) and np.isnan(ev)):
                line += f'{"blank":>11s} ' if (gv in ('',None) and np.isnan(ev)) else f'{"MISMATCH":>11s} '
            else:
                worst = max(worst, abs(gv-ev)); line += f'{gv*100:10.4f}% '
        assert abs(ag.Cells(r,3).Value - e['w']) < 1e-6, f'weight mismatch {s}'
        assert ag.Cells(r,2).Value == e['n'], f'count mismatch {s}'
        print(line)

    rt = 35
    print(f'\n--- index total (row {rt}) ---')
    print(f'  label={ag.Cells(rt,1).Value}  n={ag.Cells(rt,2).Value:.0f} (expected {exp_tot["n"]})  '
          f'weight={ag.Cells(rt,3).Value:,.0f} (expected {exp_tot["w"]:,.0f})')
    for k, c in enumerate(('revA','revB','momA','momB')):
        gv = ag.Cells(rt, 5+k).Value
        worst = max(worst, abs(gv-exp_tot[c]))
        print(f'  {c}: excel={gv*100:9.5f}%  python={exp_tot[c]*100:9.5f}%')
    print(f'\n  n columns (names actually behind each factor): '
          f'{[ag.Cells(rt,9+k).Value for k in range(4)]}')
    print(f'  coverage check: {ag.Cells(37,2).Value}')
    print(f'\nMAX |excel - python| across every weighted average: {worst:.3e}')
    bk.Close(SaveChanges=False)
finally:
    xl.Quit(); pythoncom.CoUninitialize()
