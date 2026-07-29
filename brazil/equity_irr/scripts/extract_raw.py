"""Extract full-frequency IRR panels from the three source workbooks into pickles."""
import pandas as pd
import numpy as np
from pyxlsb import open_workbook
import pickle

# ---------------- INFRA (daily) ----------------
df = pd.read_excel('Infra_Historical_IRRs.xlsx', sheet_name='IRR Summary (Daily)', header=None)
dates = pd.to_datetime(df[0], errors='coerce')
infra_cols = {'RAIL3': 2, 'HBSA3': 5, 'STBP3': 8, 'CCRO3': 11, 'ECOR3': 14}
infra = pd.DataFrame({k: pd.to_numeric(df[c], errors='coerce') for k, c in infra_cols.items()})
infra['NTNB10'] = pd.to_numeric(df[17], errors='coerce')
infra['NTNB15'] = pd.to_numeric(df[18], errors='coerce')
infra.index = dates
infra = infra[dates.notna().values]
infra = infra.dropna(how='all')

# ---------------- UTILITIES (daily) ----------------
dfu = pd.read_excel('Utilities_Historical_IRR.xlsx', sheet_name='IRR Summary', header=None)
datesu = pd.to_datetime(dfu[0], errors='coerce')
util_cols = {
    'EQTL3': 2, 'ENGI11': 7, 'CMIG4': 12, 'CPLE3': 17, 'AXIA3': 22,
    'ALUP11': 28, 'ISAE4': 33, 'TAEE11': 38, 'AURE3': 43, 'EGIE3': 48,
    'CSMG3': 53, 'SBSP3': 58, 'CPFE3': 63, 'SAPR11': 69, 'LIGT3': 74,
    'ENEV3': 79, 'ORVR3': 84,
}
util = pd.DataFrame({k: pd.to_numeric(dfu[c], errors='coerce') for k, c in util_cols.items()})
util['NTNB10'] = pd.to_numeric(dfu[94], errors='coerce')
util['NTNB15'] = pd.to_numeric(dfu[95], errors='coerce')
util.index = datesu
util = util[datesu.notna().values]
util = util.dropna(how='all')

# ---------------- MALLS (weekly, xlsb) ----------------
with open_workbook('Income_Properties_Historical_IRR.xlsb') as wb:
    with wb.get_sheet('IRR Summary') as sheet:
        rows = list(sheet.rows())
data_rows = rows[23:]
def cv(r, c):
    for cell in r:
        if cell.c == c:
            return cell.v
    return None
epoch = pd.Timestamp('1899-12-30')
mall_cols = {'MULT3': 3, 'IGTI11': 8, 'ALOS3': 13, 'NTNB10': 23, 'NTNB15': 24}
recs = []
for r in data_rows:
    d = cv(r, 1)
    if not isinstance(d, (int, float)):
        continue
    rec = {'Date': epoch + pd.Timedelta(days=d)}
    for k, c in mall_cols.items():
        rec[k] = pd.to_numeric(pd.Series([cv(r, c)]), errors='coerce').iloc[0]
    recs.append(rec)
malls = pd.DataFrame(recs).set_index('Date')
malls = malls.dropna(how='all')

for name, d in [('infra', infra), ('util', util), ('malls', malls)]:
    print(name, d.shape, d.index.min().date(), d.index.max().date())
    print('  non-null:', {c: int(d[c].notna().sum()) for c in d.columns})

with open('raw_panels.pkl', 'wb') as f:
    pickle.dump({'infra': infra, 'util': util, 'malls': malls}, f)
