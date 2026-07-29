# Ibovespa factor screen — revisions, momentum, size, by sector

A **template, not a dataset.** The workbook contains no data: every number arrives from Bloomberg
when you open it with the add-in running. There are 3,740 formulas, 961 of them BDP/BDH/BDS calls,
and zero hardcoded values on the data sheets.

Per Ibovespa constituent it pulls 3m and 6m forward-EPS revisions, 3m and 6m total-return
momentum, and market cap — then aggregates into sectors by market-cap weighted average, using a
sector classification you supply.

**File:** `ibov_factor_screen.xlsx`

## What you do

1. Open with the Bloomberg add-in running. One `BDS(...,"INDX_MEMBERS")` call spills the member
   list; everything else keys off it.
2. On `Sector Map`, type your sector next to each ticker. The ticker column fills itself.
3. Read `Sector Aggregation`. It discovers your sector names from what you typed — there is no
   sector list to maintain. Type a new one and a row appears.

That is the whole workflow. Nothing else needs touching.

## Sheets

| Sheet | Contents |
|---|---|
| `Read me` | Method, field notes, caveats, colour key |
| `Inputs` | Index ticker, as-of date, the four windows, every field mnemonic, weighting basis |
| `Sector Map` | **The only sheet you fill in.** Ticker (automatic) + sector (yours) |
| `Data` | One row per constituent: the five metrics plus flags and hidden aggregation helpers |
| `Sector Aggregation` | Market-cap weighted average per sector, plus an index total |
| `Index Members` | The single BDS call that feeds everything |

## Two decisions worth knowing about

**Weighting uses security market cap, not company market cap.** Bloomberg's `CUR_MKT_CAP` is
reported at company level, so for a name with both an ON and a PN line in the index it returns the
same figure twice — weighting by it would count Petrobras or Itaú twice over. The default basis is
`EQY_SH_OUT × PX_LAST`, the cap of the individual listed line. Both are pulled and the
`Weighting basis` input switches between them.

**The fiscal period is pinned on the revisions.** A revision is forward EPS today divided by
forward EPS at the lookback date. `BEST_FPERIOD_OVERRIDE` defaults to `BF` (blended forward 12m),
which rolls continuously. Left unpinned, every annual fiscal roll would compare two different
years and the "revision" would be an artefact of the roll rather than a change in estimates.

## Missing data cannot poison an aggregate

A security enters a weighted average only where that factor, its market cap, and its sector are
all present. The consequences are deliberately narrow:

- A name missing one factor still contributes to the other three.
- A name with no sector is excluded from every sector row **and** from the index total.
- The `n (…)` columns on `Sector Aggregation` report how many names actually stood behind each
  figure, so a thin number is visible rather than implied.
- The `Data flag` column says, per ticker, exactly what is missing.

## One thing to do once

Column A of `Sector Map` is a formula reading the live index. The Ibovespa rebalances three times
a year, and when a name enters or leaves, that formula shifts every ticker below it while your
typed sectors stay where they are — misaligning the mapping silently.

After filling the sectors: select column A, copy, Paste Special → Values. The list becomes static,
the mapping is anchored to it permanently, and any future entrant shows up as `NOT IN MAP` on
`Data` so you can add a row deliberately.

## What was verified, and what was not

**Not verified: the Bloomberg calls.** This was built without terminal access, so not one of the
961 BDP/BDH/BDS formulas has ever returned a value. Treat the first open as a test run.

`CUST_TRR_RETURN_HOLDING_PER` is the field most likely to need adjustment — it is
entitlement-sensitive and its override names vary between setups. If it fails, put `CHG_PCT_3M`
in the total-return cell on `Inputs`, clear the date overrides, and accept price-only momentum
over a fixed window. Also check two scale conventions on the first pull: whether `EQY_SH_OUT` is
in millions on your terminal, and whether the return field gives `60` or `0.60` for 60%. The
workbook assumes millions and divides the return by 100.

Every field mnemonic lives in a cell on `Inputs` precisely because of this uncertainty — a wrong
mnemonic is a one-cell fix that propagates to all constituents, not a find-and-replace.

**Verified: all the Excel-side logic.** A copy was built with only the Bloomberg calls replaced by
synthetic values — deliberately including a zero base EPS, a missing momentum figure, and an
unmapped ticker — opened in Excel 16.0 via COM, fully rebuilt, and compared against an independent
Python computation. Result: **0 error cells**, and every weighted average matched to a maximum
difference of `5.6e-17`. The edge cases behaved as intended: the zero-EPS name dropped out of that
one revision only, and the unmapped name dropped out of everything. `scripts/test_logic.py`
reproduces this.

That test caught three defects, the worst being that `INDEX` over an empty cell returns `0` rather
than blank: the 108 unused rows were resolving to a ticker of `"0 Equity"`, producing 432 error
cells and a coverage warning that claimed 109 tickers lacked a sector instead of 1.

## Scripts

| Script | Purpose |
|---|---|
| `scripts/build_ibov.py` | Builds the workbook |
| `scripts/test_logic.py` | Stubs the Bloomberg calls, checks the Excel logic against Python in real Excel |

## Note on licensing

The workbook is distributed empty by design. Bloomberg data is licensed to the terminal user and
should not be redistributed, so do not commit a copy that has been saved with values in it — use
`scripts/build_ibov.py` to regenerate a clean one instead.
