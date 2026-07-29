# Brazilian equity IRR — consolidated coverage

Historical **real (inflation-adjusted) implied equity IRR** for 25 listed Brazilian companies
across infrastructure & logistics, electric utilities, water & sanitation, and malls, alongside
the 10-year NTN-B real yield. Because both sides are real rates, the IRR and the NTN-B are
directly comparable and their difference is the implied equity risk premium of the coverage.

## Files

| File | What it is |
|---|---|
| `equity_irr_consolidated_coverage_auto.xlsx` | **Start here.** Formula-driven. An `Inputs` sheet parameterises the whole model — analysis window, NTN-B tenor, expected inflation, hurdle rate, and a per-company include switch. 6,356 formulas; the only hardcoded numbers are the three raw extracts. |
| `equity_irr_consolidated_coverage.xlsx` | The same analysis with every derived figure written as a static value. Useful if you just want the numbers without a recalculating model. |
| `scripts/` | `extract_raw.py` pulls the three source files into panels, `build_auto.py` builds the workbook, `verify_excel.py` opens it in Excel and checks it. |

## Sheets

| Sheet | Contents |
|---|---|
| `Read me` | Methodology, caveats, colour key |
| `Inputs` | *(auto version only)* The five parameters and the 25 include switches |
| `IRR Index (Monthly)` | Consolidated chain-linked IRR index, 90 months, vs NTN-B (+ chart) |
| `Company Summary` | Per-company spot IRR, P10/P25/median/P75/P90, current percentile, spread |
| `Panel (Monthly)` | Monthly average IRR, 25 companies × 90 months |
| `Chain-link Deltas` | Month-over-month change per company (paired months only) |
| `Infra (Daily)` | Raw daily series, 5 names, Dec-2022 → Jul-2026 |
| `Utilities (Daily)` | Raw daily series, 17 names, Jul-2025 → Jul-2026 |
| `Malls (Weekly)` | Raw weekly series, 3 names, Feb-2019 → May-2026 |

## The index, and why it is chain-linked

Coverage does not start at the same date: malls from Feb-2019, infrastructure from Dec-2022,
utilities only from Jul-2025. A plain cross-sectional average over whoever happens to be
available jumps whenever the roster changes — in Jul-2025 the simple average leaps from 12.8%
to 13.9% purely because 17 utilities appear, not because anything repriced.

The index therefore moves by the **average month-over-month change of only the companies present
in both months**, cumulated. A name entering or leaving never moves the level; only genuine
repricing does. The chain is anchored so the final month equals the actual simple average of the
names trading that month, then extended backwards.

Columns `Simple average` and `Simple median` are kept deliberately, so the composition jumps the
index avoids remain visible.

### Reading the gap between the index and the simple average

With a **fixed roster the two move in perfect lockstep** and the gap between them is exactly
constant — for an identical set of companies, "the average of the individual changes" and "the
change in the average" are the same quantity. Verified: across the 85 months where composition
does not change, the gap moves by `0.0000000000` p.p.

The gap therefore steps **only** at composition changes, and it is a level offset carried forward,
not a divergence happening at that date:

| Month | Roster | Event | Step | Cumulative gap |
|---|---:|---|---:|---:|
| Feb-2019 | 0→2 | series begins (MULT3, IGTI11) | +5.082 | +5.082 |
| Dec-2022 | 2→7 | the 5 infra names arrive | −3.381 | +1.701 |
| Jan-2024 | 7→8 | ALOS3 arrives | −0.107 | +1.594 |
| Jul-2025 | 8→25 | the 17 utilities arrive | −1.296 | +0.298 |
| Dec-2025 | 25→24 | STBP3 leaves | −0.075 | +0.223 |
| Jun-2026 | 24→21 | the 3 malls leave | −0.223 | 0.000 |

The gap at any past date measures **how much the coverage of that date differs from today's**,
in level terms. It closes to exactly zero at the anchor month by construction.

## How much of the history can be trusted

Propagating the standard error of each monthly step backwards from the anchor gives a band on the
index level:

| Date | Names carrying the link | Index | Band (±1.96 s.e.) |
|---|---:|---:|---:|
| Feb-2019 | 0 | 12.03% | ±7.75 p.p. |
| Jun-2022 | 2 | 15.24% | ±4.10 p.p. |
| Jun-2023 | 7 | 11.97% | ±3.50 p.p. |
| Jun-2025 | 8 | 14.41% | ±1.27 p.p. |
| Dec-2025 | 24 | 12.03% | ±0.82 p.p. |
| Jul-2026 | 21 | 12.22% | ±0.25 p.p. |

**Before Dec-2022 the level does not stand up.** For 46 months — more than half the series — the
entire chain is carried by two mall stocks, and a ±7.75 p.p. band around the Feb-2019 level is
noise rather than a measurement. In that stretch the *shape* is not reliable either, because the
shape simply is the price action of MULT3 and IGTI11.

For this reason the auto workbook **defaults its analysis window to 2022-12-01**. Set it to
`2019-02-01` on the `Inputs` sheet to see the full series, but do not quote the early level.

The companies are not a random sample of a population, so this band is a heuristic gauge of
precision rather than a formal confidence interval. It is good for orders of magnitude, which is
what the decision here turns on.

## Where the NTN-B comes from

It is not an external series. Each source file carries its own NTN-B columns, and those are what
is used here:

| Source | Columns | Label in the file |
|---|---|---|
| Infra | `R` / `S` | `10-year NTN-B` / `15-year NTN-B` |
| Utilities | `CQ` / `CR` | `10-year NTN-B` / `15-year NTN-B` |
| Malls | `X` / `Y` | `10-year NTN-B (2035) Yield` / `15-year NTN-B (2040) Yield` |

On `Company Summary` each company is paired with the NTN-B from **its own file**, as of its own
last date, so the spread is measured on matching dates. On the index sheet the precedence is
infra, then utilities, then malls.

## Caveats

- **Percentile windows differ by sector.** Utilities percentiles rest on ~13 months of history;
  infrastructure on 3.5 years; malls on nearly 7. A high utilities percentile is a much weaker
  claim than the same number for a mall or infra name. Check the observation count and date
  range on `Company Summary` before quoting one.
- **Santos Brasil (STBP3)** ends in Nov-2025, consistent with the take-private. It ships as
  `Include = No` in the auto workbook and is excluded from every aggregate in both.
- **Snapshot dates differ** — 24-Jul-2026 (infra), 27-Jul-2026 (utilities), 22/29-May-2026
  (malls). The index is calendar-aligned, so its final month covers the 21 names still reporting.
- **The three files disagree on the NTN-B.** Across the 389 days where infra and utilities
  overlap they never once agree: utilities runs on average **+0.33 p.p.** above infra (range
  +0.16 to +0.65). The malls file pins a specific bond (NTN-B 2035) while the other two say only
  "10-year", which would explain the drift, though none of the files states its method.
  *Practical impact is small.* At the current snapshot dates the gap is **+0.15 p.p.** for
  utilities and +0.04 to +0.08 p.p. for malls. Re-pricing every name against a single NTN-B
  source moves **no company more than one rank**, changes **no sign** (22 of 24 spreads stay
  positive), and is about 1% of the 14.9 p.p. spread between the widest and narrowest name.
  Within a sector, comparisons are entirely unaffected.
- **The malls NTN-B is wrong from Apr-2022 to Apr-2023.** It reads 3.3–4.8% for those 13 months
  against ~5.7% before and after. Where infra overlaps, this is demonstrably wrong rather than
  merely odd: Dec-2022 reads 3.76% in malls against 6.06% in infra. For the index the damage is
  confined to **Apr–Nov 2022**, the window where malls is the only available source; from
  Dec-2022 the index switches to infra. This is outside the default analysis window. The IRR
  series themselves are unaffected.
- **Spreads are percentage points.** Formatted as `%`, but they are the difference between two
  rates: `+7.34%` means 7.34 p.p. over the bond.
- **Excluding a company also clears its statistics.** Setting `Include = No` blanks that name's
  column in the monthly panel, so its percentiles and median disappear along with its
  contribution to the index. This is why the Santos Brasil row shows no distribution.
- The source data contains **no dividend / growth / re-rating decomposition** and no
  exit-multiple sensitivity, so neither can be derived from this workbook.

## Verification

The auto workbook was opened in Excel 16.0 through COM, fully rebuilt, and checked:

- **0 error cells** across 6,356 formulas.
- **Cross-validated end-to-end.** With the window widened to Feb-2019 and all 25 names included,
  the spreadsheet reproduces an independent Python implementation of the same chain-link to a
  maximum difference of `1.1e-16` over all 90 months — identical to floating-point precision.
  Spot IRRs match the raw source to `5.6e-17`.
- Input responsiveness was exercised: switching the NTN-B tenor, excluding a company, and
  widening the window all propagate correctly, and restoring the defaults returns the original
  values exactly.

One bug was found and fixed this way: pairing the NTN-B with `INDEX`/`MATCH` returned `0.00%` for
IGTI11, because the yield cell on that company's own last row happens to be empty and `INDEX`
over an empty cell returns zero rather than the last available value. It now uses a lookup that
requires a non-empty yield, which restored the correct +2.87 p.p. spread.

## Source

Derived from XP Research historical IRR files. IRRs are the output of third-party sell-side
valuation models and are reproduced here as received; they are not independently verified.
Provided for research and educational purposes, not as investment advice.
