# Brazilian equity IRR — consolidated coverage

Historical **real (inflation-adjusted) implied equity IRR** for 25 listed Brazilian companies
across infrastructure & logistics, electric utilities, water & sanitation, and malls, alongside
the 10-year NTN-B real yield. Because both sides are real rates, the IRR and the NTN-B are
directly comparable and their difference is the implied equity risk premium of the coverage.

**File:** `equity_irr_consolidated_coverage.xlsx`

## Sheets

| Sheet | Contents |
|---|---|
| `Read me` | Methodology, caveats, colour key |
| `IRR Index (Monthly)` | Consolidated chain-linked IRR index, 90 months, vs NTN-B (+ chart) |
| `Company Summary` | Per-company current IRR, P10/P25/median/P75/P90, current percentile, spread |
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
in both months**, cumulated forward. A name entering or leaving never moves the level; only
genuine repricing does. The chain is anchored so the final month equals the actual simple
average of the names trading that month, then extended backwards.

Columns `Simple average` and `Simple median` are kept deliberately, so the composition jumps the
index avoids remain visible.

## Caveats

- **Percentile windows differ by sector.** Utilities percentiles rest on ~13 months of history;
  infrastructure on 3.5 years; malls on nearly 7. A high utilities percentile is a much weaker
  claim than the same number for a mall or infra name. Check the observation count and date
  range on `Company Summary` before quoting one.
- **Santos Brasil (STBP3)** ends in Nov-2025, consistent with the take-private. Its row is
  shaded and flagged, and it is excluded from every aggregate.
- **Snapshot dates differ** — 24-Jul-2026 (infra), 27-Jul-2026 (utilities), 22/29-May-2026
  (malls). The index is calendar-aligned, so its final month covers the 21 names still reporting.
- **Q2-2022 NTN-B is unreliable.** The only source that far back dips to ~3-4% for four months,
  against ~5.7% around it, so the index spread is not meaningful for Apr–Jul 2022. The IRR
  series themselves are unaffected.
- **Spreads are percentage points.** Formatted as `%`, but they are the difference between two
  rates: `+7.34%` means 7.34 p.p. over the bond.
- **Values, not live formulas.** Derived figures are written as computed values. The raw sheets
  carry the complete source series, so everything is reproducible and auditable.
- The source data contains **no dividend / growth / re-rating decomposition** and no
  exit-multiple sensitivity, so neither can be derived from this workbook.

## Source

Derived from XP Research historical IRR files. IRRs are the output of third-party sell-side
valuation models and are reproduced here as received; they are not independently verified.
Provided for research and educational purposes, not as investment advice.
