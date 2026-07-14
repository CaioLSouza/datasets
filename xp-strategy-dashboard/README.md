# xp-strategy-dashboard — data sources

Public data catalog for the XP equity-strategy dashboard. This folder hosts the
**sample datasets** and the **catalog** that maps each logical source to (a) its
corporate network path (used in production) and (b) a public GitHub raw URL (used
for development on a machine without corporate network access).

## Files

- `files_config.xlsx` — the original dictionary provided by the user
  (`caminho` = corporate UNC path, `link` = S3 sample, `descrição`).
- `catalog.csv` / `catalog.json` — machine-readable manifest the dashboard reads.
  One row per source with:
  - `key` — logical source name used across the code.
  - `prod_path` — corporate UNC path (`\\xpdocs\...`), the official source.
  - `prod_filetype` — file type in production (`parquet`, `xlsx`, `xlsm`, `csv`).
  - `sample_file` — file name under `samples/`.
  - `sample_filetype` — file type of the sample (samples are `xlsx`/`csv`, even
    when production is `parquet` — the loader must switch by mode).
  - `csv_sep` — separator for CSV sources (`,` or `;`). Empty for non-CSV.
  - `sheets` — relevant sheet names for Excel sources.
  - `github_raw_url` — public raw URL to pull the sample from this repo.
  - `description`.
- `samples/` — 15 small sample datasets (same schema as production, few rows).

## How the dashboard uses this

- `DATA_SOURCE=prod`   → read `prod_path` from the corporate network.
- `DATA_SOURCE=github` → read `github_raw_url` (default when off-network).

> ⚠️ For the `github` mode to work on a personal machine **without a token**,
> this repository must be **public**.
