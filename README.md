# finance_report

This repo generates Excel reports for profit center analysis from monthly PL
and X-charge source files. The entry point is `src/profit_center_analysis.py`,
which reads source workbooks, filters by a target profit center, and writes
summary and per-month comparison reports with highlights.

## Main flow

- `src/profit_center_analysis.py` configures:
  - `base_directory`: root folder containing year subfolders with source files.
  - `years_to_process`: list of years to process (e.g., `[2025]`).
  - `target_pc`: profit center name key (e.g., `Analytics`).
- It loads data with:
  - `process_all_pl_files(...)` for `PL_*.xlsx` (sheet `pl_projects`).
  - `process_all_x_charge_files(...)` for
    `Secured Rev_Profit centers_*.xlsx`.
- It writes reports with:
  - `write_monthly_with_highlights(...)` to produce per-month sheets with
    NEW/MODIFIED/DELETED highlights.
  - `save_summary_with_format(...)` to write a summary report.
- Output file names are derived from `REPORT_PATH` in
  `src/utils/constants.py`, which places reports in the `report/` directory and
  ensures the folder exists.

## Inputs

Expected folder layout under `base_directory`:

```
<base_directory>/
  2024/
    PL_02 2024.xlsx
    Secured Rev_Profit centers_03.xlsx
  2025/
    PL_01 2025.xlsx
    Secured Rev_Profit centers_01.xlsx
```

Required sheets and columns:
- PL files: sheet `pl_projects` with a profit center column and finance fields.
- X-charge files: sheet `Secured Rev - Profit centers`.

Column mappings for different years are handled in
`src/utils/get_headers.py` via `pl_header(...)` and `secured_rev_header(...)`.

## Outputs

Reports are written to the `report/` directory with a timestamped file name:

```
report/
  finance_report [2025]-YYYY-MM-DD-HH-MM-SS.xlsx
  finance_report_PL_projects [2025]-YYYY-MM-DD-HH-MM-SS.xlsx
  finance_report_Xcharge_taker [2025]-YYYY-MM-DD-HH-MM-SS.xlsx
  finance_report_Xcharge_giver [2025]-YYYY-MM-DD-HH-MM-SS.xlsx
```

The per-month report includes a "New Record?" column and highlights:
- NEW: entire row green.
- MODIFIED: only changed cells yellow.
- DELETED: entire row red.

## Key modules

- `src/profit_center_analysis.py`: entry point and report orchestration.
- `src/utils/utils.py`: data extraction, comparison, and formatting logic.
- `src/utils/get_headers.py`: header mapping by year/month.
- `src/utils/constants.py`: report path and log path setup.
- `src/utils/service_utils.py`: directory creation helpers.

## Run

From the repo root:

```powershell
python src/profit_center_analysis.py
```
