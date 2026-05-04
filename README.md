# JMeter JSON to Excel Report Generator

This Streamlit app converts JMeter static `statistics.json` files into Excel reports.


## Final Refinements Included

- `Insights` includes **Top 10 Error APIs** table and chart.
- `APIs` highlights only response-time cells that breach SLA, not the full row.
- `Track_Comparison` excludes any track containing `Select customer`.

## Output Sheets

- `Insights` - KPI summary and charts
- `Track_Comparison` - side-by-side track comparison for all uploaded JSON files
- `Transactions` - only transaction rows starting with `T01`, `T02`, etc.
- `Errors` - rows where `errorCount > 0`
- `APIs` - non-transaction API rows only; original `transaction` column removed
- `Comparison` - raw API comparison when two or more JSON files are uploaded

## Track_Comparison Logic

Track = first part of transaction before `/`.

For each track, the sheet shows three metric rows:
- `Avg` uses `meanResTime`
- `Min` uses `minResTime`
- `Max` uses `maxResTime`

Percentages use **API count**, not sample count.

### Buckets

For tracks where Feature starts with `AskAI`:
- `0-10s`
- `10-20s`
- `20-30s`
- `>30s`

For all other tracks:
- `0-2s`
- `3-4s`
- `4-6s`
- `>6s`

Each uploaded JSON file appears as a side-by-side block with:
- Bucket 1 %
- Bucket 2 %
- Bucket 3 %
- Bucket 4 %
- Max Seconds

## SLA Logic

- AskAI Feature: SLA `< 10 sec`
- Assets, Assessments, Home, Settings and Support Features: SLA `< 2 sec`

PASS/FAIL in the `APIs` sheet is based on Avg Response Time in seconds.

## Run Locally

```bash
python3 -m venv venv
source venv/bin/activate
python -m pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Cloud

1. Replace your GitHub repo files with:
   - `app.py`
   - `main.py`
   - `requirements.txt`
   - `README.md`

2. Commit and push to GitHub.

3. In Streamlit Cloud, redeploy your app.

## Team Usage

1. Open the Streamlit URL.
2. Upload one JSON for normal report.
3. Upload two or more JSONs for side-by-side comparison.
4. Download Excel.


## Latest v13 Updates

- Streamlit UI wording updated for SLA rules and track buckets.
- `Track_Comparison` now includes embedded charts for top slow tracks.
- API percentage calculation remains API-count based.


## v14 Final Polish

- Removed M, N, O and P SLA-helper columns from the `APIs` sheet.
- `APIs` sheet highlights only breaching response-time cells, not full rows.
- `Track_Comparison` removes `Total` and any track containing `Select customer`.
- `Track_Comparison` charts also exclude `Total` and `Select customer`.
- Insights chart sizing/placement improved to reduce title/graph collisions.


## v15 Comparison-Focused Update

When two or more JSON files are uploaded, the generated workbook contains only:
- `Insights`
- `Track_Comparison`
- `APIs_Comparison`

`APIs_Comparison` shows Feature, Scenario and Endpoint separately for every uploaded report, plus side-by-side API metrics and baseline-vs-latest diff columns.


## v16 Final Fixes

- APIs sheet removes M, N, O, P SLA-helper columns.
- APIs sheet highlights only response-time cells that breach SLA.
- Insights charts use external cell titles to avoid title/graph collision.
- Top 10 Error APIs excludes `Total//` and any `Select customer` rows.
- Track_Comparison and its charts exclude `Total` and any track containing `Select customer`.


## v17 Corrections

- `APIs` sheet defensively removes SLA helper columns M, N, O and P.
- `Track_Comparison` Max Seconds is now metric-specific:
  - Avg row uses max Avg response time for that track.
  - Min row uses max Min response time for that track.
  - Max row uses max Max response time for that track.
- `Total` and `Select customer` remain excluded from Track_Comparison and its charts.


## v18 Chart and UI Polish

- Removed red highlighting from `Max Seconds` columns in `Track_Comparison`.
- Added value labels to Insights and Track_Comparison charts.
- Removed x/y axis titles from charts to avoid label collisions.
- Added a custom gradient background and styled upload/download controls in Streamlit UI.


## v19 Axis Titles and Samples

- Added `Total Samples` count in Insights.
- API SLA pie chart displays values and percentages.
- Restored x/y axis titles with larger chart sizes and spacing to avoid collisions.
- Track_Comparison charts include values and x/y axis titles with extra spacing.


## v20 Visible Chart Values

- Added Total Samples in Insights.
- SLA chart source table now includes PASS/FAIL count and percent.
- API SLA pie chart displays values and percentages.
- Restored X/Y axis titles with larger chart dimensions and more spacing.
- Track_Comparison charts include visible values and axis titles with more spacing.


## v21 Chart Titles and Visible Values

- API SLA pie chart now has an internal title again.
- Added visible PASS/FAIL/TOTAL count and percentage summary next to the SLA pie chart.
- Top 10 Slow APIs, Top 10 Error APIs, and Track_Comparison charts have internal titles.
- Chart sizes and positions increased to reduce title/plot collision.


## v22 Next-Level Dashboard

- Rebuilt Insights as an executive dashboard with KPI cards, health score, SLA breakdown, ranked slow APIs, ranked error APIs, and top slow tracks.
- Charts now use short rank labels to avoid long API-name collisions.
- Full API/track names are shown in readable side tables next to each chart.
- Track_Comparison charts now use rank-based labels and side tables for readability.


## v23 Clean Track Tables

- Removed Track charts from `Insights`.
- Removed Track charts from `Track_Comparison`.
- Added clean Track-wise summary tables instead.
- Removed cluttered bar chart value labels; values are shown clearly in the tables.
- Slow/Error API charts use Rank numbers while full API details are available in tables.


## v24 Clean Dashboard

- Removed the `Track-wise Slow Summary` table from `Insights`.
- Removed the extra `Track-wise Slow Summary` table from `Track_Comparison`.
- Removed the SLA pie chart because it was visually cluttered.
- SLA PASS/FAIL/TOTAL values remain as a clean table.


## v25 Clean Pie Chart

- Re-added `API SLA Pass vs Fail` pie chart.
- Removed overlapping pie labels from the chart.
- PASS/FAIL/TOTAL values are shown clearly in the SLA Breakdown table next to the chart.
- Track-wise summary remains removed from Insights and Track_Comparison.


## v26 Code and UI Fix

- Clean SLA pie chart code is included in `main.py`.
- Pie chart does not use overlapping internal labels.
- PASS/FAIL/TOTAL values remain visible in the SLA Breakdown table.
- Streamlit UI title changed to `CiscoIQ-SaaS-Support-Services Performance Dashboard`.
- Streamlit title font size reduced.


## v27 Title and Report Context

- Excel Insights title changed to `CiscoIQ-SaaS-Support-Services Performance Dashboard`.
- Streamlit title changed to `CiscoIQ-SaaS-Support-Services Performance Dashboard`.
- Streamlit title font size reduced.
- Insights now includes Report Context parsed from the uploaded filename:
  - Concurrent Users
  - Devices Count
  - Date
  - Duration
  - Region


## v28 Dashboard Final Updates

- Removed Report File from the Insights page.
- Added Health Score explanation in Insights.
- Moved pie chart title outside the chart to prevent overlap.
- Streamlit UI title is centered and smaller.
- Added Generate Report button after upload.
- APIs sheet column names changed:
  - `Feature` -> `Tracks`
  - `Scenario` -> `Transactions`


## v29 Next-Level Comparison Dashboard

- Track_Comparison now separates AskAI tracks and non-AskAI tracks into clear sections.
- AskAI section columns:
  - 0-10sec %
  - 10-20sec %
  - 20-30sec %
  - >30sec %
  - Max Seconds
- Other tracks section columns:
  - 0-2sec %
  - 3-4sec %
  - 4-6sec %
  - >6sec %
  - Max Seconds
- Transactions and Errors sheets now use capitalized column names:
  - Transaction
  - SampleCount
  - ErrorCount
  - ErrorPct


## v30 Verified Track Comparison

This package has been verified from the actual project code.

### Track_Comparison
The workbook now has separate sections:
- AskAI Tracks
  - 0-10sec %
  - 10-20sec %
  - 20-30sec %
  - >30sec %
  - Max Seconds

- Assets / Assessments / Home / Settings / Support Tracks
  - 0-2sec %
  - 3-4sec %
  - 4-6sec %
  - >6sec %
  - Max Seconds

### Transactions / Errors headers
- Transaction
- SampleCount
- ErrorCount
- ErrorPct
