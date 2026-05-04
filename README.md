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
