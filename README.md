# JMeter JSON to Excel Report Generator with SLA, Charts, and Comparison

This Streamlit app converts JMeter static `statistics.json` files into Excel reports.

## Output Sheets

- `Insights` - KPI summary, SLA chart, top slow APIs chart, top error features chart
- `Transactions` - only transaction rows starting with `T01`, `T02`, etc.
- `Errors` - rows where `errorCount > 0`
- `APIs` - non-transaction API rows only
- `Comparison` - added when two or more JSON files are uploaded

## Important API Sheet Change

The `APIs` sheet does **not** include the original `transaction` column anymore.

Instead, it uses:
- `Feature`
- `Scenario`
- `Endpoint`

## SLA Rules

- If `Feature` starts with `AskAI` then SLA is `< 10 sec`
- All other APIs have SLA `< 2 sec`

The app adds:
- `SLA Sec`
- `SLA Rule`
- `SLA Status`
- `SLA Breach Sec`

## Column Cleanup

Removed from all sheets:
- `medianResTime`
- `throughput`
- `receivedKBytesPerSec`
- `sentKBytesPerSec`

Converted from milliseconds to seconds and renamed:
- `meanResTime` -> `Avg ResTime in sec`
- `minResTime` -> `Min ResTime in sec`
- `maxResTime` -> `MaxRes Time in sec`
- `pct1ResTime` -> `90thPercentile Resp Time in Sec`
- `pct2ResTime` -> `95thPercentile Resp Time in Sec`
- `pct3ResTime` -> `99thPercentile Resp Time in Sec`

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
3. Upload two/more JSONs for comparison.
4. Download Excel.
