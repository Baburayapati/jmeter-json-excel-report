# JMeter JSON to Excel Report Generator

This project converts a JMeter static `statistics.json` report into an Excel workbook.

## Output Sheets

1. `Transactions`
   - Contains only transaction controller rows starting with `T01`, `T02`, etc.
   - Does not include split Feature/Scenario/Endpoint columns.

2. `Errors`
   - Contains rows where `errorCount > 0`.

3. `APIs`
   - Formerly `All_Results`.
   - Excludes transaction controller rows.
   - Splits API names into:
     - `Feature`
     - `Scenario`
     - `Endpoint`

## Column Rules

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

## Run from Command Line

```bash
pip install -r requirements.txt
python main.py statistics.json JMeter_Report.xlsx
```

## Run UI

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the browser URL shown by Streamlit, upload `statistics.json`, and download the Excel report.

## Share with Team

Option 1: Zip this folder and share it with your team.

Option 2: Push the files to GitHub:
```bash
git init
git add .
git commit -m "JMeter JSON Excel report generator"
git branch -M main
git remote add origin <your-github-repo-url>
git push -u origin main
```

Your team can clone the repo and run:
```bash
pip install -r requirements.txt
streamlit run app.py
```
