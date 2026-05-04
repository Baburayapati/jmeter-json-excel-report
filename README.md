# JMeter JSON to Excel Report Generator with Comparison

This Streamlit app converts JMeter static `statistics.json` files into Excel reports.

## Upload Modes

### Single JSON
Upload one `statistics.json` file.

Output sheets:
- `Transactions`
- `Errors`
- `APIs`

### Multiple JSON Files
Upload two or more `statistics.json` files.

Output sheets:
- `Transactions` from latest uploaded file
- `Errors` from latest uploaded file
- `APIs` from latest uploaded file
- `Comparison`

## Comparison Logic

The app compares:
- First uploaded JSON = Baseline
- Last uploaded JSON = Latest

The `Comparison` sheet includes:
- Avg response time difference in seconds
- Avg response time difference %
- 90th percentile response time difference
- 90th percentile response time difference %
- Error count difference
- Side-by-side sample count, error count, error %, avg, 90th, 95th, and 99th values for each uploaded file

## Run Locally

```bash
python3 -m venv venv
source venv/bin/activate
python -m pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Cloud

1. Push these files to GitHub:
   - `app.py`
   - `main.py`
   - `requirements.txt`
   - `README.md`

2. Go to Streamlit Cloud.

3. Deploy:
   - Repo: your GitHub repo
   - Branch: `main`
   - Main file path: `app.py`

4. Share the generated URL with your team.

## Team Usage

1. Open the app URL.
2. Upload one JSON for normal report.
3. Upload two/more JSONs for comparison.
4. Download Excel.
