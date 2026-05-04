from pathlib import Path
import tempfile

import streamlit as st

from main import build_report, build_comparison_report


st.set_page_config(page_title="CiscoIQ-SaaS-Support-Services Performance Dashboard", layout="centered")


st.markdown(
    """
<style>
    .stApp {
        background: linear-gradient(135deg, #eef7ff 0%, #f6f2ff 45%, #ecfff4 100%);
        color: #172033;
    }

    [data-testid="stHeader"] {
        background: rgba(255, 255, 255, 0);
    }

    .block-container {
        padding-top: 3rem;
        padding-bottom: 3rem;
        max-width: 950px;
    }

    h1 {
        background: linear-gradient(90deg, #154c79, #7b2cbf);
        color: white !important;
        padding: 14px 20px;
        border-radius: 16px;
        box-shadow: 0 8px 24px rgba(21, 76, 121, 0.18);
        font-size: 30px !important;
        line-height: 1.25 !important;
    }

    h3 {
        color: #154c79 !important;
        margin-top: 1.5rem;
    }

    [data-testid="stFileUploader"] {
        background: rgba(255, 255, 255, 0.82);
        border: 1px solid rgba(21, 76, 121, 0.18);
        border-radius: 16px;
        padding: 18px;
        box-shadow: 0 8px 22px rgba(0, 0, 0, 0.06);
    }

    .stAlert {
        border-radius: 14px;
    }

    .stDownloadButton button {
        background: linear-gradient(90deg, #154c79, #2a9d8f);
        color: white;
        border: 0;
        border-radius: 12px;
        padding: 0.65rem 1rem;
        font-weight: 700;
        box-shadow: 0 6px 18px rgba(42, 157, 143, 0.25);
    }

    .stDownloadButton button:hover {
        border: 0;
        color: white;
        filter: brightness(1.04);
    }
</style>
    """,
    unsafe_allow_html=True,
)


st.title("CiscoIQ-SaaS-Support-Services Performance Dashboard")
st.write(
    "Upload one JMeter `statistics.json` file for the normal report. "
    "Upload two or more files to generate a focused comparison report with Insights, Track_Comparison, and APIs_Comparison."
)

st.markdown(
    """
### SLA Rules
- **AskAI APIs**: SLA is **< 10 sec**
- **Assets, Assessments, Home, Settings and Support APIs**: SLA is **< 2 sec**

### Track Comparison Buckets
- **AskAI tracks**: `0-10s`, `10-20s`, `20-30s`, `>30s`
- **Assets, Assessments, Home, Settings and Support tracks**: `0-2s`, `3-4s`, `4-6s`, `>6s`
"""
)

uploaded_files = st.file_uploader(
    "Upload statistics.json file(s)",
    type=["json"],
    accept_multiple_files=True,
)

if uploaded_files:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        json_paths = []
        labels = []

        for idx, uploaded_file in enumerate(uploaded_files, start=1):
            clean_name = uploaded_file.name.replace(" ", "_")
            path = tmpdir / f"{idx}_{clean_name}"
            path.write_bytes(uploaded_file.getvalue())
            json_paths.append(path)
            labels.append(Path(uploaded_file.name).stem)

        output_path = tmpdir / "JMeter_Report.xlsx"

        try:
            if len(json_paths) == 1:
                build_report(json_paths[0], output_path)
                st.success("Single report generated successfully.")
            else:
                build_comparison_report(json_paths, labels, output_path)
                st.success("Comparison report generated successfully.")
                st.info(
                    "Track_Comparison shows every uploaded run side-by-side. "
                    "The first file is baseline and last file is latest for the raw Comparison sheet."
                )

            st.download_button(
                label="Download Excel Report",
                data=output_path.read_bytes(),
                file_name="JMeter_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(f"Failed to generate report: {exc}")
