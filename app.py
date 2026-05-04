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
        background: none !important;
        padding: 0 !important;
        box-shadow: none !important;
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

    .dashboard-subtitle {
        text-align: center;
        font-size: 12px;
        color: #27364a;
        margin-bottom: 0.8rem;
    }

    .rules-section h3 {
        color: #154c79 !important;
        margin-top: 0.8rem !important;
        margin-bottom: 0.3rem !important;
        font-size: 16px !important;
        line-height: 1.15 !important;
        font-weight: 600 !important;
    }

    .rules-section ul {
        margin-top: 0.2rem !important;
        margin-bottom: 0.85rem !important;
        padding-left: 1.2rem !important;
    }

    .rules-section li {
        font-size: 13px !important;
        line-height: 1.3 !important;
        font-weight: 400 !important;
        margin-bottom: 0.25rem !important;
    }

    .metric-pill {
        color: #2e7d32 !important;
        background: rgba(46, 125, 50, 0.08);
        border-radius: 7px;
        padding: 1px 6px;
        font-weight: 500;
        white-space: nowrap;
    }


    .dashboard-title {
        display: table;
        margin: 0 auto 0.55rem auto;
        background: linear-gradient(90deg, #154c79, #7b2cbf);
        color: white;
        padding: 11px 18px;
        border-radius: 15px;
        box-shadow: 0 8px 24px rgba(21, 76, 121, 0.18);
        font-size: 18px;
        line-height: 1.2;
        font-weight: 700;
        text-align: center;
        width: auto;
        max-width: fit-content;
    }
</style>
    """,
    unsafe_allow_html=True,
)


st.markdown("<div class='dashboard-title'>CiscoIQ-SaaS-Support-Services Performance Dashboard</div>", unsafe_allow_html=True)
st.markdown(
    "<div class='dashboard-subtitle'>Upload one JMeter <code>statistics.json</code> file for a normal dashboard report. Upload two or more files for comparison.</div>",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="rules-section">
  <h3>SLA Rules</h3>
  <ul>
    <li>AskAI APIs: SLA is &lt; 10 sec</li>
    <li>Assets, Assessments, Home, Settings and Support APIs: SLA is &lt; 2 sec</li>
  </ul>

  <h3>Track Comparison Metrics</h3>
  <ul>
    <li>AskAI tracks:
      <span class="metric-pill">0-10s</span>,
      <span class="metric-pill">10-20s</span>,
      <span class="metric-pill">20-30s</span>,
      <span class="metric-pill">&gt;30s</span>
    </li>
    <li>Assets, Assessments, Home, Settings and Support tracks:
      <span class="metric-pill">0-2s</span>,
      <span class="metric-pill">3-4s</span>,
      <span class="metric-pill">4-6s</span>,
      <span class="metric-pill">&gt;6s</span>
    </li>
  </ul>
</div>
""",
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader(
    "Upload statistics.json file(s)",
    type=["json"],
    accept_multiple_files=True,
)

if uploaded_files:
    generate_clicked = st.button("Generate Report", type="primary")
else:
    generate_clicked = False

if uploaded_files and generate_clicked:
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
