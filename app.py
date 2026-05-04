from pathlib import Path
import tempfile

import streamlit as st

from main import build_report, build_comparison_report


st.set_page_config(page_title="JMeter JSON to Excel Report", layout="centered")

st.title("JMeter JSON to Excel Report")
st.write(
    "Upload one JMeter `statistics.json` file for the normal report. "
    "Upload two or more files to add a `Comparison` sheet."
)

st.markdown(
    """
### SLA Rules
- APIs where **Feature starts with `AskAI`**: SLA is **< 10 sec**
- All other APIs: SLA is **< 2 sec**
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
                st.info("Comparison uses the first uploaded JSON as baseline and the last uploaded JSON as latest.")

            st.download_button(
                label="Download Excel Report",
                data=output_path.read_bytes(),
                file_name="JMeter_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(f"Failed to generate report: {exc}")
