from io import BytesIO
from pathlib import Path
import tempfile

import streamlit as st

from main import build_report


st.set_page_config(page_title="JMeter JSON to Excel Report", layout="centered")

st.title("JMeter JSON to Excel Report")
st.write("Upload your JMeter `statistics.json` file and download the formatted Excel report.")

uploaded_file = st.file_uploader("Upload statistics.json", type=["json"])

if uploaded_file is not None:
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = Path(tmpdir) / "statistics.json"
        output_path = Path(tmpdir) / "JMeter_Report.xlsx"

        input_path.write_bytes(uploaded_file.getvalue())

        try:
            build_report(input_path, output_path)
            st.success("Report generated successfully.")

            st.download_button(
                label="Download Excel Report",
                data=output_path.read_bytes(),
                file_name="JMeter_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(f"Failed to generate report: {exc}")
