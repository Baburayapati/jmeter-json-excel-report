import json
import re
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd


REMOVE_COLUMNS = [
    "medianResTime",
    "throughput",
    "receivedKBytesPerSec",
    "sentKBytesPerSec",
]

TIME_COLUMNS_MS_TO_SEC = {
    "meanResTime": "Avg ResTime in sec",
    "minResTime": "Min ResTime in sec",
    "maxResTime": "MaxRes Time in sec",
    "pct1ResTime": "90thPercentile Resp Time in Sec",
    "pct2ResTime": "95thPercentile Resp Time in Sec",
    "pct3ResTime": "99thPercentile Resp Time in Sec",
}


def is_transaction(name: str) -> bool:
    """Return True for JMeter transaction controller rows like T01_..., T10_..., etc."""
    return bool(re.match(r"^T\d{2}", str(name).strip()))


def split_api_name(name: str) -> Tuple[str, str, str]:
    """
    Example:
    AskAIAssessmentConfiguration/Click AskAI/tag/auth/ntpagetag.gif-1,371-0

    Returns:
    Feature = AskAIAssessmentConfiguration
    Scenario = Click AskAI
    Endpoint = tag/auth/ntpagetag.gif-1,371-0
    """
    parts = str(name).split("/")
    if len(parts) >= 3:
        return parts[0], parts[1], "/".join(parts[2:])
    if len(parts) == 2:
        return parts[0], parts[1], ""
    return str(name), "", ""


def load_statistics_json(json_path: str | Path) -> pd.DataFrame:
    with open(json_path, "r", encoding="utf-8") as file:
        data: Dict[str, Dict[str, Any]] = json.load(file)

    rows: List[Dict[str, Any]] = []
    for key, value in data.items():
        row = dict(value)
        row["transaction"] = row.get("transaction", key)
        rows.append(row)

    return pd.DataFrame(rows)


def clean_and_rename_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    for col in REMOVE_COLUMNS:
        if col in df.columns:
            df = df.drop(columns=[col])

    for old_col, new_col in TIME_COLUMNS_MS_TO_SEC.items():
        if old_col in df.columns:
            df[old_col] = pd.to_numeric(df[old_col], errors="coerce") / 1000
            df = df.rename(columns={old_col: new_col})

    preferred_order = [
        "Feature",
        "Scenario",
        "Endpoint",
        "transaction",
        "sampleCount",
        "errorCount",
        "errorPct",
        "Avg ResTime in sec",
        "Min ResTime in sec",
        "MaxRes Time in sec",
        "90thPercentile Resp Time in Sec",
        "95thPercentile Resp Time in Sec",
        "99thPercentile Resp Time in Sec",
    ]
    ordered = [c for c in preferred_order if c in df.columns]
    remaining = [c for c in df.columns if c not in ordered]
    return df[ordered + remaining]


def build_report(json_path: str | Path, output_excel_path: str | Path) -> None:
    df = load_statistics_json(json_path)

    transactions_df = df[df["transaction"].apply(is_transaction)].copy()
    apis_df = df[~df["transaction"].apply(is_transaction)].copy()
    errors_df = df[pd.to_numeric(df.get("errorCount", 0), errors="coerce").fillna(0) > 0].copy()

    # Split API name only in APIs sheet.
    split_df = apis_df["transaction"].apply(lambda x: pd.Series(split_api_name(x)))
    split_df.columns = ["Feature", "Scenario", "Endpoint"]
    apis_df = pd.concat([split_df, apis_df], axis=1)

    transactions_df = clean_and_rename_columns(transactions_df)
    errors_df = clean_and_rename_columns(errors_df)
    apis_df = clean_and_rename_columns(apis_df)

    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        transactions_df.to_excel(writer, sheet_name="Transactions", index=False)
        errors_df.to_excel(writer, sheet_name="Errors", index=False)
        apis_df.to_excel(writer, sheet_name="APIs", index=False)

        for sheet_name in ["Transactions", "Errors", "APIs"]:
            sheet = writer.sheets[sheet_name]
            sheet.freeze_panes = "A2"

            for cell in sheet[1]:
                cell.font = cell.font.copy(bold=True)
                cell.alignment = cell.alignment.copy(horizontal="center", vertical="center", wrap_text=True)

            for column_cells in sheet.columns:
                max_length = 0
                col_letter = column_cells[0].column_letter
                for cell in column_cells:
                    value = "" if cell.value is None else str(cell.value)
                    max_length = max(max_length, len(value))
                sheet.column_dimensions[col_letter].width = min(max(max_length + 2, 12), 45)

            for row in sheet.iter_rows(min_row=2):
                for cell in row:
                    header = sheet.cell(row=1, column=cell.column).value
                    if header and "sec" in str(header).lower():
                        cell.number_format = "0.000"


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Generate Excel report from JMeter statistics.json")
    parser.add_argument("json_file", help="Path to JMeter statistics.json")
    parser.add_argument("output_excel", help="Path to output .xlsx file")
    args = parser.parse_args()

    build_report(args.json_file, args.output_excel)
