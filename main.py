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
    return bool(re.match(r"^T\d{2}", str(name).strip()))


def split_api_name(name: str) -> Tuple[str, str, str]:
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


def build_single_report_frames(json_path: str | Path):
    df = load_statistics_json(json_path)

    transactions_df = df[df["transaction"].apply(is_transaction)].copy()
    apis_df = df[~df["transaction"].apply(is_transaction)].copy()
    errors_df = df[pd.to_numeric(df.get("errorCount", 0), errors="coerce").fillna(0) > 0].copy()

    split_df = apis_df["transaction"].apply(lambda x: pd.Series(split_api_name(x)))
    split_df.columns = ["Feature", "Scenario", "Endpoint"]
    apis_df = pd.concat([split_df, apis_df], axis=1)

    return {
        "Transactions": clean_and_rename_columns(transactions_df),
        "Errors": clean_and_rename_columns(errors_df),
        "APIs": clean_and_rename_columns(apis_df),
    }


def build_report(json_path: str | Path, output_excel_path: str | Path) -> None:
    frames = build_single_report_frames(json_path)
    write_excel(frames, output_excel_path)


def prepare_compare_frame(df: pd.DataFrame, label: str) -> pd.DataFrame:
    base = df.copy()
    base["transaction"] = base["transaction"].astype(str)
    needed = ["transaction", "sampleCount", "errorCount", "errorPct", "meanResTime", "pct1ResTime", "pct2ResTime", "pct3ResTime"]
    available = [c for c in needed if c in base.columns]
    base = base[available].copy()

    for col in ["meanResTime", "pct1ResTime", "pct2ResTime", "pct3ResTime"]:
        if col in base.columns:
            base[col] = pd.to_numeric(base[col], errors="coerce") / 1000

    rename_map = {
        "sampleCount": f"{label} Sample Count",
        "errorCount": f"{label} Error Count",
        "errorPct": f"{label} Error %",
        "meanResTime": f"{label} Avg Sec",
        "pct1ResTime": f"{label} 90th Sec",
        "pct2ResTime": f"{label} 95th Sec",
        "pct3ResTime": f"{label} 99th Sec",
    }
    return base.rename(columns=rename_map)


def build_comparison(json_paths: List[str | Path], labels: List[str]) -> pd.DataFrame:
    if len(json_paths) < 2:
        return pd.DataFrame()

    compare_df = prepare_compare_frame(load_statistics_json(json_paths[0]), labels[0])

    for path, label in zip(json_paths[1:], labels[1:]):
        next_df = prepare_compare_frame(load_statistics_json(path), label)
        compare_df = compare_df.merge(next_df, on="transaction", how="outer")

    baseline = labels[0]
    latest = labels[-1]

    if f"{baseline} Avg Sec" in compare_df.columns and f"{latest} Avg Sec" in compare_df.columns:
        compare_df["Avg Sec Diff"] = compare_df[f"{latest} Avg Sec"] - compare_df[f"{baseline} Avg Sec"]
        compare_df["Avg Sec Diff %"] = (compare_df["Avg Sec Diff"] / compare_df[f"{baseline} Avg Sec"]) * 100

    if f"{baseline} 90th Sec" in compare_df.columns and f"{latest} 90th Sec" in compare_df.columns:
        compare_df["90th Sec Diff"] = compare_df[f"{latest} 90th Sec"] - compare_df[f"{baseline} 90th Sec"]
        compare_df["90th Sec Diff %"] = (compare_df["90th Sec Diff"] / compare_df[f"{baseline} 90th Sec"]) * 100

    if f"{baseline} Error Count" in compare_df.columns and f"{latest} Error Count" in compare_df.columns:
        compare_df["Error Count Diff"] = compare_df[f"{latest} Error Count"] - compare_df[f"{baseline} Error Count"]

    # Add split columns for non-transaction/API rows too.
    split_df = compare_df["transaction"].apply(lambda x: pd.Series(split_api_name(x)))
    split_df.columns = ["Feature", "Scenario", "Endpoint"]
    compare_df = pd.concat([split_df, compare_df], axis=1)

    first_cols = ["Feature", "Scenario", "Endpoint", "transaction"]
    other_cols = [c for c in compare_df.columns if c not in first_cols]
    return compare_df[first_cols + other_cols]


def build_comparison_report(json_paths: List[str | Path], labels: List[str], output_excel_path: str | Path) -> None:
    frames = build_single_report_frames(json_paths[-1])
    frames["Comparison"] = build_comparison(json_paths, labels)
    write_excel(frames, output_excel_path)


def write_excel(frames: Dict[str, pd.DataFrame], output_excel_path: str | Path) -> None:
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        for sheet_name, df in frames.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)

            sheet = writer.sheets[safe_name]
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
                    if header and ("sec" in str(header).lower() or "diff %" in str(header).lower() or "error %" in str(header).lower()):
                        cell.number_format = "0.000"


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Generate Excel report from JMeter statistics.json")
    parser.add_argument("json_files", nargs="+", help="One or more JMeter statistics.json files")
    parser.add_argument("--output", default="JMeter_Report.xlsx", help="Path to output .xlsx file")
    args = parser.parse_args()

    labels = [Path(p).stem for p in args.json_files]

    if len(args.json_files) == 1:
        build_report(args.json_files[0], args.output)
    else:
        build_comparison_report(args.json_files, labels, args.output)
