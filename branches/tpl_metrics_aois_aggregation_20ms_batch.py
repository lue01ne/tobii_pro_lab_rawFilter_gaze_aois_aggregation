import pandas as pd
from pathlib import Path

"""
Copyright (c) 2026 Yue "Lou" Lyu

Research Use Only License:
Permission is granted for non-commercial research and academic use only.
Commercial use is prohibited without prior written permission.

See the LICENSE file in the project root for full license terms.
"""

BASE_DIR = Path(__file__).resolve().parent

INPUT_FOLDER = BASE_DIR / "input_metrics_data"
OUTPUT_FOLDER = BASE_DIR / "output_data"

SHEET_NAME = "TPL_rawFilter_metrics"
DURATION_TARGET = 20

# Aggregation happens only when all of those match. 
GROUP_COLS = [
    "Recording", "Participant", "Position", "TOI", "Interval", "Event_type", "Validity"
]


def merge_consecutive_aoi(df: pd.DataFrame):
    # Convert time columns to numeric
    for c in ["Start", "Stop", "Duration"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # Split into 20ms bins and non-20 rows
    df20 = df[df["Duration"] == DURATION_TARGET].copy()
    df_non20 = df[df["Duration"] != DURATION_TARGET].copy()

    # Sort for run detection (only for the 20ms stream)
    df20 = df20.sort_values(GROUP_COLS + ["Start", "Stop"]).reset_index(drop=True)

    def compute_new_run(g: pd.DataFrame) -> pd.Series:
        same_aoi = g["AOI"].eq(g["AOI"].shift(1))

        # continuity: Start == previous Stop OR Start step == 20
        contiguous_by_stop = g["Start"].eq(g["Stop"].shift(1))
        contiguous_by_step = (g["Start"] - g["Start"].shift(1)).eq(DURATION_TARGET)

        contiguous = contiguous_by_stop | contiguous_by_step

        # new run when AOI changes or continuity breaks
        return ~(same_aoi & contiguous)

    df20["new_run"] = (
        df20.groupby(GROUP_COLS, group_keys=False)
        .apply(compute_new_run)
        .reset_index(drop=True)
    )

    df20["run_id"] = df20.groupby(GROUP_COLS)["new_run"].cumsum()
    df20["SegmentsMerged"] = 1

    merged_runs = (
        df20.groupby(GROUP_COLS + ["run_id"], as_index=False)
        .agg(
            EventIndex=("EventIndex", "first"),
            Start=("Start", "min"),
            Stop=("Stop", "max"),
            Duration=("Duration", "sum"),
            AOI=("AOI", "first"),
            SegmentsMerged=("SegmentsMerged", "sum"),
        )
    ).sort_values(GROUP_COLS + ["Start", "Stop"]).reset_index(drop=True)

    return merged_runs, df20, df_non20


def process_one_file(input_xlsx: Path):
    df = pd.read_excel(input_xlsx, sheet_name=SHEET_NAME)

    merged_runs, df20_rows, non20_rows = merge_consecutive_aoi(df)

    # Combine merged 20ms runs + raw non-20 rows into ONE sheet
    merged_runs_out = merged_runs.copy()
    merged_runs_out["Source"] = "Merged20msRun"

    non20_out = non20_rows.copy()
    non20_out["Source"] = "RawNon20Row"
    non20_out["OriginalRowIndex"] = non20_out.index

    all_cols = list(dict.fromkeys(list(merged_runs_out.columns) + list(non20_out.columns)))
    merged_runs_out = merged_runs_out.reindex(columns=all_cols)
    non20_out = non20_out.reindex(columns=all_cols)

    combined = (
        pd.concat([merged_runs_out, non20_out], ignore_index=True)
        .sort_values(GROUP_COLS + ["Start", "Stop"], kind="mergesort")
        .reset_index(drop=True)
    )

    # AOI summaries (based on original 20ms rows)
    aoi_summary_overall = (
        df20_rows.groupby("AOI", as_index=False)
        .agg(
            Rows=("AOI", "size"),
            TotalDuration=("Duration", "sum"),
            FirstStart=("Start", "min"),
            LastStop=("Stop", "max"),
        )
        .sort_values("TotalDuration", ascending=False)
    )

    aoi_by_group = (
        df20_rows.groupby(GROUP_COLS + ["AOI"], as_index=False)
        .agg(
            Rows=("AOI", "size"),
            TotalDuration=("Duration", "sum"),
            FirstStart=("Start", "min"),
            LastStop=("Stop", "max"),
        )
    )

    output_xlsx = OUTPUT_FOLDER / f"{input_xlsx.stem}_aggregated.xlsx"

    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        combined.to_excel(writer, sheet_name="Timeline_Combined", index=False)
        merged_runs.to_excel(writer, sheet_name="MergedRuns", index=False)
        aoi_summary_overall.to_excel(writer, sheet_name="AOI_Summary", index=False)
        aoi_by_group.to_excel(writer, sheet_name="AOI_ByGroup", index=False)

    print(f"✅ Processed: {input_xlsx.name}  ->  {output_xlsx.name}")


def main():
    if not INPUT_FOLDER.exists():
        raise FileNotFoundError(f"Input folder not found: {INPUT_FOLDER}")

    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    input_files = sorted(INPUT_FOLDER.glob("*.xlsx"))

    # Skip temporary Excel lock files like "~$something.xlsx"
    input_files = [f for f in input_files if not f.name.startswith("~$")]

    if not input_files:
        raise FileNotFoundError(f"No .xlsx files found in: {INPUT_FOLDER}")

    print(f"Found {len(input_files)} Excel files in: {INPUT_FOLDER}")

    for input_xlsx in input_files:
        process_one_file(input_xlsx)

    print("🦕 Done. All files processed.")


if __name__ == "__main__":
    main()
