import pandas as pd
from pathlib import Path

"""
Copyright (c) 2026 Yue "Lou" Lyu

Research Use Only License:
Permission is granted for non-commercial research and academic use only.
Commercial use is prohibited without prior written permission.

See the LICENSE file in the project root for full license terms.
"""

# ============================================================
# PATH CONFIG (relative folders)
# ============================================================
BASE_DIR = Path(__file__).resolve().parent

INPUT_FOLDER = BASE_DIR / "input_metrics_data"
OUTPUT_FOLDER = BASE_DIR / "output_data"

SHEET_NAME = "TPL_rawFilter_metrics"
DURATION_TARGET = 20  # merge rows with Duration <= 20

# Group context: rows must match these columns to be considered for the same run stream
GROUP_COLS = [
    "Recording",
    "Participant",
    "Position",
    "TOI",
    "Interval",
    "Event_type",
    "Validity",
]


def merge_consecutive_aoi_duration_le20(df: pd.DataFrame):
    """
    Build AOI runs from rows where Duration <= 20, merging consecutive rows when:
      - AOI is the same as the previous row, AND
      - Start == previous Stop  (works for any duration <= 20)
    Optional fallback continuity (only for 20ms-to-20ms rows):
      - Start - previous Start == 20

    Returns:
      merged_runs: aggregated runs built from rows with Duration <= 20
      le20_rows:   the original rows with Duration <= 20 (used for summaries/debug)
      gt20_rows:   the original rows with Duration > 20 (kept raw; not merged)
    """
    # Ensure numeric columns
    for c in ["Start", "Stop", "Duration"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Split streams
    le20_rows = df[df["Duration"].le(DURATION_TARGET)].copy()
    gt20_rows = df[df["Duration"].gt(DURATION_TARGET)].copy()

    # Sort for run detection
    le20_rows = le20_rows.sort_values(GROUP_COLS + ["Start", "Stop"]).reset_index(drop=True)

    # -------- Vectorized new_run detection (no groupby.apply) --------
    prev_aoi = le20_rows.groupby(GROUP_COLS)["AOI"].shift(1)
    prev_stop = le20_rows.groupby(GROUP_COLS)["Stop"].shift(1)
    prev_start = le20_rows.groupby(GROUP_COLS)["Start"].shift(1)
    prev_dur = le20_rows.groupby(GROUP_COLS)["Duration"].shift(1)

    same_aoi = le20_rows["AOI"].eq(prev_aoi)

    # Primary continuity rule: Start == previous Stop (works for ANY duration <= 20)
    contiguous_by_stop = le20_rows["Start"].eq(prev_stop)

    # Optional fallback only for 20ms-to-20ms rows
    contiguous_by_step20 = (
        (le20_rows["Start"] - prev_start).eq(DURATION_TARGET)
        & le20_rows["Duration"].eq(DURATION_TARGET)
        & prev_dur.eq(DURATION_TARGET)
    )

    contiguous = contiguous_by_stop | contiguous_by_step20

    # New run when AOI changes OR continuity breaks
    le20_rows["new_run"] = ~(same_aoi & contiguous)

    # Run id within each GROUP_COLS group
    le20_rows["run_id"] = le20_rows.groupby(GROUP_COLS)["new_run"].cumsum()
    le20_rows["SegmentsMerged"] = 1

    # -------- Aggregate runs --------
    agg_dict = {
        "Start": ("Start", "min"),
        "Stop": ("Stop", "max"),
        "Duration": ("Duration", "sum"),
        "AOI": ("AOI", "first"),
        "SegmentsMerged": ("SegmentsMerged", "sum"),
    }
    if "EventIndex" in le20_rows.columns:
        agg_dict["EventIndex"] = ("EventIndex", "first")

    merged_runs = (
        le20_rows.groupby(GROUP_COLS + ["run_id"], as_index=False)
        .agg(**{k: v for k, v in agg_dict.items()})
    ).sort_values(GROUP_COLS + ["Start", "Stop"]).reset_index(drop=True)

    return merged_runs, le20_rows, gt20_rows


def process_one_file(input_xlsx: Path):
    print(f"\n📥 Processing: {input_xlsx.name}")

    df = pd.read_excel(input_xlsx, sheet_name=SHEET_NAME)

    merged_runs, le20_rows, gt20_rows = merge_consecutive_aoi_duration_le20(df)

    # ---- Combined timeline sheet: merged (<=20) runs + raw (>20) rows ----
    merged_runs_out = merged_runs.copy()
    merged_runs_out["Source"] = "Merged<=20msRun"

    gt20_out = gt20_rows.copy()
    gt20_out["Source"] = "Raw>20msRow"
    gt20_out["OriginalRowIndex"] = gt20_out.index

    # Make schemas compatible (union of columns)
    all_cols = list(dict.fromkeys(list(merged_runs_out.columns) + list(gt20_out.columns)))
    merged_runs_out = merged_runs_out.reindex(columns=all_cols)
    gt20_out = gt20_out.reindex(columns=all_cols)

    combined = (
        pd.concat([merged_runs_out, gt20_out], ignore_index=True)
        .sort_values(GROUP_COLS + ["Start", "Stop"], kind="mergesort")
        .reset_index(drop=True)
    )

    # AOI summaries (based on original <=20 rows)
    aoi_summary_overall = (
        le20_rows.groupby("AOI", as_index=False)
        .agg(
            Rows=("AOI", "size"),
            TotalDuration=("Duration", "sum"),
            FirstStart=("Start", "min"),
            LastStop=("Stop", "max"),
        )
        .sort_values("TotalDuration", ascending=False)
    )

    aoi_by_group = (
        le20_rows.groupby(GROUP_COLS + ["AOI"], as_index=False)
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
        # Optional debug sheets:
        le20_rows.to_excel(writer, sheet_name="Raw_Duration_le20", index=False)
        if not gt20_rows.empty:
            gt20_rows.to_excel(writer, sheet_name="Raw_Duration_gt20", index=False)

    print(f"✅ Output: {output_xlsx.name}")


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

    print("\n🎉 Done. All files processed.")


if __name__ == "__main__":
    main()
