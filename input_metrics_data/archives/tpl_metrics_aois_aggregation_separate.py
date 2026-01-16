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
#keeps only rows where Duration == 20
# merges consecutive 20 ms rows when they belong to the same AOI and are continuous in time 
# writes out:
## the merged segments (“runs”)
## summary tables per AOI
## GROUP_COLS defines the “context” that must match before rows can be merged 
## (same Recording, Participant, Interval, Event type, etc.). 
## Rows from different groups are never merged together.
# ============================================================

# ============================================================
# PATH CONFIG (relative folders)
## It reads from: ./input_metrics_data/hockey_vlad_metrics_workaround.xlsx
## It writes to: ./output_data/Hockey_workaround_aggregated.xlsx
## It reads the sheet: TPL_rawFilter_metrics
## It only processes rows with Duration == 20 (to correct)
# ============================================================
BASE_DIR = Path(__file__).resolve().parent

INPUT_FOLDER = BASE_DIR / "input_metrics_data"       # <-- changed
OUTPUT_FOLDER = BASE_DIR / "output_data"             # <-- changed

INPUT_XLSX = INPUT_FOLDER / "hockey_vlad_metrics_workaround.xlsx"
OUTPUT_XLSX = OUTPUT_FOLDER / "hockey_workaround_aggregated.xlsx"

SHEET_NAME = "TPL_rawFilter_metrics"
DURATION_TARGET = 20

# Two rows can only be merged if they have the same:
## Recording (same data file / trial)
## Participant
## Position (goalie/player position etc.)
## TOI
## Interval
## Event_type
## Validity
# Then (within that same group), the script checks:
## are they consecutive in time (Start/Stop)?
## do they have the same AOI?

GROUP_COLS = [
    "Recording", "Participant", "Position", "TOI", "Interval", "Event_type", "Validity"
]


def merge_consecutive_aoi(df: pd.DataFrame) -> pd.DataFrame:
    # Convert time columns to numeric
    for c in ["Start", "Stop", "Duration"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # Keep only rows with Duration == 20 (as requested)
    # Everything else (Duration ≠ 20) is excluded and later written to a separate sheet.
    df20 = df[df["Duration"] == DURATION_TARGET].copy()
    df_non20 = df[df["Duration"] != DURATION_TARGET].copy()

    # Sort for run detection
    df20 = df20.sort_values(GROUP_COLS + ["Start", "Stop"]).reset_index(drop=True)

    def compute_new_run(g: pd.DataFrame) -> pd.Series:
        same_aoi = g["AOI"].eq(g["AOI"].shift(1))

        # “Start (difference) values as 20 / 20-step continuitys” requirement:
        # either next Start == previous Stop OR Start increased by exactly 20
        contiguous_by_stop = g["Start"].eq(g["Stop"].shift(1))
        contiguous_by_step = (g["Start"] - g["Start"].shift(1)).eq(DURATION_TARGET)

        contiguous = contiguous_by_stop | contiguous_by_step

        # Start a new run if AOI changes or the 20-step continuity breaks
        new_run = ~(same_aoi & contiguous)
        return new_run
    
    # Rebuild the df20 index to be 0,1,2,3,… 
    # and throws away the old index.
    df20["new_run"] = (
        df20.groupby(GROUP_COLS, group_keys=False)
        .apply(compute_new_run)
        .reset_index(drop=True)
    )

    df20["run_id"] = df20.groupby(GROUP_COLS)["new_run"].cumsum()

    df20["SegmentsMerged"] = 1

    # agg = merged runs for 20ms
    # df20 = original 20ms rows (for summaries)
    # non20_rows = original non-20 rows

    agg = (
        df20.groupby(GROUP_COLS + ["run_id"], as_index=False)
        .agg(
            EventIndex=("EventIndex", "first"),
            Start=("Start", "min"),
            Stop=("Stop", "max"),
            Duration=("Duration", "sum"),
            AOI=("AOI", "first"),
            SegmentsMerged=("SegmentsMerged", "sum"),
            #Fixation_order_nr=("Fixation order nr", "first"),
            #Tot_dur_fix_SpaceIce=("Tot dur fix SpaceIce", "first"),
            #Tot_dur_fix_Puck=("Tot dur fix Puck", "first"),
        )
    )

    # Nice ordering
    agg = agg.sort_values(["Interval", "Start"]).reset_index(drop=True)
    return agg, df20, df_non20


def main():
    # Ensure folders exist
    if not INPUT_FOLDER.exists():
        raise FileNotFoundError(
            f"Input folder not found: {INPUT_FOLDER}\n"
            f"Expected file at: {INPUT_XLSX}"
        )

    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    # Read Excel
    df = pd.read_excel(INPUT_XLSX, sheet_name=SHEET_NAME)

    merged_runs, df20_rows, non20_rows = merge_consecutive_aoi(df)

    # Recompute fixation order on merged runs (within each group) based on time
    ## merged_runs = merged_runs.sort_values(GROUP_COLS + ["Start"]).reset_index(drop=True)
    ## merged_runs["Fixation_order_nr_recomputed"] = (
    ## merged_runs.groupby(GROUP_COLS).cumcount() + 1
    ## )


    # AOI summaries
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

    # Write output
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        merged_runs.to_excel(writer, sheet_name="MergedRuns", index=False)
        aoi_summary_overall.to_excel(writer, sheet_name="AOI_Summary", index=False)
        aoi_by_group.to_excel(writer, sheet_name="AOI_ByGroup", index=False)
        if not non20_rows.empty:
            non20_rows.to_excel(writer, sheet_name="Excluded_Not20ms", index=False)

    print(f"✅ Read from:  {INPUT_XLSX}")
    print(f"✅ Wrote to:   {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
