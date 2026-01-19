# Tobii Pro Glasses 3 — AOI Aggregation (≤20ms) from TSV Export (Excel Input)

This repository contains a Python script that aggregates consecutive raw gaze samples that fall on the same **AOI** (Area of Interest) within a **Tobii Pro Glasses 3** recording.

> **Important:** In this workflow, the **gaze → AOI mapping is done manually** (i.e. AOI labels are assigned by manual mapping prior to running this script).

The script reads one or more `.xlsx` files exported from an event-based TSV workflow and outputs aggregated AOI segments and summary tables into new Excel workbooks.

---

## What the script does

For each recording / participant / interval context, the script:

1. Reads raw gaze rows from an Excel file (`.xlsx`)
2. Keeps rows where **Duration ≤ 20 ms**
3. Creates continuous AOI **runs** by merging consecutive rows when:
   - AOI is the same, and
   - Start time equals the previous stop time (or a 20ms step fallback)
4. Outputs aggregated AOI runs + AOI summaries into a new Excel workbook

---

## Input / Output

### Input
- One or multiple Excel files under:
  ```
  ./input_metrics_data/
  ```
- Each `.xlsx` file is expected to contain a sheet named:
  ```
  TPL_rawFilter_metrics
  ```

### Output
- Aggregated `.xlsx` files written to:
  ```
  ./output_data/
  ```

Each input file will generate:
- `<input_filename>_aggregated.xlsx`

---

## Folder structure

Recommended repository structure:

```txt
your_repo/
  README.md
  tpl_events_tsv_aois_aggregation_le20ms_batch.py
  input_metrics_data/
    file_1.xlsx
    file_2.xlsx
  output_data/
```

> `output_data/` will be created automatically if it does not exist.

---

## Installation (Windows 11 / CMD)

### Recommendation: Create and activate a dedicated Python virtual environment

To avoid dependency conflicts, it is recommended to create a dedicated Python virtual environment for this project.

#### 1) Create a virtual environment
```bash
python -m venv YOUR_VIRTUAL_ENVIRONMENT_NAME
```

#### 2) Activate the virtual environment (Windows CMD)
Navigate to your project folder, then run:
```bat
YOUR_VIRTUAL_ENVIRONMENT_NAME\Scripts\activate
```

You should see the terminal prompt prefixed with:
```
(YOUR_VIRTUAL_ENVIRONMENT_NAME)
```

#### 3) Install dependencies
```bash
python -m pip install pandas openpyxl
```

---

## How to run

Place all input `.xlsx` files into:

```
./input_metrics_data/
```

Then run:

```bash
python tpl_events_tsv_aois_aggregation_le20ms_batch.py
```

The aggregated output files will appear under:

```
./output_data/
```

---

## Output sheets (Excel)

Each aggregated Excel output typically includes:

- `Timeline_Combined`  
  Combined view including:
  - aggregated AOI segments built from rows with **Duration ≤ 20 ms**
  - raw rows with **Duration > 20 ms** (kept as-is)

- `MergedRuns`  
  Aggregated AOI runs (one row per AOI segment)

- `AOI_Summary`  
  Total duration and counts per AOI

- `AOI_ByGroup`  
  AOI summary broken down by recording / participant / interval context

- Debug sheets (optional):
  - `Raw_Duration_le20`
  - `Raw_Duration_gt20`

---

## Notes / Extensions

- The current workflow reads `.xlsx` files exported from an event-based TSV pipeline.
- The input stage can be extended to ingest the **TSV directly** (instead of Excel).
- AOI mapping is assumed to be manual in this workflow.

---


## License

**Research Use Only (Non-Commercial).**

This script is licensed for **non-commercial research and academic use only**.

✅ Permitted:
- Use in research projects (academic / non-commercial)
- Modification for research use
- Sharing within the research community with attribution

❌ Not permitted without permission:
- Any commercial use
- Use in commercial products
- Selling, licensing, or monetizing the software

Commercial use is available under a **separate commercial license**.
Please contact:

**Yue "Lou" Lyu **  
Email: yue.lyu@linkaitech.eu

