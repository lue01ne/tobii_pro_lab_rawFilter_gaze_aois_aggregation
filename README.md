This script aggregates consecutive raw gaze samples that were manually mapped to the same AOI within a Tobii Pro Glasses 3 (G3) recording. It reads an .xlsx exported from an event-based TSV file (via Tobii Pro Lab) and writes aggregated AOI segments plus summary tables to a new Excel workbook (the input can be extended to read TSV directly as part of future work).

## Installation

[Recommendation] Create and activate a Python virtual environment 

python -m venv venv
venv\Scripts\activate

Install dependencies:

### pandas & openpyxl
pip install pandas openpyxl

## Script execution

### Input
- Place all input files in './input_metrics_data'.
- THe input Excel workbook ('.xlsx') must contain a worksheet named 'TPL_rawFilter_metrics", which includes the raw gaze metrics exported from Tobii Pro Lab https://www.tobii.com/products/software/behavior-research-software/tobii-pro-lab

### Output
- All aggregated output files will be written to './output_data/'.

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

**Yue Lyu **  
Email: yue.lyu@linkaitech.eu

