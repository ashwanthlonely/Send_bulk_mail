# merge.py User Manual

## Purpose
Merge all Excel files in a folder into a single Excel file.

## Inputs
- Folder path: `F:\Non voice data`
  - Reads all `.xlsx` and `.xls` files in this folder.

## Configuration
- Output path: `F:\Nonvoice_merged_data.xlsx`

## How to run
1) Update `folder_path` and `save_path` if needed.
2) Run:
   - `python merge.py`

## What it does
- Reads each Excel file in the folder.
- Concatenates all rows into one DataFrame.
- Writes the merged file to `save_path`.

## Outputs
- `F:\Nonvoice_merged_data.xlsx`
- Prints a success message with the save path.

## Troubleshooting / Notes
- Ensure all input files have compatible columns.
- Make sure the output file is not open while writing.
