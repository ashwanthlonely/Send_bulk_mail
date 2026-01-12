# data.py User Manual

## Purpose
Combine multiple Excel files into a single file and perform an outer merge on a specified column.

## Inputs
- Directory path: `C:\Users\ashwa\Downloads\drive-download-20240918T115343Z-001`
  - Reads all `.xlsx` files in this folder.

## Configuration
- Merge column: `Column_Name` (replace with the actual column name in your data)
- Output file: `Merged_File.xlsx` (saved in the same directory)

## How to run
1) Update `dir_path` and `merge_column`.
2) Run:
   - `python data.py`

## What it does
- Reads all Excel files in `dir_path`.
- Concatenates them into one DataFrame.
- Performs an outer merge of the DataFrame with itself on `merge_column`.
- Writes `Merged_File.xlsx` to `dir_path`.

## Outputs
- `Merged_File.xlsx` in the input directory.

## Troubleshooting / Notes
- The self-merge can create duplicate columns or a large output. If you intended a different merge, adjust the logic.
- Ensure all input files share the merge column.
