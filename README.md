# ExcelMerger
This small Python script merges multiple Excel files from a folder into a single workbook.  
Each source file is added as a separate sheet, named after the original file.

## Requirements
- Python 3.x
- [xlwings](https://pypi.org/project/xlwings/)
- Microsoft Excel installed!!

You can install the dependency with:

```bash
pip install xlwings
```

## Usage
1. Update the `path_to_files` variable in the script with the path to your Excel files.
2. Run the script:

```bash
python excel_merger.py
```
