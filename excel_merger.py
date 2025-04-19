import xlwings as xw
import os
import glob

path_to_files = r"path/to/excel/files" # e.g. r"C:\Users\Besitzer\Desktop\SEV Oberhausen Wesel"
output_file = os.path.join(path_to_files, "merged.xlsx")

# new workbook
wb_target = xw.Book()

# find all xlsx files in the directory
files = glob.glob(os.path.join(path_to_files, "*.xlsx"))
files = [file for file in files if "merged.xlsx" not in file]  # remove the merged file from the list

# reverse the list to copy in the correct order
files.reverse()

# copy each file
for file in files:
    print(f"Copy from: {file}")
    wb_source = xw.Book(file)
    source_sheet = wb_source.sheets[0]
    
    # use the filename without extension as the new sheet name
    filename = os.path.basename(file)
    sheetname = os.path.splitext(filename)[0]  # remove the file extension
    
    new_sheet = source_sheet.api.Copy(Before=wb_target.sheets[0].api)
    wb_target.sheets[0].name = sheetname  # rename the copied sheet
    
    wb_source.close()

# remove the first sheet (which is empty)
wb_target.sheets[-1].delete()

# save the merged workbook
wb_target.save(output_file)
wb_target.close()
print(f"Done! Saved in: {output_file}")
