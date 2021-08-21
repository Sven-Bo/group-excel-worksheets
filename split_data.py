from pathlib import Path  # Standard Python Module
import xlwings as xw  # pip install xlwings

"""
Iterate over all excel files in a given source directory.
For each worksheet within the excel files -> create a new workbook.
The name of the new workbook will be equal to the worksheet's name.
If the workbook already exsits, append the worksheet to the existing workbook.
"""

SOURCE_DIR = "Month_End_Data"
OUTPUT_DIR = "Output"

excel_files = list(Path(SOURCE_DIR).glob("*.xlsx"))

# Create a dict with the excel file name & path:
# excel_outputs = {'filename1':filepath1,
#                  'filename2':filepath2}
output_paths = list(Path(OUTPUT_DIR).glob("*.xlsx"))
output_names = [file.stem for file in output_paths]
excel_outputs = dict(zip(output_names, output_paths))

for excel_file in excel_files:
    wb = xw.Book(excel_file)
    for sheet in wb.sheets:
        # Check, if the sheet name already exsits in the output folder (as a separate wb)
        if sheet.name in excel_outputs:
            # Check, if that wb also contains out sheet (the wb's name)
            # If not, we will copy/append the sheet to the existing wb
            wb_tmp = xw.Book(excel_outputs[sheet.name])
            wb_tmp_sheets = [sheet.name for sheet in wb_tmp.sheets]
            if wb.name not in wb_tmp_sheets:
                sheet.copy(after=wb_tmp.sheets[0])
                sht = xw.sheets.active
                sht.name = wb.name
                wb_tmp.save()
            wb_tmp.close()
        else:
            # Create a new wb and copy the sheet to this new wb
            # Afterwards, add the new file & path to the dict 'excel_outputs'
            wb_tmp = xw.Book()
            sheet.copy(after=wb_tmp.sheets[0])
            sht = xw.sheets.active
            sht.name = wb.name
            wb_tmp.sheets[0].delete()
            output_path = Path(OUTPUT_DIR, sheet.name + ".xlsx")
            wb_tmp.save(output_path)
            excel_outputs[sheet.name] = wb_tmp.fullname
            wb_tmp.close()
    # Only quit the excel instance if no other wb is open
    if len(wb.app.books) == 1:
        wb.app.quit()
    else:
        wb.close()
