import xlwings as xw

# Load the workbook to read formula results
file_path = "PRODUCT DEV ORG-Product Custom Layout.xlsx" 
output_path = "your_file_static.xlsx"  # output file name

# Open the workbook with xlwings
app = xw.App(visible=False)
wb = app.books.open(file_path)

# Recalculate formulas (use Application.CalculateFull for compatibility)
wb.api.Application.CalculateFull()

# Replace formulas with values
for sheet in wb.sheets:
    for row in sheet.used_range:
        for cell in row:
            if cell.formula:
                cell.value = cell.value

# Save and close the workbook
wb.save(output_path)
wb.close()
app.quit()
