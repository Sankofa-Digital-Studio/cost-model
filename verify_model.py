import openpyxl

wb = openpyxl.load_workbook("West_Gold_Plant_Demolition_Rehabilitation_Model.xlsx", data_only=False)
print("Sheet names:", wb.sheetnames)

for sheetname in wb.sheetnames:
    ws = wb[sheetname]
    print(f"\n--- Checking Sheet: {sheetname} ---")
    formulas_found = 0
    ref_errors = 0
    for row in ws.iter_rows(values_only=False):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                formulas_found += 1
                if "#REF!" in cell.value:
                    print(f"ERROR: #REF! found at {cell.coordinate} in {sheetname}: {cell.value}")
                    ref_errors += 1
    print(f"Total formulas: {formulas_found}, Ref errors: {ref_errors}")

print("\nVerification complete.")
